import {
    ActivityHandler, CloudAdapter, Request,
    Response,
} from "botbuilder";
import { ConfigurationServiceClientCredentialFactory, ConfigurationBotFrameworkAuthentication, Activity, TurnContext, ActivityTypes, Middleware } from "botbuilder-core";

export class MessageExtension {
    public readonly adapter: CloudAdapter;
    public linkUnfurling?: LinkUnfurling;
    // public search: SearchInvokeOptions;
    public constructor(options: MessageExtensionOptions) {
        if (options.adapter) {
            this.adapter = options.adapter;
        } else {
            this.adapter = this.createDefaultAdapter(options.adapterConfig);
        }

        if (options.linkUnfurling?.enabled) {
            this.linkUnfurling = new LinkUnfurling(this.adapter, options.linkUnfurling);
        }
    }
    public async requestHandler(
        req: Request,
        res: Response,
        logic?: (context: TurnContext) => Promise<any>
    ): Promise<void> {
        if (logic === undefined) {
            // create empty logic
            logic = async () => { };
        }

        await this.adapter.process(req, res, logic);
    }

    private createDefaultAdapter(adapterConfig?: { [key: string]: unknown }): CloudAdapter {
        const credentialsFactory =
            adapterConfig === undefined
                ? new ConfigurationServiceClientCredentialFactory({
                    MicrosoftAppId: process.env.BOT_ID,
                    MicrosoftAppPassword: process.env.BOT_PASSWORD,
                    MicrosoftAppType: "MultiTenant",
                })
                : new ConfigurationServiceClientCredentialFactory(adapterConfig);
        const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
            {},
            credentialsFactory
        );
        const adapter = new CloudAdapter(botFrameworkAuthentication);

        // the default error handler
        adapter.onTurnError = async (context, error) => {
            // This check writes out errors to console.
            console.error(`[onTurnError] unhandled error: ${error}`);

            // Send a trace activity, which will be displayed in Bot Framework Emulator
            await context.sendTraceActivity(
                "OnTurnError Trace",
                `${error}`,
                "https://www.botframework.com/schemas/error",
                "TurnError"
            );

            // Send a message to the user
            await context.sendActivity(`The bot encountered unhandled error: ${error.message}`);
            await context.sendActivity("To continue to run this bot, please fix the bot source code.");
        };
        return adapter;
    }
}
export class LinkUnfurling {
    private readonly adapter: CloudAdapter;
    private readonly middleware: LinkUnfurlingMiddleware;

    constructor(
        adapter: CloudAdapter,
        options?: LinkUnfurlingOptions,
    ) {

        this.middleware = new LinkUnfurlingMiddleware(
            options?.links,
        );
        this.adapter = adapter.use(this.middleware);
    }

    /**
     * Register a command into the command bot.
     *
     * @param command - The command to be registered.
     */
    public registerLink(link: TeamsFxLinkUnfurlingHandler): void {
        if (link) {
            this.middleware.linkUnfurlingHandlers.push(link);
        }
    }
}
export interface MessageExtensionOptions {
    /**
     * The bot adapter. If not provided, a default adapter will be created:
     * - with `adapterConfig` as constructor parameter.
     * - with a default error handler that logs error to console, sends trace activity, and sends error message to user.
     *
     * @remarks
     * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
     */
    adapter?: CloudAdapter;

    /**
     * If `adapter` is not provided, this `adapterConfig` will be passed to the new `BotFrameworkAdapter` when created internally.
     *
     * @remarks
     * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
     */
    adapterConfig?: { [key: string]: unknown };

    linkUnfurling?: LinkUnfurlingOptions & {
        enabled?: boolean;
    };
}

export interface LinkUnfurlingOptions {
    links?: TeamsFxLinkUnfurlingHandler[];

}

export interface TeamsFxLinkUnfurlingHandler {
    /**
     * The string or regular expression patterns that can trigger this handler.
     */
    triggerLinks: TriggerLinks;

    /**
     * Handles a bot command received activity.
     *
     * @param context The bot context.
     * @param message The command message the user types from Teams.
     * @returns A `Promise` representing an activity or text to send as the command response.
     * Or no return value if developers want to send the response activity by themselves in this method.
     */
    handleLinkReceived(
        context: TurnContext,
        link: Link
    ): Promise<any>;
}

export type TriggerLinks = string | RegExp | (string | RegExp)[];

export interface Link {
    link: string;
    matches?: RegExpMatchArray;
}

export class LinkUnfurlingMiddleware implements Middleware {
    public readonly linkUnfurlingHandlers: TeamsFxLinkUnfurlingHandler[] = [];
    constructor(
        handlers?: TeamsFxLinkUnfurlingHandler[],
    ) {
        handlers = handlers ?? [];
        this.linkUnfurlingHandlers.push(...handlers);
    }


    public async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
        if (context.activity.name === 'composeExtension/queryLink') {
            // Invoke corresponding command handler for the command response

            const url = context.activity.value.url;
            let alreadyProcessed = false;

            for (const handler of this.linkUnfurlingHandlers) {
                const matchResult = this.shouldTrigger(handler.triggerLinks, url);

                // It is important to note that the command bot will stop processing handlers
                // when the first command handler is matched.
                if (!!matchResult) {
                    const message: Link = {
                        link: url,
                    };
                    message.matches = Array.isArray(matchResult) ? matchResult : void 0;
                    const response = await handler.handleLinkReceived(context, message);
                    await context.sendActivity({ type: "invokeResponse", value: { status: 200, body: response } })
                    // await this.processResponse(context, response);
                    alreadyProcessed = true;
                    break;
                }
            }

        }
        await next();
    }

    private async processResponse(context: TurnContext, response: string | void | Partial<Activity>) {
        if (typeof response === "string") {
            await context.sendActivity(response);
        } else {
            const replyActivity = response as Partial<Activity>;
            if (replyActivity) {
                await context.sendActivity(replyActivity);
            }
        }
    }

    private matchPattern(pattern: string | RegExp, text: string): boolean | RegExpMatchArray {
        if (text) {
            if (typeof pattern === "string") {
                const regExp = new RegExp(pattern as string, "i");
                return regExp.test(text);
            }

            if (pattern instanceof RegExp) {
                const matches = text.match(pattern as RegExp);
                return matches ?? false;
            }
        }

        return false;
    }

    private shouldTrigger(patterns: TriggerLinks, text: string): RegExpMatchArray | boolean {
        const expressions = Array.isArray(patterns) ? patterns : [patterns];

        for (const ex of expressions) {
            const arg = this.matchPattern(ex, text);
            if (arg) return arg;
        }

        return false;
    }

    private getActivityText(activity: Activity): string {
        let text = activity.text;
        const removedMentionText = TurnContext.removeRecipientMention(activity);
        if (removedMentionText) {
            text = removedMentionText
                .toLowerCase()
                .replace(/\n|\r\n/g, "")
                .trim();
        }

        return text;
    }
}