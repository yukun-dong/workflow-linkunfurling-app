import { TurnContext, Activity, CardFactory, MessagingExtensionResponse } from "botbuilder";
import { Link, TeamsFxLinkUnfurlingHandler, TriggerLinks } from "./sdk/messageExtension";
import * as ACData from "adaptivecards-templating";
import template from "./data/template.json";
import data from "./data/data.json";

export class MainLinkUnfurlingHandler implements TeamsFxLinkUnfurlingHandler {
    triggerLinks: TriggerLinks = ".*\.test111\.com/main/.*";
    async handleLinkReceived(context: TurnContext, link: Link): Promise<any> {
        console.log(`app received link ${link.link}`);
        const templateJson = new ACData.Template(template);
        const card=templateJson.expand({$root:data});

        const attachment ={ ...CardFactory.adaptiveCard(card), preview:  CardFactory.heroCard("test", "test") }; 
        // const attachment = CardFactory.thumbnailCard("card rendered for main", link.link, ['https://bot-framework.azureedge.net/static/341749-d71b264801/intercom-webui/v1.6.2/assets/landing-page/images/Composer_Icon.png']);
        return {
            composeExtension: {
              type: "result",
              attachmentLayout: "list",
              attachments: [attachment],
            },
          };
    }

}

export class DevLinkUnfurlingHandler implements TeamsFxLinkUnfurlingHandler {
    triggerLinks: TriggerLinks = ".*\.test111\.com/dev/.*";
    async handleLinkReceived(context: TurnContext, link: Link): Promise<any> {
        console.log(`app received link ${link.link}`);
        const attachment = CardFactory.thumbnailCard("card rendered for dev", link.link, ['https://bot-framework.azureedge.net/static/341749-d71b264801/intercom-webui/v1.6.2/assets/landing-page/images/Composer_Icon.png']);
        return {
            composeExtension: {
              type: "result",
              attachmentLayout: "list",
              attachments: [attachment],
            },
          };
    }

}