import * as restify from "restify";
import { workflowApp } from "./internal/initialize";

const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

server.post("/api/messages", async (req, res) => {
  await workflowApp.requestHandler(req, res);
});
