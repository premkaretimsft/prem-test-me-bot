// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  ActivityTypes,
} from "botbuilder";

// This bot's main dialog.
import { SearchApp } from "./searchApp";
import config from "./config";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${JSON.stringify(error)}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
const continuationParameters: {} = {};
const searchApp = new SearchApp(async () => {
  console.log(
    `Handling continuation - ${JSON.stringify(continuationParameters)}`
  );
  for (const continuationParameter of Object.values(continuationParameters)) {
    const conversationReference = (continuationParameter as any)
      .conversationReference;
    if (conversationReference.conversation.conversationType === "channel") {
      const channelConversationId = (
        conversationReference.conversation.id as string
      ).split(";")[0];
      conversationReference.conversation.id = channelConversationId;

      await adapter.continueConversationAsync(
        (continuationParameter as any).claimsIdentity,
        conversationReference,
        (continuationParameter as any).oAuthScope,
        async (context) => {
          // MicrosoftAppCredentials.trustServiceUrl(
          //   conversationReference.serviceUrl
          // );
          const continuationToken = (continuationParameter as any)
            .continuationToken;
          await context.sendActivities([
            { type: ActivityTypes.Typing },
            (continuationParameter as any).partialActivity,
          ]);
        }
      );
    } else {
      await adapter.continueConversationAsync(
        (continuationParameter as any).claimsIdentity,
        conversationReference,
        (continuationParameter as any).oAuthScope,
        async (context) => {
          // MicrosoftAppCredentials.trustServiceUrl(
          //   conversationReference.serviceUrl
          // );
          const continuationToken = (continuationParameter as any)
            .continuationToken;
          await context.sendActivities([
            {
              type: ActivityTypes.Message,
              text: "Continuing conversation from copilot...",
            },
            { type: ActivityTypes.Typing },
            { type: "delay", value: 1000 },
            {
              type: ActivityTypes.Message,
              text: `Fetching more details using the continuation token passed: ${continuationToken}`,
            },
            { type: ActivityTypes.Typing },
            { type: "delay", value: 4000 },
            {
              type: ActivityTypes.Message,
              text: `Handoff successful!`,
              attachments: [(continuationParameter as any).cardAttachment],
            },
            { type: ActivityTypes.Typing },
            { type: "delay", value: 2000 },
            {
              type: ActivityTypes.Message,
              text: `Do you need revenue or discounts details about ${(
                continuationToken as string
              ).replace("-continuation", "")}?`,
            },
          ]);
        }
      );
    }
  }
}, continuationParameters /* conversationReferences */);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await searchApp.run(context);
    console.log("------------------------");
    console.log("-------Response(Incoming Request)---------");
    console.log(res);
  });
});

server.on("after", function (req, res, route, error) {
  console.log("------------------------");
  console.log("------Request(in after method)---------");
  // console.log(req.route.path);
  console.log(req.body);
  console.log("------------------------");
  console.log("-------Response(in after method)---------");
  console.log(`Http Status: ${res.statusCode}`);
  console.log(`Has Body: ${res._hasBody}`);
  //console.log(`Response Headers: ${res.getHeaders()}`);
  console.log(`Response Headers: ${res.header('ms-cv')}`);
  console.log(res._data);
  console.log(res.body);
});

// server.on("pre", function (req, res) {
//   console.log("---------In Pre Method------------");
//   console.log("------Request---------");
//   // console.log(req.route.path);
//   console.log(req.body);
//   console.log(req.Headers);
//   console.log("-------Response---------");
//   console.log(`Http Status: ${res.statusCode}`);
//   console.log(`Has Body: ${res._hasBody}`);
//   //console.log(`Response Headers: ${res.getHeaders()}`);
//   console.log(`Alternate Response Headers: ${res.header('ms-cv')}`);
//   console.log(res._data);
//   console.log(res.body);
// });
