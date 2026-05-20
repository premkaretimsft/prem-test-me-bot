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
import { MicrosoftAppCredentials } from "botframework-connector";

// This bot's main dialog.
import { SearchApp } from "./searchApp";
import config, { CANARY_SERVICE_URL } from "./config";

// Trust Canary Bot Framework service URLs
// This is required for Canary/PPE environments where the service URL is different from production
MicrosoftAppCredentials.trustServiceUrl("https://canary.botapi.skype.com");
MicrosoftAppCredentials.trustServiceUrl(CANARY_SERVICE_URL);
// console.log("[Auth] Trusted Canary service URLs");

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
// console.log("[Auth Debug] BOT_ID:", config.botId);
// console.log("[Auth Debug] BOT_PASSWORD:", config.botPassword ? config.botPassword.substring(0, 10) + "..." : "MISSING");
// console.log("[Auth Debug] BOT_TENANT_ID:", config.botTenantId);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "SingleTenant",
  MicrosoftAppTenantId: config.botTenantId,
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

// Force outgoing responses through Canary, with fallback to the original
// serviceUrl on auth failure. This runs AFTER the adapter has validated the
// inbound JWT's `serviceurl` claim, so it doesn't break inbound auth.
adapter.use({
  async onTurn(context, next): Promise<void> {
    const originalServiceUrl = context.activity.serviceUrl;

    // applyConversationReference (inside TurnContext.sendActivities) copies
    // context.activity.serviceUrl onto every outgoing activity, so this is
    // what actually pins the egress URL.
    context.activity.serviceUrl = CANARY_SERVICE_URL;

    // Hook outbound sends so we can retry on the original URL if Canary
    // rejects the call (e.g., 401 / "ServiceUrl claim do not match").
    context.onSendActivities(async (ctx, activities, sendNext) => {
      try {
        return await sendNext();
      } catch (err) {
        const message = (err as Error)?.message || "";
        const status = (err as any)?.statusCode;
        const isServiceUrlAuth =
          status === 401 ||
          /serviceurl/i.test(message) ||
          /unauthorized/i.test(message);

        if (!isServiceUrlAuth || !originalServiceUrl) {
          throw err;
        }

        console.warn(
          `[Canary] send to ${CANARY_SERVICE_URL} failed (${message}). ` +
          `Falling back to original serviceUrl: ${originalServiceUrl}`
        );

        // Restore original serviceUrl on each activity AND on the turn context
        // so any subsequent sends this turn don't repeat the failure.
        activities.forEach((a) => {
          a.serviceUrl = originalServiceUrl;
        });
        ctx.activity.serviceUrl = originalServiceUrl;

        return await sendNext();
      }
    });

    await next();
  },
});

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
    // Debug incoming activity
  if (req.body) {
    // console.log(`\n[Incoming Activity]`);
    // console.log(`  Type: ${req.body.type}`);
    // console.log(`  ServiceUrl: ${req.body.serviceUrl}`);
    // console.log(`  From: ${req.body.from?.id}`);
    // console.log(`  Conversation: ${req.body.conversation?.id?.substring(0, 30)}...`);

    // Dynamically trust the incoming service URL (for Canary/PPE environments)
    if (req.body.serviceUrl) {
      MicrosoftAppCredentials.trustServiceUrl(req.body.serviceUrl);
      // console.log(`  [Auth] Trusted serviceUrl: ${req.body.serviceUrl}`);
    }
    // NOTE: do NOT rewrite req.body.serviceUrl here. The adapter validates the
    // JWT's `serviceurl` claim against this value before bot logic runs, so
    // changing it here causes "ServiceUrl claim do not match" (401). The
    // override-to-Canary now happens via adapter middleware (after auth) with
    // a fallback to the original URL on failure.
  }

  await adapter.process(req, res, async (context) => {
    await searchApp.run(context);
    // console.log("------------------------");
    // console.log("-------Response(Incoming Request)---------");
    // console.log(res);
  });
});

// server.on("after", function (req, res, route, error) {
//   console.log("------------------------");
//   console.log("------Request(in after method)---------");
//   // console.log(req.route.path);
//   console.log(req.body);
//   console.log("------------------------");
//   console.log("-------Response(in after method)---------");
//   console.log(`Http Status: ${res.statusCode}`);
//   console.log(`Has Body: ${res._hasBody}`);
//   //console.log(`Response Headers: ${res.getHeaders()}`);
//   console.log(`Response Headers: ${res.header('ms-cv')}`);
//   console.log(res._data);
//   console.log(res.body);
// });

server.on("pre", function (req, res) {
  // console.log("---------In Pre Method------------");
  // console.log("------Request---------");
  // // console.log(req.route.path);
  // console.log(req.body);
  // console.log(req.Headers);
  // console.log("-------Response---------");
  // console.log(`Http Status: ${res.statusCode}`);
  // console.log(`Has Body: ${res._hasBody}`);
  // //console.log(`Response Headers: ${res.getHeaders()}`);
  // console.log(`Alternate Response Headers: ${res.header('ms-cv')}`);
  // console.log(res._data);
  // console.log(res.body);
});
