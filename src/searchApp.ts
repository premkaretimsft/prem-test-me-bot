import {
  TeamsActivityHandler,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  InvokeResponse,
  AdaptiveCardInvokeResponse,
  ActivityTypes,
  Activity,
  CardFactory,
  Attachment,
  MessageFactory,
} from "botbuilder";
import productSearchCommand from "./messageExtensions/productSearchCommand";
import discountedSearchCommand from "./messageExtensions/discountSearchCommand";
import revenueSearchCommand from "./messageExtensions/revenueSearchCommand";
import actionHandler from "./adaptiveCards/cardHandler";
import { CreateActionErrorResponse, CreateInvokeResponse } from "./adaptiveCards/utils";
import { setTimeout as nodeTimeout} from "timers/promises";

const streamingPacketsText = ["Prem streaming- Second informative request that is meant to be a little longer than the first one",
  `<p>In a quiet forest, a small stream flowed peacefully through the trees. It was no ordinary stream, for it was said to hold magical powers that could grant wishes to those who drank from its waters.<br/><br/>
One day, a young girl named Lily stumbled upon the stream while wandering through the forest. She was lost and had been wandering for hours, but when she saw the stream, an overwhelming feeling of hope filled her heart.</p>`,
`<p>Desperate for a way out of the forest, Lily approached the stream and knelt down to take a drink. As she sipped the cool, clear water, she closed her eyes and made a wish with all her heart.<br/><br/>
Suddenly, the world around her began to spin and swirl. When she opened her eyes again, she was no longer in the forest. Instead, she found herself in a grand castle, with marble floors and glittering chandeliers hanging overhead.<br/><br/>
Confused but curious, Lily began to explore the castle. She wandered through grand ballrooms, ornate dining halls, and even a secret garden filled with rare flowers and exotic birds.</p>`,
`<p>As she explored, she began to realize that the castle belonged to a powerful sorceress. She had heard stories of the sorceress before, of the endless riches and magical powers she possessed.<br/><br/> Lily realized that her wish had brought her to this place, and she knew that she had to find a way to make the most of this opportunity.
For days, Lily explored the castle, learning all she could about the sorceress and her powers. She watched as the sorceress performed spells and incantations, and slowly but surely, she began to learn the ways of magic.<br/><br/>
In time, Lily became a powerful sorceress in her own right. She learned to control the elements, to summon creatures from the forest, and to cast spells that could bend reality itself. And all because of a wish she made at a magical stream in the heart of a forest.</p>`,
`<p>In time, Lily became a powerful sorceress in her own right. She learned to control the elements, to summon creatures from the forest, and to cast spells that could bend reality itself.<br/><br/> And all because of a wish she made at a magical stream in the heart of a forest.</p>`
];

export class SearchApp extends TeamsActivityHandler {
  notifyContinuationActivity: any;
  continuationParameters: any;

  constructor(
    notifyContinuationActivity: any,
    continuationParameters: any = {} /*conversations: any = {}*/
  ) {
    super();
    this.notifyContinuationActivity = notifyContinuationActivity;
    this.continuationParameters = continuationParameters;

    this.onReactionsAdded(async (context, next) => {
      const reactionsAdded = context.activity.reactionsAdded;
      if (reactionsAdded && reactionsAdded.length > 0) {
        for (let i = 0; i < reactionsAdded.length; i++) {
          const reaction = reactionsAdded[i];
          const newReaction = `You reacted with '${reaction.type}' to the following message: '${context.activity.replyToId}'`;
          // Sends an activity to the sender of the incoming activity.
          const resourceResponse = context.sendActivity(newReaction);
          // Save information about the sent message and its ID (resourceResponse.id).
        }
      }
    });
  }

  public async onMessageActivity(context: TurnContext): Promise<void> {
    if (context.activity.text.includes("stream")) {
      try {
        await this.processStreamingRequest(
          context,
          context.activity.text.includes("ac"),
          context.activity.text.includes("delay") ? 1500 : 900
        );
      } catch (error) {
        // If an error occurs during sending, inform the user
        await context.sendActivity(
          MessageFactory.text(
            "Error while sending streaming activity: " + error.message
          )
        );
        throw new Error("Error sending activity: " + error.message); // Propagate error
      }
    } else {
      this.addOrUpdateChannelPostParameters(context);
      await context.sendActivity(this.getBotAIGenActivity(context));
    }
  }

  private getStreamingActivity(
    streamType: string,
    streamId: string,
    sequence: number,
    addAttachments: boolean = false
  ): Partial<Activity> {
    let textContent = streamingPacketsText[sequence];

    if (streamType !== "informative") {
      textContent = streamingPacketsText.slice(1, sequence + 1).join(`<br/>`);
    }

    const activity = {
      type:
        streamType === "final" ? ActivityTypes.Message : ActivityTypes.Typing,
      text: textContent,
      entities: [
        {
          type: "streaminfo",
          streamType: streamType,
          streamId: streamId,
        },
      ],
    };

    if (streamType !== "final") {
      activity.entities[0]["streamSequence"] = sequence + 2;
      // activity["attachments"] = [this.getChartAdaptiveCard()];
    } else {
      if (addAttachments) {
        activity["attachments"] = [this.getChartAdaptiveCard()];
      }
    }

    return activity;
  }

  private async processStreamingRequest(
    context: TurnContext,
    addAttachments: boolean = false,
    packetDelay: number = 900
  ): Promise<void> {
    const result = await context.sendActivity({
      type: ActivityTypes.Typing,
      text: "Prem streaming- first informative request that is not too long",
      entities: [
        {
          type: "streaminfo",
          streamType: "informative", // informative or streaming; default= streaming.
          streamSequence: 1, // (required) incremental integer; must be present for start and continue streaming request, but must not be set for final streaming request.
        },
      ],
    });

    const streamId = result.id;
    console.log(`streamId: ${streamId}`);

    const streamingPacketsLength = streamingPacketsText.length;
    for (let i = 0; i < streamingPacketsLength; i++) {
      await nodeTimeout(packetDelay);
      const streamType =
        i === 0
          ? "informative"
          : i === streamingPacketsLength - 1
          ? "final"
          : "streaming";
      await context.sendActivity(
        this.getStreamingActivity(streamType, streamId, i, addAttachments)
      );
    }
  }

  public async onTeamsMessageEdit(context: TurnContext): Promise<void> {
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );

    const isChannel =
      conversationReference.conversation.conversationType === "channel";
    await context.sendActivity(
      this.getBotAIGenActivity(context, isChannel, true)
    );
  }

  private getBotAIGenActivity(
    context: TurnContext,
    isChannelPost: boolean = false,
    isEdited: boolean = false
  ): Partial<Activity> {
    const edited = isEdited ? "Edited-" : "";
    const powerpointImageObj = { name: "microsoft powerpoint" };
    const excelImageObj = { name: "microsoft excel" };
    const oneNoteImageObj = { name: "microsoft onenote" };
    const wordImageObj = { name: "microsoft word" };
    const citation1ImageObj = isChannelPost
      ? powerpointImageObj
      : excelImageObj;
    const citation2ImageObj = isChannelPost ? oneNoteImageObj : wordImageObj;

    // return {
    //   type: "message",
    //   text: `You said: ${context.activity.text}`,
    //   attachments: [this.getChartInputActionsCard()],
    //   entities: [
    //     {
    //       type: "https://schema.org/Message",
    //       "@type": "Message",
    //       "@context": "https://schema.org",
    //       additionalType: ["AIGeneratedContent"], // AI Generated label
    //     },
    //   ],
    //   channelData: {
    //     feedbackLoopEnabled: true, // Enable feedback buttons
    //   },
    // };

    return {
      type: "message",
      value: { requestId: "1234" },
      attachments: [this.getChartAdaptiveCard()],
      text: `[1] You said: ${context.activity.text} in ${
        isChannelPost ? "channel post" : "reply"
      }. From your inventory The information about chai in the Prkare Inventory indicates that it is supplied by Contoso Beverage Company of London with 349 units in stock. The stock information includes a unit price of 18 USD, an average discount of 8.6%, and an inventory valuation of 6,282 USD. There are currently 349 units in stock, with a reorder level of 25 units and a revenue this period of 12,788 USD.[1]
      From the web There are also references to a "Chai's Inventory Sorter" which is a mod for Minecraft that allows for inventory sorting and management. However, this is likely not related to your query. [2]
      If you need more detailed information or specific actions to be taken regarding the chai inventory, please let me know how I can further assist you.`,
      channelData: { feedbackLoop: { type: "default" } },
      // channelData: {
      //   feedbackLoopEnabled: true // Enable feedback buttons
      // },
      entities: [
        {
          type: "https://schema.org/Message",
          "@type": "Message",
          "@context": "https://schema.org",
          "@id": "",
          additionalType: ["AIGeneratedContent"],
          usageInfo: {
            name: `${
              isChannelPost
                ? "Company level sensitivity"
                : "Org level sensitivity"
            }`,
            description: `Please don't share outside of the ${
              isChannelPost ? "company" : "organization"
            }`,
            "@id": "1a19d03a-48bc-4359-8038-5b5f6d5847c3",
            position: 5,
          },
          citation: [
            {
              "@type": "Claim",
              position: 1,
              appearance: {
                "@type": "DigitalDocument",
                name: "Beverages data in the company inventory reference list very very long test name for testing the citation name",
                text: JSON.stringify(this.getChartAdaptiveCard().content),
                url: "https://www.microsoft.com",
                abstract: `From the web There are also references to a "Chai's Inventory Sorter" which is a mod for Minecraft that allows for inventory sorting and management. However, this is likely not related to your query. 2
                If you need more detailed information or specific actions to be taken regarding the chai inventory.`,
                encodingFormat: "application/vnd.microsoft.card.adaptive",
                image: citation1ImageObj,
                keywords: [
                  "Company Data with a longer version of keyword to test the twenty eight character",
                  "Recently Updated with a longer version of keyword to test the twenty eight character",
                  "2022-09-19 17:44:17.858167 with a longer version of keyword to test the twenty eight",
                ],
                usageInfo: {
                  "@type": "CreativeWork",
                  description: "Please don't share outside of the company",
                  name: "Company level sensitivity",
                },
              },
            },
            {
              "@type": "Claim",
              position: 2,
              appearance: {
                "@type": "DigitalDocument",
                name: "Products revenue data in the company",
                text: JSON.stringify(this.getChartAdaptiveCard().content),
                url: "https://www.microsoft.com",
                abstract: `From the web There are also references to a "Chai's Inventory Sorter" which is a mod for Minecraft that allows for inventory sorting and management. However, this is likely not related to your query. 2
                If you need more detailed information or specific actions to be taken regarding the chai inventory.`,
                encodingFormat: "application/vnd.microsoft.card.adaptive",
                image: citation2ImageObj,
                keywords: [
                  "Company Data",
                  "Recently Updated",
                  "2022-09-19 17:44:17.858167",
                ],
                usageInfo: {
                  "@type": "CreativeWork",
                  "@id": "usage-info-1",
                  description: "Please don't share outside of the company",
                  name: "Company level sensitivity",
                  position: 5,
                  pattern: {
                    "@type": "DefinedTerm",
                    inDefinedTermSet: "https://www.w3.org/TR/css-values-4/",
                    name: "color",
                    termCode: "#454545",
                  },
                },
              },
              claimInterpreter: {
                "@type": "Project",
                name: "Claim Interpreter name",
                slogan: "Claim Interpreter slogan",
                url: "https://www.example.com/claim-interpreter",
              },
            },
          ],
        },
      ],
    };
  }

  private getChartInputActionsCard(): Attachment {
    return CardFactory.adaptiveCard({
      type: "AdaptiveCard",
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.6",
      body: [
        {
          type: "Input.Text",
          placeholder: "Placeholder text",
          label: "Text input",
          id: "text",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Date",
          label: "Date input",
          id: "date",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Time",
          id: "time",
          label: "Time input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Number",
          placeholder: "Placeholder text",
          id: "number",
          label: "Number input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.ChoiceSet",
          choices: [
            {
              title: "Choice 1",
              value: "Choice 1",
            },
            {
              title: "Choice 2",
              value: "Choice 2",
            },
          ],
          placeholder: "Placeholder text",
          id: "choiceSet",
          label: "ChoiceSet input",
          isRequired: true,
          errorMessage: "Error",
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Action.Submit",
          conditionallyEnabled: true,
          associatedInputs: "auto",
        },
      ],
    });
  }

  private getChartAdaptiveCard(): Attachment {
    const groupedBarChartData = [
      {
        legend: "Outlook",
        values: [
          { x: "2023-05-01", y: 24 },
          { x: "2023-05-02", y: 27 },
          { x: "2023-05-03", y: 18 },
          { x: "2023-05-04", y: 30 },
          { x: "2023-05-05", y: 20 },
          { x: "2023-05-06", y: 35 },
          { x: "2023-05-07", y: 40 },
          { x: "2023-05-08", y: 45 },
        ],
      },
      {
        legend: "Teams",
        values: [
          { x: "2023-05-01", y: 9 },
          { x: "2023-05-02", y: 100 },
          { x: "2023-05-03", y: 22 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 30 },
          { x: "2023-05-06", y: 45 },
          { x: "2023-05-07", y: 50 },
          { x: "2023-05-08", y: 55 },
        ],
      },
      {
        legend: "Office",
        values: [
          { x: "2023-05-01", y: 10 },
          { x: "2023-05-02", y: 20 },
          { x: "2023-05-03", y: 30 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 50 },
          { x: "2023-05-06", y: 60 },
          { x: "2023-05-07", y: 70 },
          { x: "2023-05-08", y: 80 },
        ],
      },
      {
        legend: "Windows",
        values: [
          { x: "2023-05-01", y: 10 },
          { x: "2023-05-02", y: 20 },
          { x: "2023-05-03", y: 30 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 50 },
          { x: "2023-05-06", y: 60 },
          { x: "2023-05-07", y: 70 },
          { x: "2023-05-08", y: 80 },
        ],
      },
    ];

    return CardFactory.adaptiveCard({
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: "Simple",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar",
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          colorSet: "categorical",
          data: [
            { x: "Pear", y: 59 },
            { x: "Banana", y: 292 },
            { x: "Apple", y: 143 },
            { x: "Peach", y: 98 },
            { x: "Kiwi", y: 179 },
            { x: "Grapefruit", y: 20 },
            { x: "Orange", y: 212 },
            { x: "Cantaloupe", y: 68 },
            { x: "Grape", y: 102 },
            { x: "Tangerine", y: 38 },
          ],
        },
        {
          type: "TextBlock",
          text: "Grouped",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar.Grouped",
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          colorSet: "diverging",
          data: groupedBarChartData,
        },
        {
          type: "TextBlock",
          text: "Stacked",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar.Grouped",
          stacked: true,
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          data: groupedBarChartData,
        },
        {
          type: "Input.Text",
          placeholder: "Placeholder text",
          label: "Text input",
          id: "text",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Date",
          label: "Date input",
          id: "date",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Time",
          id: "time",
          label: "Time input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Number",
          placeholder: "Placeholder text",
          id: "number",
          label: "Number input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.ChoiceSet",
          choices: [
            {
              title: "Choice 1",
              value: "Choice 1",
            },
            {
              title: "Choice 2",
              value: "Choice 2",
            },
          ],
          placeholder: "Placeholder text",
          id: "choiceSet",
          label: "ChoiceSet input",
          isRequired: true,
          errorMessage: "Error",
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Action.Submit",
          conditionallyEnabled: true,
          associatedInputs: "auto",
        },
      ],
    });
  }

  private addOrUpdateChannelPostParameters(context): void {
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );

    if (conversationReference.conversation.conversationType === "channel") {
      console.log(
        `Adding continuation parameters for channel context: ${JSON.stringify(
          context
        )}`
      );
      this.continuationParameters[context.activity.from.id] = {
        claimsIdentity: context.turnState.get(context.adapter.BotIdentityKey),
        conversationReference: TurnContext.getConversationReference(
          context.activity
        ),
        oAuthScope: context.turnState.get(context.adapter.OAuthScopeKey),
        partialActivity: this.getBotAIGenActivity(context, true),
      };

      setTimeout(async () => await this.notifyContinuationActivity(), 1000);
    }
  }

  public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    try {
      switch (context.activity.name) {
        case "message/submitAction":
          return CreateInvokeResponse(200);
        case "composeExtension/query":
          return {
            status: 200,
            body: await this.handleTeamsMessagingExtensionQuery(
              context,
              context.activity.value
            ),
          };
        case "adaptiveCard/action":
          return {
            status: 200,
            body: await this.onAdaptiveCardInvoke(context),
          };
        default:
          return {
            status: 200,
            body: `Unknown invoke activity handled as default- ${context.activity.name}`,
          };
      }
    } catch (err) {
      console.log(`Error in onInvokeActivity: ${err}`);
      return {
        status: 500,
        body: `Invoke activity received- ${context.activity.name}`,
      };
    }
  }

  // Handle search message extension
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    switch (query.commandId) {
      case productSearchCommand.COMMAND_ID: {
        return productSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
      case discountedSearchCommand.COMMAND_ID: {
        return discountedSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
      case revenueSearchCommand.COMMAND_ID: {
        return revenueSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
    }
  }

  // Handle adaptive card actions
  public async onAdaptiveCardInvoke(
    context: TurnContext
  ): Promise<AdaptiveCardInvokeResponse> {
    try {
      switch (context.activity.value.action.verb) {
        case "ok": {
          return actionHandler.handleTeamsCardActionUpdateStock(context);
        }
        case "restock": {
          return actionHandler.handleTeamsCardActionRestock(context);
        }
        case "cancel": {
          return actionHandler.handleTeamsCardActionCancelRestock(context);
        }
        default:
          return CreateActionErrorResponse(
            400,
            0,
            `ActionVerbNotSupported: ${context.activity.value.action.verb} is not a supported action verb.`
          );
      }
    } catch (err) {
      return CreateActionErrorResponse(500, 0, err.message);
    }
  }
}
