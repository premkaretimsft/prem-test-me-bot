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
import { setTimeout as nodeTimeout } from "timers/promises";
import {
  markdownScenarios,
  buildCombinedMarkdown,
  streamingMarkdownChunks,
} from "./markdownContent";

/**
 * Logs the outgoing HTTP request details for sendActivity calls.
 * Captures headers and body payload for Postman replication.
 */
async function logOutgoingActivityRequest(
  context: TurnContext,
  activity: Partial<Activity>
): Promise<void> {
  try {
    const serviceUrl = context.activity.serviceUrl;
    const conversationId = context.activity.conversation.id;
    
    // Build the endpoint URL that the Bot Framework will call
    const endpoint = `${serviceUrl}v3/conversations/${encodeURIComponent(conversationId)}/activities`;
    
    // Get the connector client to access credentials info
    const connectorClient = context.turnState.get(context.adapter.ConnectorClientKey);
    
    // Build the request body (the activity payload)
    const requestBody: Partial<Activity> = {
      type: activity.type || ActivityTypes.Message,
      text: activity.text,
      entities: activity.entities,
      attachments: activity.attachments,
      from: context.activity.recipient, // Bot's identity
      recipient: context.activity.from, // User's identity
      conversation: context.activity.conversation,
      replyToId: context.activity.id,
      locale: context.activity.locale,
      ...activity
    };

    // Standard Bot Framework headers
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer <YOUR_ACCESS_TOKEN>', // Token would be obtained from credentials
      'User-Agent': 'Microsoft-BotFramework/3.1',
    };

    // Try to get the actual auth token if available
    if (connectorClient?.credentials?.getToken) {
      try {
        const tokenResponse = await connectorClient.credentials.getToken();
        if (tokenResponse) {
          headers['Authorization'] = `Bearer ${tokenResponse}`;
        }
      } catch (tokenError) {
        console.log('[LogRequest] Could not retrieve actual token, using placeholder');
      }
    }

    console.log('\n' + '='.repeat(80));
    console.log('📤 OUTGOING SENDACTIVITY REQUEST - POSTMAN EXPORT');
    console.log('='.repeat(80));
    console.log('\n🔗 ENDPOINT (POST):');
    console.log(endpoint);
    console.log('\n📋 HEADERS:');
    console.log(JSON.stringify(headers, null, 2));
    console.log('\n📦 REQUEST BODY:');
    console.log(JSON.stringify(requestBody, null, 2));
    console.log('\n📝 CURL COMMAND:');
    console.log(`curl -X POST "${endpoint}" \\`);
    console.log(`  -H "Content-Type: application/json" \\`);
    console.log(`  -H "Authorization: ${headers['Authorization']}" \\`);
    console.log(`  -d '${JSON.stringify(requestBody)}'`);
    console.log('\n' + '='.repeat(80) + '\n');
  } catch (error) {
    console.error('[LogRequest] Error logging outgoing request:', error);
  }
}

export class SearchApp extends TeamsActivityHandler {
  notifyContinuationActivity: any;
  continuationParameters: any;

  // Order of options shown in the numbered menu. Pure markdown scenarios first,
  // then "All Combined", then the modifier choices (Citations / Mentions /
  // Streaming) which mix one of the markdown scenarios with the corresponding
  // entity / streaming behavior.
  private menuOrder: string[] = [
    "Basic Formatting",
    "Links and Code",
    "Headings and Structure",
    "Lists and Emojis",
    "Performance Table",
    "Project Roadmap",
    "TypeScript Code",
    "KaTeX Math Equations",
    "Visual Elements",
    "Adaptive Card",
    "All Combined",
    "Citations",
    "Mentions",
    "Streaming",
    "AC",
    "20 Citations",
    "50 Citations",
    "51 Citations",
    "Suggested Action Invoke",
    "EM Mention (premk)",
    "Non-EM Mention (premk)",
  ];

  // Fixed user identity used by the "EM Mention (premk)" / "Non-EM Mention (premk)" menu
  // entries to probe how mention entities flow through the EM vs. non-EM bot→APX paths.
  private static readonly PREMK_MENTION = {
    id: "premk@testtenant3226.onmicrosoft.com",
    name: "premk",
  };

  constructor(
    notifyContinuationActivity: any,
    continuationParameters: any = {} /*conversations: any = {}*/
  ) {
    super();
    this.notifyContinuationActivity = notifyContinuationActivity;
    this.continuationParameters = continuationParameters;

    // Add the "All Combined" entry derived from every other markdown scenario.
    if (!markdownScenarios.has("All Combined")) {
      markdownScenarios.set("All Combined", buildCombinedMarkdown());
    }
  }

  public async onMessageActivity(context: TurnContext): Promise<void> {
    const userText = this.getCleanedUserText(context);

    // Show the menu on greeting, on `list`/`menu`/`help`, or on empty input.
    if (
      !userText ||
      userText === "list" ||
      userText === "menu" ||
      userText === "help" ||
      userText === "hi" ||
      userText === "hello"
    ) {
      await context.sendActivity(this.buildMenuActivity(context));
      return;
    }

    // Numeric selection from the menu.
    const num = parseInt(userText, 10);
    if (!isNaN(num) && num >= 1 && num <= this.menuOrder.length) {
      const selection = this.menuOrder[num - 1];
      try {
        await this.handleMenuSelection(context, selection);
      } catch (err) {
        await context.sendActivity(
          MessageFactory.text(
            `Error while sending '${selection}' response: ${(err as Error).message}`
          )
        );
        throw err;
      }
      return;
    }

    // Anything else — re-show the menu prefixed with what the user said.
    await context.sendActivity(
      this.buildMenuActivity(
        context,
        `I didn't recognize \`${context.activity.text}\`. Pick a number from below:\n\n`
      )
    );
  }

  private buildMenuActivity(context: TurnContext, prefix: string = ""): Partial<Activity> {
    const items = this.menuOrder
      .map((name, i) => {
        const note =
          name === "Citations"
            ? " — markdown body + citation entities"
            : name === "Mentions"
            ? " — markdown body + an @mention of you"
            : name === "Streaming"
            ? " — markdown delivered as a streamed response"
            : name === "All Combined"
            ? " — every markdown scenario in one message"
            : name === "AC"
            ? " — regular (non-EM) message with just an adaptive card attachment"
            : name === "Suggested Action Invoke"
            ? " — regular (non-EM) message carrying Action.Submit suggested actions that invoke the bot (per teams-modular-packages PR #1472218)"
            : name === "20 Citations"
            ? " — regular (non-EM) plain-text message with 20 citation entities (at the default MaxAllowedCitations of 20 → expects success)"
            : name === "50 Citations"
            ? " — regular (non-EM) plain-text message with 50 citation entities (rejected at default 20; succeeds once ECS sets MaxAllowedCitations=50)"
            : name === "51 Citations"
            ? " — regular (non-EM) plain-text message with 51 citation entities (over MaxAllowedCitations=50 → expects rejection)"
            : name === "EM Mention (premk)"
            ? ` — ExtendedMarkdown message with an @mention of ${SearchApp.PREMK_MENTION.id}`
            : name === "Non-EM Mention (premk)"
            ? ` — regular (non-EM) message with an @mention of ${SearchApp.PREMK_MENTION.id}`
            : "";
        return `${i + 1}. **${name}**${note}`;
      })
      .join("\n");

    const text =
      `${prefix}# 🎯 Markdown Test Bot\n\n` +
      `Reply with a number (1–${this.menuOrder.length}) to see that response. ` +
      `Type \`list\` anytime to see this menu again.\n\n${items}`;

    return this.buildExtendedMarkdownActivity(context, text, {});
  }

  private async handleMenuSelection(
    context: TurnContext,
    selection: string
  ): Promise<void> {
    if (selection === "Citations") {
      // Mix a markdown scenario with two inline citation references and the
      // citation entity payload. The bracketed `[1]`/`[2]` markers are what
      // Teams replaces with citation chips.
      const baseScenario = this.pickRandomMarkdownScenario([
        "Citations", "Mentions", "Streaming",
      ]);
      const baseMd = markdownScenarios.get(baseScenario)!;
      const md =
        `# ${baseScenario} *(with citations)*\n\n` +
        `${baseMd}\n\n---\n\n` +
        `**Inline citation references:** the heading example is backed by [1] ` +
        `and the supporting detail is sourced from [2].`;
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, md, { withCitations: true })
      );
      return;
    }

    if (selection === "Mentions") {
      const senderName = context.activity.from?.name || "there";
      const senderId =
        context.activity.from?.aadObjectId || context.activity.from?.id || "";
      const baseScenario = this.pickRandomMarkdownScenario([
        "Citations", "Mentions", "Streaming",
      ]);
      const baseMd = markdownScenarios.get(baseScenario)!;
      const md =
        `Hi <at>${senderName}</at>! Here's a quick **${baseScenario}** demo for you:\n\n${baseMd}`;
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, md, {
          withMention: { id: senderId, name: senderName },
        })
      );
      return;
    }

    if (selection === "EM Mention (premk)") {
      // ExtendedMarkdown message that @-mentions the fixed premk UPN — used to
      // probe how an EM-formatted bot→APX activity carries the mention entity
      // and the inline <at> tag through the EM encoder path.
      const { id, name } = SearchApp.PREMK_MENTION;
      const md =
        `Hi <at>${name}</at>! This is an **ExtendedMarkdown** message that @-mentions you.`;
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, md, {
          withMention: { id, name },
        })
      );
      return;
    }

    if (selection === "Non-EM Mention (premk)") {
      // Regular (non-EM) message that @-mentions the fixed premk UPN — used to
      // probe the standard bot→APX mention flow (Markdig + plainTextMentions
      // regex) for parity with the EM path above. textFormat is intentionally
      // left unset so APX takes the default markdown→HTML conversion route.
      const { id, name } = SearchApp.PREMK_MENTION;
      await context.sendActivity({
        type: "message",
        text: `Hi <at>${name}</at>! This is a **regular (non-EM)** message that @-mentions you.`,
        entities: [
          {
            type: "mention",
            mentioned: { id, name },
            text: `<at>${name}</at>`,
          },
        ],
      });
      return;
    }

    if (selection === "Streaming") {
      await this.processMarkdownStreamingRequest(context);
      return;
    }

    if (selection === "AC") {
      // Regular (non-EM) message — no text, just the customer-picker adaptive card.
      // Intentionally does NOT set textFormat=extendedmarkdown.
      await context.sendActivity({
        type: "message",
        attachments: [this.getCustomerPickerAdaptiveCardAttachment()],
      });
      return;
    }

    if (selection === "Suggested Action Invoke") {
      // Regular (non-EM) message carrying Action.Submit suggested-action buttons
      // per teams-modular-packages PR #1472218 (SuggestedActionInvoke):
      //   - "Approve" / "Reject" → no `value.name` → client invokes with the
      //     default name "suggestedActions/submit".
      //   - "Approve (custom name)" → carries `value.name: "voteInvoke"` →
      //     client invokes with that custom name instead.
      // Both paths are handled in onInvokeActivity below.
      const fromId = context.activity.from?.id;
      await context.sendActivity({
        type: "message",
        text: "Please review this request and choose an action:",
        suggestedActions: {
          to: fromId ? [fromId] : [],
          actions: [
            {
              type: "Action.Submit" as any,
              title: "Approve",
              value: { vote: "approve" },
            },
            {
              type: "Action.Submit" as any,
              title: "Reject",
              value: { vote: "reject" },
            },
            {
              type: "Action.Submit" as any,
              title: "Approve (custom invoke name)",
              value: { name: "voteInvoke", vote: "approve" },
            },
          ],
        },
      });
      return;
    }

    // "<N> Citations" — regular (non-EM) plain-text message with N inline
    // citation markers (`[1]`..`[N]`) and N Claim entities. Used to probe
    // the APX MaxAllowedCitations limit at various boundaries.
    const citationsMatch = /^(\d+) Citations$/.exec(selection);
    if (citationsMatch) {
      const count = parseInt(citationsMatch[1], 10);
      const text =
        `Probe message with ${count} citations: ` +
        Array.from({ length: count }, (_, i) => `claim ${i + 1} [${i + 1}]`).join("; ") +
        ".";
      await context.sendActivity({
        type: "message",
        text,
        entities: [
          {
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "https://schema.org",
            "@id": "",
            additionalType: ["AIGeneratedContent"],
            citation: this.generateCitations(count),
          } as any,
        ],
      });
      return;
    }

    // Plain markdown scenario (Basic Formatting, Lists and Emojis, ..., All Combined).
    const md = markdownScenarios.get(selection) || `(no content for ${selection})`;
    const titled =
      selection === "All Combined" ? md : `# ${selection}\n\n${md}`;
    await context.sendActivity(
      this.buildExtendedMarkdownActivity(context, titled, {})
    );
  }

  // Generates `count` Claim citation entities, one per position 1..count.
  // Each one is a small unique fake source so APX has distinct items to count
  // against the MaxAllowedCitations limit.
  private generateCitations(count: number): any[] {
    const sourceImages = ["microsoft excel", "microsoft word", "microsoft powerpoint", "microsoft onenote"];
    return Array.from({ length: count }, (_, i) => {
      const position = i + 1;
      return {
        "@type": "Claim",
        position,
        appearance: {
          "@type": "DigitalDocument",
          name: `Sample source #${position}`,
          url: `https://www.example.com/source-${position}`,
          abstract: `Auto-generated abstract for sample source #${position}. This is filler text used purely to give the citation entity a non-empty body for testing the max-citations limit.`,
          image: { name: sourceImages[i % sourceImages.length] },
          keywords: [`source-${position}`, "test-data", "max-citations-probe"],
          usageInfo: {
            "@type": "CreativeWork",
            description: "Please don't share outside of the company",
            name: "Company level sensitivity",
          },
        },
      };
    });
  }

  // Returns the customer-picker adaptive card as an Attachment.
  // Used by the "AC" menu option to send a regular (non-EM) message that
  // carries this card as its only payload (no text content).
  private getCustomerPickerAdaptiveCardAttachment(): Attachment {
    return {
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        type: "AdaptiveCard",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            size: "medium",
            weight: "bolder",
            text: "Select a Customer",
            wrap: true,
            choices: null,
            placeholder: null,
          },
          {
            type: "Input.ChoiceSet",
            id: "customerTPID",
            style: "compact",
            isMultiSelect: false,
            choices: [
              {
                title: "1+78XD+3396 : U N U M LIFE INSURANCE COMPANY",
                value: "1+78XD+3396",
              },
              {
                title: "1A0DPEP : CA-STATE GOVERNMENT",
                value: "1A0DPEP",
              },
              {
                title: "1-J6DC0L : AXIS SPECIALTY US SERVICES INC",
                value: "1-J6DC0L",
              },
            ],
            placeholder: "Select a customer (Account Number: Account Name)",
            size: null,
            text: null,
            weight: null,
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            data: {
              actionSubmitId: "Submit",
            },
            title: "Submit",
          },
        ],
      },
    };
  }

  // In group chats / channels Teams requires the user to @-mention the bot
  // before any input ("@PrkareInventory 19"), which surfaces in
  // activity.text as "<at>Prkare Inventory</at> 19". The Bot Framework's
  // TurnContext.removeRecipientMention uses the mention entities on the
  // activity to strip the bot's own <at>...</at> tag, leaving just the user's
  // input. In 1:1 chats the activity has no bot mention entity, so this is a
  // no-op and falls through to the raw text.
  private getCleanedUserText(context: TurnContext): string {
    const rawText = context.activity.text || "";
    if (!rawText || !context.activity.recipient?.id) {
      return rawText.toLowerCase().trim();
    }
    const stripped = TurnContext.removeRecipientMention(context.activity) || rawText;
    return stripped.toLowerCase().trim();
  }

  private pickRandomMarkdownScenario(exclude: string[]): string {
    const candidates = Array.from(markdownScenarios.keys()).filter(
      (k) => !exclude.includes(k)
    );
    return candidates[Math.floor(Math.random() * candidates.length)];
  }

  private buildExtendedMarkdownActivity(
    context: TurnContext,
    text: string,
    opts: {
      withCitations?: boolean;
      withMention?: { id: string; name: string };
    }
  ): Partial<Activity> {
    const schemaEntity: any = {
      type: "https://schema.org/Message",
      "@type": "Message",
      "@context": "https://schema.org",
      "@id": "",
      additionalType: ["AIGeneratedContent"],
    };

    if (opts.withCitations) {
      schemaEntity.usageInfo = {
        name: "Org level sensitivity",
        description: "Please don't share outside of the organization",
      };
      schemaEntity.citation = this.getSampleCitations();
    }

    const entities: any[] = [schemaEntity];

    if (opts.withMention) {
      entities.push({
        type: "mention",
        mentioned: {
          id: opts.withMention.id,
          name: opts.withMention.name,
        },
        text: `<at>${opts.withMention.name}</at>`,
      });
    }

    const activity: Partial<Activity> = {
      type: "message",
      // serviceUrl is intentionally NOT set here. TurnContext.applyConversationReference
      // overwrites it from context.activity.serviceUrl on every send, so the
      // egress URL is pinned by the adapter middleware in index.ts.
      text,
      channelData: { feedbackLoop: { type: "default" } },
      entities,
    };
    // textFormat: "extendedmarkdown" is a Teams-specific value not yet in the
    // botbuilder Activity typings, so it's attached via index access.
    (activity as any).textFormat = "extendedmarkdown";
    return activity;
  }

  private getSampleCitations(): any[] {
    return [
      {
        "@type": "Claim",
        position: 1,
        appearance: {
          "@type": "DigitalDocument",
          name: "Beverages data in the company",
          url: "https://www.microsoft.com",
          abstract:
            `From the web There are also references to a "Chai's Inventory Sorter" which is a mod for Minecraft ` +
            `that allows for inventory sorting and management. However, this is likely not related to your query.`,
          image: { name: "microsoft excel" },
          keywords: [
            "Company Data",
            "Recently Updated",
            "2026-04-01 09:00:00",
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
          url: "https://www.microsoft.com",
          abstract:
            `Quarterly revenue summary across business units. Use this source for verifying any revenue or ` +
            `growth-related figures cited in the response.`,
          image: { name: "microsoft word" },
          keywords: ["Revenue", "Quarterly", "2026"],
          usageInfo: {
            "@type": "CreativeWork",
            description: "Please don't share outside of the company",
            name: "Company level sensitivity",
          },
        },
        claimInterpreter: {
          "@type": "Project",
          name: "Claim Interpreter name",
          slogan: "Claim Interpreter slogan",
          url: "https://www.example.com/claim-interpreter",
        },
      },
    ];
  }

  // Streams a markdown response in chunks. Each chunk is appended to the
  // accumulated text and sent as a Typing activity (streamType=streaming),
  // and the last one is sent as a Message activity (streamType=final).
  private async processMarkdownStreamingRequest(
    context: TurnContext,
    packetDelay: number = 900
  ): Promise<void> {
    const initialActivity: Partial<Activity> = {
      type: ActivityTypes.Typing,
      text: "🔍 Preparing streamed markdown response...",
      entities: [
        {
          type: "streaminfo",
          streamType: "informative",
          streamSequence: 1,
        } as any,
      ],
    };
    (initialActivity as any).textFormat = "extendedmarkdown";

    await logOutgoingActivityRequest(context, initialActivity);
    const result = await context.sendActivity(initialActivity);
    const streamId = result?.id;
    console.log(`streamId: ${streamId}`);

    let accumulated = "";
    for (let i = 0; i < streamingMarkdownChunks.length; i++) {
      await nodeTimeout(packetDelay);
      accumulated += (i === 0 ? "" : "\n\n") + streamingMarkdownChunks[i];

      const isFinal = i === streamingMarkdownChunks.length - 1;
      const streamEntity: any = {
        type: "streaminfo",
        streamType: isFinal ? "final" : "streaming",
        streamId,
      };
      if (!isFinal) streamEntity.streamSequence = i + 2;

      const activity: Partial<Activity> = {
        type: isFinal ? ActivityTypes.Message : ActivityTypes.Typing,
        text: accumulated,
        entities: [streamEntity],
      };
      (activity as any).textFormat = "extendedmarkdown";

      await context.sendActivity(activity);
    }
  }

  public async onTeamsMessageEdit(context: TurnContext): Promise<void> {
    const editedText = context.activity.text || "(empty)";
    const md =
      `📝 I noticed you edited a message: \`${editedText}\`.\n\n` +
      `Type a number 1–${this.menuOrder.length} or \`list\` to see the menu.`;
    await context.sendActivity(this.buildExtendedMarkdownActivity(context, md, {}));
  }

  private getChartInputActionsCard(): Attachment {
    return CardFactory.adaptiveCard({
      type: "AdaptiveCard",
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.6",
      body: [
        {
          type: "TextBlock",
          text: "TextBlock element to show dynamic content with some input elements below followed by actions",
          size: "large",
          separator: true,
          spacing: "large",
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
          data: {
            hiddenKey: 456.12,
            msteams: {
              type: "invoke"
            }
          }
        },
        {
          type: "Action.Submit",
          title: "Diff Action.Submit",
          data: {
            hiddenKey: 123.45,
            msteams: {
              type: "invoke"
            }
          }
        },
        {
          type: "Action.OpenUrl",
          title: "Action.OpenUrl",
          url: "https://www.microsoft.com",
        },
        {
          type: "Action.OpenUrl",
          title: "OpenUrl Info",
          url: "https://youtu.be/ceV3RsG946s?si=7Z4eSo2Ak8ZYXVZW",
        },
        {
          type: "Action.Execute",
          id: "executeAction",
          title: "Action.Execute",
          verb: "submitVerb",
          data: {
            key1: "value1",
            key2: "value2",
            msTeams: {
              type: "imBack",
              value: "value3",
            },
          },
        }
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
          data: {
            hiddenKey: 456.12,
            msteams: {
              type: "invoke"
            }
          }
        },
        {
          type: "Action.Submit",
          title: "Diff Action.Submit",
          data: {
            hiddenKey: 123.45,
            msteams: {
              type: "invoke"
            }
          }
        },
        {
          type: "Action.OpenUrl",
          title: "Action.OpenUrl",
          url: "https://www.microsoft.com",
        },
        {
          type: "Action.OpenUrl",
          title: "OpenUrl Info",
          url: "https://youtu.be/ceV3RsG946s?si=7Z4eSo2Ak8ZYXVZW",
        },
        {
          type: "Action.Execute",
          id: "executeAction",
          title: "Action.Execute",
          verb: "submitVerb",
          data: {
            key1: "value1",
            key2: "value2",
            msTeams: {
              type: "imBack",
              value: "value3",
            },
          },
        }
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
        partialActivity: this.buildMenuActivity(context),
      };

      setTimeout(async () => await this.notifyContinuationActivity(), 1000);
    }
  }

  public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    try {
      switch (context.activity.name) {
        case "suggestedActions/submit":
        case "voteInvoke": {
          // Both invoke names are produced by the "Suggested Action Invoke" menu
          // option (teams-modular-packages PR #1472218):
          //   - "suggestedActions/submit" is the default when the bot's
          //     suggested-action `value` has no `name` field.
          //   - "voteInvoke" is the custom name the bot embedded in `value.name`
          //     to override the default.
          // activity.value carries the flat payload sent by the bot, e.g.
          //   { vote: "approve" }  or  { name: "voteInvoke", vote: "approve" }.
          console.log(`Received '${context.activity.name}' invoke. value=${JSON.stringify(context.activity.value)} fullActivity=${JSON.stringify(context.activity)}`);
          // Return 200 with empty body per the bot framework invoke pattern.
          // The bot can optionally send follow-up messages using context.sendActivity().
          return { status: 200 };
        }
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
