import {
  TeamsActivityHandler,
  TeamsInfo,
  teamsGetChannelId,
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
import config from "./config";
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
import {
  streamingRedPngB64,
  streamingGreenPngB64,
  streamingBluePngB64,
} from "./streamingBase64Images";

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
    "EM Private Entities",
    "Non-EM AI + Sensitivity + Feedback",
    "EM Credit Card (DLP)",
    "Non-EM Credit Card (DLP)",
    "Non-EM AC + Text",
    "EM AC + Text",
    "Non-EM Citations Table (repro)",
    "EM Citations Table (repro)",
    "Non-EM Citations Table (no sensitivity)",
    "EM Citations Table (no sensitivity)",
    "EM Inline Images (url + base64)",
    "Non-EM Inline Images (url + base64)",
    "Non-EM Signin Popup (tab vs popup)",
    "EM Streaming Base64 Images (partial packets)",
    "EM Stop Streaming Test (10 packets)",
    "EM Suggested Actions (HITL)",
    "Non-EM Suggested Actions (HITL)",
    "Non-EM Citation Special Chars (IcM 875738424)",
    "EM Citation Special Chars (IcM 875738424)",
  ];

  // Fixed user identity used by the "EM Mention (premk)" / "Non-EM Mention (premk)" menu
  // entries to probe how mention entities flow through the EM vs. non-EM bot→APX paths.
  private static readonly PREMK_MENTION = {
    id: "premk@testtenant3226.onmicrosoft.com",
    name: "premk",
  };

  // imBack values sent by the HITL suggested-action buttons, so a click is acknowledged
  // instead of falling through to "I didn't recognize".
  private static readonly HITL_REPLY_PREFIX = "hitl: ";

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

    if (userText.startsWith(SearchApp.HITL_REPLY_PREFIX)) {
      await context.sendActivity(
        MessageFactory.text(`Got suggested action: **${userText.substring(SearchApp.HITL_REPLY_PREFIX.length)}**`)
      );
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
            : name === "EM Private Entities"
            ? " — ExtendedMarkdown message carrying a sample `privateEntities` entity (value = stringified JSON); APX stamps it onto the message properties when ExtendedMarkdownPrivateEntitiesEnabled is on"
            : name === "Non-EM AI + Sensitivity + Feedback"
            ? " — regular (non-EM) message with the AI-generated label, a root-level sensitivity label, per-citation sensitivity labels (2 citations), and the feedback-loop opt-in"
            : name === "EM Credit Card (DLP)"
            ? " — ExtendedMarkdown message with fictitious credit-card details (Luhn-valid Visa test number + brand/expiry/CVV) to trip the tenant's Purview financial-info DLP policy"
            : name === "Non-EM Credit Card (DLP)"
            ? " — the SAME fictitious credit-card content as #24 but as a regular (non-EM / HTML) message (no textFormat), to compare DLP behavior EM vs non-EM"
            : name === "Non-EM AC + Text"
            ? " — regular (non-EM) message carrying BOTH text content and a simple adaptive card attachment in the same activity (no textFormat)"
            : name === "EM AC + Text"
            ? " — ExtendedMarkdown message carrying BOTH text content and a simple adaptive card attachment in the same activity (textFormat=extendedmarkdown), to compare card+text behavior EM vs non-EM"
            : name === "Non-EM Citations Table (repro)"
            ? " — regular (non-EM) message with the exact bug #4630824 repro payload: a markdown table whose cells carry citation markers [1] (x5), [2], [3], [4] plus 4 AI-generated citations. Reproduces the client's partial citation rendering inside table cells"
            : name === "EM Citations Table (repro)"
            ? " — the SAME citations-in-a-table payload as the non-EM repro but as an ExtendedMarkdown message (textFormat=extendedmarkdown), to confirm EM renders all citations inside the table"
            : name === "Non-EM Citations Table (no sensitivity)"
            ? " — same payload as #28 (Non-EM Citations Table) but with NO sensitivity/usageInfo at either the message level or the citation level, to isolate whether sensitivity affects citation rendering"
            : name === "EM Citations Table (no sensitivity)"
            ? " — same payload as #29 (EM Citations Table) but with NO sensitivity/usageInfo at either the message level or the citation level"
            : name === "Non-EM Signin Popup (tab vs popup)"
            ? " — regular (non-EM) message with an adaptive card carrying TWO `Action.Submit` buttons that use `msteams.type=\"signin\"` (the partner repro). One sign-in URL has no size params, the other adds `?width=600&height=800`. Desktop pops out for both; in the browser only the second should open as a popup instead of a new tab"
            : name === "EM Suggested Actions (HITL)"
            ? " — ExtendedMarkdown HITL approval prompt with suggested-action buttons (2 imBack + 1 Action.Submit). Identical to the next option except textFormat"
            : name === "Non-EM Suggested Actions (HITL)"
            ? " — the SAME HITL prompt and buttons as a regular (non-EM) message, as the control"
            : name === "Non-EM Citation Special Chars (IcM 875738424)"
            ? " — regular (non-EM) message with one citation per abstract variation: [1] the incident abstract, [2]-[12] raw, encoded, numeric and double-encoded `&`, `<` `>`, markup, quotes, unknown entities and the 160-char cut-off, [13]-[15] lookalike workarounds (fullwidth ＆ U+FF06, small ﹠ U+FE60, fullwidth ＜ ＞)"
            : name === "EM Citation Special Chars (IcM 875738424)"
            ? " — the SAME citations as the previous option but as an ExtendedMarkdown message, to confirm EM shows the same thing"
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

    if (selection === "EM Private Entities") {
      // ExtendedMarkdown message that carries a sample `privateEntities` entity.
      // On the APX side, the EM encoder's processPrivateEntities looks for an
      // entity whose type is "privateEntities" (case-insensitive), reads its
      // stringified-JSON `value`, and stamps it onto the outgoing ChatService
      // message properties (MessagePropertiesKeys.PrivateEntities) — gated by the
      // ExtendedMarkdownPrivateEntitiesEnabled flag (sub-gated under EM). Per the
      // bot→APX contract the `value` is a stringified JSON object at the root.
      const samplePrivateEntities = {
        schemaVersion: "1.0",
        source: "prem-test-me-bot",
        entities: [
          { id: "ent-001", kind: "product", name: "Chai", sku: "NWND-1" },
          { id: "ent-002", kind: "person", displayName: "Nancy Davolio" },
        ],
        metadata: { generatedAt: new Date().toISOString(), confidential: true },
      };
      const md =
        `# Private Entities demo\n\n` +
        `This is an **ExtendedMarkdown** message that carries a sample ` +
        `\`privateEntities\` entity in its payload. The entity body is **not** ` +
        `rendered in the message — APX extracts its \`value\` and stamps it onto ` +
        `the message properties when \`ExtendedMarkdownPrivateEntitiesEnabled\` is on.`;
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, md, {
          withPrivateEntities: samplePrivateEntities,
        })
      );
      return;
    }

    if (selection === "EM Credit Card (DLP)") {
      // ExtendedMarkdown (textFormat=extendedmarkdown) message carrying the shared
      // fictitious credit-card content (see dlpCreditCardMarkdown). Pairs with
      // "Non-EM Credit Card (DLP)" (#25) — identical content, EM transport — so the
      // tenant's Purview financial-info DLP behavior can be A/B'd EM vs non-EM.
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, this.dlpCreditCardMarkdown(), {})
      );
      return;
    }

    if (selection === "EM Inline Images (url + base64)") {
      // ExtendedMarkdown message carrying BOTH inline image kinds so the outgoing payload can be
      // compared side by side: the https URLs should reach the client URLP-wrapped
      // (.../v1/url/content?url=...) with no AMS document created, while the base64 image should be
      // uploaded to AMS (.../v1/objects/{id}/views/img_preview) and listed in amsreferences.
      // Uses the five images from the reported issue, two of which are the upload.wikimedia.org URLs
      // that the current EM path fails to fetch (HTTP 403, no User-Agent) and leaves as raw URLs.
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, this.inlineImagesMarkdown("ExtendedMarkdown"), {})
      );
      return;
    }

    if (selection === "Non-EM Inline Images (url + base64)") {
      // Same five images as the EM option, but expressed the way a regular message carries images:
      // real <img> tags, since non-EM content is HTML. textFormat is intentionally left unset so APX
      // takes the default markdown->HTML path, which passes the raw tags through untouched.
      await context.sendActivity({
        type: "message",
        text: this.inlineImagesHtml("Non-ExtendedMarkdown (HTML)"),
      });
      return;
    }

    if (selection === "Non-EM Credit Card (DLP)") {
      // Byte-identical credit-card content to #24, but sent as a REGULAR (non-EM)
      // message: textFormat is intentionally left unset, so APX takes the default
      // markdown->HTML (Markdig) path instead of the ExtendedMarkdown path. Isolates
      // the EM vs non-EM variable for the DLP comparison.
      await context.sendActivity({
        type: "message",
        text: this.dlpCreditCardMarkdown(),
      });
      return;
    }

    if (selection === "Non-EM AI + Sensitivity + Feedback") {
      // 1) Immediate response in the current conversation — 1:1 chat, group chat,
      //    OR the channel reply chain where the bot was invoked. Unchanged behavior.
      await context.sendActivity(this.buildNonEmAiSensitivityFeedbackActivity());

      // 2) ONLY when invoked from a channel, ALSO post the same message as a new
      //    top-level post in that channel (its own thread), in addition to the
      //    reply above. Personal / group-chat experience is left untouched.
      if (context.activity.conversation?.conversationType === "channel") {
        const channelId = teamsGetChannelId(context.activity);
        const botAppId = config.botId;
        if (channelId && botAppId) {
          // sendMessageToTeamsChannel creates a brand-new conversation (top-level
          // post) in the channel, distinct from the reply sent above.
          await TeamsInfo.sendMessageToTeamsChannel(
            context,
            this.buildNonEmAiSensitivityFeedbackActivity() as Activity,
            channelId,
            botAppId
          );
        } else {
          console.warn(
            `[Non-EM AI] channel post skipped — channelId=${channelId}, botAppId set=${!!botAppId}`
          );
        }
      }
      return;
    }

    if (selection === "Streaming") {
      await this.processMarkdownStreamingRequest(context);
      return;
    }

    if (selection === "EM Streaming Base64 Images (partial packets)") {
      await this.processBase64StreamingRequest(context);
      return;
    }

    if (selection === "EM Stop Streaming Test (10 packets)") {
      await this.processStopStreamingTestRequest(context);
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

    if (selection === "Non-EM Signin Popup (tab vs popup)") {
      // Replicates the partner scenario: a regular (non-EM) message whose adaptive
      // card carries Action.Submit buttons with `msteams: { type: "signin", value: <url> }`.
      //
      // The Teams client routes this through signInFn -> launchAuthWindow ->
      // window.open(url, "appAuthWindow", features), where `features` is built ONLY
      // when both `width` and `height` are present in the sign-in URL's query string.
      // No features => browsers open a new TAB; with features => a real popup.
      // Teams desktop pops out either way. Two buttons here so both can be A/B'd
      // in a single card.
      //
      // The target is just a plain public page, not a real sign-in endpoint — this
      // option only exists to observe the window shape, not to complete an auth flow.
      await context.sendActivity({
        type: "message",
        text:
          `Sign-in popup repro. Click each button and note whether the page opens ` +
          `as a **popup window** or a **new browser tab**.`,
        attachments: [this.getSignInPopupAdaptiveCardAttachment()],
      });
      return;
    }

    if (selection === "Non-EM AC + Text") {
      // Regular (non-EM) message that carries BOTH text content and a simple
      // adaptive card attachment in the SAME activity. textFormat is intentionally
      // left unset so APX takes the default markdown->HTML path for the text while
      // the card rides along as an attachment. Pairs with "EM AC + Text" so the
      // card+text combo behavior can be A/B'd EM vs non-EM.
      await context.sendActivity({
        type: "message",
        text:
          `Here's a **regular (non-EM)** message with text content **and** a simple ` +
          `adaptive card attachment in the same activity.`,
        attachments: [this.getSimpleAdaptiveCardAttachment()],
      });
      return;
    }

    if (selection === "EM AC + Text") {
      // ExtendedMarkdown (textFormat=extendedmarkdown) message that carries BOTH text
      // content and a simple adaptive card attachment in the SAME activity. Pairs with
      // "Non-EM AC + Text" — identical intent, EM transport — to isolate the EM vs
      // non-EM variable for the card+text combo.
      const md =
        `Here's an **ExtendedMarkdown** message with text content **and** a simple ` +
        `adaptive card attachment in the same activity.`;
      await context.sendActivity(
        this.buildExtendedMarkdownActivity(context, md, {
          attachments: [this.getSimpleAdaptiveCardAttachment()],
        })
      );
      return;
    }

    if (selection === "Non-EM Citations Table (repro)") {
      // Exact repro payload from bug #4630824 (Corina's comment) sent as a regular
      // (non-EM / markdown) message — no textFormat. This is the failing case: the
      // client renders the citation chips only partially inside the table cells.
      await context.sendActivity(this.buildCitationTableReproActivity(false));
      return;
    }

    if (selection === "EM Citations Table (repro)") {
      // Byte-identical citations-in-a-table payload to "Non-EM Citations Table (repro)"
      // but sent as an ExtendedMarkdown message (textFormat=extendedmarkdown). The only
      // variable under test is textFormat, to confirm the EM path renders every citation
      // inside the table.
      await context.sendActivity(this.buildCitationTableReproActivity(true));
      return;
    }

    if (selection === "Non-EM Citations Table (no sensitivity)") {
      // Same citations-in-a-table payload as "Non-EM Citations Table (repro)" but with
      // sensitivity/usageInfo explicitly omitted at both the message level and every
      // citation level, to eliminate sensitivity as a possible cause of the client's
      // citation-rendering issue. Regular (non-EM / markdown) message — no textFormat.
      await context.sendActivity(
        this.buildCitationTableReproActivity(false, { includeSensitivity: false })
      );
      return;
    }

    if (selection === "EM Citations Table (no sensitivity)") {
      // Same as the no-sensitivity non-EM variant above but sent as an ExtendedMarkdown
      // message (textFormat=extendedmarkdown). Sensitivity/usageInfo omitted at both the
      // message level and every citation level.
      await context.sendActivity(
        this.buildCitationTableReproActivity(true, { includeSensitivity: false })
      );
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

    if (selection === "EM Suggested Actions (HITL)" || selection === "Non-EM Suggested Actions (HITL)") {
      await context.sendActivity(
        this.buildSuggestedActionsActivity(context, selection.startsWith("EM "))
      );
      return;
    }

    if (
      selection === "Non-EM Citation Special Chars (IcM 875738424)" ||
      selection === "EM Citation Special Chars (IcM 875738424)"
    ) {
      await context.sendActivity(this.buildCitationSpecialCharsActivity(selection.startsWith("EM ")));
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

  // Repro payload from bug #4630824 (Corina's comment): a markdown table whose cells
  // carry citation markers [1] (x5), [2], [3], [4] together with 4 AI-generated
  // citations (positions 1-4). The Web client renders the citation chips only partially
  // inside table cells for a regular (non-EM / markdown) message. Sending the identical
  // content as ExtendedMarkdown lets us confirm the EM path renders them all.
  private citationTableReproText(): string {
    return (
      "| Item | Info |\n| --- | --- |\n" +
      "| Program name | CT 03b - International Women's Day gift |\n" +
      "| Gift | 01 Blindbox STYLE drink (SKU 40004083) [1] |\n" +
      "| Store count | 67 selected stores [1] |\n" +
      "| Slots per store | 10 first female customers [1] |\n" +
      "| Total gifts | 670 units (10 x 67) |\n" +
      "| Invoice condition | From 1,000,000d and above [1] |\n" +
      "| Time window | Only 1 day - 08/03/2026 [2] |\n" +
      "| Cost code (IO) | 260C090042 - Retail Mar_QTPN 8.3 [3] |\n" +
      "| Gift method | Manual at store, noted on gift invoice [1] |\n" +
      "| Participation limit | Each customer once, verified via phone/CCCD [4] |"
    );
  }

  // The 4 Claim citations (positions 1-4) attached to the repro payload. Mirrors the
  // SDK repro's CitationAppearance(name=`Source {pos}`, abstract=`Ref {pos}.`,
  // url=`https://example.com/s{pos}`).
  private citationTableReproCitations(): any[] {
    return Array.from({ length: 4 }, (_, i) => {
      const position = i + 1;
      return {
        "@type": "Claim",
        position,
        appearance: {
          "@type": "DigitalDocument",
          name: `Source ${position}`,
          abstract: `Ref ${position}.`,
          url: `https://example.com/s${position}`,
        },
      };
    });
  }

  // IcM 875738424 repro: one abstract variation per citation, so each can be checked before and
  // after the client fix. [13]-[15] are lookalike workarounds whose look depends on the font.
  private buildCitationSpecialCharsActivity(extendedMarkdown: boolean): Partial<Activity> {
    const citation = (position: number, name: string, abstract: string, keywords: string[]) => ({
      "@type": "Claim",
      position,
      appearance: {
        "@type": "DigitalDocument",
        name,
        url: `https://www.example.com/special-chars-${position}`,
        abstract,
        image: { name: "microsoft word" },
        keywords,
      },
    });

    const activity: Partial<Activity> = {
      type: "message",
      text:
        "Citation abstract variations (IcM 875738424): incident abstract [1], raw ampersand [2], " +
        "encoded ampersand [3], numeric ampersands [4], double-encoded ampersand [5], " +
        "raw angle brackets [6], encoded angle brackets [7], encoded markup [8], real tags [9], " +
        "quotes and other entities [10], unknown entities [11], ampersand at the length limit [12]. " +
        "Workarounds: fullwidth ampersand [13], small ampersand [14], fullwidth angle brackets [15].",
      entities: [
        {
          type: "https://schema.org/Message",
          "@type": "Message",
          "@context": "https://schema.org",
          "@id": "",
          additionalType: ["AIGeneratedContent"],
          citation: [
            citation(1, "Incident abstract", "D&A stands for Data and Analytics. D&amp;A.", ["D&A", "R&D"]),
            citation(2, "Raw ampersand", "AT&T, R&D and D&A", ["variation"]),
            citation(3, "Encoded ampersand", "D&amp;A stands for Data and Analytics.", ["variation"]),
            citation(4, "Numeric ampersands", "D&#38;A and D&#x26;A", ["variation"]),
            citation(5, "Double-encoded ampersand", "Show the entity itself: D&amp;amp;A", ["variation"]),
            citation(6, "Raw angle brackets", "Q3 < Q4 > Q2", ["variation"]),
            citation(7, "Encoded angle brackets", "Q3 &lt; Q4 &gt; Q2", ["variation"]),
            citation(8, "Encoded markup", "&lt;b&gt;bold&lt;/b&gt; and &lt;img src=x&gt;", ["variation"]),
            citation(9, "Real tags", "<b>Bold</b> and <i>italic</i> tags are stripped", ["variation"]),
            citation(10, "Quotes and other entities", "&quot;Quoted&quot;, it&#39;s caf&eacute; &copy; 2026 and \"raw\" it's", ["variation"]),
            citation(11, "Unknown entities", "&foo; stays as typed, &bar", ["variation"]),
            citation(12, "Ampersand at the length limit", "Lorem ipsum ".repeat(13) + "R&D and more text past the limit.", ["variation"]),
            citation(13, "Workaround: fullwidth ampersand (U+FF06)", "D\uFF06A stands for Data and Analytics. D and A.", ["workaround"]),
            citation(14, "Workaround: small ampersand (U+FE60)", "D\uFE60A stands for Data and Analytics.", ["workaround"]),
            citation(15, "Workaround: fullwidth angle brackets", "Q3 \uFF1C Q4 \uFF1E Q2", ["workaround"]),
          ],
        } as any,
      ],
    };

    if (extendedMarkdown) {
      // textFormat: "extendedmarkdown" is a Teams-specific value not yet in the
      // botbuilder TextFormatTypes union.
      (activity as any).textFormat = "extendedmarkdown";
    }

    return activity;
  }

  // Builds the bug #4630824 repro activity. The EM and non-EM variants are byte-identical
  // except for textFormat, so the only variable under test is the message type.
  private buildCitationTableReproActivity(
    extendedMarkdown: boolean,
    opts: { includeSensitivity?: boolean } = {}
  ): Partial<Activity> {
    const { includeSensitivity = true } = opts;

    const citations = this.citationTableReproCitations();
    const messageEntity: any = {
      type: "https://schema.org/Message",
      "@type": "Message",
      "@context": "https://schema.org",
      "@id": "",
      additionalType: ["AIGeneratedContent"],
      citation: citations,
    };

    if (!includeSensitivity) {
      // No-sensitivity variant: guarantee the bot emits zero sensitivity/usageInfo at
      // BOTH the citation level and the message level. (The base repro already carries
      // none, so this is a defensive strip that stays correct if the repro changes.)
      for (const c of citations) {
        if (c?.appearance) {
          delete c.appearance.usageInfo;
        }
        delete c?.usageInfo;
      }
      delete messageEntity.usageInfo;
    }

    const activity: Partial<Activity> = {
      type: "message",
      text: this.citationTableReproText(),
      entities: [messageEntity],
    };
    if (extendedMarkdown) {
      // textFormat: "extendedmarkdown" is a Teams-specific value not yet in the
      // botbuilder Activity typings, so it's attached via index access.
      (activity as any).textFormat = "extendedmarkdown";
    }
    return activity;
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

  // Returns a minimal "simple" adaptive card as an Attachment. Used by the
  // "Non-EM AC + Text" and "EM AC + Text" menu options, which send this card as an
  // attachment alongside text content in the SAME activity (non-EM vs EM transport).
  private getSimpleAdaptiveCardAttachment(): Attachment {
    return {
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            size: "Medium",
            weight: "Bolder",
            text: "Simple Adaptive Card",
            wrap: true,
          },
          {
            type: "TextBlock",
            text:
              "This card was sent as an attachment alongside text content in the same activity.",
            wrap: true,
          },
        ],
      },
    };
  }

  // Card for the "Non-EM Signin Popup (tab vs popup)" option. Carries two
  // Action.Submit buttons using the `msteams.type = "signin"` shape the partner
  // team reported. The only difference between them is whether the sign-in URL
  // includes `width`/`height` query params, which is what decides popup vs tab
  // in a browser (Teams desktop pops out regardless).
  //
  // The target is a plain public page, not a real sign-in endpoint — this option
  // only exists to observe the window shape.
  //
  // The sign-in URL host must be listed in the manifest's `validDomains`
  // ("example.com"), otherwise the client blocks the action with AppNotExist
  // ("This action can't be performed since the app does not exist...") before
  // opening any window.
  private getSignInPopupAdaptiveCardAttachment(): Attachment {
    const baseUrl = "https://example.com/";
    return {
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            size: "Medium",
            weight: "Bolder",
            text: "Sign-in popup: tab vs popup",
            wrap: true,
          },
          {
            type: "TextBlock",
            text:
              "Both buttons use the same `msteams.type=\"signin\"` action. Only the " +
              "sign-in URL differs.",
            wrap: true,
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            title: "Connect (no size params)",
            data: {
              msteams: {
                type: "signin",
                value: baseUrl,
              },
            },
          },
          {
            type: "Action.Submit",
            title: "Connect (width=600&height=800)",
            data: {
              msteams: {
                type: "signin",
                value: `${baseUrl}?width=600&height=800`,
              },
            },
          },
        ],
      },
    };
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

  // HITL approval prompt for the "EM / Non-EM Suggested Actions (HITL)" options. Both options
  // send exactly this payload; textFormat is the only difference, so a missing button row on
  // the EM one points at the EM path rather than the payload. Mixes imBack (the usual HITL
  // shape) with Action.Submit so both client render paths are exercised.
  private buildSuggestedActionsActivity(context: TurnContext, isExtendedMarkdown: boolean): Partial<Activity> {
    const fromId = context.activity.from?.id;
    const activity: Partial<Activity> = {
      type: "message",
      text:
        "**Approval needed** — deploy `build 2026.09.29.1` to Canary?\n\n" +
        "- Scope: EUS2 ring\n" +
        "- Risk: low\n\n" +
        "Pick an action below.",
      suggestedActions: {
        // The client hides the buttons unless `to` includes the viewer, so omit it rather
        // than send an empty list when the sender is unknown.
        ...(fromId ? { to: [fromId] } : {}),
        actions: [
          { type: "imBack" as any, title: "Approve", value: SearchApp.HITL_REPLY_PREFIX + "approve" },
          { type: "imBack" as any, title: "Reject", value: SearchApp.HITL_REPLY_PREFIX + "reject" },
          {
            type: "Action.Submit" as any,
            title: "Approve (invoke)",
            value: { vote: "approve" },
          },
        ],
      } as any,
    };

    if (isExtendedMarkdown) {
      // textFormat: "extendedmarkdown" is a Teams-specific value not yet in the
      // botbuilder TextFormatTypes union.
      (activity as any).textFormat = "extendedmarkdown";
    }

    return activity;
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
      withPrivateEntities?: Record<string, unknown>;
      attachments?: Attachment[];
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

    if (opts.withPrivateEntities) {
      // APX reads the stringified-JSON `value` off this entity (case-insensitive)
      // and stamps it onto the ChatService message properties under the
      // PrivateEntities key. The value must be a stringified JSON object at the
      // root level, so JSON.stringify the sample object here.
      entities.push({
        type: "privateEntities",
        value: JSON.stringify(opts.withPrivateEntities),
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

    if (opts.attachments && opts.attachments.length > 0) {
      activity.attachments = opts.attachments;
    }

    return activity;
  }

  // Fictitious credit-card content shared by "EM Credit Card (DLP)" (#24) and
  // "Non-EM Credit Card (DLP)" (#25) so both send byte-identical content and the only
  // variable under test is EM vs non-EM (textFormat). The number is Luhn-valid and
  // Visa-format but NOT a real card (and deliberately not the commonly-excluded test
  // number 4111111111111111); the brand + expiry + CVV within 300 chars push Purview's
  // built-in "Credit Card Number" sensitive info type to high confidence.
  // Ref: https://learn.microsoft.com/en-us/purview/sit-defn-credit-card-number
  // The five images from the reported issue, shared by the EM and non-EM "Inline Images" menu
  // options so both send exactly the same sources and the only variable is how they are expressed:
  // markdown image syntax for EM (whose content stays raw markdown) vs. real <img> tags for non-EM
  // (whose content is HTML), matching how each was sent in the report.
  private static readonly ISSUE_IMAGES = {
    fronalpstock:
      "https://upload.wikimedia.org/wikipedia/commons/3/3f/Fronalpstock_big.jpg",
    pinterest:
      "https://i.pinimg.com/474x/87/54/71/8754710a5c3df0bd5d24c0f194cff85e.jpg",
    example: "https://upload.wikimedia.org/wikipedia/commons/a/a9/Example.jpg",
    w3c: "https://www.w3.org/Icons/w3c_home.png",
    inlineBase64Png:
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAKAAAABgCAIAAAAVRe7OAAABGUlEQVR42u3dsQ2AIBQEUEawcgO3snMEJ6B2ImeyN7HSng5iAT/vcgscr/+kV0IneQLAAliGAF62q6rPPjd3OnNzj/Vubu3GEfcCBgwYMGDAgAEDBgwYMGDAgAEDBgwYcM/AkGLsBQwYMGDAgAEDBgwYMGDAgAEDBgwYMGDAgAEDBgwYMOASGFKMvYABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwZcAkOKsRcwYMCAAQMGDBgwYMCAAQMGDBgwYMCAAQMGDBgwYMCAHQR38R0wYMCAAQMGDBgwYMCAAQMGDBgwYMCAAQMGDBgwYMD+D4bk4jtgwIABAwYMGDBgwIABAwYMGDBgwIAB/wEs8QIYsACWbvMBOtL0y/mKwtwAAAAASUVORK5CYII=",
  };

  // EM payload for the "EM Inline Images" menu option. ExtendedMarkdown content stays raw markdown
  // (no HTML conversion), so images are expressed with markdown image syntax.
  //
  // These are the exact five images from the reported issue, in the order they were reported.
  // Two of them matter especially: upload.wikimedia.org rejects requests carrying no User-Agent with
  // HTTP 403, so on the current EM path (which downloads each URL and re-uploads it to AMS) images 1
  // and 3 fail and their raw URLs are left in the text, while images 2 and 4 become AMS documents.
  private inlineImagesMarkdown(label: string): string {
    const img = SearchApp.ISSUE_IMAGES;

    return (
      `# Inline image handling — ${label}\n\n` +
      `The same five images from the reported issue, inline with text, so the outgoing payload ` +
      `shows how each kind is handled.\n\n` +
      `**1. Wikimedia JPEG (Fronalpstock, ~14.7 MB):** ![fronalpstock](${img.fronalpstock}) — the ` +
      `first of the two that failed in the report.\n\n` +
      `**2. Pinterest JPEG:** ![pinterest](${img.pinterest}) — this one succeeded in the report.\n\n` +
      `**3. Wikimedia JPEG (Example.jpg):** ![example](${img.example}) — the second one that failed.\n\n` +
      `**4. W3C home PNG:** ![w3c](${img.w3c}) — this one succeeded in the report.\n\n` +
      `**5. Inline base64 PNG:** ![inline base64](${img.inlineBase64Png}) — expected to become an AMS object.\n\n` +
      `---\n\n` +
      `Today (pre-fix) images 1 and 3 stay as raw \`upload.wikimedia.org\` URLs because APX's ` +
      `download gets a 403, 2 and 4 become \`/v1/objects/{id}/views/img_preview\`, and 5 becomes an ` +
      `AMS object — with no \`ams_references\`, so all three expire after 7 days. After the fix all ` +
      `four URLs become \`/v1/url/content?url=\` and only 5 creates a document, which is claimed.`
    );
  }

  // Non-EM payload for the "Non-EM Inline Images" menu option. A regular message's text is converted
  // to HTML, so inline images belong in real <img> tags here rather than markdown image syntax —
  // this is the "HTML Message" shape from the report. textFormat is intentionally left unset, and
  // Markdig passes raw HTML through untouched, so these tags reach the client verbatim.
  //
  // APX itself does not rewrite these URLs; MFE does. MFE's MessageSanitizer runs only for
  // RichText/Html and, for bot-initiated messages, sets PrefixUntrustedDomainsWithUrlp, which
  // rewrites every untrusted <img src> to the regional URLP endpoint. That is why non-EM URLs arrive
  // URLP-wrapped while EM (RichText/ExtendedMarkdown) is skipped by that sanitizer entirely.
  private inlineImagesHtml(label: string): string {
    const img = SearchApp.ISSUE_IMAGES;

    return (
      `<h1>Inline image handling — ${label}</h1>\n\n` +
      `<p>The same five images from the reported issue, inline with text, sent as real ` +
      `<code>&lt;img&gt;</code> tags so the outgoing payload shows how each kind is handled.</p>\n\n` +
      `<p><b>1. Wikimedia JPEG (Fronalpstock, ~14.7 MB):</b> ` +
      `<img src="${img.fronalpstock}" alt="fronalpstock"> — the first of the two that failed in the report.</p>\n\n` +
      `<p><b>2. Pinterest JPEG:</b> <img src="${img.pinterest}" alt="pinterest"> — this one succeeded in the report.</p>\n\n` +
      `<p><b>3. Wikimedia JPEG (Example.jpg):</b> ` +
      `<img src="${img.example}" alt="example"> — the second one that failed.</p>\n\n` +
      `<p><b>4. W3C home PNG:</b> <img src="${img.w3c}" alt="w3c"> — this one succeeded in the report.</p>\n\n` +
      `<p><b>5. Inline base64 PNG:</b> ` +
      `<img src="${img.inlineBase64Png}" alt="inline base64"> — expected to become an AMS object.</p>\n\n` +
      `<hr>\n\n` +
      `<p>Expected: MFE's sanitizer rewrites the four http URLs to ` +
      `<code>/urlp/v1/url/content?url=</code>, and image 5 arrives as ` +
      `<code>/v1/objects/{id}/views/img_preview</code> because APX uploaded it to AMS.</p>`
    );
  }

  private dlpCreditCardMarkdown(): string {
    return (
      `# 🔐 DLP test — sensitive financial data\n\n` +
      `This is a **fictitious test message** from prem-test-me-bot to exercise the ` +
      `tenant's Microsoft Purview DLP policy for financial information. The card ` +
      `details below are **not real**.\n\n` +
      `- **Cardholder:** Test User\n` +
      `- **Card type:** Visa\n` +
      `- **Card number:** 4532 0151 1283 0978\n` +
      `- **Expires:** 11/27\n` +
      `- **CVV:** 123\n\n` +
      `Expected: your DLP policy should flag or block this message.`
    );
  }

  // Builds the non-EM "AI + sensitivity + feedback" message used by menu option
  // #23. Returns a FRESH activity each call (getSampleCitations() allocates new
  // arrays) so the same payload can be sent both as the immediate reply and as a
  // separate top-level channel post without sharing mutable state.
  private buildNonEmAiSensitivityFeedbackActivity(): Partial<Activity> {
    const aiEntity: any = {
      type: "https://schema.org/Message",
      "@type": "Message",
      "@context": "https://schema.org",
      "@id": "",
      additionalType: ["AIGeneratedContent"],
      usageInfo: {
        "@type": "CreativeWork",
        name: "Org level sensitivity",
        description: "Please don't share outside of the organization",
      },
      citation: this.getSampleCitations(),
    };
    return {
      type: "message",
      text:
        `**AI-generated summary.** Chai inventory looks healthy this quarter [1], ` +
        `and revenue is trending up versus last period [2].`,
      channelData: { feedbackLoop: { type: "default" } },
      entities: [aiEntity],
    };
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

  // Builds the delta sequence for the "EM Streaming Base64 Images" option.
  //
  // Unlike streamingMarkdownChunks these are raw deltas, concatenated with no
  // separator, so a single data URI can be cut across several packets. That is the
  // whole point of this scenario: APX does not touch images in intermediate
  // streaming packets, so the client sees a data URI that is still arriving.
  //
  // The sequence deliberately mixes the cases that render differently:
  //   1-4  a data URI split mid-base64 across three packets, then completed
  //   5-7  a data URI inside a fenced block (should stay literal text, never an image)
  //   8    a whole image arriving in one packet
  //   9-10 an image split across a fenced markdown block boundary
  //   11   trailing text after the last image
  private static base64StreamingDeltas(): string[] {
    const red = streamingRedPngB64;
    const green = streamingGreenPngB64;
    const blue = streamingBluePngB64;

    // Split red into three uneven pieces so the cut lands mid-base64 both times.
    const r1 = red.slice(0, 120);
    const r2 = red.slice(120, 400);
    const r3 = red.slice(400);

    // Split blue across the fenced-block boundary.
    const b1 = blue.slice(0, 260);
    const b2 = blue.slice(260);

    return [
      "## Streaming base64 images\n\n" +
        "Packet 1 is plain text - no image yet. Watch each packet render as it arrives.\n\n",

      "**Case A - data URI cut mid-base64.** The image below is incomplete until packet 4:\n\n" +
        `![red stripes](data:image/png;base64,${r1}`,

      r2,

      `${r3})\n\n`,

      "Packet 5: text after the completed image. Case A should now be a red striped image.\n\n",

      "**Case B - data URI inside a fenced block.** This must stay literal text, never an image:\n\n" +
        "```text\n" +
        `data:image/png;base64,${green.slice(0, 200)}`,

      `${green.slice(200, 380)}`,

      "\n```\n\n" +
        "**Case C - a whole image in a single packet:**\n\n" +
        `![green stripes](data:image/png;base64,${green})\n\n`,

      "**Case D - image split across a fenced markdown block.** The fence opens here and closes in the next packet:\n\n" +
        "```markdown\n" +
        `![blue stripes](data:image/png;base64,${b1}`,

      `${b2})\n` + "```\n\n",

      "Packet 11 (final): trailing text. Cases A and C should be images; B and D should be code.\n",
    ];
  }

  // Streams the base64 image scenario. Mirrors processMarkdownStreamingRequest
  // (informative packet, then streaming Typing packets, then a final Message) but
  // appends each delta verbatim instead of joining with blank lines, so a data URI
  // can be delivered half-finished in an intermediate packet.
  private async processBase64StreamingRequest(
    context: TurnContext,
    packetDelay: number = 1200
  ): Promise<void> {
    const deltas = SearchApp.base64StreamingDeltas();

    const initialActivity: Partial<Activity> = {
      type: ActivityTypes.Typing,
      text: "🖼️ Streaming base64 images in intermediate packets...",
      entities: [
        {
          type: "streaminfo",
          streamType: "informative",
          streamSequence: 1,
        } as any,
      ],
    };
    (initialActivity as any).textFormat = "extendedmarkdown";

    const result = await context.sendActivity(initialActivity);
    const streamId = result?.id;
    console.log(`base64 streaming streamId: ${streamId}`);

    let accumulated = "";
    for (let i = 0; i < deltas.length; i++) {
      await nodeTimeout(packetDelay);
      accumulated += deltas[i];

      const isFinal = i === deltas.length - 1;
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

  // Content for the "EM Stop Streaming Test" option. Ten self-contained blocks of
  // plain text and fenced code, joined with blank lines, so the message still reads
  // sensibly whenever the user stops it part way through.
  //
  // Deliberately contains NO inline images: an image in ExtendedMarkdown streaming
  // makes the final packet fail Thread Service content-stream validation (the final
  // rewrites image srcs, the intermediate packets do not), which would mask the
  // stop-button behaviour this option is meant to exercise.
  private static stopStreamingTestChunks(): string[] {
    return [
      "## Stop streaming test\n\n" +
        "This streams in 10 packets with a ~3s gap between each one. " +
        "Hit the **Stop** button while it is still streaming.",

      "### Step 1 - plain text\n\n" +
        "Streaming has started. Each packet appends one more block to the message.",

      "### Step 2 - a fenced code block\n\n" +
        "```ts\n" +
        "export function greet(name: string): string {\n" +
        "  return `Hello, ${name}!`;\n" +
        "}\n" +
        "```",

      "Text after the first fenced block. The stream is still going.",

      "### Step 3 - a second fenced block\n\n" +
        "```json\n" +
        "{\n" +
        '  "streamType": "streaming",\n' +
        '  "packet": 5\n' +
        "}\n" +
        "```",

      "### Step 4 - a list\n\n" + "- First item\n" + "- Second item\n" + "- Third item",

      "More text between blocks. If you have not hit **Stop** yet, now is a good time.",

      "### Step 5 - a shell block\n\n" +
        "```bash\n" +
        "curl -s https://example.com/health | jq .status\n" +
        "```",

      "Almost done - one more packet after this one.",

      "### Done\n\n" +
        "If you can read this, the stream finished on its own. " +
        "Run the option again and hit **Stop** earlier.",
    ];
  }

  // Streams the stop-button test. Same shape as processMarkdownStreamingRequest
  // (informative packet, then streaming Typing packets, then a final Message) but
  // with its own content and a slower cadence so there is time to press Stop.
  //
  // Pressing Stop makes the client PUT the message back with
  // streamingMetadata.streamType=Final and messagetype RichText/ExtendedMarkdown,
  // which is the path MFE PR 1533537 unblocked for the TeamsClient partner.
  private async processStopStreamingTestRequest(
    context: TurnContext,
    packetDelay: number = 3000
  ): Promise<void> {
    const chunks = SearchApp.stopStreamingTestChunks();

    const initialActivity: Partial<Activity> = {
      type: ActivityTypes.Typing,
      text: "⏳ Starting a 10-packet stream - press Stop any time...",
      entities: [
        {
          type: "streaminfo",
          streamType: "informative",
          streamSequence: 1,
        } as any,
      ],
    };
    (initialActivity as any).textFormat = "extendedmarkdown";

    const result = await context.sendActivity(initialActivity);
    const streamId = result?.id;
    console.log(`stop-streaming test streamId: ${streamId}`);

    let accumulated = "";
    for (let i = 0; i < chunks.length; i++) {
      await nodeTimeout(packetDelay);
      accumulated += (i === 0 ? "" : "\n\n") + chunks[i];

      const isFinal = i === chunks.length - 1;
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
