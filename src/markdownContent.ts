// Markdown content samples used by SearchApp to demonstrate ExtendedMarkdown
// rendering in Teams. Adapted from the Teams-Markdown-Sampler bot.

export const markdownScenarios = new Map<string, string>([
  [
    "Basic Formatting",
    `**Bold Text Example:** This text should appear bold and emphasized!

_Italic Text Example:_ This text should appear slanted and stylized with elegant cursive-like rendering

***Bold Italic Combined:*** This combines both bold AND italic formatting for maximum emphasis!

~~Strikethrough Example:~~ This text should appear crossed out and deprecated`,
  ],

  [
    "Links and Code",
    `**Link Example:** Click here to visit [Microsoft Teams Developer Platform Documentation](https://docs.microsoft.com/microsoftteams/platform) for comprehensive guides

**Code Span Example:** \`const apiEndpoint = "https://graph.microsoft.com/v1.0/me"; // inline code formatting\`

**Special Characters:** Handle special characters: @#$%^&*(){}[]|\\;':",./<>?`,
  ],

  [
    "Headings and Structure",
    `# 🚀 Main Heading: Advanced Markdown Rendering Test Suite

## 🎯 Subheading: React-Markdown Component Performance Analysis

### 📈 Sub-subheading: Real-time Rendering Statistics Dashboard

> **Quote Example:** "Imagination is more important than knowledge. Knowledge is limited, whereas imagination embraces the entire world, stimulating progress, giving birth to evolution." - Albert Einstein
>
> This multi-line blockquote demonstrates proper indentation and styling.`,
  ],

  [
    "Lists and Emojis",
    `**🛠️ Feature Testing Checklist:**

- 🔥 Process complex nested markdown structures
- ⚡ Handle special characters and symbols
- 🎯 Render emojis and Unicode: 🌟✨🎉🚀💡🔥⭐
- 🌍 Support international text: Héllo Wörld! 你好世界! مرحبا بالعالم!
- ✨ Format mathematical expressions: E=mc² and π≈3.14159

**🔬 Numbered Process:**

1. 📊 Initialize test environment with sample data
2. 🧪 Execute controlled experiments
3. 📈 Collect performance metrics
4. 🔍 Analyze results using statistical tests
5. 📝 Generate comprehensive reports`,
  ],

  [
    "Performance Table",
    `📊 **Performance Metrics Dashboard:**

Component | Render Time (ms) | Memory Usage (MB) | CPU Load (%) | Status | Optimization Score
--- | --- | --- | --- | --- | ---
Header Navigation | 12.5 | 2.3 | 1.2 | 🟢 Optimal | 95/100
Content Renderer | 45.8 | 8.7 | 4.5 | 🟡 Good | 87/100
Markdown Parser | 23.1 | 5.2 | 2.8 | 🟢 Excellent | 98/100
Image Processor | 156.3 | 15.9 | 12.7 | 🟠 Moderate | 73/100
Table Generator | 34.7 | 6.1 | 3.2 | 🟢 Great | 91/100
Syntax Highlighter | 67.2 | 11.4 | 7.8 | 🟢 Good | 85/100`,
  ],

  [
    "Project Roadmap",
    `**🚀 Comprehensive Project Roadmap:**

- [x] ✅ Initialize markdown parsing engine with TypeScript support
- [x] 🎨 Implement custom styling for Teams-specific components
- [x] 📱 Test responsive design across different screen sizes
- [x] 🔧 Configure webpack optimization for production builds
- [ ] 🧪 Conduct A/B testing with focus groups (target: 500 users)
- [ ] 🌍 Add internationalization support for 12+ languages
- [ ] 🔒 Implement advanced security measures and data encryption
- [ ] 📈 Deploy analytics tracking for user engagement metrics
- [ ] 🚀 Launch beta version to Microsoft Teams App Store
- [ ] 🏆 Achieve 4.8+ star rating and 10,000+ active installations`,
  ],

  [
    "TypeScript Code",
    `**💻 Advanced Code Example - Teams Bot Implementation:**

\`\`\`typescript
// Advanced Teams Bot with Adaptive Cards and Graph API integration
import { TeamsActivityHandler, CardFactory, MessageFactory } from 'botbuilder';
import { Client } from '@microsoft/microsoft-graph-client';

interface TeamsMember {
  id: string;
  displayName: string;
  email: string;
  roles: string[];
  lastActive: Date;
}

class AdvancedTeamsBot extends TeamsActivityHandler {
  private graphClient: Client;

  constructor(graphClient: Client) {
    super();
    this.graphClient = graphClient;
  }

  protected async onMessageActivity(context: TurnContext): Promise<void> {
    const userMessage = context.activity.text?.toLowerCase();

    if (userMessage?.includes('analytics')) {
      const analyticsCard = this.createAnalyticsCard();
      await context.sendActivity(MessageFactory.attachment(analyticsCard));
    }
  }

  private createAnalyticsCard(): Attachment {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Advanced Analytics Dashboard',
        weight: 'Bolder',
        size: 'Large'
      }]
    });
  }
}

export { AdvancedTeamsBot, TeamsMember };
\`\`\``,
  ],

  [
    "KaTeX Math Equations",
    `**🧮 Mathematical Expressions with KaTeX:**

**Quadratic Formula:** The solution to $ax^2 + bx + c = 0$ is given by:

$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$

**Euler's Identity:** One of the most beautiful equations in mathematics:

$$e^{i\\pi} + 1 = 0$$

**Fourier Transform:** Converting from time domain to frequency domain:

$$F(\\omega) = \\int_{-\\infty}^{\\infty} f(t) e^{-i\\omega t} dt$$

**Inline Math Examples:**
- Area of circle: $A = \\pi r^2$
- Pythagorean theorem: $a^2 + b^2 = c^2$
- Derivative: $\\frac{d}{dx}(x^n) = nx^{n-1}$

**Matrix Operations:**

$$\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix} \\begin{pmatrix} x \\\\ y \\end{pmatrix} = \\begin{pmatrix} ax + by \\\\ cx + dy \\end{pmatrix}$$`,
  ],

  [
    "Visual Elements",
    `**🎨 Visual Separator with Custom Styling:**

---

**🌈 Multi-line Formatting Demonstration:**

🎭 **Current Status:** Testing react-markdown rendering capabilities
⚡ **Processing Speed:** 2,847 operations per second
🎯 **Accuracy Rate:** 99.7% successful markdown transformations
🔮 **Next Phase:** Advanced interactive component integration
✨ **Final Goal:** Seamless Teams client markdown experience!

**🖼️ Dynamic Test Image:** ![Markdown Test Visualization](https://via.placeholder.com/400x200/FF6B6B/FFFFFF?text=React+Markdown+Test+Suite)`,
  ],

  [
    "Adaptive Card",
    `**🎴 Interactive Adaptive Card Example:**

This example demonstrates how Teams can render rich, interactive Adaptive Cards using the \`adaptivecard\` fenced block syntax:

\`\`\`adaptivecard
{
  "type": "AdaptiveCard",
  "version": "1.6",
  "body": [
    {
      "type": "TextBlock",
      "size": "Medium",
      "weight": "Bolder",
      "text": "Publish Adaptive Card schema"
    },
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "items": [
            {
              "type": "Image",
              "style": "Person",
              "url": "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
              "size": "Small"
            }
          ],
          "width": "auto"
        },
        {
          "type": "Column",
          "items": [
            { "type": "TextBlock", "weight": "Bolder", "text": "Matt Hidinger", "wrap": true },
            { "type": "TextBlock", "spacing": "None", "text": "Created {{DATE(2017-02-14T06:08:39Z,SHORT)}}", "isSubtle": true, "wrap": true }
          ],
          "width": "stretch"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Board:", "value": "Adaptive Card" },
        { "title": "List:", "value": "Backlog" },
        { "title": "Assigned to:", "value": "Matt Hidinger" },
        { "title": "Due date:", "value": "Not set" }
      ]
    }
  ]
}
\`\`\``,
  ],
]);

// Build a single combined markdown document from every base scenario above.
export function buildCombinedMarkdown(): string {
  const all = Array.from(markdownScenarios.entries())
    .map(([title, content]) => `# 📋 ${title}\n\n${content}`)
    .join("\n\n---\n\n");
  return `# 🎯 Complete Markdown Test Suite — All Scenarios Combined\n\n${all}`;
}

// Markdown chunks used by the streaming demo. Each chunk is appended to the
// previous one (cumulative) and the bot sends typing/final activities until
// the full document is delivered.
export const streamingMarkdownChunks: string[] = [
  `## 🎬 Streaming markdown demo

Walking you through a few markdown features delivered as a **streamed** response.`,
  `## ✨ Formatting

Bold (**bold**), italic (*italic*), strikethrough (~~old~~), and inline \`code\` all render as you'd expect.`,
  `## 📋 Lists

- 🔥 First item
- ⚡ Second item with **bold** inside
- 🎯 Third item

1. Step one
2. Step two
3. Step three`,
  `## 💻 Code blocks

\`\`\`typescript
async function streamMarkdown(): Promise<void> {
  console.log("Streaming markdown is fun!");
}
\`\`\``,
  `## 🧮 A bit of math

The quadratic formula:

$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$

✅ Stream complete!`,
];
