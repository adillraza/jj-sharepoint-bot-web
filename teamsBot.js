const { TeamsActivityHandler, CardFactory, TurnContext } = require('botbuilder');
const { DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');
const { SharePointGraphClient } = require('./graph');
const { DIALOG_ID } = require('./mainDialog');

// Environment variables
const CONNECTION_NAME = process.env.ConnectionName || "GraphConnection";
const CLIENT_ID = process.env.MicrosoftAppId;
const TENANT_ID = process.env.MicrosoftAppTenantId;

class SharePointBot extends TeamsActivityHandler {
    constructor(conversationState, mainDialog) {
        super();
        
        this.conversationState = conversationState;
        this.dialogState = this.conversationState.createProperty("DialogState");
        
        // Host the dialog
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(mainDialog);

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded || [];
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    const welcomeText = `
👋 **Welcome to SharePoint Document Assistant!**

I can help you find and read your SharePoint documents. Try these commands:

📋 **Get Started:**
• \`help\` - Show all commands
• \`signin\` - Connect to Microsoft 365

🔍 **Find Documents:**
• \`recent\` - Your recent files
• \`search [keyword]\` - Find documents

❓ **Ask Questions:**
• "What's in the project plan?"
• "Show me Excel files"

Type \`help\` to see all available commands!
                    `;
                    await context.sendActivity(welcomeText);
                }
            }
            await next();
        });

        this.onMessage(async (context, next) => {
            // Remove @mentions for Teams compatibility
            const text = TurnContext.removeRecipientMention(context.activity)?.text?.trim() || 
                         context.activity.text?.trim() || '';

            console.log(`SharePointBot received: "${text}"`);
            
            const lowerText = text.toLowerCase().trim();
            
            if (lowerText === 'signin' || lowerText === 'login' || lowerText === 'connect') {
                // TEMPORARILY DISABLE OAuth for debugging
                console.log('🔐 OAuth temporarily disabled for debugging');
                await context.sendActivity('🚧 **OAuth temporarily disabled for debugging**\n\n' +
                    'This confirms the bot is working. The OAuth issue needs to be resolved at the Azure configuration level.\n\n' +
                    'Try these basic commands:\n' +
                    '• `test` - Test bot functionality\n' +
                    '• `help` - See all commands');
                return;
                
                // ORIGINAL CODE (disabled):
                // console.log('🔐 Starting OAuth dialog...');
                // const dc = await this.dialogs.createContext(context);
                // await dc.beginDialog(DIALOG_ID);
                        } else if (lowerText === 'token') {
                // TEMPORARILY DISABLE token check for debugging
                await context.sendActivity('🚧 **Token check temporarily disabled for debugging**\n\n' +
                    'This avoids the OAuth 404 error. The issue is in the Bot Framework OAuth configuration.');
            } else {
                try {
                    await this.handleUserMessage(context, text);
                } catch (error) {
                    console.error('Error in SharePointBot:', error);
                    await context.sendActivity('Sorry, I encountered an error processing your request. Please try again.');
                }
            }

            await next();
        });

        // Handle OAuth token responses
        this.onTokenResponseEvent(async (context, next) => {
            console.log('Token response received');
            await next();
        });
    }

    async run(context) {
        const dc = await this.dialogs.createContext(context);
        const result = await dc.continueDialog();
        if (result.status === DialogTurnStatus.empty) {
            // no-op; normal onMessage handled above
        }
        await this.conversationState.saveChanges(context, false);
        await super.run(context);
    }

    async handleUserMessage(context, text) {
        const lowerText = text.toLowerCase().trim();
        
        // Test command (no auth needed)
        if (lowerText === 'test' || lowerText === 'ping') {
            await context.sendActivity('✅ **Bot is working!**\n\nBasic functionality confirmed. Environment variables:\n' +
                `• Client ID: ${CLIENT_ID ? 'SET' : 'MISSING'}\n` +
                `• Tenant ID: ${TENANT_ID ? 'SET' : 'MISSING'}\n` +
                `• Connection: ${CONNECTION_NAME}`);
            return;
        }
        

        
        // Help command
        if (lowerText === 'help' || lowerText === 'commands') {
            const helpText = `
**📋 SharePoint Document Assistant Commands:**

**🔐 Authentication:**
• \`signin\` - Sign in to Microsoft 365
• \`logout\` - Sign out

**📁 Document Discovery:**
• \`recent\` - Show recent documents
• \`search [query]\` - Search SharePoint documents
• \`find [filename]\` - Find specific files

**📄 Document Actions:**
• \`read [filename]\` - Read document content
• \`open [filename]\` - Get document link

**❓ Q&A:**
• Just ask questions about your documents!
• "What's in the project plan?"
• "Show me budget documents"
• "Find Excel files about sales"

**💡 Examples:**
• \`recent\` - See your latest files
• \`search budget\` - Find budget-related docs
• \`read project-plan.docx\` - Read document content
            `;
            
            await context.sendActivity(helpText);
            return;
        }

        

        // Sign-out command
        if (lowerText === 'logout' || lowerText === 'signout') {
            try {
                await context.adapter.signOutUser(context, CONNECTION_NAME);
                await context.sendActivity('✅ You have been signed out. Type `signin` to connect again.');
            } catch (error) {
                console.error('Sign-out error:', error);
                await context.sendActivity('You are now signed out.');
            }
            return;
        }

        // Get user token using Bot Framework OAuth
        const tokenResponse = await context.adapter.getUserToken(context, CONNECTION_NAME);
        if (!tokenResponse || !tokenResponse.token) {
            await context.sendActivity('🔐 **Please sign in first**\n\nType `signin` to connect to Microsoft 365 and access your SharePoint documents.');
            return;
        }
        
        const token = tokenResponse.token;
        console.log(`🔍 Token found for user, length: ${token.length}`);

        const graphClient = new SharePointGraphClient(token);

        // Recent documents
        if (lowerText === 'recent' || lowerText === 'recent files') {
            try {
                await context.sendActivity('🔍 Fetching your recent documents...');
                const recentDocs = await graphClient.getRecentDocuments();
                
                if (recentDocs.value && recentDocs.value.length > 0) {
                    const docList = recentDocs.value.map((doc, i) => 
                        `${i + 1}. **${doc.name}**\n   📁 [Open file](${doc.webUrl})\n   📅 ${new Date(doc.lastModifiedDateTime).toLocaleDateString()}`
                    ).join('\n\n');
                    
                    await context.sendActivity(`**📂 Your recent documents:**\n\n${docList}`);
                } else {
                    await context.sendActivity('📁 No recent documents found.');
                }
            } catch (error) {
                console.error('Error fetching recent documents:', error);
                await context.sendActivity('❌ Sorry, I couldn\'t retrieve your recent documents. Please try again.');
            }
            return;
        }

        // Search commands
        if (lowerText.startsWith('search ') || lowerText.startsWith('find ')) {
            const query = text.substring(text.indexOf(' ') + 1);
            try {
                await context.sendActivity(`🔍 Searching for "${query}"...`);
                const results = await graphClient.getRecentDocuments();
                const filtered = results.value?.filter(doc => 
                    doc.name.toLowerCase().includes(query.toLowerCase())
                ) || [];
                
                if (filtered.length > 0) {
                    const docList = filtered.map((doc, i) => 
                        `${i + 1}. **${doc.name}**\n   📁 [Open file](${doc.webUrl})\n   📅 Modified: ${new Date(doc.lastModifiedDateTime).toLocaleDateString()}`
                    ).join('\n\n');
                    
                    await context.sendActivity(`**🎯 Found ${filtered.length} documents matching "${query}":**\n\n${docList}`);
                } else {
                    await context.sendActivity(`❌ No documents found matching "${query}".\n\n💡 Try:\n• Different keywords\n• Broader search terms\n• \`recent\` to see all your files`);
                }
            } catch (error) {
                console.error('Error searching documents:', error);
                await context.sendActivity('❌ Sorry, I couldn\'t search documents at this time. Please try again.');
            }
            return;
        }

        // Default: treat as a question about documents
        await context.sendActivity(`💭 I understand you're asking: "${text}"\n\n🚧 **Document Q&A is coming soon!**\n\nFor now, try:\n• \`recent\` - See your recent files\n• \`search [keyword]\` - Find documents\n• \`help\` - See all commands`);
    }
}

module.exports.SharePointBot = SharePointBot;