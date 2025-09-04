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
        
        console.log('🔍 SharePointBot constructor - Starting...');
        console.log('🔍 SharePointBot constructor - mainDialog:', mainDialog ? 'PROVIDED' : 'MISSING');
        
        this.conversationState = conversationState;
        this.dialogState = this.conversationState.createProperty("DialogState");
        
        // Host the dialog
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(mainDialog);
        
        console.log('✅ SharePointBot constructor - Dialog added successfully');

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
                // Start the OAuth dialog (re-enabled with new client secret)
                console.log('🔐 Starting OAuth dialog with new client secret...');
                console.log('🔍 CONNECTION_NAME:', CONNECTION_NAME);
                console.log('🔍 CLIENT_ID:', CLIENT_ID);
                console.log('🔍 TENANT_ID:', TENANT_ID);
                console.log('🔍 DIALOG_ID:', DIALOG_ID);
                
                try {
                    const dc = await this.dialogs.createContext(context);
                    console.log('✅ Dialog context created successfully');
                    await dc.beginDialog(DIALOG_ID);
                    console.log('✅ Dialog started successfully');
                } catch (error) {
                    console.error('❌ OAuth Dialog Error:', error);
                    console.error('❌ Error stack:', error.stack);
                    await context.sendActivity('❌ **OAuth Error**\n\nFailed to start sign-in dialog. Check logs for details.');
                }
            } else if (lowerText === 'token') {
                // Try to get a cached token (re-enabled with new client secret)
                try {
                    console.log('🔍 Checking for cached token with new client secret...');
                    const token = await context.adapter.getUserToken(context, CONNECTION_NAME);
                    if (token?.token) {
                        await context.sendActivity(`🔐 **Token available**\n\nFirst 20 chars: ${token.token.substring(0, 20)}...\n\nYou can now use commands like \`recent\` or \`search\`.`);
                    } else {
                        await context.sendActivity('❌ **No token found**\n\nType `signin` first to authenticate.');
                    }
                } catch (error) {
                    console.error('Token check error:', error);
                    await context.sendActivity('❌ **Error checking token**\n\nType `signin` to authenticate.');
                }
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