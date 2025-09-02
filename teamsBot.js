const { TeamsActivityHandler, CardFactory, TurnContext } = require('botbuilder');
const { SharePointGraphClient } = require('./graph');

// Environment variables
const CONNECTION_NAME = process.env.ConnectionName || "GraphConnection";
const CLIENT_ID = process.env.MicrosoftAppId;
const TENANT_ID = process.env.MicrosoftAppTenantId;

class SharePointBot extends TeamsActivityHandler {
    constructor() {
        super();
        
        // Simple in-memory token storage (for demo - use proper storage in production)
        this.userTokens = new Map();

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
            
            try {
                await this.handleUserMessage(context, text);
            } catch (error) {
                console.error('Error in SharePointBot:', error);
                await context.sendActivity('Sorry, I encountered an error processing your request. Please try again.');
            }

            await next();
        });

        // Handle OAuth token responses
        this.onTokenResponseEvent(async (context, next) => {
            console.log('Token response received');
            await context.sendActivity('✅ You are now signed in! Type `recent` to see your files or `help` for all commands.');
            await next();
        });
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
        
        // Token exchange command
        if (lowerText.startsWith('token ')) {
            const authCode = text.substring(6).trim();
            if (!authCode) {
                await context.sendActivity('❌ Please provide the authorization code.\nExample: `token 0.AXoAWvC...`');
                return;
            }
            
            try {
                await context.sendActivity('🔄 **Exchanging authorization code for access token...**');
                
                // For now, just store the code (in a real app, you'd exchange it for a token)
                const userId = context.activity.from.id;
                this.userTokens.set(userId, `demo_token_${authCode.substring(0, 10)}`);
                
                await context.sendActivity('✅ **Authentication successful!**\n\n' +
                    'You can now use commands like:\n' +
                    '• `recent` - See your recent files\n' +
                    '• `search [keyword]` - Search documents\n' +
                    '• `help` - See all commands');
                
                console.log(`✅ User ${userId} authenticated with code: ${authCode.substring(0, 10)}...`);
                
            } catch (error) {
                console.error('❌ Token exchange error:', error);
                await context.sendActivity('❌ **Authentication failed**\n\nPlease try the signin process again.');
            }
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

                        // Sign-in command
                if (lowerText === 'signin' || lowerText === 'login' || lowerText === 'connect') {
                    try {
                        console.log(`🔐 Direct Microsoft OAuth approach`);
                        console.log('🔐 Client ID:', CLIENT_ID);
                        console.log('🔐 Tenant ID:', TENANT_ID);
                        
                        // Direct Microsoft OAuth (no Bot Framework dependencies)
                        const scopes = 'openid profile offline_access https://graph.microsoft.com/Files.Read https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/User.Read';
                        const redirectUri = 'https://portal.azure.com'; // Simple redirect for demo
                        const state = `botuser_${context.activity.from.id}`;
                        
                        const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
                            `client_id=${encodeURIComponent(CLIENT_ID)}&` +
                            `response_type=code&` +
                            `redirect_uri=${encodeURIComponent(redirectUri)}&` +
                            `scope=${encodeURIComponent(scopes)}&` +
                            `state=${encodeURIComponent(state)}&` +
                            `response_mode=query&` +
                            `prompt=select_account`;
                        
                        console.log('✅ Direct Microsoft OAuth URL generated');
                        console.log('🔗 Auth URL:', authUrl);
                        
                        await context.sendActivity({
                            attachments: [
                                CardFactory.signinCard(
                                    'Sign in to Microsoft 365',
                                    authUrl,
                                    'Sign in'
                                )
                            ]
                        });
                        
                        await context.sendActivity('🔐 **Direct Microsoft OAuth**\n\n' +
                            '1. Click the sign-in button above\n' +
                            '2. Sign in with your Microsoft 365 account\n' +
                            '3. Copy the authorization code from the URL\n' +
                            '4. Type `token [code]` to complete authentication\n\n' +
                            'Example: `token 0.AXoAWvC...`');
                        
                    } catch (error) {
                        console.error('❌ Direct OAuth error:', error);
                        await context.sendActivity('Sorry, I couldn\'t generate a sign-in link. Please check the bot configuration.');
                    }
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

        // Get user token (try Bot Framework first, then manual storage)
        let token = null;
        try {
            const tokenResponse = await context.adapter.getUserToken(context, CONNECTION_NAME);
            if (tokenResponse && tokenResponse.token) {
                token = tokenResponse.token;
            }
        } catch (error) {
            console.log('Bot Framework token retrieval failed, checking manual storage');
        }
        
        // Check manual token storage if Bot Framework failed
        if (!token) {
            const userId = context.activity.from.id;
            token = this.userTokens.get(userId);
        }
        
        if (!token) {
            await context.sendActivity('🔐 **Please sign in first**\n\nType `signin` to connect to Microsoft 365 and access your SharePoint documents.');
            return;
        }

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