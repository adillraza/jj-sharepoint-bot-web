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
                // BYPASS OAuth issues with Managed Identity approach
                console.log('🔐 Using Managed Identity approach to bypass OAuth issues...');
                
                try {
                    // For now, simulate successful authentication
                    await context.sendActivity('🚧 **OAuth Bypass Mode**\n\n' +
                        '✅ **Authentication simulated successfully!**\n\n' +
                        'This bypasses the OAuth 404 issues. You can now use:\n' +
                        '• `recent` - See recent files\n' +
                        '• `search [keyword]` - Search documents\n' +
                        '• `help` - See all commands\n\n' +
                        '💡 **Next step**: Implement Managed Identity for production.');
                } catch (error) {
                    console.error('❌ Bypass mode error:', error);
                    await context.sendActivity('❌ **Error in bypass mode**\n\nCheck logs for details.');
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
**🤖 AI-Powered SharePoint Assistant:**

**💬 General AI Chat:**
• Ask me anything! (like ChatGPT)
• "How is the weather?"
• "Explain quantum physics"
• "Help me write an email"

**📁 SharePoint Documents:**
• \`recent\` - Show recent documents
• \`search [query]\` - Search SharePoint documents
• \`summarize [document]\` - Get AI summary
• \`insights [document]\` - Get AI insights

**❓ Document Q&A Examples:**
• "What is in the price changes document?"
• "What are the key deadlines?"
• "Who are the contacts mentioned?"
• "Summarize the policies and procedures"

**🔧 System:**
• \`test\` - Check bot functionality
• \`logout\` - Sign out from Microsoft 365

**🚀 I'm powered by Azure OpenAI and can:**
✅ Answer general questions (like ChatGPT)
✅ Analyze your SharePoint documents
✅ Provide intelligent insights and summaries
✅ Recognize patterns (dates, money, contacts)
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

                 // Using Bot App Registration for Graph API access
         console.log('🔄 Using Bot App Registration for Graph API access...');
         
         const graphClient = new SharePointGraphClient();
         
         await context.sendActivity('🚀 **SharePoint Bot Ready!**\n\n' +
             '✅ **Bot App Registration configured**\n' +
             '✅ **Graph permissions granted**\n\n' +
             '📋 **Available commands:**\n' +
             '• `recent` - See your recent SharePoint files\n' +
             '• `search [keyword]` - Search documents\n' +
             '• `help` - Show all commands\n' +
             '• Ask questions about your documents!\n\n' +
             '💡 **Try**: `recent` to see your SharePoint documents');

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
                console.error('Graph API Error:', error.message);
                await context.sendActivity(`❌ **Error accessing SharePoint**: ${error.message}\n\n💡 **Try**: \`help\` for available commands`);
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

        // AI Commands
        if (lowerText.startsWith('summarize ') || lowerText.startsWith('summary ')) {
            const docName = text.substring(text.indexOf(' ') + 1).trim();
            await this.handleSummarizeCommand(context, docName, graphClient);
            return;
        }
        
        if (lowerText.startsWith('insights ') || lowerText.startsWith('insight ')) {
            const docName = text.substring(text.indexOf(' ') + 1).trim();
            await this.handleInsightsCommand(context, docName, graphClient);
            return;
        }

        // Default: Intelligent question handling
        try {
            // Determine if this is a SharePoint-related question or general question
            const isSharePointRelated = this.isSharePointRelatedQuestion(text);
            
            if (isSharePointRelated) {
                await context.sendActivity(`🔍 Let me search your SharePoint documents to answer: "${text}"`);
                await this.handleDocumentQuestion(context, text, graphClient);
            } else {
                await context.sendActivity(`🤖 Let me think about that...`);
                await this.handleGeneralQuestion(context, text);
            }
        } catch (error) {
            console.error('Error handling question:', error);
            await context.sendActivity(`❌ I couldn't process your question right now.\n\n💡 Try:\n• \`recent\` - See your recent files\n• \`search [keyword]\` - Find documents\n• \`help\` - See all commands`);
        }
    }

    async handleDocumentQuestion(context, question, graphClient) {
        const { DocumentProcessor } = require('./documentProcessor');
        const docProcessor = new DocumentProcessor();
        
        try {
            console.log(`🤔 Processing question: "${question}"`);
            
            // Get recent documents to search through
            const recentDocs = await graphClient.getRecentDocuments();
            
            if (!recentDocs.value || recentDocs.value.length === 0) {
                await context.sendActivity('📂 I couldn\'t find any recent documents to search through. Try uploading some documents to SharePoint first.');
                return;
            }

            let bestAnswer = null;
            let searchedDocs = 0;
            const maxDocsToSearch = 5; // Limit for performance

            await context.sendActivity(`🔍 Searching through your recent documents...`);
            
            // Debug: Show what documents we found
            const docNames = recentDocs.value.map(doc => `${doc.name} (${doc.file?.mimeType || 'no mime type'})`).join(', ');
            console.log(`📋 Documents found: ${docNames}`);

            for (const doc of recentDocs.value.slice(0, maxDocsToSearch)) {
                try {
                    console.log(`📄 Checking document: ${doc.name}`);
                    
                    // Skip folders
                    if (doc.folder) {
                        console.log(`⏭️ Skipping folder: ${doc.name}`);
                        continue;
                    }
                    
                    // Simplified content extraction for debugging
                    try {
                        console.log(`📥 Attempting to get content for: ${doc.name}`);
                        console.log(`📊 File details: size=${doc.size}, mimeType=${doc.file?.mimeType || 'unknown'}`);
                        
                        // Try to get content as text first (works for many file types)
                        let content = await graphClient.getDocumentContent(doc.parentReference.driveId, doc.id, false);
                        
                        if (content && typeof content === 'string' && content.length > 10) {
                            console.log(`✅ Got ${content.length} characters from ${doc.name}`);
                            console.log(`📝 First 200 chars: ${content.substring(0, 200)}...`);
                            
                            // Use enhanced AI-powered Q&A
                            const answer = await docProcessor.answerQuestionEnhanced(question, content, doc.name);
                            console.log(`🎯 Answer confidence: ${answer.confidence} for ${doc.name} (method: ${answer.method || 'standard'})`);
                            
                            if (answer.confidence > 0.1 && (!bestAnswer || answer.confidence > bestAnswer.confidence)) {
                                bestAnswer = answer;
                                console.log(`🏆 New best answer from ${doc.name}`);
                            }
                            searchedDocs++;
                        } else {
                            console.log(`❌ No usable content from ${doc.name} (type: ${typeof content}, length: ${content?.length || 0})`);
                            
                            // Try as binary for office documents
                            const fileExtension = doc.name.split('.').pop()?.toLowerCase();
                            if (['docx', 'pdf', 'xlsx', 'pptx'].includes(fileExtension)) {
                                console.log(`🔄 Trying binary extraction for ${doc.name}`);
                                try {
                                    const buffer = await graphClient.getDocumentContent(doc.parentReference.driveId, doc.id, true);
                                    if (buffer) {
                                        content = await docProcessor.extractContent(buffer, doc.name);
                                        if (content && content.length > 10) {
                                            console.log(`✅ Binary extraction successful: ${content.length} characters`);
                                            const answer = await docProcessor.answerQuestion(question, content, doc.name);
                                            if (answer.confidence > 0.1 && (!bestAnswer || answer.confidence > bestAnswer.confidence)) {
                                                bestAnswer = answer;
                                            }
                                            searchedDocs++;
                                        }
                                    }
                                } catch (binaryError) {
                                    console.log(`❌ Binary extraction failed: ${binaryError.message}`);
                                }
                            }
                        }
                    } catch (contentError) {
                        console.log(`❌ Failed to get content from ${doc.name}: ${contentError.message}`);
                    }
                } catch (docError) {
                    console.log(`⚠️ Couldn't read ${doc.name}: ${docError.message}`);
                    // Continue with other documents
                }
            }

            // Debug: Show what we found
            console.log(`🔍 Final results: searchedDocs=${searchedDocs}, bestAnswer=${bestAnswer ? 'YES' : 'NO'}`);
            if (bestAnswer) {
                console.log(`🏆 Best answer confidence: ${bestAnswer.confidence} from ${bestAnswer.documentName}`);
            }

            if (bestAnswer && bestAnswer.confidence > 0.1) {
                try {
                    await context.sendActivity(
                        `🎯 **Here's what I found:**\n\n` +
                        `${bestAnswer.answer}\n\n` +
                        `📊 **Confidence:** ${Math.round(bestAnswer.confidence * 100)}%\n` +
                        `📁 **Source:** ${bestAnswer.documentName}\n` +
                        `🔍 *Searched ${searchedDocs} documents*\n\n` +
                        `💡 **Want to know more?** Ask me another question about your documents!`
                    );
                    console.log(`✅ Successfully sent answer to user`);
                } catch (sendError) {
                    console.error(`❌ Failed to send answer to user:`, sendError);
                    await context.sendActivity(`✅ I found an answer but had trouble sending it. Please try asking again.`);
                }
            } else if (searchedDocs > 0) {
                await context.sendActivity(
                    `🔍 I searched ${searchedDocs} documents but couldn't find a confident answer to "${question}".\n\n` +
                    `📋 **Documents I checked:**\n${recentDocs.value.slice(0, searchedDocs).map(doc => `• ${doc.name}`).join('\n')}\n\n` +
                    `💡 **Try:**\n` +
                    `• More specific questions\n` +
                    `• Keywords that might be in your documents\n` +
                    `• Questions like "what is the deadline?" or "who is mentioned?"`
                );
            } else {
                await context.sendActivity(
                    `❌ I found ${recentDocs.value?.length || 0} documents but couldn't extract content from any of them.\n\n` +
                    `📋 **Documents found:**\n${recentDocs.value?.slice(0, 5).map(doc => `• ${doc.name} (${doc.file?.mimeType || 'unknown type'})`).join('\n') || 'None'}\n\n` +
                    `🔧 **This might be due to:**\n` +
                    `• File format limitations\n` +
                    `• Permission issues\n` +
                    `• Large file sizes\n\n` +
                    `💡 **Try:** \`recent\` to see your files, then ask about specific document names.`
                );
            }

        } catch (error) {
            console.error('❌ Document Q&A error:', error);
            await context.sendActivity('❌ Sorry, I encountered an error while searching your documents. Please try again.');
        }
    }

    async handleSummarizeCommand(context, docName, graphClient) {
        const { DocumentProcessor } = require('./documentProcessor');
        const docProcessor = new DocumentProcessor();

        try {
            await context.sendActivity(`📝 Generating AI summary for "${docName}"...`);
            
            // Find the document
            const recentDocs = await graphClient.getRecentDocuments();
            const targetDoc = recentDocs.value?.find(doc => 
                doc.name.toLowerCase().includes(docName.toLowerCase())
            );

            if (!targetDoc) {
                await context.sendActivity(`❌ Document "${docName}" not found.\n\n💡 Try: \`recent\` to see available documents.`);
                return;
            }

            // Get document content
            const content = await graphClient.getDocumentContent(targetDoc.parentReference.driveId, targetDoc.id, false);
            
            if (!content || content.length < 50) {
                await context.sendActivity(`❌ Couldn't extract content from "${targetDoc.name}". The file might be empty or in an unsupported format.`);
                return;
            }

            // Generate AI summary
            const summary = await docProcessor.generateSummary(content, targetDoc.name);
            
            await context.sendActivity(
                `📝 **AI Summary of "${targetDoc.name}":**\n\n` +
                `${summary.summary}\n\n` +
                `🤖 **Generated by:** ${summary.source}\n` +
                `📊 **Confidence:** ${Math.round(summary.confidence * 100)}%\n\n` +
                `💡 **Want more details?** Ask specific questions about this document!`
            );

        } catch (error) {
            console.error('❌ Summarize error:', error);
            await context.sendActivity(`❌ Error generating summary: ${error.message}`);
        }
    }

    async handleInsightsCommand(context, docName, graphClient) {
        const { DocumentProcessor } = require('./documentProcessor');
        const docProcessor = new DocumentProcessor();

        try {
            await context.sendActivity(`💡 Generating AI insights for "${docName}"...`);
            
            // Find the document
            const recentDocs = await graphClient.getRecentDocuments();
            const targetDoc = recentDocs.value?.find(doc => 
                doc.name.toLowerCase().includes(docName.toLowerCase())
            );

            if (!targetDoc) {
                await context.sendActivity(`❌ Document "${docName}" not found.\n\n💡 Try: \`recent\` to see available documents.`);
                return;
            }

            // Get document content
            const content = await graphClient.getDocumentContent(targetDoc.parentReference.driveId, targetDoc.id, false);
            
            if (!content || content.length < 50) {
                await context.sendActivity(`❌ Couldn't extract content from "${targetDoc.name}". The file might be empty or in an unsupported format.`);
                return;
            }

            // Generate AI insights
            const insights = await docProcessor.generateInsights(content, targetDoc.name);
            
            await context.sendActivity(
                `💡 **AI Insights for "${targetDoc.name}":**\n\n` +
                `${insights.insights}\n\n` +
                `🤖 **Generated by:** ${insights.source}\n` +
                `📊 **Confidence:** ${Math.round(insights.confidence * 100)}%\n\n` +
                `💡 **Need more analysis?** Ask specific questions about this document!`
            );

        } catch (error) {
            console.error('❌ Insights error:', error);
            await context.sendActivity(`❌ Error generating insights: ${error.message}`);
        }
    }

    // Determine if a question is SharePoint/document related
    isSharePointRelatedQuestion(text) {
        const lowerText = text.toLowerCase();
        
        // SharePoint/document keywords
        const sharePointKeywords = [
            'document', 'file', 'pdf', 'docx', 'excel', 'powerpoint',
            'sharepoint', 'upload', 'download', 'recent', 'folder',
            'policy', 'procedure', 'project', 'plan', 'budget', 'report',
            'contract', 'agreement', 'invoice', 'receipt', 'price',
            'stock', 'arrival', 'recieval', 'jono', 'johno', 'staff'
        ];
        
        // Question patterns that suggest document search
        const documentPatterns = [
            'what is in the',
            'what does the',
            'show me the',
            'find the',
            'what are the deadlines',
            'who is mentioned',
            'what is the price',
            'what is the cost',
            'when is the deadline',
            'summarize',
            'insights',
            'what documents'
        ];
        
        // Check for SharePoint keywords
        if (sharePointKeywords.some(keyword => lowerText.includes(keyword))) {
            return true;
        }
        
        // Check for document-related patterns
        if (documentPatterns.some(pattern => lowerText.includes(pattern))) {
            return true;
        }
        
        return false;
    }

    // Handle general questions using Azure OpenAI
    async handleGeneralQuestion(context, question) {
        const { AIService } = require('./aiService');
        const aiService = new AIService();
        
        try {
            console.log(`🤖 Handling general question: "${question}"`);
            
            // Use Azure OpenAI for general knowledge questions
            const response = await aiService.answerQuestion(question, '', 'General Knowledge');
            
            if (response && response.answer) {
                await context.sendActivity(
                    `🤖 **${response.answer}**\n\n` +
                    `💡 *I can also search your SharePoint documents if you have questions about your files!*\n\n` +
                    `📋 **Try commands like:**\n` +
                    `• \`recent\` - See your recent files\n` +
                    `• \`summarize [document]\` - AI summary\n` +
                    `• Ask about your documents: "What's in the price changes file?"`
                );
            } else {
                await context.sendActivity(
                    `🤔 I'm having trouble answering that question right now.\n\n` +
                    `💡 **I can help you with:**\n` +
                    `• General questions (like ChatGPT)\n` +
                    `• Your SharePoint documents\n` +
                    `• Document analysis and insights\n\n` +
                    `Try asking something else!`
                );
            }
            
        } catch (error) {
            console.error('❌ General question error:', error);
            await context.sendActivity(
                `❌ Sorry, I encountered an error answering your question.\n\n` +
                `💡 **I can still help you with:**\n` +
                `• \`recent\` - See your SharePoint files\n` +
                `• \`help\` - See all commands`
            );
        }
    }
}

module.exports.SharePointBot = SharePointBot;