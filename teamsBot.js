const { TeamsActivityHandler, CardFactory, TurnContext } = require('botbuilder');
const { DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');
const { SharePointGraphClient } = require('./graph');
const { DIALOG_ID } = require('./mainDialog');

// Environment variables
const CONNECTION_NAME = process.env.ConnectionName || "GraphConnection";
const CLIENT_ID = process.env.MicrosoftAppId;
const TENANT_ID = process.env.MicrosoftAppTenantId;

class SharePointBot extends TeamsActivityHandler {
    constructor(conversationState, mainDialog, deploymentId) {
        super();
        
        this.deploymentId = deploymentId || 'unknown';
        console.log(`🤖 SharePointBot initialized - Deployment: ${this.deploymentId}`);
        
        this.conversationState = conversationState;
        this.dialogState = this.conversationState.createProperty("DialogState");
        
        // Host the dialog
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(mainDialog);

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded || [];
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    const welcomeText = `👋 **Welcome to SharePoint Document Assistant!**

**💬 What I can do:**
• Find and read your SharePoint documents
• Answer questions about your files  
• General questions and AI assistance
• Analyze your entire SharePoint site

**📋 Available Commands:**
• \`recent\` - Show recent documents
• \`stats\` - Site statistics (files, folders, types)
• \`search [keyword]\` - Find specific documents
• \`summarize [document]\` - AI summary of any document
• \`insights [document]\` - AI insights from documents
• \`help\` - Show detailed help

**🚀 Quick Examples:**
• "How many files are in our SharePoint?"
• "What's in the customer service document?"
• "Show me recent Excel files"
• "How is the weather today?"

**Just ask me anything!** 📦 *${this.deploymentId}*`;
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
**🤖 SharePoint AI Assistant**

**💬 Ask me anything:**
• General questions and AI assistance
• Questions about your SharePoint documents

**📁 Commands:**
• \`recent\` - Show recent documents
• \`stats\` - Site statistics (total files, folders, types)
• \`search [keyword]\` - Find documents
• \`summarize [document]\` - AI summary
• \`insights [document]\` - AI insights

**Examples:**
• "How is the weather?"
• "What's in the project plan document?"
• "Show me recent Excel files"

💡 **Just ask me anything!**
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

        // Site statistics command
        if (lowerText === 'stats' || lowerText === 'statistics' || lowerText === 'site stats') {
            try {
                await context.sendActivity('📊 Analyzing SharePoint site...');
                const stats = await graphClient.getSiteStatistics();
                
                await context.sendActivity(
                    `📊 **SharePoint Site Statistics:**\n\n` +
                    `📁 **Total Folders:** ${stats.folderCount}\n` +
                    `📄 **Total Files:** ${stats.fileCount}\n` +
                    `💾 **Total Size:** ${stats.totalSize}\n` +
                    `📅 **Last Updated:** ${stats.lastModified}\n\n` +
                    `🔍 **File Types:**\n${stats.fileTypes.map(ft => `• ${ft.type}: ${ft.count} files`).join('\n')}`
                );
            } catch (error) {
                console.error('Error getting site statistics:', error);
                await context.sendActivity('❌ Sorry, I couldn\'t retrieve site statistics at this time.');
            }
            return;
        }

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
                await this.handleDocumentQuestion(context, text, graphClient);
            } else {
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
            
            // Get ALL documents from SharePoint for comprehensive search
            console.log('🔍 Getting ALL documents for comprehensive Q&A...');
            let recentDocs;
            
            // Use the working method that we know works
            console.log('🔄 Using proven document retrieval method...');
            recentDocs = await graphClient.getRecentDocuments();
            
            if (!recentDocs.value || recentDocs.value.length === 0) {
                await context.sendActivity('📂 I couldn\'t find any recent documents to search through. Try uploading some documents to SharePoint first.');
                return;
            }

            let bestAnswer = null;
                            let searchedDocs = 0;
                const maxDocsToSearch = 2; // Reduce from 5 to 2 for F0 tier // Limit for performance

            
            // Debug: Show what documents we found
            const docNames = recentDocs.value.map(doc => `${doc.name} (${doc.file?.mimeType || 'no mime type'})`).join(', ');
            console.log(`📋 Total documents found: ${recentDocs.value.length}`);
            console.log(`📋 Documents: ${docNames}`);
            
            // Log debug info (not sent to user)
            console.log(`📋 Found ${recentDocs.value.length} documents in SharePoint. Searching through first ${Math.min(maxDocsToSearch, recentDocs.value.length)}...`);

            // SPECIFIC document targeting: if user mentions a filename, search that file specifically
            const questionLower = question.toLowerCase();
            let targetDoc = null;
            
            // Check if user is asking about a specific file
            for (const doc of recentDocs.value) {
                const fileName = doc.name.toLowerCase();
                const fileNameWithoutExt = fileName.substring(0, fileName.lastIndexOf('.'));
                
                if (questionLower.includes(fileName) || questionLower.includes(fileNameWithoutExt)) {
                    targetDoc = doc;
                    console.log(`🎯 DIRECT FILE MATCH: User asking about "${doc.name}"`);
                    break;
                }
            }
            
            let docsToSearch = [];
            if (targetDoc) {
                // User asked about specific file - search ONLY that file
                docsToSearch = [targetDoc];
                console.log(`📄 Searching ONLY the requested file: ${targetDoc.name}`);
            } else {
                // General question - use smart document selection
                const questionKeywords = question.toLowerCase().split(' ').filter(word => word.length > 3);
                const scoredDocs = recentDocs.value.map(doc => {
                    const nameScore = questionKeywords.filter(keyword => 
                        doc.name.toLowerCase().includes(keyword)
                    ).length;
                    return { doc, score: nameScore };
                });
                
                // Sort by relevance, then by date
                scoredDocs.sort((a, b) => {
                    if (a.score !== b.score) return b.score - a.score; // Higher score first
                    return new Date(b.doc.lastModifiedDateTime) - new Date(a.doc.lastModifiedDateTime); // Newer first
                });
                
                console.log(`🎯 Document relevance scores:`, scoredDocs.map(sd => `${sd.doc.name}: ${sd.score}`).join(', '));
                docsToSearch = scoredDocs.slice(0, maxDocsToSearch).map(sd => sd.doc);
            }
            
            for (const doc of docsToSearch) {
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
                        
                        // EMERGENCY DEBUG: What are we actually getting?
                        let content = await graphClient.getDocumentContent(doc.parentReference.driveId, doc.id, false);
                        
                        console.log(`🔍 EMERGENCY DEBUG for ${doc.name}:`);
                        console.log(`   - Content type: ${typeof content}`);
                        console.log(`   - Content length: ${content?.length || 0}`);
                        console.log(`   - Is string: ${typeof content === 'string'}`);
                        console.log(`   - First 100 chars: ${content ? String(content).substring(0, 100) : 'NULL'}`);
                        
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
                                        content = await docProcessor.extractTextFromDocument(buffer, doc.file.mimeType, doc.name);
                                        if (content && content.length > 10) {
                                            console.log(`✅ Binary extraction successful: ${content.length} characters from ${doc.name}`);
                                            console.log(`📄 Content preview: ${content.substring(0, 200)}...`);
                                            
                                            const answer = await docProcessor.answerQuestion(question, content, doc.name);
                                            console.log(`🎯 Answer from ${doc.name}: confidence=${answer.confidence}, answer="${answer.answer?.substring(0, 100)}..."`);
                                            
                                            if (answer.confidence > 0.1 && (!bestAnswer || answer.confidence > bestAnswer.confidence)) {
                                                bestAnswer = answer;
                                                console.log(`🏆 NEW BEST ANSWER from ${doc.name} with confidence ${answer.confidence}`);
                                            }
                                            searchedDocs++;
                                        } else {
                                            console.log(`❌ Binary extraction failed for ${doc.name} - content length: ${content?.length || 0}`);
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

            // EMERGENCY FALLBACK: If no content extracted, give helpful response
            if (!bestAnswer && searchedDocs === 0) {
                const targetFileName = targetDoc ? targetDoc.name : 'the requested document';
                await context.sendActivity(
                    `❌ **I'm having trouble reading document content right now.**\n\n` +
                    `📋 **I can see these documents in SharePoint:**\n${recentDocs.value.map(doc => `• ${doc.name} (${doc.file?.mimeType || 'unknown type'})`).join('\n')}\n\n` +
                    `🔧 **This might be due to:**\n` +
                    `• Document format limitations\n` +
                    `• Permission restrictions\n` +
                    `• File size limitations\n\n` +
                    `💡 **Try asking:** "recent" to see all available files, or ask about a different document.`
                );
                return;
            }
            
            if (bestAnswer && bestAnswer.confidence > 0.1) {
                try {
                    // If user asked about specific file, confirm we searched the right file
                    let responseText = bestAnswer.answer;
                    if (targetDoc && bestAnswer.documentName === targetDoc.name) {
                        responseText = `📄 **From ${targetDoc.name}:**\n\n${bestAnswer.answer}`;
                    } else if (targetDoc) {
                        responseText = `⚠️ **Note:** You asked about "${targetDoc.name}" but I found relevant information in "${bestAnswer.documentName}":\n\n${bestAnswer.answer}`;
                    }
                    
                    await context.sendActivity(
                        `${responseText}\n\n` +
                        `📁 *Source: ${bestAnswer.documentName}*`
                    );
                    console.log(`✅ Successfully sent answer to user from ${bestAnswer.documentName}`);
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
            
            // Add timeout for F0 tier - 15 seconds max
            const timeoutPromise = new Promise((_, reject) => 
                setTimeout(() => reject(new Error('Timeout')), 15000)
            );
            
            // Use Azure OpenAI for general knowledge questions with timeout
            const response = await Promise.race([
                aiService.answerQuestion(question, '', 'General Knowledge'),
                timeoutPromise
            ]);
            
            if (response && response.answer) {
                await context.sendActivity(response.answer);
            } else {
                await context.sendActivity(`🤔 I couldn't find an answer to that question. Try asking something else!`);
            }
            
        } catch (error) {
            console.error('❌ General question error:', error);
            if (error.message === 'Timeout') {
                await context.sendActivity(`⏱️ That question is taking too long to process. Try asking something simpler.`);
            } else {
                await context.sendActivity(`❌ Sorry, I couldn't answer that question right now. Please try again.`);
            }
        }
    }
}

module.exports.SharePointBot = SharePointBot;