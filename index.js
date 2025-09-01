// index.js
const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');

// Environment variables - try multiple possible names
const MICROSOFT_APP_ID = process.env.MicrosoftAppId || process.env.MicrosoftAppid || "";
const MICROSOFT_APP_PASSWORD = process.env.MicrosoftAppPassword || process.env.MicrosoftApppassword || "";

console.log('=== Bot Startup ===');
console.log(`Node.js version: ${process.version}`);
console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
console.log(`MicrosoftAppId present: ${MICROSOFT_APP_ID ? 'YES' : 'NO'}`);
console.log(`MicrosoftAppId value: ${MICROSOFT_APP_ID ? MICROSOFT_APP_ID.substring(0, 8) + '...' : 'EMPTY'}`);
console.log(`MicrosoftAppPassword present: ${MICROSOFT_APP_PASSWORD ? 'YES' : 'NO'}`);
console.log(`All Microsoft env vars:`, Object.keys(process.env).filter(k => k.toLowerCase().includes('microsoft')));
console.log('==================');

// Create adapter
const adapter = new BotFrameworkAdapter({
    appId: MICROSOFT_APP_ID,
    appPassword: MICROSOFT_APP_PASSWORD
});

// State management
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// Catch-all for errors
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] Unhandled error: ${error}`);
    console.error('Error details:', error.message);
    
    // Only try to send error message if it's a real error, not a network issue
    if (error.message && !error.message.includes('RestError')) {
        try {
            await context.sendActivity('Oops! Something went wrong.');
        } catch (sendError) {
            console.error('Failed to send error message:', sendError.message);
        }
    }
};

// Create server
const server = restify.createServer();
const PORT = process.env.PORT || 3978;
server.listen(PORT, () => {
    console.log(`\nBot Started, listening on http://localhost:${PORT}`);
});

// Bot logic
async function handleMessage(context) {
    try {
        console.log("===== Incoming Activity =====");
        console.log(JSON.stringify(context.activity, null, 2));

        if (context.activity.type === 'message') {
            const userMessage = context.activity.text || '';
            console.log(`User said: ${userMessage}`);

            const response = `You said: ${userMessage}`;
            console.log(`Bot responding: ${response}`);
            console.log(`Sending to conversation: ${context.activity.conversation?.id}`);
            console.log(`Channel: ${context.activity.channelId}`);
            
            const result = await context.sendActivity(response);
            console.log('✅ Response sent successfully, result:', result);
        } else {
            const response = `[${context.activity.type} event detected]`;
            console.log(`Bot responding: ${response}`);
            await context.sendActivity(response);
            console.log('Response sent successfully');
        }
    } catch (error) {
        console.error('Error in handleMessage:', error.message);
        // Don't re-throw network errors, just log them
        if (!error.message?.includes('RestError')) {
            throw error;
        }
    }
}

// Entry point: async handler for Restify
server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
        await handleMessage(context);
    });
});
