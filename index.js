// index.js
const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');

// Environment variables
const MICROSOFT_APP_ID = process.env.MicrosoftAppId || "";
const MICROSOFT_APP_PASSWORD = process.env.MicrosoftAppPassword || "";

console.log('=== Bot Startup ===');
console.log(`MicrosoftAppId present: ${MICROSOFT_APP_ID ? 'YES' : 'NO'}`);
console.log(`MicrosoftAppPassword present: ${MICROSOFT_APP_PASSWORD ? 'YES' : 'NO'}`);
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
    await context.sendActivity('Oops! Something went wrong.');
};

// Create server
const server = restify.createServer();
const PORT = process.env.PORT || 3978;
server.listen(PORT, () => {
    console.log(`\nBot Started, listening on http://localhost:${PORT}`);
});

// Bot logic
async function handleMessage(context) {
    console.log("===== Incoming Activity =====");
    console.log(JSON.stringify(context.activity, null, 2));

    if (context.activity.type === 'message') {
        const userMessage = context.activity.text;
        console.log(`User said: ${userMessage}`);

        await context.sendActivity(`You said: ${userMessage}`);
    } else {
        await context.sendActivity(`[${context.activity.type} event detected]`);
    }
}

// Entry point: async handler for Restify
server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
        await handleMessage(context);
    });
});
