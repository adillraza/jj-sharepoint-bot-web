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

// Create adapter - allow empty credentials for development
const adapter = new BotFrameworkAdapter({
    appId: MICROSOFT_APP_ID || undefined,
    appPassword: MICROSOFT_APP_PASSWORD || undefined
});

// Log adapter configuration
console.log('Adapter created with:');
console.log(`- appId: ${MICROSOFT_APP_ID ? 'SET' : 'EMPTY (development mode)'}`);
console.log(`- appPassword: ${MICROSOFT_APP_PASSWORD ? 'SET' : 'EMPTY (development mode)'}`);
console.log('===================');

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
// Add a simple health endpoint to verify deployment
server.get('/', (req, res, next) => {
    const status = {
        status: 'Bot is running',
        timestamp: new Date().toISOString(),
        nodeVersion: process.version,
        appId: MICROSOFT_APP_ID ? 'SET' : 'MISSING',
        appPassword: MICROSOFT_APP_PASSWORD ? 'SET' : 'MISSING'
    };
    res.send(200, status);
    return next();
});

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

            console.log(`About to send response...`);
            
            // Try the simplest possible response
            try {
                await context.sendActivity('Hello! I received your message.');
                console.log('✅ Simple response sent successfully');
            } catch (sendError) {
                console.error('❌ Failed to send simple response:', sendError.message);
                console.error('Full error:', sendError);
            }
        } else {
            const response = `[${context.activity.type} event detected]`;
            console.log(`Bot responding: ${response}`);
            await context.sendActivity(response);
            console.log('Response sent successfully');
        }
    } catch (error) {
        console.error('❌ Error in handleMessage:', error.message);
        console.error('Error type:', error.constructor.name);
        console.error('Error details:', error);
        // Don't re-throw any errors - just log them
    }
}

// Entry point: async handler for Restify (VERSION 2)
server.post('/api/messages', async (req, res) => {
    console.log('🚀 [VERSION 2] Processing /api/messages request...');
    await adapter.processActivity(req, res, async (context) => {
      console.log("===== Incoming Activity =====");
      console.log(JSON.stringify(context.activity, null, 2));
  
      if (context.activity.type === 'message') {
        const text = context.activity.text || '';
        console.log(`Sending echo to conversation: ${context.activity.conversation?.id}`);
        try {
          await context.sendActivity(`You said: ${text}`);
          console.log('✅ Response sent successfully!');
        } catch (err) {
          console.error('❌ [sendActivity error]', err.message);
          console.error('Error status:', err.statusCode);
          console.error('Error code:', err.code);
          console.error('Full error details:', JSON.stringify(err, null, 2));
        }
      } else {
        // Don't reply to typing/other events to keep the connector happy
        console.log(`Non-message activity (${context.activity.type}) received; no reply sent.`);
      }
    });
});
  
