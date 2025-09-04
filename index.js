// index.js
const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { SharePointBot } = require('./teamsBot');
const { MainDialog } = require('./mainDialog');

// Environment variables - try multiple possible names
const MICROSOFT_APP_ID = process.env.MicrosoftAppId || process.env.MicrosoftAppid || "";
const MICROSOFT_APP_PASSWORD = process.env.MicrosoftAppPassword || process.env.MicrosoftApppassword || "";
const CONNECTION_NAME = process.env.ConnectionName || "GraphConnection";

console.log('=== Bot Startup ===');
console.log(`Node.js version: ${process.version}`);
console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
console.log(`MicrosoftAppId present: ${MICROSOFT_APP_ID ? 'YES' : 'NO'}`);
console.log(`MicrosoftAppId value: ${MICROSOFT_APP_ID ? MICROSOFT_APP_ID.substring(0, 8) + '...' : 'EMPTY'}`);
console.log(`MicrosoftAppPassword present: ${MICROSOFT_APP_PASSWORD ? 'YES' : 'NO'}`);
console.log(`ConnectionName: ${CONNECTION_NAME}`);
console.log(`All Microsoft env vars:`, Object.keys(process.env).filter(k => k.toLowerCase().includes('microsoft')));
console.log('==================');

// Create adapter with explicit validation
if (!MICROSOFT_APP_ID || !MICROSOFT_APP_PASSWORD) {
    console.error('❌ CRITICAL: Missing bot credentials!');
    console.error('- MicrosoftAppId:', MICROSOFT_APP_ID ? 'SET' : 'MISSING');
    console.error('- MicrosoftAppPassword:', MICROSOFT_APP_PASSWORD ? 'SET' : 'MISSING');
    process.exit(1);
}

// Create adapter with explicit authentication settings
const adapter = new BotFrameworkAdapter({
    appId: MICROSOFT_APP_ID,
    appPassword: MICROSOFT_APP_PASSWORD,
    channelAuthTenant: process.env.MicrosoftAppTenantId || 'common',
    oAuthEndpoint: 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token'
});

// Log adapter configuration
console.log('✅ Adapter created with valid credentials:');
console.log(`- appId: ${MICROSOFT_APP_ID.substring(0, 8)}...`);
console.log(`- appPassword: ${MICROSOFT_APP_PASSWORD.substring(0, 8)}...`);
console.log('===================');

// State management
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// Dialog & Bot
console.log('🔍 Creating MainDialog...');
const mainDialog = new MainDialog();
console.log('✅ MainDialog created successfully');

console.log('🔍 Creating SharePointBot...');
const bot = new SharePointBot(conversationState, mainDialog);
console.log('✅ SharePointBot created successfully');

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





// Entry point: async handler for Restify
server.post('/api/messages', async (req, res) => {
    console.log('🚀 Processing /api/messages request...');
    await adapter.processActivity(req, res, async (context) => {
        console.log("===== Incoming Activity =====");
        console.log(JSON.stringify(context.activity, null, 2));
        
        // Run the bot
        await bot.run(context);
        
        // Save conversation state
        await conversationState.saveChanges(context, false);
    });
});
  
