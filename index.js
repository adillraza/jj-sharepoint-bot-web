// index.js
const restify = require('restify');
const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  ActivityTypes,
  TurnContext,
} = require('botbuilder');

require('dotenv').config();

// --- Env checks (only logs; uses App Service env vars) ---
const APP_ID = process.env.MicrosoftAppId || '';
const APP_PASSWORD = process.env.MicrosoftAppPassword || '';
const APP_TENANT_ID = process.env.MicrosoftAppTenantId || '';
const APP_TYPE = process.env.MicrosoftAppType || 'SingleTenant'; // ok for single-tenant

console.log('[startup] MicrosoftAppId present:', !!APP_ID);
console.log('[startup] MicrosoftAppPassword present:', !!APP_PASSWORD);
console.log('[startup] MicrosoftAppTenantId present:', !!APP_TENANT_ID);

// --- Auth + Adapter ---
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: APP_ID,
  MicrosoftAppPassword: APP_PASSWORD,
  MicrosoftAppTenantId: APP_TENANT_ID,
  MicrosoftAppType: APP_TYPE, // SingleTenant or MultiTenant
});

const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// log unhandled errors so we see them in Log Stream
adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  try {
    await context.sendActivity('Sorry—something went wrong.');
  } catch { /* ignore */ }
};

// --- Very simple bot logic for now ---
async function handleTurn(context) {
  if (context.activity.type === ActivityTypes.Message) {
    // Remove @mentions for Teams compatibility
    const text = TurnContext.removeRecipientMention(context.activity)?.text?.trim() || 
                 context.activity.text?.trim() || '';
    
    if (text.toLowerCase().includes('hello')) {
      await context.sendActivity('Hello! 👋 I am alive and responding.');
    } else {
      await context.sendActivity(`You said: ${text || '(empty message)'}`);
    }
  } else if (context.activity.type === ActivityTypes.ConversationUpdate) {
    // Greet in Web Chat/Emulator/Teams
    if (context.activity.membersAdded?.some(m => m.id !== context.activity.recipient?.id)) {
      await context.sendActivity('Hi! I am alive ✅ Say "hello" to test me.');
    }
  } else {
    await context.sendActivity(`(${context.activity.type}) event received.`);
  }
}

// --- Restify server ---
const server = restify.createServer();

// ✅ This is the critical piece: parse JSON bodies
server.use(restify.plugins.bodyParser({ mapParams: false }));

// Health check
server.get('/', (_req, res, _next) => {
  res.send(200, { status: 'ok', path: '/' });
});

// Bot messages endpoint
server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context) => handleTurn(context));
});

// Start server
const port = process.env.PORT || 3978;
server.listen(port, () => {
  console.log(`Bot Started, restify listening on http://localhost:${port}`);
});
