// index.js
// Minimal Bot Framework v4 app for Azure App Service (Node 20)
// Uses CloudAdapter + Restify and reads creds from env vars:
//   MicrosoftAppId, MicrosoftAppPassword, MicrosoftAppTenantId (optional)

const restify = require('restify');
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  TurnContext,
} = require('botbuilder');

// --- Optional: load .env when running locally --- //
try {
  if (process.env.NODE_ENV !== 'production') {
    require('dotenv').config();
  }
} catch (_) { /* noop */ }

// --- Verify required environment variables --- //
const APP_ID = process.env.MicrosoftAppId || '';
const APP_PASSWORD = process.env.MicrosoftAppPassword || '';
const APP_TENANT_ID = process.env.MicrosoftAppTenantId || ''; // needed for single-tenant apps

if (!APP_ID || !APP_PASSWORD) {
  console.warn(
    '[startup] Missing MicrosoftAppId and/or MicrosoftAppPassword — the bot will return 401 to the connector.'
  );
}

// --- Create adapter with proper credentials --- //
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: APP_ID,
  MicrosoftAppPassword: APP_PASSWORD,
  MicrosoftAppTenantId: APP_TENANT_ID, // safe to include even if empty
});

const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(
  null,
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Global error handler (shows up in App Service log stream)
adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError] Unhandled error:', error);

  // Send a trace activity for Emulator / inspection
  await context.sendTraceActivity(
    'onTurnError Trace',
    `${error}`,
    'https://www.botframework.com/schemas/error',
    'TurnError'
  );

  // Friendly message to the user (fails silently if auth failed)
  try {
    await context.sendActivity('Oops, something went wrong processing your message.');
  } catch (sendErr) {
    console.error('[onTurnError] Failed to notify user:', sendErr);
  }
};

// --- A very small bot --- //
class SharePointBot extends ActivityHandler {
  constructor() {
    super();

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded || [];
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            `Hi! I'm online. Say "hello" to check I'm responding.`
          );
        }
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      const text = (TurnContext.removeRecipientMention(context.activity)?.text ||
        context.activity.text ||
        '').trim();

      if (!text) {
        await context.sendActivity('I received your message.');
      } else if (/^hello\b/i.test(text)) {
        await context.sendActivity('Hello! 👋 I’m alive and connected.');
      } else {
        await context.sendActivity(`You said: "${text}"`);
      }

      await next();
    });
  }
}

const bot = new SharePointBot();

// --- Restify server --- //
const server = restify.createServer();
const port = process.env.PORT || 8080;

// Health pings (GET /) — returns 200 JSON
server.get('/', (req, res, next) => {
  res.send(200, { status: 'ok', bot: 'jj-sharepoint-bot-web' });
  return next();
});

// **IMPORTANT**: CloudAdapter expects an async handler, so we pass async (req, res) => ...
server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

// Start the server
server.listen(port, () => {
  console.log(`[startup] Server listening on http://localhost:${port}`);
  console.log(`[startup] MicrosoftAppId present: ${APP_ID ? 'yes' : 'no'}`);
  console.log(`[startup] MicrosoftAppTenantId present: ${APP_TENANT_ID ? 'yes' : 'no'}`);
});
