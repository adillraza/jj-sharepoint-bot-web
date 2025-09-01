const restify = require('restify');
const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  ActivityTypes,
  StatusCodes,
  TurnContext,
  CardFactory
} = require('botbuilder');
const { graphGet } = require('./graph');

require('dotenv').config();

// ---- Config
const APP_ID = process.env.MicrosoftAppId || '';
const APP_PASSWORD = process.env.MicrosoftAppPassword || '';
const APP_TENANT_ID = process.env.MicrosoftAppTenantId || '';
const APP_TYPE = process.env.MicrosoftAppType || 'SingleTenant';
const CONNECTION_NAME = process.env.ConnectionName || 'GraphConnection';

console.log('[startup] MicrosoftAppId present:', !!APP_ID);
console.log('[startup] MicrosoftAppTenantId present:', !!APP_TENANT_ID);
console.log('[startup] ConnectionName:', CONNECTION_NAME);

// ---- Auth + Adapter
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: APP_ID,
  MicrosoftAppPassword: APP_PASSWORD,
  MicrosoftAppTenantId: APP_TENANT_ID,
  MicrosoftAppType: APP_TYPE
});
const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('Sorry—something went wrong.');
};

// ---- Bot logic
async function handleTurn(context) {
  if (context.activity.type === ActivityTypes.Message) {
    const text = (context.activity.text || '').trim().toLowerCase();

    if (text === 'signin' || text === 'login' || text === 'connect') {
      // ask user to sign-in
      const signInLink = await adapter.getSignInLink(context, CONNECTION_NAME);
      await context.sendActivity({
        attachments: [
          CardFactory.signinCard(
            'Sign in to Microsoft 365',
            signInLink,
            'Continue'
          )
        ]
      });
      return;
    }

    if (text === 'recent' || text === 'files') {
      // try to get token
      const token = await adapter.getUserToken(context, CONNECTION_NAME);
      if (!token) {
        const signInLink = await adapter.getSignInLink(context, CONNECTION_NAME);
        await context.sendActivity({
          attachments: [
            CardFactory.signinCard(
              'Please sign in to view your recent SharePoint/OneDrive files',
              signInLink,
              'Sign in'
            )
          ]
        });
        return;
      }

      try {
        // Call Graph for recent files across OneDrive/SharePoint
        // You can swap for a SharePoint site collection later (Part 4C)
        const data = await graphGet('/v1.0/me/drive/recent', token.token);
        const items = (data?.value || []).slice(0, 5);
        if (items.length === 0) {
          await context.sendActivity('No recent files found.');
        } else {
          const lines = items.map((it, i) => `${i + 1}. **${it.name}**\n   📁 [Open file](${it.webUrl})`);
          await context.sendActivity(`**Your recent files:**\n\n${lines.join('\n\n')}`);
        }
      } catch (e) {
        console.error('Graph error:', e);
        await context.sendActivity('Could not fetch files from Microsoft Graph.');
      }
      return;
    }

    // default echo
    await context.sendActivity(`You said: ${context.activity.text}`);
  }

  else if (context.activity.type === ActivityTypes.Event &&
           context.activity.name === 'tokens/response') {
    // Token response from OAuth flow (Teams/WebChat)
    await context.sendActivity('You are now signed in. Type "recent" to see your files.');
  }

  else if (context.activity.type === ActivityTypes.ConversationUpdate) {
    if (context.activity.membersAdded?.some(m => m.id !== context.activity.recipient?.id)) {
      await context.sendActivity('Hi! I am alive ✅ Say "recent" to fetch your files, or "signin" to connect.');
    }
  }
  else {
    await context.sendActivity(`(${context.activity.type}) event received.`);
  }
}

// ---- Restify server
const server = restify.createServer();
server.use(restify.plugins.bodyParser({ mapParams: false }));

server.get('/', (_req, res) => res.send(200, { status: 'ok', path: '/' }));
server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (ctx) => handleTurn(ctx));
});

const port = process.env.PORT || 3978;
server.listen(port, () => {
  console.log(`Bot Started, restify listening on http://localhost:${port}`);
});
