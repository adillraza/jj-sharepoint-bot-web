// index.js
const restify = require('restify');
const path = require('path');
// remove this line if not using .env
// require('dotenv').config({ path: path.join(__dirname, '.env') });


const {
  BotFrameworkAdapter, // If you're on botbuilder <= 4.20
} = require('botbuilder');

const { TeamsBot } = require('./teamsBot');

// --- App credentials from env vars ---
const APP_ID = process.env.MicrosoftAppId || '';
const APP_PASSWORD = process.env.MicrosoftAppPassword || '';
const TENANT_ID = process.env.MicrosoftAppTenantId || ''; // optional but fine to keep

// Create adapter
const adapter = new BotFrameworkAdapter({
  appId: APP_ID,
  appPassword: APP_PASSWORD,
});

// Global onTurnError handler
adapter.onTurnError = async (context, error) => {
  console.error('### onTurnError:', error);
  try {
    await context.sendActivity('The bot encountered an error.');
  } catch (_) {}
};

// Create bot
const bot = new TeamsBot();

// Create Restify server
const server = restify.createServer();
const PORT = process.env.PORT || 3978;
server.use(restify.plugins.bodyParser());
server.listen(PORT, () => {
  console.log(`Bot Started, restify listening to http://localhost:${PORT}`);
});

// IMPORTANT: Do NOT call res.send() / res.end() / next() here.
// Bot Framework will complete the response.
server.post('/api/messages', async (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Optional: simple GET for root to avoid “ResourceNotFound” confusion
server.get('/', (_, res, next) => {
  res.send(200, { ok: true, message: 'Bot is running' });
  return next();
});
