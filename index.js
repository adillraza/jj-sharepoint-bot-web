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
// index.js
// Minimal Bot Framework + Restify host for Azure Web App

// ---- Load env (locally). Azure App Service uses App Settings, so .env is optional.
try {
    require('dotenv').config();
  } catch (_) { /* ignore if not present */ }
  
  const restify = require('restify');
  const { BotFrameworkAdapter } = require('botbuilder');
  const pkg = require('./package.json');
  
  // Your bot logic (the class you already have in teamsBot.js)
  const { TeamsBot } = require('./teamsBot');
  
  // ---- Basic sanity logging of required envs
  const APP_ID = process.env.MicrosoftAppId || process.env.MICROSOFT_APP_ID || '';
  const APP_PASSWORD = process.env.MicrosoftAppPassword || process.env.MICROSOFT_APP_PASSWORD || '';
  const TENANT_ID = process.env.MicrosoftAppTenantId || process.env.MICROSOFT_APP_TENANT_ID || '';
  
  console.log(`[startup] package: ${pkg.name}@${pkg.version}`);
  console.log(`[startup] Node: ${process.version}`);
  console.log(`[startup] MicrosoftAppId present: ${APP_ID ? 'yes' : 'no'}`);
  console.log(`[startup] MicrosoftAppTenantId present: ${TENANT_ID ? 'yes' : 'no'}`);
  
  // ---- Create Restify server
  const server = restify.createServer({ name: pkg.name });
  server.use(restify.plugins.bodyParser({ mapParams: false }));
  server.use(restify.plugins.queryParser());
  
  const PORT = process.env.PORT || process.env.port || 3978;
  server.listen(PORT, () => {
    console.log(`${server.name} listening on http://localhost:${PORT}`);
  });
  
  // ---- Health + root routes (handy for Azure probes & quick checks)
  server.get('/', (_req, res, _next) => {
    res.send(200, { name: pkg.name, version: pkg.version, ok: true });
  });
  server.get('/api/health', (_req, res, _next) => res.send(200, { status: 'Healthy' }));
  
  // ---- Bot Framework adapter
  const adapter = new BotFrameworkAdapter({
    appId: APP_ID,
    appPassword: APP_PASSWORD,
  });
  
  // Catch-all error handler so Web Chat doesn’t “mysteriously” fail
  adapter.onTurnError = async (context, error) => {
    console.error('[onTurnError]', error);
    try {
      await context.sendActivity('Sorry, something went wrong.');
    } catch (e) {
      console.error('Failed to send error to user:', e);
    }
  };
  
  // ---- Create bot instance
  const bot = new TeamsBot();
  
  // ---- Messages endpoint (IMPORTANT: async (req,res) — no `next`)
  server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
      await bot.run(context);
    });
  });
  
  // (Optional) support preflight in some custom reverse-proxy setups
  server.opts('/api/messages', (req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'authorization, content-type');
    res.header('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.send(200);
    return next();
  });
  
  // ---- Graceful shutdown (helps on swap/restart)
  process.on('SIGTERM', () => {
    console.log('SIGTERM received. Shutting down server…');
    try {
      server.close(() => process.exit(0));
      setTimeout(() => process.exit(0), 2000);
    } catch {
      process.exit(0);
    }
  });
  