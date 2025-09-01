// index.js
// Minimal Bot Framework bot running on Restify with verbose logging

const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState, ActivityHandler } = require('botbuilder');
require('dotenv').config();

// ----- Basic startup logging (helps in Log Stream) -----
console.log('Starting bot...');
console.log(`Node.js: ${process.version}`);
console.log(`PORT (if provided by Azure): ${process.env.PORT || '(none)'}\n`);

const appId = process.env.MicrosoftAppId || '';
const appPassword = process.env.MicrosoftAppPassword || '';

console.log('Startup credential check:');
console.log(`  MicrosoftAppId present: ${appId ? 'yes' : 'NO'}`);
console.log(`  MicrosoftAppPassword present: ${appPassword ? 'yes' : 'NO'}`);
console.log('');

// ----- Restify server -----
const server = restify.createServer();
server.use(restify.plugins.bodyParser()); // important for /api/messages
server.listen(process.env.PORT || 3978, () => {
  console.log(`${server.name} listening on ${server.url}`);
});

// Health check (use correct Restify signature: (req, res, next))
server.get('/', (req, res, next) => {
  res.send(200, { status: 'Bot is running', time: new Date().toISOString() });
  return next();
});

// ----- Adapter -----
const adapter = new BotFrameworkAdapter({
  appId,
  appPassword
});

// Catch-all for errors
adapter.onTurnError = async (context, error) => {
  console.error('\n[onTurnError]', error);
  try { await context.sendActivity('Oops, something went wrong.'); } catch {}
};

// ----- State (in-memory for now) -----
const conversationState = new ConversationState(new MemoryStorage());

// ----- Bot implementation -----
class EchoBot extends ActivityHandler {
  constructor() {
    super();

    // Log *every* activity that arrives (useful in Azure Log Stream)
    this.use(async (context, next) => {
      const a = context.activity;
      console.log('--- Incoming Activity --------------------------------');
      console.log(`type:        ${a.type}`);
      console.log(`channelId:   ${a.channelId}`);
      console.log(`conversation:${a.conversation?.id}`);
      console.log(`from:        ${a.from?.id} (${a.from?.name || ''})`);
      console.log(`recipient:   ${a.recipient?.id}`);
      if (a.type === 'message') {
        console.log(`text:        ${JSON.stringify(a.text)}`);
      }
      console.log('------------------------------------------------------\n');
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded || [];
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity('Hi! ✅ I am alive. Say "hello" to test me.');
        }
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim();
      // simple echo
      await context.sendActivity(`You said: ${text || '(no text)'}`);
      await next();
    });
  }

  // Middleware hook to enable the logging "use" above
  use(mw) {
    const prev = this.run.bind(this);
    this.run = async (context) => {
      await mw(context, async () => prev(context));
    };
  }
}

const bot = new EchoBot();

// ----- Messaging endpoint (async handler; no "next") -----
server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
    await conversationState.saveChanges(context, false);
  });
});
