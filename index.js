const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');
const { TeamsBot } = require('./teamsBot');

// Create server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Create adapter
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Catch-all for errors
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError]: ${ error }`);
    await context.sendActivity('Oops. Something went wrong!');
};

// Create bot
const bot = new TeamsBot();

// Listen for incoming requests
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        await bot.run(context);
    });
});