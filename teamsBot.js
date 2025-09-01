const { TeamsActivityHandler, CardFactory } = require('botbuilder');

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMessage(async (context, next) => {
            const text = context.activity.text.trim().toLowerCase();

            if (text === 'profile') {
                // Try to get a Graph token via the configured OAuth connection
                const tokenResponse = await context.adapter.getUserToken(context, process.env.ConnectionName);

                if (!tokenResponse) {
                    await context.sendActivity('Please sign in so I can access your profile.');
                } else {
                    await context.sendActivity(`I got a token! (truncated): ${tokenResponse.token.substring(0, 25)}...`);
                }
            } else {
                await context.sendActivity(`You said: "${context.activity.text}"`);
            }

            await next();
        });
    }
}

module.exports.TeamsBot = TeamsBot;