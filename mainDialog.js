// mainDialog.js
const {
  ComponentDialog,
  WaterfallDialog,
  OAuthPrompt
} = require("botbuilder-dialogs");

const DIALOG_ID = "MAIN_DIALOG";
const OAUTH_PROMPT_ID = "OAUTH_PROMPT";
const WATERFALL_ID = "WF";

const CONNECTION_NAME = process.env.ConnectionName || "GraphConnection";

class MainDialog extends ComponentDialog {
  constructor() {
    super(DIALOG_ID);

    console.log('🔍 MainDialog constructor - CONNECTION_NAME:', CONNECTION_NAME);
    console.log('🔍 MainDialog constructor - Creating OAuthPrompt...');

    try {
      this.addDialog(
        new OAuthPrompt(OAUTH_PROMPT_ID, {
          connectionName: CONNECTION_NAME,
          text: "Please sign in to Microsoft 365 to access your SharePoint documents.",
          title: "Sign in to Microsoft 365",
          timeout: 300000 // 5 minutes
        })
      );
      console.log('✅ MainDialog constructor - OAuthPrompt created successfully');
    } catch (error) {
      console.error('❌ MainDialog constructor - OAuthPrompt creation failed:', error);
    }

    this.addDialog(
      new WaterfallDialog(WATERFALL_ID, [
        this.promptStep.bind(this),
        this.tokenStep.bind(this)
      ])
    );

    this.initialDialogId = WATERFALL_ID;
  }

  async promptStep(step) {
    // Start OAuth flow
    console.log('🔍 MainDialog promptStep - CONNECTION_NAME:', CONNECTION_NAME);
    console.log('🔍 MainDialog promptStep - Starting OAuthPrompt...');
    try {
      const result = await step.beginDialog(OAUTH_PROMPT_ID);
      console.log('✅ MainDialog promptStep - OAuthPrompt started successfully');
      return result;
    } catch (error) {
      console.error('❌ MainDialog promptStep Error:', error);
      console.error('❌ Error stack:', error.stack);
      await step.context.sendActivity('❌ **OAuth Prompt Error**\n\nFailed to create sign-in prompt. Check configuration.');
      return await step.endDialog();
    }
  }

  async tokenStep(step) {
    const tokenResponse = step.result; // TokenResponse or undefined
    if (tokenResponse && tokenResponse.token) {
      await step.context.sendActivity("✅ **You are now signed in to Microsoft 365!**\n\n" +
        "You can now use commands like:\n" +
        "• `recent` - See your recent files\n" +
        "• `search [keyword]` - Search documents\n" +
        "• `help` - See all commands");
    } else {
      await step.context.sendActivity("⚠️ **Sign-in was not completed.**\n\nPlease try again by typing `signin`.");
    }
    return await step.endDialog();
  }
}

module.exports = { MainDialog, DIALOG_ID };
