# Azure OpenAI Setup Instructions

## 1. Get Credentials from Azure Portal

### From your `chainsaw-azure-openai` resource:
1. Go to **"Keys and Endpoint"** in the left menu
2. Copy **Key 1** (your API key)
3. Copy **Endpoint** URL (should look like: `https://chainsaw-azure-openai.openai.azure.com/`)

## 2. Deploy Model in Azure AI Foundry

1. Click **"Explore Azure AI Foundry portal"** 
2. Go to **"Deployments"** → **"Create new deployment"**
3. Configure:
   - **Model**: `gpt-35-turbo`
   - **Deployment name**: `gpt-35-turbo`
   - **Version**: Latest
   - **Rate limit**: 30K tokens/minute

## 3. Add Environment Variables to Azure App Service

Go to your Azure App Service `jj-sharepoint-bot-web` and add these environment variables:

```
AZURE_OPENAI_ENDPOINT=https://chainsaw-azure-openai.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key-here
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-35-turbo
```

### Steps to add environment variables:
1. Go to your App Service `jj-sharepoint-bot-web`
2. Click **"Configuration"** in the left menu
3. Click **"+ New application setting"** for each variable
4. Click **"Save"** at the top
5. Click **"Continue"** when prompted to restart

## 4. Test the Integration

After adding the environment variables and restarting your app:

1. Wait 2-3 minutes for restart
2. Test in Web Chat:
   - `help` - Should show AI features
   - `summarize [document]` - Test AI summarization
   - Ask any question about your documents

## 5. Expected Improvements

With Azure OpenAI, you should see:
- Much better document summaries
- More intelligent insights
- Context-aware question answering
- Professional-quality responses

## Troubleshooting

If you get errors:
1. Check environment variables are correctly set
2. Verify API key is valid
3. Ensure deployment name matches exactly
4. Check Azure OpenAI resource is in same region as your bot
