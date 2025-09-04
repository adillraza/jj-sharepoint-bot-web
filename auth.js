// auth.js - Azure Managed Identity Authentication
const { DefaultAzureCredential } = require('@azure/identity');

class ManagedIdentityAuth {
    constructor() {
        this.credential = new DefaultAzureCredential();
        this.graphScope = 'https://graph.microsoft.com/.default';
    }

    async getAccessToken() {
        try {
            console.log('🔐 Getting access token using Managed Identity...');
            const tokenResponse = await this.credential.getToken(this.graphScope);
            console.log('✅ Access token obtained successfully');
            return tokenResponse.token;
        } catch (error) {
            console.error('❌ Failed to get access token:', error);
            throw error;
        }
    }

    // Fallback to client credentials for development
    async getAccessTokenFallback() {
        const { Client } = require('@microsoft/microsoft-graph-client');
        const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
        
        try {
            const authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
                scopes: [this.graphScope]
            });
            
            const graphClient = Client.initWithMiddleware({ authProvider });
            return graphClient;
        } catch (error) {
            console.error('❌ Fallback authentication failed:', error);
            throw error;
        }
    }
}

module.exports = { ManagedIdentityAuth };
