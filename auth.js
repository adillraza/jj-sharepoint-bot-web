// auth.js - Bot App Registration Authentication
const axios = require('axios');

class BotAppAuth {
    constructor() {
        this.clientId = process.env.MicrosoftAppId;
        this.clientSecret = process.env.MicrosoftAppPassword;
        this.tenantId = process.env.MicrosoftAppTenantId;
        this.scope = 'https://graph.microsoft.com/.default';
    }

    async getAccessToken() {
        try {
            console.log('🔐 Getting access token using Bot App Registration...');
            console.log(`📋 Using Client ID: ${this.clientId?.substring(0, 8)}...`);
            console.log(`📋 Using Tenant ID: ${this.tenantId}`);
            console.log(`📋 Requesting scope: ${this.scope}`);
            
            const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
            const params = new URLSearchParams();
            params.append('client_id', this.clientId);
            params.append('client_secret', this.clientSecret);
            params.append('scope', this.scope);
            params.append('grant_type', 'client_credentials');

            const response = await axios.post(tokenUrl, params, {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });

            console.log('✅ Access token obtained successfully using Bot App Registration');
            
            // Decode token to see what permissions we have
            if (response.data.access_token) {
                try {
                    const tokenParts = response.data.access_token.split('.');
                    const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
                    console.log(`🔍 Token app ID: ${payload.appid}`);
                    console.log(`🔍 Token roles: ${payload.roles ? payload.roles.join(', ') : 'None'}`);
                } catch (decodeError) {
                    console.log('⚠️ Could not decode token for debugging');
                }
            }
            
            return response.data.access_token;
        } catch (error) {
            console.error('❌ Failed to get access token:', error.response?.data || error.message);
            throw error;
        }
    }

}

module.exports = { BotAppAuth };
