  // graph.js - SharePoint Document Reader
  const axios = require('axios');
  const { BotAppAuth } = require('./auth');
  
  class SharePointGraphClient {
    constructor(accessToken = null) {
      this.accessToken = accessToken;
      this.baseURL = 'https://graph.microsoft.com/v1.0';
      this.auth = new BotAppAuth();
    }
  
    async ensureToken() {
      if (!this.accessToken || this.accessToken === 'TEST_MODE') {
        console.log('🔄 Using Bot App Registration to get access token...');
        this.accessToken = await this.auth.getAccessToken();
      }
      return this.accessToken;
    }

  async request(endpoint) {
    try {
      await this.ensureToken();
      const response = await axios.get(`${this.baseURL}${endpoint}`, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json'
        }
      });
      return response.data;
    } catch (error) {
      console.error('Graph API Error:', error.response?.data || error.message);
      throw new Error(`Graph API failed: ${error.response?.status} ${error.response?.statusText}`);
    }
  }

  // Search for SharePoint sites
  async searchSites(query) {
    const endpoint = `/sites?search=${encodeURIComponent(query)}`;
    return await this.request(endpoint);
  }

  // Get site drive (document library)
  async getSiteDrive(siteId) {
    const endpoint = `/sites/${siteId}/drive`;
    return await this.request(endpoint);
  }

  // Search documents in a site
  async searchDocuments(siteId, query) {
    const endpoint = `/sites/${siteId}/drive/root/search(q='${encodeURIComponent(query)}')`;
    return await this.request(endpoint);
  }

  // Get recent documents from SharePoint sites (using application permissions)
  async getRecentDocuments() {
    try {
      // First, get all sites
      const sitesEndpoint = '/sites?$top=10';
      const sitesResponse = await this.request(sitesEndpoint);
      
      if (!sitesResponse.value || sitesResponse.value.length === 0) {
        return { value: [] };
      }

      // Get documents from the first available site
      const firstSite = sitesResponse.value[0];
      const documentsEndpoint = `/sites/${firstSite.id}/drive/root/children?$top=10&$orderby=lastModifiedDateTime desc`;
      return await this.request(documentsEndpoint);
    } catch (error) {
      console.error('Error getting recent documents:', error.message);
      // Fallback: try to get from root site
      try {
        const rootSiteEndpoint = '/sites/root/drive/root/children?$top=10&$orderby=lastModifiedDateTime desc';
        return await this.request(rootSiteEndpoint);
      } catch (fallbackError) {
        console.error('Fallback also failed:', fallbackError.message);
        throw error;
      }
    }
  }

  // Get document content (for text files)
  async getDocumentContent(driveId, itemId) {
    const endpoint = `/drives/${driveId}/items/${itemId}/content`;
    try {
      const response = await axios.get(`${this.baseURL}${endpoint}`, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`
        },
        responseType: 'text'
      });
      return response.data;
    } catch (error) {
      throw new Error(`Failed to get document content: ${error.message}`);
    }
  }

  // Get document metadata
  async getDocumentMetadata(driveId, itemId) {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    return await this.request(endpoint);
  }
}

// Legacy function for backward compatibility
async function graphGet(path, accessToken) {
  const client = new SharePointGraphClient(accessToken);
  return await client.request(path);
}

module.exports = { SharePointGraphClient, graphGet };
