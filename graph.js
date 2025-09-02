// graph.js - SharePoint Document Reader
const axios = require('axios');

class SharePointGraphClient {
  constructor(accessToken) {
    this.accessToken = accessToken;
    this.baseURL = 'https://graph.microsoft.com/v1.0';
  }

  async request(endpoint) {
    try {
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

  // Get recent documents from user's OneDrive/SharePoint
  async getRecentDocuments() {
    const endpoint = '/me/drive/recent?$top=10';
    return await this.request(endpoint);
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
