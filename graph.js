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

  // Get recent documents from specific SharePoint site
  async getRecentDocuments() {
    try {
      console.log(`🎯 Attempting to access SharePoint with current token...`);
      
      // Try the most basic approach first - root site
      try {
        console.log(`🔄 Step 1: Trying root site access...`);
        const rootSiteResponse = await this.request('/sites/root');
        console.log(`✅ Root site accessible: ${rootSiteResponse.displayName}`);
        
        // Get files from root site
        const rootDocsEndpoint = `/sites/${rootSiteResponse.id}/drive/root/children?$top=20`;
        const rootDocsResponse = await this.request(rootDocsEndpoint);
        
        if (rootDocsResponse.value && rootDocsResponse.value.length > 0) {
          console.log(`📄 Found ${rootDocsResponse.value.length} items in root site`);
          return rootDocsResponse;
        }
      } catch (rootError) {
        console.log(`❌ Root site access failed: ${rootError.message}`);
      }
      
      // Try the specific OnlineCustomerServiceTeam site
      const siteUrl = 'jonoandjohno.sharepoint.com:/sites/OnlineCustomerServiceTeam859';
      console.log(`🔄 Step 2: Targeting specific site: ${siteUrl}`);
      
      const siteEndpoint = `/sites/${siteUrl}`;
      const siteResponse = await this.request(siteEndpoint);
      console.log(`✅ Found site: ${siteResponse.displayName}`);
      
      // Get documents from the site's document library
      const documentsEndpoint = `/sites/${siteResponse.id}/drive/root/children?$top=20&$orderby=lastModifiedDateTime desc&$expand=children($top=5)`;
      const documentsResponse = await this.request(documentsEndpoint);
      
      console.log(`📄 Found ${documentsResponse.value?.length || 0} items in the site`);
      
      // Also try to get more files by looking in subfolders
      const allFiles = [];
      if (documentsResponse.value) {
        for (const item of documentsResponse.value) {
          if (item.file) {
            // It's a file, add it
            allFiles.push(item);
          } else if (item.folder && !item.name.startsWith('.')) {
            // It's a folder, get files from it
            try {
              console.log(`🔍 Looking inside folder: ${item.name}`);
              const folderFilesEndpoint = `/sites/${siteResponse.id}/drive/items/${item.id}/children?$top=10&$filter=file ne null`;
              const folderFiles = await this.request(folderFilesEndpoint);
              if (folderFiles.value) {
                allFiles.push(...folderFiles.value);
                console.log(`📁 Found ${folderFiles.value.length} files in ${item.name}`);
              }
            } catch (folderError) {
              console.log(`❌ Couldn't access folder ${item.name}: ${folderError.message}`);
            }
          }
        }
      }
      
      console.log(`📄 Total files found: ${allFiles.length}`);
      return { value: allFiles };
      
    } catch (error) {
      console.error('Error getting documents from OnlineCustomerServiceTeam859 site:', error.message);
      
      // Fallback: try alternative site URL format
      try {
        console.log('🔄 Trying alternative site URL format...');
        const altSiteEndpoint = '/sites/jonoandjohno.sharepoint.com,/sites/OnlineCustomerServiceTeam859';
        const altSiteResponse = await this.request(altSiteEndpoint);
        
        const documentsEndpoint = `/sites/${altSiteResponse.id}/drive/root/children?$top=20&$orderby=lastModifiedDateTime desc`;
        return await this.request(documentsEndpoint);
        
      } catch (fallbackError) {
        console.error('Alternative format also failed:', fallbackError.message);
        
        // Final fallback: search all sites for the one we want
        try {
          console.log('🔍 Searching all sites for JonoJohno-allstaff...');
          const allSitesEndpoint = '/sites?$filter=displayName eq \'JonoJohno-allstaff\' or name eq \'JonoJohno-allstaff\'&$top=10';
          const sitesResponse = await this.request(allSitesEndpoint);
          
          if (sitesResponse.value && sitesResponse.value.length > 0) {
            const targetSite = sitesResponse.value[0];
            console.log(`🎯 Found target site: ${targetSite.displayName}`);
            
            const documentsEndpoint = `/sites/${targetSite.id}/drive/root/children?$top=20&$orderby=lastModifiedDateTime desc`;
            return await this.request(documentsEndpoint);
          } else {
            console.log('❌ Could not find JonoJohno-allstaff site');
            throw new Error('Could not find the JonoJohno-allstaff SharePoint site');
          }
        } catch (finalError) {
          console.error('Final fallback failed:', finalError.message);
          throw error;
        }
      }
    }
  }

  // Get document content (handles both text and binary files)
  async getDocumentContent(driveId, itemId, asBinary = false) {
    const endpoint = `/drives/${driveId}/items/${itemId}/content`;
    try {
      await this.ensureToken();
      const response = await axios.get(`${this.baseURL}${endpoint}`, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`
        },
        responseType: asBinary ? 'arraybuffer' : 'text',
        timeout: 30000 // 30 second timeout
      });
      
      if (asBinary) {
        return Buffer.from(response.data);
      }
      return response.data;
    } catch (error) {
      console.error(`Error getting document content for ${itemId}: ${error.message}`);
      if (error.response?.status === 413) {
        console.error('File too large to process');
      } else if (error.response?.status === 404) {
        console.error('File not found or no permission');
      }
      return null;
    }
  }

  // Search documents in the OnlineCustomerServiceTeam859 site
  async searchDocumentsInSite(query) {
    try {
      // Target the specific OnlineCustomerServiceTeam859 site
      const siteUrl = 'jonoandjohno.sharepoint.com:/sites/OnlineCustomerServiceTeam859';
      console.log(`🔍 Searching in site: ${siteUrl} for query: "${query}"`);
      
      // Get the specific site
      const siteEndpoint = `/sites/${siteUrl}`;
      const siteResponse = await this.request(siteEndpoint);
      
      // Search within the site
      const searchEndpoint = `/sites/${siteResponse.id}/drive/root/search(q='${encodeURIComponent(query)}')?$top=20`;
      const searchResults = await this.request(searchEndpoint);
      
      console.log(`🎯 Search found ${searchResults.value?.length || 0} results`);
      return searchResults;
      
    } catch (error) {
      console.error('Error searching in JonoJohno-allstaff site:', error.message);
      throw error;
    }
  }

  // Get document metadata
  async getDocumentMetadata(driveId, itemId) {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    return await this.request(endpoint);
  }

  // Get comprehensive site statistics
  async getSiteStatistics() {
    try {
      const siteUrl = 'jonoandjohno.sharepoint.com:/sites/OnlineCustomerServiceTeam859';
      console.log(`📊 Getting statistics for site: ${siteUrl}`);
      
      // Get the site
      const siteEndpoint = `/sites/${siteUrl}`;
      const siteResponse = await this.request(siteEndpoint);
      
      // Get all items recursively
      const allItems = await this.getAllItemsRecursively(siteResponse.id);
      
      // Analyze the items
      const stats = this.analyzeItems(allItems);
      
      return stats;
    } catch (error) {
      console.error('Error getting site statistics:', error);
      throw error;
    }
  }

  // Recursively get all items from all folders
  async getAllItemsRecursively(siteId, folderId = 'root', allItems = []) {
    try {
      const endpoint = `/sites/${siteId}/drive/${folderId}/children?$top=999`;
      const response = await this.request(endpoint);
      
      for (const item of response.value || []) {
        allItems.push(item);
        
        // If it's a folder, recursively get its contents
        if (item.folder) {
          await this.getAllItemsRecursively(siteId, item.id, allItems);
        }
      }
      
      return allItems;
    } catch (error) {
      console.error(`Error getting items from folder ${folderId}:`, error);
      return allItems; // Return what we have so far
    }
  }

  // Analyze items to generate statistics
  analyzeItems(items) {
    let folderCount = 0;
    let fileCount = 0;
    let totalSize = 0;
    const fileTypes = {};
    let lastModified = new Date(0);

    for (const item of items) {
      if (item.folder) {
        folderCount++;
      } else if (item.file) {
        fileCount++;
        totalSize += item.size || 0;
        
        // Track file types
        const extension = item.name.split('.').pop()?.toLowerCase() || 'no extension';
        fileTypes[extension] = (fileTypes[extension] || 0) + 1;
        
        // Track latest modification
        const itemModified = new Date(item.lastModifiedDateTime);
        if (itemModified > lastModified) {
          lastModified = itemModified;
        }
      }
    }

    // Convert file types to sorted array
    const fileTypesArray = Object.entries(fileTypes)
      .map(([type, count]) => ({ type, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10); // Top 10 file types

    return {
      folderCount,
      fileCount,
      totalSize: this.formatFileSize(totalSize),
      lastModified: lastModified.toLocaleDateString(),
      fileTypes: fileTypesArray
    };
  }

  // Format file size in human-readable format
  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
}

// Legacy function for backward compatibility
async function graphGet(path, accessToken) {
  const client = new SharePointGraphClient(accessToken);
  return await client.request(path);
}

module.exports = { SharePointGraphClient, graphGet };
