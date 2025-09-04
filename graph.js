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
      // Try multiple approaches to find the SharePoint site
      console.log(`🎯 Attempting to connect to SharePoint...`);
      
      // First, try to get all sites to see what's available
      let siteResponse = null;
      
      try {
        console.log(`🔍 Step 1: Getting all available sites...`);
        const allSitesResponse = await this.request('/sites?$top=10');
        console.log(`📋 Found ${allSitesResponse.value?.length || 0} sites available`);
        
        if (allSitesResponse.value && allSitesResponse.value.length > 0) {
          console.log(`📋 Available sites:`);
          allSitesResponse.value.forEach((site, i) => {
            console.log(`  ${i + 1}. ${site.displayName} - ${site.webUrl}`);
          });
          
          // Look for JonoJohno site
          const targetSite = allSitesResponse.value.find(site => 
            site.displayName?.toLowerCase().includes('jono') || 
            site.webUrl?.toLowerCase().includes('jono') ||
            site.webUrl?.toLowerCase().includes('allstaff')
          );
          
          if (targetSite) {
            console.log(`✅ Found target site: ${targetSite.displayName}`);
            siteResponse = targetSite;
          } else {
            console.log(`⚠️ No JonoJohno site found, using first available site: ${allSitesResponse.value[0].displayName}`);
            siteResponse = allSitesResponse.value[0];
          }
        } else {
          throw new Error('No sites found');
        }
      } catch (sitesError) {
        console.log(`❌ Couldn't get sites list: ${sitesError.message}`);
        
        // Fallback: Try specific site URL formats
        const siteUrls = [
          'jonoandjohno.sharepoint.com:/sites/JonoJohno-allstaff',
          'jonoandjohno.sharepoint.com,80ee117e-949a-4cc2-9d56-b0c4923a47f2,c9a6ce8e-ff75-4451-8666-7c4f5ee30d34',
          'jonoandjohno.sharepoint.com:/sites/JonoAndJohno-allstaff'
        ];
        
        for (const siteUrl of siteUrls) {
          try {
            console.log(`🔄 Trying site URL: ${siteUrl}`);
            const siteEndpoint = `/sites/${siteUrl}`;
            siteResponse = await this.request(siteEndpoint);
            console.log(`✅ Successfully connected to: ${siteResponse.displayName}`);
            break;
          } catch (urlError) {
            console.log(`❌ Failed with URL ${siteUrl}: ${urlError.message}`);
          }
        }
        
        if (!siteResponse) {
          throw new Error('Could not connect to any SharePoint site');
        }
      }
      
      if (!siteResponse) {
        throw new Error('Could not find any SharePoint site');
      }
      
      console.log(`✅ Using site: ${siteResponse.displayName}`);
      
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
      console.error('Error getting documents from JonoJohno-allstaff site:', error.message);
      
      // Fallback: try alternative site URL format
      try {
        console.log('🔄 Trying alternative site URL format...');
        const altSiteEndpoint = '/sites/jonoandjohno.sharepoint.com,/sites/JonoJohno-allstaff';
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

  // Search documents in the JonoJohno-allstaff site
  async searchDocumentsInSite(query) {
    try {
      // Target the specific JonoJohno-allstaff site
      const siteUrl = 'jonoandjohno.sharepoint.com:/sites/JonoJohno-allstaff';
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
}

// Legacy function for backward compatibility
async function graphGet(path, accessToken) {
  const client = new SharePointGraphClient(accessToken);
  return await client.request(path);
}

module.exports = { SharePointGraphClient, graphGet };
