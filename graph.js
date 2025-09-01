// graph.js
const https = require('https');

function graphGet(path, accessToken) {
  const options = {
    hostname: 'graph.microsoft.com',
    path,
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  };
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', d => (data += d));
      res.on('end', () => {
        try {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            resolve(JSON.parse(data || '{}'));
          } else {
            reject(new Error(`Graph API ${res.statusCode}: ${data}`));
          }
        } catch (parseError) {
          reject(new Error(`Graph API response parsing failed: ${parseError.message}`));
        }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

module.exports = { graphGet };
