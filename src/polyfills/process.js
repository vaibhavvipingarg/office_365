// Minimal process polyfill for browser environment
window.process = {
  env: {
    NODE_ENV: 'development',
    SF_LOGIN_URL: 'https://login.salesforce.com',
    SF_CLIENT_ID: 'mock_client_id',
    SF_CLIENT_SECRET: 'mock_client_secret',
    SF_REDIRECT_URI: 'https://localhost:3001/oauth/callback'
  }
};

export default window.process; 