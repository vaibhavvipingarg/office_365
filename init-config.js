// Initialize Salesforce configuration from window.SalesforceConfig
(function() {
  // Wait for DOM to be ready
  document.addEventListener('DOMContentLoaded', function() {
    // Check if SalesforceConfig is available from the config.js file
    if (window.SalesforceConfig) {
      console.log('Initializing Salesforce configuration from config.js');
      
      // Store the client ID in localStorage for the application to use
      if (window.SalesforceConfig.clientId) {
        localStorage.setItem('sf_client_id', window.SalesforceConfig.clientId);
        console.log('Stored Salesforce client ID in localStorage');
      }
      
      // Store the login URL in localStorage
      if (window.SalesforceConfig.loginUrl) {
        localStorage.setItem('sf_login_url', window.SalesforceConfig.loginUrl);
        console.log('Stored Salesforce login URL in localStorage');
      }
      
      // Store the redirect URI in localStorage
      if (window.SalesforceConfig.redirectUri) {
        localStorage.setItem('sf_redirect_uri', window.SalesforceConfig.redirectUri);
        console.log('Stored Salesforce redirect URI in localStorage');
      }
    } else {
      console.warn('SalesforceConfig not found. Make sure config.js is loaded before this script.');
    }
  });
})(); 