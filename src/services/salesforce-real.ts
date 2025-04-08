// Real Salesforce integration using OAuth and REST API
// This implementation works in the browser environment

export class SalesforceAuth {
  private static instance: SalesforceAuth;
  private isConnected: boolean = false;
  private accessToken: string | null = null;
  private instanceUrl: string | null = null;
  private tokenExpiry: number | null = null; // Timestamp when token expires

  // Salesforce OAuth configuration - get from localStorage if available
  private static get CLIENT_ID(): string {
    return localStorage.getItem('sf_client_id') || 'YOUR_SALESFORCE_CLIENT_ID';
  }
  private static set CLIENT_ID(value: string) {
    localStorage.setItem('sf_client_id', value);
  }
  private static readonly REDIRECT_URI = window.location.origin + '/office_365/oauth-callback.html';
  private static readonly LOGIN_URL = 'https://sdb42com6.test13.my.pc-rnd.salesforce.com';

  private constructor() {
    console.log("Real Salesforce service initialized");
    // Check if we have a token in local storage
    this.loadSession();
  }

  static getInstance(): SalesforceAuth {
    if (!SalesforceAuth.instance) {
      SalesforceAuth.instance = new SalesforceAuth();
    }
    return SalesforceAuth.instance;
  }

  static async checkAuth(): Promise<boolean> {
    const instance = SalesforceAuth.getInstance();
    console.log("Salesforce: Checking auth status");
    
    // Check if token is expired
    if (instance.isConnected && instance.tokenExpiry) {
      // Add 5 minute buffer before actual expiration
      const isExpired = Date.now() > (instance.tokenExpiry - (5 * 60 * 1000));
      if (isExpired) {
        console.log("Salesforce: Token expired, needs refresh");
        instance.isConnected = false;
        
        // Clear the token but keep the client ID
        sessionStorage.removeItem('sf_access_token');
        sessionStorage.removeItem('sf_instance_url');
        sessionStorage.removeItem('sf_token_expiry');
        
        // Don't clear localStorage since we want to keep the client ID
        instance.accessToken = null;
        instance.instanceUrl = null;
        instance.tokenExpiry = null;
      } else {
        console.log(`Salesforce: Token valid for ${Math.floor((instance.tokenExpiry - Date.now()) / 60000)} more minutes`);
      }
    }
    
    return instance.isConnected;
  }

  private loadSession(): void {
    const token = localStorage.getItem('sf_access_token');
    const url = localStorage.getItem('sf_instance_url');
    const expiryStr = localStorage.getItem('sf_token_expiry');
    
    if (token && url) {
      this.accessToken = token;
      this.instanceUrl = url;
      this.tokenExpiry = expiryStr ? parseInt(expiryStr, 10) : null;
      
      // Only set connected if we have a token and it's not expired
      if (this.tokenExpiry && Date.now() < (this.tokenExpiry - (5 * 60 * 1000))) {
        this.isConnected = true;
        console.log(`Salesforce: Restored session from storage. Token valid for ${Math.floor((this.tokenExpiry - Date.now()) / 60000)} more minutes`);
      } else if (!this.tokenExpiry) {
        // If we don't have expiry info, assume it's still valid
        this.isConnected = true;
        console.log("Salesforce: Restored session from storage (no expiry info)");
      } else {
        console.log("Salesforce: Found expired token in storage");
        this.isConnected = false;
      }
    }
  }

  private saveSession(): void {
    if (this.accessToken && this.instanceUrl) {
      localStorage.setItem('sf_access_token', this.accessToken);
      localStorage.setItem('sf_instance_url', this.instanceUrl);
      if (this.tokenExpiry) {
        localStorage.setItem('sf_token_expiry', this.tokenExpiry.toString());
      }
      console.log("Salesforce: Saved session to local storage");
    }
  }

  static async login(): Promise<void> {
    const instance = SalesforceAuth.getInstance();
    console.log("Salesforce: Starting authentication flow");
    
    // Check if we already have a valid token
    if (await SalesforceAuth.checkAuth()) {
      console.log("Salesforce: Already authenticated with valid token");
      return;
    }

    // Check if we have a client ID
    const clientId = localStorage.getItem('sf_client_id');
    if (!clientId) {
      console.error("Salesforce: No client ID found");
      throw new Error('No Salesforce client ID configured. Please configure your Salesforce connection.');
    }

    // For quick local testing without OAuth, use this flag
    const skipOAuthInLocalMode = false; // Set to true to skip real OAuth in local mode

    // For local development mode, we can either use mock data or try real OAuth
    if (!window.hasOwnProperty('Office') && skipOAuthInLocalMode) {
      console.log("Local mode: Using mock authentication");
      await new Promise(resolve => setTimeout(resolve, 1000));
      instance.isConnected = true;
      instance.accessToken = "mock_token";
      instance.instanceUrl = "https://mock.salesforce.com";
      // Set mock token to expire in 2 hours
      instance.tokenExpiry = Date.now() + (2 * 60 * 60 * 1000);
      instance.saveSession();
      return;
    }
    
    try {
      // Start OAuth flow using popup window
      const authUrl = `${SalesforceAuth.LOGIN_URL}/services/oauth2/authorize?` +
        `client_id=${encodeURIComponent(SalesforceAuth.CLIENT_ID)}` +
        `&redirect_uri=${encodeURIComponent(SalesforceAuth.REDIRECT_URI)}` +
        `&response_type=token`;
      
      console.log(`Opening OAuth URL: ${authUrl}`);
      console.log(`Redirect URI: ${SalesforceAuth.REDIRECT_URI}`);
      console.log(`Make sure this matches the Callback URL in your Salesforce Connected App settings`);
      
      // Open popup window for authentication
      const width = 600;
      const height = 700;
      const left = window.screen.width / 2 - width / 2;
      const top = window.screen.height / 2 - height / 2;
      
      const popup = window.open(
        authUrl,
        'SalesforceOAuth',
        `width=${width},height=${height},left=${left},top=${top}`
      );
      
      if (!popup) {
        throw new Error('Could not open popup window for authentication. Please check your popup blocker settings.');
      }
      
      // Wait for the OAuth redirect to complete
      return new Promise((resolve, reject) => {
        const timer = setInterval(() => {
          if (popup.closed) {
            clearInterval(timer);
            // Check if we have the token after popup is closed
            if (instance.isConnected) {
              resolve();
            } else {
              reject(new Error('Authentication cancelled or failed'));
            }
          }
        }, 500);
        
        // Listen for OAuth response from popup
        window.addEventListener('message', (event) => {
          console.log('Received message:', event.origin, event.data);
          if (event.origin !== window.location.origin) return;
          
          if (event.data && event.data.type === 'salesforce_oauth_response') {
            clearInterval(timer);
            popup.close();
            
            if (event.data.error) {
              console.error('OAuth error:', event.data.error);
              reject(new Error(event.data.error));
            } else {
              console.log('OAuth successful:', event.data.access_token?.substring(0, 10) + '...');
              instance.accessToken = event.data.access_token;
              instance.instanceUrl = event.data.instance_url;
              
              // Calculate token expiry (default to 2 hours if not provided)
              // Salesforce tokens typically expire in 2 hours
              const expiresIn = event.data.expires_in || 7200; // 2 hours in seconds
              instance.tokenExpiry = Date.now() + (expiresIn * 1000);
              console.log(`Token will expire at: ${new Date(instance.tokenExpiry).toLocaleString()}`);
              
              instance.isConnected = true;
              instance.saveSession();
              resolve();
            }
          }
        }, false);
      });
    } catch (error) {
      console.error('Authentication error:', error);
      throw error;
    }
  }

  static async getData(): Promise<any[]> {
    const instance = SalesforceAuth.getInstance();
    
    if (!instance.isConnected) {
      console.error("Salesforce: Not authenticated");
      throw new Error('Not authenticated with Salesforce');
    }

    console.log("Salesforce: Fetching dashboard data");
    
    // Only use mock data if we don't have a real access token
    if (!instance.accessToken || !instance.instanceUrl) {
      console.log("No valid Salesforce token: Returning mock dashboard data");
      
      // Simulate network delay
      await new Promise(resolve => setTimeout(resolve, 800));

      // Return mock dashboard data
      return [
        {
          Id: '0DyUA00000004GX0AY',
          DeveloperName: 'New_Dashboard',
          CreatedBy: { Name: 'Admin User' },
          Description: 'A dashboard',
          CreatedDate: '2024-08-19T16:36:05.000+0000',
          MasterLabel: 'New Dashboard',
          AnalyticsWorkspaceId: '1DyUA00000004GX0AY',
          AnalyticsWorkspace: { MasterLabel: 'sonia test workspace' }
        },
        {
          Id: '0DyUA00000002p1UAA',
          DeveloperName: 'NE_Sales',
          CreatedBy: { Name: 'Admin User' },
          Description: 'A dashboard',
          CreatedDate: '2024-08-19T17:32:55.000+0000',
          MasterLabel: 'NE Sales',
          AnalyticsWorkspaceId: '1DyUA00000002p0AA',
          AnalyticsWorkspace: { MasterLabel: 'Test workspace' }
        },
        {
          Id: '0DyUA00000002p2UAA',
          DeveloperName: 'New_Dashboard_2',
          CreatedBy: { Name: 'Admin User' },
          Description: 'A dashboard',
          CreatedDate: '2024-08-19T18:58:49.000+0000',
          MasterLabel: 'New Dashboard',
          AnalyticsWorkspaceId: '1DyUA00000002p0AA',
          AnalyticsWorkspace: { MasterLabel: 'Test workspace' }
        },
        {
          Id: '0DyUA0000000BDB0A2',
          DeveloperName: 'BK_Dashboard',
          CreatedBy: { Name: 'Admin User' },
          Description: 'A dashboard',
          CreatedDate: '2024-08-21T09:18:45.000+0000',
          MasterLabel: 'BK Dashboard',
          AnalyticsWorkspaceId: '1DyUA0000000BDB0A2',
          AnalyticsWorkspace: { MasterLabel: 'BK Workspace' }
        },
        {
          Id: '0DyUA0000000BOT0A2',
          DeveloperName: 'New_Dashboard_3',
          CreatedBy: { Name: 'Admin User' },
          Description: 'A dashboard',
          CreatedDate: '2024-08-21T09:55:45.000+0000',
          MasterLabel: 'New Dashboard',
          AnalyticsWorkspaceId: '1DyUA0000000BOT0A2',
          AnalyticsWorkspace: { MasterLabel: 'Test Blitz' }
        }
      ];
    }

    try {
      // Real Salesforce API call
      console.log(`Making API call to ${instance.instanceUrl}`);
      
      const response = await fetch(`${instance.instanceUrl}/services/data/v64.0/query?q=${encodeURIComponent(
        'SELECT Id, DeveloperName, CreatedBy.Name, Description, CreatedDate, MasterLabel, AnalyticsWorkspaceId, AnalyticsWorkspace.MasterLabel FROM AnalyticsDashboard ORDER BY LastModifiedDate DESC LIMIT 5'
      )}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${instance.accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Salesforce API error: ${response.status}`, errorText);
        throw new Error(`Salesforce API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      console.log('Received dashboard data from Salesforce API:', data);
      
      // Transform the data to match our expected format
      return data.records.map((record: any) => {
        console.log('Processing dashboard record:', record);
        
        // Extract the actual data from the record, ignoring the attributes field
        const {
          Id,
          DeveloperName,
          CreatedBy,
          Description,
          CreatedDate,
          MasterLabel,
          AnalyticsWorkspaceId,
          AnalyticsWorkspace
        } = record;
        
        // Format the creation date
        const formattedDate = CreatedDate ? 
          new Date(CreatedDate).toLocaleDateString() : 'N/A';
        
        return {
          Id: Id || `dashboard-${Math.random().toString(36).substring(2, 10)}`,
          DeveloperName: DeveloperName || MasterLabel || 'Untitled Dashboard',
          CreatedBy: CreatedBy || { Name: 'Unknown User' },
          Description: Description || 'No description',
          CreatedDate: CreatedDate,
          MasterLabel: MasterLabel || DeveloperName || 'Untitled Dashboard',
          AnalyticsWorkspaceId: AnalyticsWorkspaceId || '',
          AnalyticsWorkspace: AnalyticsWorkspace || { MasterLabel: 'Unknown Workspace' },
          FormattedDate: formattedDate
        };
      });
    } catch (error) {
      console.error('Salesforce dashboard data fetch error:', error);
      throw error;
    }
  }

  // Get the current user ID from the OAuth session
  private static async getCurrentUserId(): Promise<string> {
    const instance = SalesforceAuth.getInstance();
    
    if (!instance.isConnected || !instance.accessToken) {
      throw new Error('Not authenticated with Salesforce');
    }

    try {
      // Extract user ID from the ID URL in the OAuth response if available
      const userId = localStorage.getItem('sf_user_id');
      if (userId) {
        return userId;
      }

      // If we don't have the user ID stored, make an API call to get current user info
      const response = await fetch(`${instance.instanceUrl}/services/data/v57.0/chatter/users/me`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${instance.accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`Failed to get user info: ${response.status}`);
      }

      const userInfo = await response.json();
      const newUserId = userInfo.id;
      
      // Store user ID for future use
      localStorage.setItem('sf_user_id', newUserId);
      
      return newUserId;
    } catch (error) {
      console.error('Error getting user ID:', error);
      // Return a placeholder ID for development
      return '005xx000001234AAA';
    }
  }

  // Get metrics data from Tableau Analytics API
  static async getMetricsData(): Promise<any[]> {
    const instance = SalesforceAuth.getInstance();
    
    if (!instance.isConnected) {
      console.error("Salesforce: Not authenticated");
      throw new Error('Not authenticated with Salesforce');
    }

    console.log("Salesforce: Fetching metrics data");
    
    // Only use mock data if we don't have a real access token
    if (!instance.accessToken || !instance.instanceUrl) {
      console.log("No valid Salesforce token: Returning mock metrics data");
      
      // Simulate network delay
      await new Promise(resolve => setTimeout(resolve, 800));

      // Return mock metrics data based on the screenshot
      const mockMetric = {
        id: "1HUUA0000001F8H4AU",
        name: "sub_metric_b6fcf61a_5871_4d4d_9917_cd762b02b435",
        label: "Swim Count",
        metadata: {
          root: {
            asset: {
              createdBy: { Name: "Admin User" },
              createdDate: "2025-03-21T20:21:48.000Z",
              id: "1HUUA0000001F8H4AU",
              label: "Swim Count",
              lastModifiedBy: { Name: "Admin User" },
              lastModifiedDate: "2025-03-21T20:21:48.000Z",
              metricFilterSummary: "Last 30 days",
              metricFilters: [],
              name: "sub_metric_b6fcf61a_5871_4d4d_9917_cd762b02b435"
            }
          }
        }
      };
      return [mockMetric];
    }

    try {
      // Get current user ID
      const userId = await this.getCurrentUserId();
      
      // First, get the list of metrics
      console.log(`Making metrics API call to ${instance.instanceUrl} for user ${userId}`);
      
      const response = await fetch(`${instance.instanceUrl}/services/data/v64.0/tableau/follow/followers/${userId}/followed-assets`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${instance.accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Salesforce API error: ${response.status}`, errorText);
        throw new Error(`Salesforce API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      console.log('Received metrics list from Salesforce API:', data);
      
      // For each metric, fetch its metadata
      if (data.followedAssets && Array.isArray(data.followedAssets)) {
        const metricsWithMetadata = await Promise.all(
          data.followedAssets.map(async (metric: any) => {
            try {
              console.log(`Fetching metadata for metric ${metric.id}`);
              const metadataResponse = await fetch(`${instance.instanceUrl}/services/data/v64.0/tableau/download?metadataOnly=false`, {
                method: 'POST',
                headers: {
                  'Authorization': `Bearer ${instance.accessToken}`,
                  'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                  asset: {
                    type: "Submetric",
                    assetId: metric.id,
                    timeRange: "eyJmaWVsZE5hbWUiOiIiLCJvcGVyYXRvciI6Ikxhc3RORGF5cyIsInZhbHVlcyI6WzMwXX0"
                  }
                })
              });

              if (!metadataResponse.ok) {
                throw new Error(`Failed to fetch metric metadata: ${metadataResponse.statusText}`);
              }

              const metadata = await metadataResponse.json();
              console.log(`Received metadata for metric ${metric.id}:`, metadata);

              // Return a simplified object with the metric data
              return {
                id: metric.id,
                name: metric.name,
                label: metric.label || metadata.root?.asset?.label,
                hasError: false,
                metadata: metadata,
                metricValue: metadata.root?.asset?.metricValue,
                metricChange: metadata.root?.asset?.metricChange,
                metricSentiment: metadata.root?.asset?.metricSentiment,
                metricInsight: metadata.root?.asset?.metricInsight,
                metricFilterSummary: metadata.root?.asset?.metricFilterSummary,
                createdDate: metadata.root?.asset?.createdDate,
                lastModifiedDate: metadata.root?.asset?.lastModifiedDate,
                createdBy: metadata.root?.asset?.createdBy,
                lastModifiedBy: metadata.root?.asset?.lastModifiedBy
              };
            } catch (error) {
              console.warn(`Failed to fetch metadata for metric ${metric.id}:`, error);
              // Return the metric with error flag but keep basic info
              return {
                id: metric.id,
                name: metric.name,
                label: metric.label,
                hasError: true,
                errorMessage: error instanceof Error ? error.message : 'Failed to load metric details',
                metadata: null
              };
            }
          })
        );

        // Filter out completely failed metrics but keep ones with partial data
        return metricsWithMetadata.filter(metric => metric !== null);
      }
      
      return [];
    } catch (error) {
      console.error('Salesforce metrics data fetch error:', error);
      throw error;
    }
  }
} 