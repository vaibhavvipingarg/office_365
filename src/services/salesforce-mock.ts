// Mock implementation of Salesforce integration
// This mock doesn't rely on any Node.js modules

export class SalesforceAuth {
  private static instance: SalesforceAuth;
  private isConnected: boolean = false;

  private constructor() {
    console.log("Mock Salesforce service initialized");
  }

  static getInstance(): SalesforceAuth {
    if (!SalesforceAuth.instance) {
      SalesforceAuth.instance = new SalesforceAuth();
    }
    return SalesforceAuth.instance;
  }

  static async checkAuth(): Promise<boolean> {
    const instance = SalesforceAuth.getInstance();
    console.log("Mock Salesforce: Checking auth status");
    return instance.isConnected;
  }

  static async login(): Promise<void> {
    const instance = SalesforceAuth.getInstance();
    console.log("Mock Salesforce: Starting authentication flow");
    
    // Simulate OAuth flow with a delay
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    instance.isConnected = true;
    console.log("Mock Salesforce: Authentication successful");
  }

  static async getData(): Promise<any[]> {
    const instance = SalesforceAuth.getInstance();
    
    if (!instance.isConnected) {
      console.error("Mock Salesforce: Not authenticated");
      throw new Error('Not authenticated with Salesforce');
    }

    console.log("Mock Salesforce: Fetching data");
    
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 800));

    // Return mock data
    return [
      {
        Id: '001xx000003DGb1AAG',
        Name: 'Acme Corporation',
        Industry: 'Manufacturing',
        AnnualRevenue: '$4.3M',
        LastModifiedDate: '2023-12-15',
        Status: 'Active',
        Phone: '(555) 123-4567',
        Website: 'https://www.acme.example.com'
      },
      {
        Id: '001xx000003DGb2AAG',
        Name: 'Globex',
        Industry: 'Technology',
        AnnualRevenue: '$8.7M',
        LastModifiedDate: '2024-01-22',
        Status: 'Active',
        Phone: '(555) 765-4321',
        Website: 'https://www.globex.example.com'
      },
      {
        Id: '001xx000003DGb3AAG',
        Name: 'Soylent Corp',
        Industry: 'Food & Beverage',
        AnnualRevenue: '$2.1M',
        LastModifiedDate: '2023-11-03',
        Status: 'Inactive',
        Phone: '(555) 888-9999',
        Website: 'https://www.soylent.example.com'
      },
      {
        Id: '001xx000003DGb4AAG',
        Name: 'Initech',
        Industry: 'Software',
        AnnualRevenue: '$1.4M',
        LastModifiedDate: '2024-02-18',
        Status: 'Active',
        Phone: '(555) 555-5555',
        Website: 'https://www.initech.example.com'
      },
      {
        Id: '001xx000003DGb5AAG',
        Name: 'Umbrella Corporation',
        Industry: 'Pharmaceuticals',
        AnnualRevenue: '$12.9M',
        LastModifiedDate: '2024-03-07',
        Status: 'Active',
        Phone: '(555) 666-7777',
        Website: 'https://www.umbrella.example.com'
      },
      {
        Id: '001xx000003DGb6AAG',
        Name: 'Wayne Enterprises',
        Industry: 'Technology',
        AnnualRevenue: '$45.2M',
        LastModifiedDate: '2024-03-10',
        Status: 'Active',
        Phone: '(555) 999-8888',
        Website: 'https://www.wayne.example.com'
      },
      {
        Id: '001xx000003DGb7AAG',
        Name: 'Stark Industries',
        Industry: 'Manufacturing',
        AnnualRevenue: '$57.8M',
        LastModifiedDate: '2024-01-15',
        Status: 'Active',
        Phone: '(555) 111-2222',
        Website: 'https://www.stark.example.com'
      }
    ];
  }
} 