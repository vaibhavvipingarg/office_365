import { Connection, Auth } from '@salesforce/core';

export class SalesforceAuth {
  private static instance: SalesforceAuth;
  private connection: Connection | null = null;

  private constructor() {}

  static getInstance(): SalesforceAuth {
    if (!SalesforceAuth.instance) {
      SalesforceAuth.instance = new SalesforceAuth();
    }
    return SalesforceAuth.instance;
  }

  static async checkAuth(): Promise<boolean> {
    try {
      const instance = SalesforceAuth.getInstance();
      return instance.connection !== null;
    } catch (error) {
      return false;
    }
  }

  static async login(): Promise<void> {
    try {
      const instance = SalesforceAuth.getInstance();
      const auth = await Auth.oauth2({
        loginUrl: process.env.SF_LOGIN_URL || 'https://login.salesforce.com',
        clientId: process.env.SF_CLIENT_ID || '',
        clientSecret: process.env.SF_CLIENT_SECRET || '',
        redirectUri: process.env.SF_REDIRECT_URI || 'https://localhost:3000/oauth/callback'
      });
      
      instance.connection = await Connection.create({ auth });
    } catch (error) {
      console.error('Salesforce login error:', error);
      throw error;
    }
  }

  static async getData(): Promise<any> {
    try {
      const instance = SalesforceAuth.getInstance();
      if (!instance.connection) {
        throw new Error('Not authenticated with Salesforce');
      }

      // Example: Fetch accounts
      const result = await instance.connection.query('SELECT Id, Name, Industry FROM Account LIMIT 10');
      return result.records;
    } catch (error) {
      console.error('Salesforce data fetch error:', error);
      throw error;
    }
  }
} 