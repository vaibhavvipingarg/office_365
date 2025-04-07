# Office 365 Salesforce Add-in - Quickstart Guide

This is a specialized Office 365 add-in that integrates Salesforce data into Word, PowerPoint, and Outlook. Below are the essential steps to get you started.

## Project Structure Overview

The project uses:
- React for the UI components
- Office.js API for Office integration
- Salesforce API for data integration
- Webpack for bundling

## Quick Setup

1. **Prerequisites**:
   - Node.js & npm need to be installed first
   - Office 365 account
   - Salesforce developer account

2. **First-time setup**:
   ```bash
   # Install Node.js from https://nodejs.org/

   # Install dependencies
   npm install

   # Generate SSL certificates for development
   npx office-addin-dev-certs install
   ```

3. **Configure Salesforce**:
   - Create a Connected App in Salesforce
   - Update `.env` file with your Salesforce credentials

4. **Run the add-in**:
   ```bash
   # Start the development server
   npm run dev

   # In a new terminal, start the add-in
   npm start
   ```

## Features Implemented

- ✅ Salesforce authentication via OAuth
- ✅ Data retrieval from Salesforce
- ✅ Integration with Word document content
- ✅ Integration with PowerPoint slides
- ✅ Integration with Outlook compose window
- ✅ Interactive HTML/iframe insertion

## How to Use

1. Launch the add-in in your Office application
2. Click "Connect to Salesforce" to authenticate
3. Once connected, you can insert Salesforce content into your document
4. The content will be formatted according to the current context

## Troubleshooting

If you encounter issues:
- Check browser console for errors
- Verify Salesforce credentials
- Ensure SSL certificates are properly installed

## Next Steps

To extend this add-in:
- Add more Salesforce object types
- Customize the UI with your branding
- Implement caching for offline use
- Add advanced data visualization 