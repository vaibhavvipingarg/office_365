# Using the Add-in with Office Web Apps

Since you don't have the desktop version of Microsoft Office installed, we'll set up the add-in to work with Office Online (Word Online, PowerPoint Online, etc.).

## Setup for Office Web Apps

1. **Make sure the development server is running**:
   ```bash
   npm run dev
   ```

2. **Start the add-in in web mode**:
   ```bash
   npm run start:web
   ```

3. **Select the Office application**:
   When prompted, select which Office application you want to use (Word, PowerPoint, etc.)

4. **This will open your browser with Office Online**:
   You'll need to sign in with your Microsoft account or Office 365 account

## Manual Sideloading (Alternative Method)

If the automated method doesn't work, you can manually sideload the add-in:

1. Go to [Office.com](https://www.office.com/)

2. Sign in with your Microsoft account

3. Open Word Online, PowerPoint Online, or Outlook Online

4. For Word/PowerPoint:
   - Create a new document
   - Click on "Insert" tab
   - Click "Add-ins" or "Office Add-ins"
   - Choose "Upload My Add-in"
   - Select the manifest.xml file from this project

5. For Outlook:
   - Click the gear icon (Settings)
   - Click "Manage add-ins"
   - Click "Add from file..."
   - Select the manifest.xml file

## Troubleshooting

- Make sure your development server is running at https://localhost:3000
- Ensure you're signed in with a Microsoft account that supports add-ins
- If you get certificate warnings in the browser, accept the self-signed certificate
- If the add-in doesn't load, check the browser console for errors

## Testing Features

When using the web versions:

1. After the add-in loads, the "Show Salesforce" button should appear in the ribbon
2. Click it to open the add-in panel
3. Click "Connect to Salesforce" to authenticate
4. Once connected, you can insert Salesforce data into your document

Note that some features may behave slightly differently in the web version compared to the desktop version of Office. 