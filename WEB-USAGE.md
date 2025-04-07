# Using Salesforce Add-in with Office Web Apps

This guide will help you use the Salesforce add-in with Office web applications.

## Starting the Add-in

We've created specific commands for each Office web application:

```bash
# For Word Online
npm run start:word

# For PowerPoint Online
npm run start:powerpoint

# For Excel Online
npm run start:excel

# For Outlook Online
npm run start:outlook
```

Make sure your development server is running first:

```bash
npm run dev
```

## Authentication

1. When the add-in loads, you'll see a **Connect to Salesforce** button
2. Click this button to start the OAuth flow
3. A popup window will appear for Salesforce authentication
4. Log in with your Salesforce credentials
5. Allow the requested permissions
6. The popup will close and you'll be authenticated

## Features Available in Web Apps

### In Word Online
- Insert Salesforce data as text
- Add formatted HTML content from Salesforce
- Insert images linked from Salesforce

### In PowerPoint Online
- Add Salesforce data to slides
- Insert charts and visualizations
- Embed interactive content

### In Outlook Online
- View Salesforce contact information
- Add Salesforce content to emails
- Track emails in Salesforce

## Known Limitations in Web Apps

- Some Office.js APIs have limited functionality in web versions
- Performance may differ from desktop versions
- Popup blockers can interfere with the OAuth flow
- Browsers have restrictions on cross-origin requests

## Troubleshooting Web-Specific Issues

1. **Add-in doesn't load**
   - Make sure you're signed in to Office Online
   - Check if your browser is blocking popups
   - Verify the development server is running

2. **Authentication fails**
   - Ensure your Salesforce credentials are correct
   - Check that your Salesforce Connected App configuration is correct
   - Verify the redirect URI matches what's in your .env file

3. **Content doesn't insert properly**
   - Different Office web apps have different capabilities
   - Try simpler content types first (plain text)
   - Check browser console for errors

4. **Security Warnings**
   - Accept certificates when prompted
   - Use a modern browser (Chrome, Edge, Firefox) 