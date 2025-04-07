# Getting Started with Office 365 Salesforce Add-in

## Your Current Status

Based on the environment check:
- Node.js is not installed on your system
- You need to install Node.js before proceeding

## Complete Setup Process

### Step 1: Install Prerequisites

1. **Install Node.js**
   - Download and install from [https://nodejs.org/](https://nodejs.org/)
   - Choose the LTS (Long Term Support) version

2. **Verify Installation**
   - After installation, open a new terminal
   - Run `node -v` and `npm -v` to confirm installation

### Step 2: Set Up the Add-in

1. **Run the Installation Script**
   ```bash
   ./install.sh
   ```
   
   This script will:
   - Install all required dependencies
   - Generate SSL certificates for development
   - Prepare your environment

2. **Configure Salesforce**
   - Create a Connected App in Salesforce
     - Go to Setup > App Manager > New Connected App
     - Enable OAuth settings
     - Set callback URL to: https://localhost:3000/oauth/callback
     - Select scopes: api, refresh_token, offline_access
   
   - Update your `.env` file with:
     ```
     SF_LOGIN_URL=https://login.salesforce.com
     SF_CLIENT_ID=your_client_id_here
     SF_CLIENT_SECRET=your_client_secret_here
     SF_REDIRECT_URI=https://localhost:3000/oauth/callback
     ```

### Step 3: Run the Add-in

1. **Start the Development Server**
   ```bash
   npm run dev
   ```

2. **Launch the Add-in**
   - Open a new terminal window
   - Run:
     ```bash
     npm start
     ```
   - This will open your default Office application with the add-in

### Step 4: Using the Add-in

1. The add-in appears in the sidebar of your Office application
2. Click "Connect to Salesforce" to authenticate
3. Once connected, you can view and insert Salesforce data
4. Use the various insertion options (text, HTML, images)

## Troubleshooting

- **Certificate Issues**: Run `npx office-addin-dev-certs install --machine` to install certificates machine-wide
- **Authentication Failures**: Double-check your Salesforce credentials in the `.env` file
- **Add-in Not Loading**: Check browser console for errors (the add-in runs in a web environment)

## Development Notes

- Edit files in the `src` folder to modify functionality
- Changes are automatically applied when you save (hot reloading)
- Use browser developer tools to debug
- See `README.md` for full documentation

## Need Help?

- Check [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- Visit [Salesforce API Documentation](https://developer.salesforce.com/docs/) 