# Setup Instructions for Office 365 Salesforce Add-in

Before you can run the add-in, you need to set up your development environment.

## Prerequisites Installation

1. **Install Node.js and npm**:
   - Visit [Node.js website](https://nodejs.org/) and download the LTS version
   - Follow the installation instructions for your operating system (macOS)
   - After installation, verify by opening a new terminal and running:
     ```
     node -v
     npm -v
     ```

2. **Install dependencies**:
   - Once Node.js is installed, navigate to the project directory in your terminal
   - Run the following command to install all dependencies:
     ```
     npm install
     ```

3. **Generate SSL certificates for development**:
   - After installing dependencies, run:
     ```
     npx office-addin-dev-certs install
     ```
   - This will create the necessary certificates in the `./certs` directory

4. **Configure Salesforce**:
   - Create a Connected App in your Salesforce Developer account
   - Update the `.env` file with your Salesforce client credentials
   - Make sure you set the correct redirect URI in your Salesforce Connected App settings

## Running the Add-in

1. **Start the development server**:
   ```
   npm run dev
   ```

2. **In a separate terminal, start the add-in**:
   ```
   npm start
   ```

3. This will launch your default Office application (Word, PowerPoint, or Outlook)
   with the add-in loaded in the sidebar.

## Troubleshooting

- If you encounter certificate issues, make sure you've correctly installed the dev certificates
- If you have issues connecting to Salesforce, verify your client ID and secret in the `.env` file
- For more complex debugging, check the browser console (the add-in runs in a browser environment)

## Next Steps

Once you have the development environment set up:

1. Test the basic authentication flow with Salesforce
2. Customize the data retrieval to match your specific Salesforce objects
3. Enhance the UI to match your application's design
4. Implement additional features like caching or offline support 