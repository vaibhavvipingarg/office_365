# Office 365 Salesforce Integration Add-in

This Office 365 add-in provides seamless integration with Salesforce, allowing you to access and insert Salesforce data into Word, PowerPoint, and Outlook documents.

## 🚀 Features

- **Salesforce OAuth Integration**: Securely authenticate with your Salesforce account
- **Cross-Platform Support**: Works with Word, PowerPoint, and Outlook
- **Dynamic Content**: Insert text, HTML, and images from Salesforce
- **Modern UI**: Clean, responsive interface using Fluent UI
- **Interactive Elements**: Add HTML/iFrames to enhance your documents

## 📋 Prerequisites

- Node.js (v14 or later) & npm
- Office 365 account
- Salesforce account with API access

## 🛠️ Setup Options

### Option 1: Easy Setup (Requires Node.js)

Run the installation script:

```bash
./install.sh
```

### Option 2: Manual Setup

1. Install dependencies:
   ```bash
   npm install
   ```

2. Generate SSL certificates:
   ```bash
   npx office-addin-dev-certs install
   ```

3. Configure Salesforce:
   - Edit the `.env` file with your Salesforce credentials

## 🏃‍♂️ Running the Add-in

### Local Development Mode (No Office Required)

You can develop and test the add-in UI without Office applications:

```bash
npm run local
```

This will open the add-in UI in your browser for faster development.
See [LOCAL-DEVELOPMENT.md](./LOCAL-DEVELOPMENT.md) for more details.

### For Office Web Apps (Online)

1. Start the development server:
   ```bash
   npm run dev
   ```

2. In a new terminal, launch the add-in in your preferred Office web app:
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

See [WEB-USAGE.md](./WEB-USAGE.md) for detailed instructions on using the add-in with Office web apps.

### For Office Desktop Apps

1. Start the development server:
   ```bash
   npm run dev
   ```

2. In a new terminal, start the add-in:
   ```bash
   npm run start:desktop
   ```

## 🧰 Using the Add-in

1. Open Word, PowerPoint, or Outlook (desktop or online)
2. Click the "Show Salesforce" button in the ribbon
3. Connect to Salesforce when prompted
4. Browse and insert Salesforce data into your document

## 🔄 Development Workflow

- **Source Code**: All code is in the `src` directory
- **Build Process**: Webpack bundles the code
- **Debugging**: Use browser developer tools
- **Manifest**: `manifest.xml` defines the add-in's properties

## 📁 Project Structure

```
.
├── assets/            # Static assets and icons
├── certs/             # SSL certificates for development
├── src/
│   ├── components/    # React components
│   ├── services/      # Service modules (Salesforce, Office)
│   ├── types/         # TypeScript type definitions
│   ├── index.tsx      # Main entry point
│   └── commands.ts    # Add-in commands
├── .babelrc           # Babel configuration
├── .env               # Environment variables
├── manifest.xml       # Add-in manifest
├── package.json       # NPM dependencies
├── tsconfig.json      # TypeScript configuration
└── webpack.config.js  # Webpack configuration
```

## 🔒 Security Considerations

- All communication uses HTTPS
- OAuth 2.0 for secure authentication
- No storage of credentials in the add-in

## 🛣️ Roadmap

- Advanced data visualization
- Offline capability
- Custom templates
- Batch operations

## 📚 Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Salesforce API Documentation](https://developer.salesforce.com/docs/atlas.en-us.api_rest.meta/api_rest/)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui#/controls/web)

## 📝 License

MIT 