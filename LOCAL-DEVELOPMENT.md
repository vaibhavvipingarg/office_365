# Local Development Mode

This document explains how to develop and test the Salesforce add-in UI without connecting to Office applications.

## Running in Local Mode

To run the add-in in local development mode:

```bash
# Start the development server in local mode
npm run local
```

This will:
1. Start the development server
2. Open the add-in UI in your default browser
3. Load the UI without Office.js dependencies
4. Display a simulated Office task pane

## Benefits of Local Development

- **Faster Development**: No need to wait for Office applications to load
- **Easier Debugging**: Directly use browser dev tools
- **Focus on UI**: Develop the interface without Office.js integration concerns
- **Test Anywhere**: No need for Office 365 subscription during UI development

## How Local Mode Works

When in local mode:
- The app detects the absence of Office.js and switches to local mode
- Salesforce authentication still works normally
- Instead of inserting content into Office documents, data is displayed in the UI
- Office-specific API calls are mocked

## Testing Features

1. **Authentication**
   - Click "Connect to Salesforce" to test the authentication flow
   - The mock implementation will simulate successful authentication

2. **Data Retrieval**
   - Once authenticated, click "Insert Salesforce Content"
   - In local mode, this will display the mock data in the UI
   - This lets you verify data formatting and display logic

3. **UI Components**
   - All UI components should render and behave normally
   - Styling should match the Office task pane dimensions

## Limitations

- Office-specific APIs are not available
- Document insertion cannot be fully tested
- Office context-specific behaviors need to be tested in the real Office environment

## Switching Between Modes

- Use `npm run local` for local UI development
- Use `npm run start:word` (or other Office apps) for full Office integration testing

## Browser Support

The local development mode works in any modern browser:
- Chrome (recommended for development)
- Firefox
- Safari
- Edge 