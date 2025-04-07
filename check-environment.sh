#!/bin/bash

echo "Checking environment for Office 365 Salesforce Add-in"
echo "==================================================="

# Check system information
echo "System: $(uname -s)"
echo "Architecture: $(uname -m)"

# Check for Node.js
if command -v node &> /dev/null; then
    echo "✅ Node.js is installed: $(node -v)"
else
    echo "❌ Node.js is not installed"
    echo "Please install Node.js from https://nodejs.org/"
fi

# Check for npm
if command -v npm &> /dev/null; then
    echo "✅ npm is installed: $(npm -v)"
else
    echo "❌ npm is not installed"
fi

# Check for Office applications
echo "Checking for Office applications..."

if [ -d "/Applications/Microsoft Word.app" ]; then
    echo "✅ Microsoft Word is installed"
else
    echo "❓ Microsoft Word not found in standard location"
fi

if [ -d "/Applications/Microsoft PowerPoint.app" ]; then
    echo "✅ Microsoft PowerPoint is installed"
else
    echo "❓ Microsoft PowerPoint not found in standard location"
fi

if [ -d "/Applications/Microsoft Outlook.app" ]; then
    echo "✅ Microsoft Outlook is installed"
else
    echo "❓ Microsoft Outlook not found in standard location"
fi

echo ""
echo "Next steps:"
echo "1. Install Node.js if not already installed"
echo "2. Run './install.sh' to set up the add-in"
echo "3. Update .env with your Salesforce credentials"
echo "4. Run 'npm run dev' and 'npm start' to launch the add-in" 