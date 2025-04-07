#!/bin/bash

echo "Office 365 Salesforce Add-in Setup"
echo "=================================="

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install it from https://nodejs.org/"
    exit 1
fi

# Check if npm is installed
if ! command -v npm &> /dev/null; then
    echo "❌ npm is not installed. It should come with Node.js."
    exit 1
fi

# Display Node.js and npm versions
echo "✅ Node.js version: $(node -v)"
echo "✅ npm version: $(npm -v)"

# Install dependencies
echo "📦 Installing dependencies..."
npm install

# Check if installation was successful
if [ $? -ne 0 ]; then
    echo "❌ Failed to install dependencies."
    exit 1
fi

# Generate SSL certificates
echo "🔐 Generating SSL certificates for development..."
npx office-addin-dev-certs install

# Check if certificate generation was successful
if [ $? -ne 0 ]; then
    echo "❌ Failed to generate SSL certificates."
    echo "You may need to run this command manually: npx office-addin-dev-certs install"
    exit 1
fi

echo "✅ Setup completed successfully!"
echo ""
echo "Next steps:"
echo "1. Update your .env file with your Salesforce credentials"
echo "2. Start the development server: npm run dev"
echo "3. In a separate terminal, start the add-in: npm start"
echo ""
echo "Happy coding! 🚀"

chmod +x install.sh 