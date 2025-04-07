import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './components/App';
import './index.css';

// Set up the app in standalone mode
console.log('Starting app in standalone mode');

// Initialize the application
const container = document.getElementById('app-container');
if (container) {
  const root = createRoot(container);
  root.render(<App isLocalMode={true} />);
} else {
  console.error("Container element with id 'app-container' not found");
} 