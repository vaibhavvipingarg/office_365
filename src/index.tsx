import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './components/App';
import './index.css';

// Check if we're in local development mode
const isLocalMode = !window.hasOwnProperty('Office');

// Function to initialize the application
const initializeApp = () => {
  // In local mode, use app-container; in Office mode, use root
  const containerId = isLocalMode ? 'app-container' : 'root';
  const container = document.getElementById(containerId);
  
  if (container) {
    console.log(`Initializing app in ${isLocalMode ? 'local' : 'Office'} mode`);
    const root = createRoot(container);
    root.render(<App isLocalMode={isLocalMode} />);
  } else {
    console.error(`Container element with id '${containerId}' not found`);
  }
};

// If Office.js is available, wait for it to initialize
if (!isLocalMode) {
  // Office.js is loaded from CDN in index.html
  console.log('Waiting for Office.js to initialize...');
  Office.onReady(() => {
    console.log('Office.js initialized');
    initializeApp();
  });
} else {
  // In local mode, initialize immediately
  console.log('Running in local development mode (without Office.js)');
  initializeApp();
} 