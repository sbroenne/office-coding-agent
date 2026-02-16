import React from 'react';
import { createRoot } from 'react-dom/client';
import '../styles/globals.css';
import { App } from './App';

/* global Office */

void Office.onReady(() => {
  const container = document.getElementById('root');
  if (!container) {
    throw new Error('Root element not found');
  }

  const root = createRoot(container);
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
});
