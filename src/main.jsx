import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import { SpreadsheetDataProvider } from './context/SpreadsheetDataContext';
import App from './App.jsx'
import { BrowserRouter } from 'react-router-dom'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <BrowserRouter 
      future={{
        v7_startTransition: true,
        v7_relativeSplatPath: true,
        v7_prependBasename: true
      }}
    >
      <SpreadsheetDataProvider>
        <App />
      </SpreadsheetDataProvider>
    </BrowserRouter>
  </StrictMode>
)