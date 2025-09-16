import React from 'react';
import { Routes, Route } from "react-router-dom"; 
import { Toaster } from 'sonner';
import { AuthProvider } from './context/AuthContext';
import { SpreadsheetProvider } from './context/SpreadsheetContext';
import ProtectedRoute from './components/ProtectedRoute';
import SpreadsheetModel from './component/SpreadsheetModel';
import SpreadsheetDashboard from './component/SpreadsheetDashboard';
import PowerBIDashboard from './component/PowerBIDashboard';
import StrixChart from './component/StrixChart';
import Menu from './component/Menubar';
import ImageLink from './component/Image&Link';
import BIDashboard from './component/BIDashboard';
import CreateChart from './component/CreateChart';
import StrixAuth from './component/StrixAuth';
import ForgotPassword from "./component/ForgetPassword";

const App = () => {
  return (
    <AuthProvider>
      <SpreadsheetProvider>
        <div>
          <div style={{ display: "flex", flexDirection: "column", minHeight: "100vh" }}>
            <div style={{ flex: "1" }}>
              <Routes>
                {/* Public Routes */}
                <Route path="/" element={<StrixAuth/>}/>
                <Route path="/forgot-password" element={<ForgotPassword/>} />
                
                {/* Protected Routes */}
                <Route path="/integration" element={
                  <ProtectedRoute>
                    <SpreadsheetModel />
                  </ProtectedRoute>
                } />
                <Route path="/spreadsheet-dashboard" element={
                  <ProtectedRoute>
                    <SpreadsheetDashboard />
                  </ProtectedRoute>
                } />
                <Route path="/dashboard" element={
                  <ProtectedRoute>
                    <BIDashboard/>
                  </ProtectedRoute>
                }/>
                <Route path="/sheet" element={
                  <ProtectedRoute>
                    <SpreadsheetModel />
                  </ProtectedRoute>
                } />
                <Route path="/charts" element={
                  <ProtectedRoute>
                    <StrixChart />
                  </ProtectedRoute>
                } />
                <Route path="/create-chart" element={
                  <ProtectedRoute>
                    <CreateChart />
                  </ProtectedRoute>
                } />
                <Route path="/menu" element={
                  <ProtectedRoute>
                    <Menu />
                  </ProtectedRoute>
                } />
                <Route path="/spreadsheet-model" element={
                  <ProtectedRoute>
                    <SpreadsheetModel/>
                  </ProtectedRoute>
                }/>
                <Route path="/powerbi" element={
                  <ProtectedRoute>
                    <PowerBIDashboard/>
                  </ProtectedRoute>
                }/>
                <Route path="/image-link" element={
                  <ProtectedRoute>
                    <ImageLink/>
                  </ProtectedRoute>
                }/>
              </Routes>
            </div>
            <Toaster richColors position="top-center" />
          </div>
        </div>
      </SpreadsheetProvider>
    </AuthProvider>
  )
}

export default App;