import React from 'react';
import { Routes, Route } from "react-router-dom"; 
import { Toaster } from 'sonner';
import { AuthProvider } from './context/AuthContext';
import { SpreadsheetProvider } from './context/SpreadsheetContext';
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
                <Route path="/" element={<StrixAuth/>}/>
                <Route path="/integration" element={<SpreadsheetModel />} />
                <Route path="/spreadsheet-dashboard" element={<SpreadsheetDashboard />} />
                <Route path="/dashboard" element={<BIDashboard/>}/>
                <Route path="/sheet" element={<SpreadsheetModel />} />
                <Route path="/charts" element={<StrixChart />} />
                <Route path="/create-chart" element={<CreateChart />} />
                <Route path="/menu" element={<Menu />} />
                <Route path="/forgot-password" element={<ForgotPassword/>} />
                <Route path="/spreadsheet-model" element={<SpreadsheetModel/>}/>
                <Route path="/powerbi" element={<PowerBIDashboard/>}/>
                <Route path="/image-link" element={<ImageLink/>}/>
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