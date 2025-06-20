import React from 'react'
import {Route, Routes} from "react-router-dom";
import { Toaster } from 'sonner'
import Test from './component/SpreadsheetModel'
import Test2 from './component/FormulaModel'
import Test3 from './component/FormulaWorking'
import Test4 from './component/Integration'
import SpreadsheetPage from './component/SpreadsheetPage';
import Test5 from './component/PowerBIDashboard';
import Test6 from './component/StrixChart';
import Menu from './component/Menubar';
import Test7 from './component/Image&Link';



const App = () => {
  return (
    <div>
      <div style={{ display: "flex", flexDirection: "column", minHeight: "100vh" }}>
      <div style={{ flex: "1" }}>
      <Routes>
        <Route path="/" element={<Test4/>}/>
        <Route path="/sheet" element={<SpreadsheetPage />} />
        <Route path="/charts" element={<Test6 />} />
        <Route path="/menu" element={<Menu />} />
        <Route path="/test" element={<Test/>}/>
        <Route path="/test2" element={<Test2/>}/>
        <Route path="/test3" element={<Test3/>}/>
        <Route path="/test5" element={<Test5/>}/>
        <Route path="/test6" element={<Test6/>}/>
        <Route path="/test7" element={<Test7/>}/>
      </Routes>
      </div>
      <Toaster richColors position="top-center" />
      </div>
    </div>
    
  )
}

export default App;