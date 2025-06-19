import React from 'react'
import {Route, Routes} from "react-router-dom";
import { Toaster } from 'sonner'
import Test from './Pages/Test'
import Test2 from './Pages/Test2'
import Test3 from './Pages/Test3'
import Test4 from './Pages/Integration'
import SpreadsheetPage from './Pages/SpreadsheetPage';
import Test5 from './Pages/Test5';
import Test6 from './Pages/Test6';
import Menu from './Pages/Menubar';



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
      </Routes>
      </div>
      <Toaster richColors position="top-center" />
      </div>
    </div>
    
  )
}

export default App;