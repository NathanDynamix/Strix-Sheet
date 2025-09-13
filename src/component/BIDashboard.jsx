import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  BarChart, 
  Filter, 
  Plus, 
  Grid3X3,
  Upload,
  ChevronRight
} from 'lucide-react';

const BIDashboard = () => {
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState('dashboard');
  const [chartType, setChartType] = useState('bar');
  const [xAxisColumn, setXAxisColumn] = useState('');
  const [yAxisColumn, setYAxisColumn] = useState('');

  const handleLogout = () => {
    localStorage.clear();
    window.location.href = '/';
  };

  return (
    <div className="h-screen flex flex-col bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b border-gray-200">
        {/* Top Header Bar */}
        <div className="flex items-center justify-between px-4 py-2 bg-white border-b border-gray-200">
          <div className="flex items-center space-x-4">
            <button
              onClick={() => navigate('/integration')}
              className="p-2 hover:bg-gray-100 rounded-md transition-colors"
              title="Back to Spreadsheet"
            >
              <svg className="w-5 h-5 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 19l-7-7m0 0l7-7m-7 7h18" />
              </svg>
            </button>
            <h1 className="text-2xl font-bold bg-gradient-to-r from-indigo-600 to-purple-600 bg-clip-text text-transparent">
              Strix Sheets
            </h1>
          </div>
          
          {/* Logout Button */}
          <button 
            onClick={handleLogout}
            className="px-4 py-2 bg-red-600 text-white rounded-md text-sm font-medium hover:bg-red-700 flex items-center space-x-2"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 19l-7-7m0 0l7-7m-7 7h18" />
            </svg>
            <span>Logout</span>
          </button>
        </div>
      </div>

      {/* Main Dashboard Layout */}
      <div className="flex flex-1">
        {/* Left Sidebar */}
        <div className="w-80 bg-white border-r border-gray-200 flex flex-col">
          {/* BI Dashboard Header */}
          <div className="p-4 border-b border-gray-200">
            <div className="flex items-center space-x-2 mb-1">
              <BarChart size={20} className="text-blue-600" />
              <h2 className="text-lg font-semibold text-gray-800">BI Dashboard</h2>
            </div>
            <p className="text-sm text-gray-600">Advanced Business Intelligence</p>
          </div>

          {/* Upload Data Section */}
          <div className="p-4 border-b border-gray-200">
            <button className="w-full bg-blue-50 text-blue-600 px-4 py-3 rounded-lg border border-blue-200 hover:bg-blue-100 transition-colors flex items-center justify-center space-x-2">
              <Upload size={20} />
              <span className="font-medium">Upload Data</span>
            </button>
          </div>

          {/* Chart Builder Section */}
          <div className="p-4 flex-1">
            <div className="flex items-center space-x-2 mb-4">
              <Plus size={16} className="text-gray-600" />
              <h3 className="font-medium text-gray-800">Chart Builder</h3>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Chart Type</label>
                <select 
                  value={chartType}
                  onChange={(e) => setChartType(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="bar">Bar Chart</option>
                  <option value="line">Line Chart</option>
                  <option value="pie">Pie Chart</option>
                  <option value="area">Area Chart</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">X-Axis</label>
                <select 
                  value={xAxisColumn}
                  onChange={(e) => setXAxisColumn(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">Select Column</option>
                  {Array.from({ length: 40 }, (_, i) => {
                    const columnName = String.fromCharCode(65 + i);
                    return (
                      <option key={i} value={columnName}>
                        {columnName}
                      </option>
                    );
                  })}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Y-Axis</label>
                <select 
                  value={yAxisColumn}
                  onChange={(e) => setYAxisColumn(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">Select Column</option>
                  {Array.from({ length: 40 }, (_, i) => {
                    const columnName = String.fromCharCode(65 + i);
                    return (
                      <option key={i} value={columnName}>
                        {columnName}
                      </option>
                    );
                  })}
                </select>
              </div>

              <button className="w-full bg-blue-600 text-white px-4 py-2 rounded-md hover:bg-blue-700 transition-colors font-medium">
                + Add Chart
              </button>
            </div>
          </div>

          {/* Formula Bar (Bottom) */}
          <div className="p-4 border-t border-gray-200">
            <div className="flex items-center space-x-2 text-sm text-gray-600">
              <Grid3X3 size={16} />
              <span>Formula Bar</span>
              <ChevronRight size={16} className="ml-auto" />
            </div>
          </div>
        </div>

        {/* Main Content */}
        <div className="flex-1 flex flex-col">
          {/* Tabs */}
          <div className="bg-white border-b border-gray-200">
            <div className="flex">
              <button
                onClick={() => setActiveTab('dashboard')}
                className={`px-6 py-3 text-sm font-medium border-b-2 ${
                  activeTab === 'dashboard'
                    ? 'border-blue-500 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-gray-700'
                }`}
              >
                Dashboard
              </button>
              <button
                onClick={() => setActiveTab('data')}
                className={`px-6 py-3 text-sm font-medium border-b-2 ${
                  activeTab === 'data'
                    ? 'border-blue-500 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-gray-700'
                }`}
              >
                Data
              </button>
            </div>
          </div>

          {/* Main Content Area */}
          <div className="flex-1 flex items-center justify-center bg-gray-50">
            <div className="text-center">
              <div className="w-16 h-16 bg-gray-200 rounded-lg flex items-center justify-center mx-auto mb-4">
                <BarChart size={32} className="text-gray-400" />
              </div>
              <h3 className="text-lg font-medium text-gray-900 mb-2">No charts yet</h3>
              <p className="text-gray-500">Use the Chart Builder to create visualizations</p>
            </div>
          </div>
        </div>

        {/* Right Sidebar */}
        <div className="w-80 bg-white border-l border-gray-200">
          <div className="p-4">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center space-x-2">
                <Filter size={16} className="text-gray-600" />
                <h3 className="font-medium text-gray-800">Smart Filters</h3>
              </div>
              <button className="text-sm text-blue-600 hover:text-blue-700">Clear All</button>
            </div>
            
            <div className="text-center py-8">
              <p className="text-gray-500 text-sm">No filters applied</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default BIDashboard;
