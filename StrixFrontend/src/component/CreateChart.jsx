import React, { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell, 
  ScatterChart, Scatter, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, 
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from 'recharts';
import { 
  Upload, Download, Settings, BarChart2, LineChart as LineChartIcon, 
  PieChart as PieChartIcon, ScatterChart as ScatterChartIcon, AreaChart as AreaChartIcon,
  Gauge as GaugeIcon, Target, DollarSign, Award, AlertTriangle,
  CheckCircle, Layers, ChevronsUp, ChevronsDown, TrendingUp, TrendingDown,
  Activity, Grid, Table, Database, Sliders, Filter, FileText
} from 'lucide-react';

const CreateChart = () => {
  const navigate = useNavigate();
  
  // State management
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [chartType, setChartType] = useState('bar');
  const [fileName, setFileName] = useState('Sample Data');
  const [dataView, setDataView] = useState('chart');
  const [filters, setFilters] = useState({});
  
  const [chartConfig, setChartConfig] = useState({
    xAxis: 'category',
    yAxis: ['sales', 'profit'],
    colors: ['#8884d8', '#82ca9d']
  });

  // Sample data
  const sampleData = [
    { category: 'Jan', sales: 2300, profit: 3900 }
  ];

  // Initialize with sample data
  React.useEffect(() => {
    setData(sampleData);
    setColumns(['category', 'sales', 'profit']);
  }, []);

  // Chart type categories
  const chartCategories = {
    'Cards': [
      { type: 'number', name: 'Number Card', icon: <Target size={16} /> },
      { type: 'kpi', name: 'KPI Card', icon: <Award size={16} /> },
      { type: 'gauge', name: 'Gauge Card', icon: <GaugeIcon size={16} /> },
      { type: 'metric', name: 'Metric Card', icon: <DollarSign size={16} /> },
      { type: 'trend', name: 'Trend Card', icon: <TrendingUp size={16} /> },
      { type: 'status', name: 'Status Card', icon: <CheckCircle size={16} /> }
    ],
    'Bar/Column': [
      { type: 'bar', name: 'Bar Chart', icon: <BarChart2 size={16} /> },
      { type: 'stacked-bar', name: 'Stacked Bar', icon: <Layers size={16} /> },
      { type: 'horizontal-bar', name: 'Horizontal Bar', icon: <ChevronsUp size={16} /> },
      { type: 'grouped-bar', name: 'Grouped Bar', icon: <Grid size={16} /> },
      { type: 'waterfall', name: 'Waterfall Chart', icon: <ChevronsDown size={16} /> }
    ],
    'Line/Area': [
      { type: 'line', name: 'Line Chart', icon: <LineChartIcon size={16} /> },
      { type: 'multi-line', name: 'Multi-Line Chart', icon: <Activity size={16} /> },
      { type: 'area', name: 'Area Chart', icon: <AreaChartIcon size={16} /> },
      { type: 'stacked-area', name: 'Stacked Area', icon: <Layers size={16} /> }
    ],
    'Pie/Doughnut': [
      { type: 'pie', name: 'Pie Chart', icon: <PieChartIcon size={16} /> },
      { type: 'doughnut', name: 'Doughnut Chart', icon: <PieChartIcon size={16} /> },
      { type: 'radial-bar', name: 'Radial Bar', icon: <Target size={16} /> }
    ],
    'Scatter/Bubble': [
      { type: 'scatter', name: 'Scatter Plot', icon: <ScatterChartIcon size={16} /> },
      { type: 'bubble', name: 'Bubble Chart', icon: <Activity size={16} /> }
    ],
    'Special': [
      { type: 'radar', name: 'Radar Chart', icon: <Target size={16} /> },
      { type: 'treemap', name: 'Treemap', icon: <Grid size={16} /> },
      { type: 'sunburst', name: 'Sunburst', icon: <PieChartIcon size={16} /> },
      { type: 'sankey', name: 'Sankey', icon: <Activity size={16} /> },
      { type: 'funnel', name: 'Funnel Chart', icon: <ChevronsDown size={16} /> },
      { type: 'heatmap', name: 'Heatmap', icon: <Grid size={16} /> },
      { type: 'gauge', name: 'Gauge Chart', icon: <GaugeIcon size={16} /> }
    ]
  };

  // Handle file upload
  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      // Here you would process the file and update data/columns
    }
  };

  // Handle filter changes
  const handleFilterChange = (column, value) => {
    setFilters(prev => ({
      ...prev,
      [column]: value
    }));
  };

  // Handle chart configuration changes
  const handleConfigChange = (type, value) => {
    setChartConfig(prev => ({
      ...prev,
      [type]: value
    }));
  };

  // Render chart based on type
  const renderChart = () => {
    const { xAxis, yAxis, colors } = chartConfig;
    
    switch (chartType) {
      case 'bar':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart data={data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {yAxis.map((axis, index) => (
                <Bar 
                  key={axis} 
                  dataKey={axis} 
                  fill={colors[index % colors.length]} 
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        );
      case 'line':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <LineChart data={data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {yAxis.map((axis, index) => (
                <Line 
                  key={axis} 
                  type="monotone" 
                  dataKey={axis} 
                  stroke={colors[index % colors.length]} 
                />
              ))}
            </LineChart>
          </ResponsiveContainer>
        );
      case 'pie':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <PieChart>
              <Pie
                data={data}
                dataKey={yAxis[0]}
                nameKey={xAxis}
                cx="50%"
                cy="50%"
                outerRadius={80}
                fill="#8884d8"
              >
                {data.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        );
      default:
        return (
          <div className="flex items-center justify-center h-96 text-gray-500">
            <p>Select a chart type to preview</p>
          </div>
        );
    }
  };

  const handleLogout = () => {
    localStorage.clear();
    window.location.href = '/';
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8 flex justify-between items-center">
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
            <div>
              <h1 className="text-2xl font-bold text-gray-900 flex items-center">
                <BarChart2 className="mr-3 text-blue-600" size={28} />
                Strix Charts
              </h1>
              <p className="text-gray-600 mt-1">Transform your data into actionable insights</p>
            </div>
          </div>
          <div className="flex items-center space-x-4">
            <button
              className="inline-flex items-center px-4 py-2 border border-gray-300 shadow-sm text-sm font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50"
            >
              <Download className="mr-2" size={16} />
              Export
            </button>
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
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 py-6 sm:px-6 lg:px-8">
        {/* Top Control Section */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
          {/* Data Source */}
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
            <h3 className="text-lg font-semibold text-gray-900 mb-3 flex items-center">
              <Upload className="mr-2 text-blue-500" size={20} />
              Data Source
            </h3>
            <label className="block w-full px-4 py-3 bg-blue-50 border-2 border-dashed border-blue-200 rounded-lg cursor-pointer hover:bg-blue-100 text-center">
              <input
                type="file"
                className="hidden"
                accept=".csv,.xlsx,.xls"
                onChange={handleFileUpload}
              />
              <div className="flex flex-col items-center">
                <Upload className="text-blue-500 mb-2" size={24} />
                <span className="text-blue-700 font-medium">Choose File</span>
                <p className="text-xs text-gray-500 mt-1">{fileName}</p>
              </div>
            </label>
            <p className="text-xs text-gray-500 text-center mt-2">
              Supports CSV, Excel files. Max 10MB.
            </p>
          </div>

          {/* Data Configuration */}
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200 lg:col-span-2">
            <h3 className="text-lg font-semibold text-gray-900 mb-3 flex items-center">
              <Sliders className="mr-2 text-blue-500" size={20} />
              Data Configuration
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* X-Axis */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">X-Axis (Category)</label>
                <div className="flex flex-wrap gap-2">
                  {columns.map((col) => (
                    <button
                      key={col}
                      onClick={() => handleConfigChange('xAxis', col)}
                      className={`px-3 py-1 rounded-full text-sm font-medium transition-colors ${
                        chartConfig.xAxis === col
                          ? 'bg-blue-100 text-blue-800 border border-blue-200'
                          : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                      }`}
                    >
                      {chartConfig.xAxis === col && <CheckCircle className="inline w-3 h-3 mr-1" />}
                      {col}
                    </button>
                  ))}
                </div>
              </div>

              {/* Y-Axis */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Y-Axis (Values)</label>
                <div className="flex flex-wrap gap-2">
                  {columns.map((col) => (
                    <button
                      key={col}
                      onClick={() => {
                        const newYAxis = chartConfig.yAxis.includes(col)
                          ? chartConfig.yAxis.filter(y => y !== col)
                          : [...chartConfig.yAxis, col];
                        handleConfigChange('yAxis', newYAxis);
                      }}
                      className={`px-3 py-1 rounded-full text-sm font-medium transition-colors ${
                        chartConfig.yAxis.includes(col)
                          ? 'bg-blue-100 text-blue-800 border border-blue-200'
                          : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                      }`}
                    >
                      {chartConfig.yAxis.includes(col) && <CheckCircle className="inline w-3 h-3 mr-1" />}
                      {col}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Chart Type Sidebar */}
          <div className="lg:col-span-1">
            <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                <BarChart2 className="mr-2 text-blue-500" size={20} />
                Chart Type
              </h3>
              
              {Object.entries(chartCategories).map(([category, charts]) => (
                <div key={category} className="mb-6">
                  <h4 className="text-sm font-medium text-gray-700 mb-2">{category}</h4>
                  <div className="space-y-1">
                    {charts.map((chart) => (
                      <button
                        key={chart.type}
                        onClick={() => setChartType(chart.type)}
                        className={`w-full flex items-center px-3 py-2 text-sm rounded-md transition-colors ${
                          chartType === chart.type
                            ? 'bg-blue-100 text-blue-800 border border-blue-200'
                            : 'text-gray-700 hover:bg-gray-100'
                        }`}
                      >
                        {chart.icon}
                        <span className="ml-2">{chart.name}</span>
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Main Chart Area */}
          <div className="lg:col-span-3">
            <div className="bg-white rounded-lg shadow border border-gray-200">
              {/* Tabs */}
              <div className="border-b border-gray-200">
                <div className="flex items-center justify-between px-4 py-3">
                  <div className="flex space-x-8">
                    <button
                      onClick={() => setDataView('chart')}
                      className={`flex items-center px-3 py-2 text-sm font-medium border-b-2 transition-colors ${
                        dataView === 'chart'
                          ? 'border-blue-500 text-blue-600'
                          : 'border-transparent text-gray-500 hover:text-gray-700'
                      }`}
                    >
                      <BarChart2 className="mr-2" size={16} />
                      Chart View
                    </button>
                    <button
                      onClick={() => setDataView('table')}
                      className={`flex items-center px-3 py-2 text-sm font-medium border-b-2 transition-colors ${
                        dataView === 'table'
                          ? 'border-blue-500 text-blue-600'
                          : 'border-transparent text-gray-500 hover:text-gray-700'
                      }`}
                    >
                      <Table className="mr-2" size={16} />
                      Data Table
                    </button>
                  </div>
                  <button className="flex items-center px-3 py-2 text-sm text-gray-600 hover:text-gray-800">
                    <Settings className="mr-2" size={16} />
                    Settings
                  </button>
                </div>
              </div>

              {/* Content */}
              <div className="p-6">
                {dataView === 'chart' ? (
                  <div>
                    {/* Filters */}
                    <div className="mb-6 flex gap-4">
                      {columns.map((col) => (
                        <div key={col} className="flex-1">
                          <label className="block text-sm font-medium text-gray-700 mb-1">
                            {col}
                          </label>
                          <input
                            type="text"
                            placeholder={`Filter ${col}...`}
                            value={filters[col] || ''}
                            onChange={(e) => handleFilterChange(col, e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                      ))}
                    </div>

                    {/* Chart */}
                    <div className="bg-gray-50 rounded-lg p-4">
                      {renderChart()}
                    </div>
                  </div>
                ) : (
                  <div>
                    <div className="overflow-x-auto">
                      <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50">
                          <tr>
                            {columns.map((col) => (
                              <th key={col} className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                                {col}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                          {data.map((row, index) => (
                            <tr key={index}>
                              {columns.map((col) => (
                                <td key={col} className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                  {row[col]}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Data Summary */}
        <div className="mt-6">
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
            <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
              <FileText className="mr-2 text-blue-500" size={20} />
              Data Summary
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="bg-blue-50 p-3 rounded-lg">
                <p className="text-sm text-blue-800">Total Records</p>
                <p className="text-xl font-bold text-blue-900">{data.length}</p>
              </div>
              <div className="bg-green-50 p-3 rounded-lg">
                <p className="text-sm text-green-800">Columns</p>
                <p className="text-xl font-bold text-green-900">{columns.length}</p>
              </div>
              <div className="bg-purple-50 p-3 rounded-lg">
                <p className="text-sm text-purple-800">File</p>
                <p className="text-sm font-medium text-purple-900 truncate">{fileName}</p>
              </div>
              <div className="bg-yellow-50 p-3 rounded-lg">
                <p className="text-sm text-yellow-800">Last Updated</p>
                <p className="text-sm font-medium text-yellow-900">
                  {new Date().toLocaleString()}
                </p>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
};

export default CreateChart;
