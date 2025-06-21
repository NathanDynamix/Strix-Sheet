import React, { useState, useEffect, useMemo ,useRef} from 'react';
import html2canvas from 'html2canvas';
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell, 
  ScatterChart, Scatter, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, 
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, FunnelChart, Funnel, LabelList, Treemap, RadialBarChart, 
  RadialBar
} from 'recharts';
import { Sankey } from 'recharts';
import { ResponsiveSunburst } from '@nivo/sunburst';
import GaugeChart from 'react-gauge-chart';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { useSpreadsheetData } from '../context/SpreadsheetDataContext';
import { 
  Upload, FileText, Download, Settings, Eye, BarChart2, LineChart as LineChartIcon, 
  PieChart as PieChartIcon, ScatterChart as ScatterChartIcon, AreaChart as AreaChartIcon,
  Gauge as GaugeIcon, Target, DollarSign,  Award, AlertTriangle,
  CheckCircle,  Layers, ChevronsUp, ChevronsDown, TrendingUp, TrendingDown,
  Activity,  Grid, Table, Database, Sliders, Filter
} from 'lucide-react';


const ChartLibrary = () => {
  const [data, setData] = useState([]);
  
  const [localData, setLocalData] = useState([]);


  const [chartData, setChartData] = useState([]);

const { spreadsheetData, source } = useSpreadsheetData();
  
  // Unified data state
  const [displayData, setDisplayData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [chartType, setChartType] = useState('bar');
  const [fileName, setFileName] = useState('');
  const [loading, setLoading] = useState(false);
  const [filters, setFilters] = useState({});
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [dataView, setDataView] = useState('chart');
  
  const [chartConfig, setChartConfig] = useState({
    xAxis: '',
    yAxis: [],
    colors: ['#8884d8', '#82ca9d', '#ffc658']
  });

  // Sample data fallback
  const sampleData = [
    { product: 'apple', price: 500, category: 'Jan', sales: 4000, profit: 2400, expenses: 1600 },
    // ... other sample data
  ];

  // Initialize data
  useEffect(() => {
    if (spreadsheetData && spreadsheetData.length > 0) {
      setDisplayData(spreadsheetData);
      setColumns(Object.keys(spreadsheetData[0]));
      setFileName(source === 'spreadsheet' ? 'Current Spreadsheet' : 'Imported Data');
      
      // Auto-configure chart
      const firstCol = Object.keys(spreadsheetData[0])[0];
      const secondCol = Object.keys(spreadsheetData[0])[1] || firstCol;
      
      setChartConfig({
        xAxis: firstCol,
        yAxis: [secondCol],
        colors: ['#8884d8', '#82ca9d']
      });
    } else {
      // Fallback to sample data
      setDisplayData(sampleData);
      setColumns(['product', 'price', 'category', 'sales', 'profit', 'expenses']);
      setFileName('Sample Data');
      setChartConfig({
        xAxis: 'category',
        yAxis: ['sales', 'profit'],
        colors: ['#8884d8', '#82ca9d']
      });
    }
  }, [spreadsheetData, source]);

  // Process data with filters and sorting
  const processedData = useMemo(() => {
    let result = [...displayData];
    
    // Apply filters
    Object.entries(filters).forEach(([key, value]) => {
      if (value) {
        result = result.filter(item => 
          String(item[key]).toLowerCase().includes(String(value).toLowerCase())
        );
      }
    });
    
    // Apply sorting
    if (sortConfig.key) {
      result.sort((a, b) => {
        if (a[sortConfig.key] < b[sortConfig.key]) {
          return sortConfig.direction === 'asc' ? -1 : 1;
        }
        if (a[sortConfig.key] > b[sortConfig.key]) {
          return sortConfig.direction === 'asc' ? 1 : -1;
        }
        return 0;
      });
    }
    
    return result;
  }, [displayData, filters, sortConfig]);
  useEffect(() => {
    if (spreadsheetData && spreadsheetData.length > 0) {
      setChartData(spreadsheetData);
      setColumns(Object.keys(spreadsheetData[0]));
      
      // Auto-configure chart with first two columns
      const firstCol = Object.keys(spreadsheetData[0])[0];
      const secondCol = Object.keys(spreadsheetData[0])[1] || firstCol;
      
      setChartConfig({
        xAxis: firstCol,
        yAxis: [secondCol],
        colors: ['#8884d8', '#82ca9d']
      });
    } else {
      // Fallback to sample data
      const sample = [
        { category: 'Jan', value: 400 },
        { category: 'Feb', value: 300 },
        { category: 'Mar', value: 600 }
      ];
      setChartData(sample);
      setColumns(['category', 'value']);
      setChartConfig({
        xAxis: 'category',
        yAxis: ['value'],
        colors: ['#8884d8']
      });
    }
  }, [spreadsheetData]);
  useEffect(() => {
    if (chartData) {
      setLocalData(chartData);
      setFileName(source === 'spreadsheet' ? 'Current Spreadsheet' : 'Sample Data');
      
      if (chartData.length > 0) {
        const firstCol = Object.keys(chartData[0])[0];
        setColumns(Object.keys(chartData[0]));
        setChartConfig(prev => ({
          ...prev,
          xAxis: firstCol || '',
          yAxis: prev.yAxis.length ? prev.yAxis : [Object.keys(chartData[0])[1] || '']
        }));
      }
    } else {
      // Fallback to sample data
      setLocalData(sampleData);
      setColumns(['product','price','category', 'sales', 'profit', 'expenses']);
      setChartConfig(prev => ({
        ...prev,
        xAxis: 'category',
        yAxis: ['sales', 'profit']
      }));
    }
  }, [chartData, source]);
useEffect(() => {
  console.log('Received data:', location.state?.chartData);
}, [location.state]);
const processIncomingData = (rawData) => {
  if (!rawData || !Array.isArray(rawData) || rawData.length === 0) {
    return sampleData; // Fallback to sample data
  }

  // Ensure all rows have consistent structure
  const allKeys = new Set();
  rawData.forEach(row => Object.keys(row).forEach(key => allKeys.add(key)));
  
  return rawData.map(row => {
    const processedRow = {};
    allKeys.forEach(key => {
      // Convert numeric strings to numbers
      const value = row[key];
      processedRow[key] = typeof value === 'string' && !isNaN(value) ? 
        parseFloat(value) : 
        value || '';
    });
    return processedRow;
  });
};
  // Sample data for demonstration
//   const sampleData = [
//     { product:'apple',price:500,category: 'Jan', sales: 4000, profit: 2400, expenses: 1600 },
//     { product:'orange',price:200,category: 'Feb', sales: 3000, profit: 1398, expenses: 1602 },
//     { product:'grapes',price:800,category: 'Mar', sales: 2000, profit: 9800, expenses: 800 },
//     { product:'mango',price:1000,category: 'Apr', sales: 2780, profit: 3908, expenses: 872 },
//     { product:'guva',price:300,category: 'May', sales: 1890, profit: 4800, expenses: 1090 },
//     { product:'pineapple',price:398,category: 'Jun', sales: 2390, profit: 3800, expenses: 590 }
//   ];

  // Initialize with sample data
  useEffect(() => {
    if (data.length === 0) {
      setData(sampleData);
      setColumns(['product','price','category', 'sales', 'profit', 'expenses']);
      setChartConfig(prev => ({
  ...prev,
  xAxis: 'category', // ✅ use one string column
  yAxis: ['sales', 'profit']
}));

    }
  }, [data.length]);


  // Enhanced chart types with categories
  const chartTypes = [
    // Cards
    { id: 'numberCard', name: 'Number Card', icon: DollarSign, category: 'Cards' },
    { id: 'kpiCard', name: 'KPI Card', icon: Target, category: 'Cards' },
    { id: 'gaugeCard', name: 'Gauge Card', icon: GaugeIcon, category: 'Cards' },
    { id: 'metricCard', name: 'Metric Card', icon: Award, category: 'Cards' },
    { id: 'trendCard', name: 'Trend Card', icon: TrendingUp, category: 'Cards' },
    { id: 'statusCard', name: 'Status Card', icon: Activity, category: 'Cards' },
    
    // Bar/Column Charts
    { id: 'bar', name: 'Bar Chart', icon: BarChart2, category: 'Bar/Column' },
    { id: 'stackedBar', name: 'Stacked Bar', icon: BarChart2, category: 'Bar/Column' },
    { id: 'horizontalBar', name: 'Horizontal Bar', icon: BarChart2, category: 'Bar/Column' },
    { id: 'groupedBar', name: 'Grouped Bar', icon: BarChart2, category: 'Bar/Column' },
    { id: 'waterfall', name: 'Waterfall Chart', icon: BarChart2, category: 'Bar/Column' },
    
    // Line/Area Charts
    { id: 'line', name: 'Line Chart', icon: LineChartIcon, category: 'Line/Area' },
    { id: 'multiLine', name: 'Multi-Line Chart', icon: LineChartIcon, category: 'Line/Area' },
    { id: 'area', name: 'Area Chart', icon: AreaChartIcon, category: 'Line/Area' },
    { id: 'stackedArea', name: 'Stacked Area', icon: AreaChartIcon, category: 'Line/Area' },
    
    // Pie/Doughnut Charts
    { id: 'pie', name: 'Pie Chart', icon: PieChartIcon, category: 'Pie/Doughnut' },
    { id: 'doughnut', name: 'Doughnut Chart', icon: PieChartIcon, category: 'Pie/Doughnut' },
    { id: 'radialBar', name: 'Radial Bar', icon: PieChartIcon, category: 'Pie/Doughnut' },
    
    // Scatter/Bubble
    { id: 'scatter', name: 'Scatter Plot', icon: ScatterChartIcon, category: 'Scatter/Bubble' },
    { id: 'bubble', name: 'Bubble Chart', icon: ScatterChartIcon, category: 'Scatter/Bubble' },
    
    // Special Charts
    { id: 'radar', name: 'Radar Chart', icon: Radar, category: 'Special' },
    { id: 'treemap', name: 'Treemap', icon: Layers, category: 'Special' },
    { id: 'sunburst', name: 'Sunburst', icon: Layers, category: 'Special' },
    { id: 'sankey', name: 'Sankey', icon: Layers, category: 'Special' },
    { id: 'funnel', name: 'Funnel Chart', icon: Filter, category: 'Special' },
    { id: 'heatmap', name: 'Heatmap', icon: Grid, category: 'Special' },
    { id: 'gauge', name: 'Gauge Chart', icon: GaugeIcon, category: 'Special' }
  ];

  // Calculate summary statistics
  const summaryStats = useMemo(() => {
    if (!data.length || !chartConfig.yAxis.length) return {};
    
    const stats = {};
    chartConfig.yAxis.forEach(col => {
      const values = data.map(row => parseFloat(row[col])).filter(val => !isNaN(val));
      if (values.length > 0) {
        stats[col] = {
          sum: values.reduce((a, b) => a + b, 0),
          avg: values.reduce((a, b) => a + b, 0) / values.length,
          min: Math.min(...values),
          max: Math.max(...values),
          count: values.length,
          last: values[values.length - 1],
          first: values[0],
          trend: values.length > 1 ? (values[values.length - 1] - values[0]) / values[0] : 0
        };
      }
    });
    return stats;
  }, [data, chartConfig.yAxis]);

  // Filter and sort data
  

  // Handle file upload
  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);
    
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    if (fileExtension === 'csv') {
      Papa.parse(file, {
        complete: (result) => {
          processData(result.data);
        },
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true
      });
    } else if (['xlsx', 'xls'].includes(fileExtension)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        processData(jsonData);
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const processData = (rawData) => {
    if (rawData.length === 0) {
      setLoading(false);
      return;
    }
    
    // Clean data (remove empty rows)
    const cleanData = rawData.filter(row => 
      Object.values(row).some(val => val !== null && val !== undefined && val !== '')
    );
    
    setData(cleanData);
    setColumns(Object.keys(cleanData[0] || {}));
    setLoading(false);
    
    // Auto-set first column as x-axis if available
    if (cleanData.length > 0) {
      const firstCol = Object.keys(cleanData[0])[0];
      setChartConfig(prev => ({
        ...prev,
        xAxis: firstCol || ''
      }));
    }
  };

  const handleColumnSelect = (column, type) => {
    if (type === 'x') {
      setChartConfig(prev => ({ ...prev, xAxis: column }));
    } else if (type === 'y') {
      setChartConfig(prev => ({
        ...prev,
        yAxis: prev.yAxis.includes(column) 
          ? prev.yAxis.filter(col => col !== column)
          : [...prev.yAxis, column]
      }));
    }
  };

  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const handleFilterChange = (key, value) => {
    setFilters(prev => ({
      ...prev,
      [key]: value
    }));
  };

  const getChartData = () => {
    if (!processedData.length || !chartConfig.xAxis) return [];
    
    return processedData.map(row => {
      const item = { [chartConfig.xAxis]: row[chartConfig.xAxis] };
      chartConfig.yAxis.forEach(col => {
        item[col] = parseFloat(row[col]) || 0;
      });
      return item;
    }).filter(item => item[chartConfig.xAxis]);
  };

  // Enhanced card components
 const renderCard = (type, data, title, value, subtitle) => {
  const formatValue = (val) => {
    if (typeof val === 'number') {
      return val > 1000000 ? `${(val / 1000000).toFixed(1)}M` : 
             val > 1000 ? `${(val / 1000).toFixed(1)}K` : 
             val.toFixed(2);
    }
    return val;
  };


    const cardClasses = "bg-white rounded-lg shadow p-4 border-l-4 flex flex-col h-full";
    
    switch(type) {
      case 'number':
        return (
         <div key={title} className={`${cardClasses} border-l-blue-500`}>
  <div className="flex items-center mb-2">
    <DollarSign className="text-blue-500 mr-2" size={18} />
    <h3 className="text-sm font-medium text-gray-600">{title}</h3>
  </div>
  <p className="text-2xl font-bold text-gray-900 mb-1">{formatValue(value)}</p>
  {subtitle && <p className="text-xs text-gray-500">{subtitle}</p>}
</div>

        );
      
      case 'kpi':
        const isPositive = value >= 0;
        return (
  <div key={title} className={`${cardClasses} border-l-green-500`}>
    <div className="flex justify-between items-center mb-2">
      <h3 className="text-sm font-medium text-gray-600">{title}</h3>
      <Target className="text-green-500" size={18} />
    </div>
    <p className="text-2xl font-bold text-green-600 mb-1">{formatValue(value)}</p>
    <div className="flex items-center text-xs">
      {isPositive ? (
        <ChevronsUp className="text-green-500 mr-1" size={14} />
      ) : (
        <ChevronsDown className="text-red-500 mr-1" size={14} />
      )}
      <span className={isPositive ? 'text-green-500' : 'text-red-500'}>
        {Math.abs(value).toFixed(2)}%
      </span>
      <span className="text-gray-500 ml-1">{subtitle}</span>
    </div>
  </div>
);

// ...

case 'gauge':
  const percentage = Math.min((value / (summaryStats[chartConfig.yAxis[0]]?.max || 100)) * 100, 100);
  return (
    <div key={title} className={`${cardClasses} border-l-purple-500`}>
      <h3 className="text-sm font-medium text-gray-600 mb-2">{title}</h3>
      <div className="relative pt-1 flex-grow flex flex-col justify-center">
        <div className="overflow-hidden h-2 mb-2 text-xs flex rounded bg-purple-200">
          <div 
            style={{ width: `${percentage}%` }}
            className="shadow-none flex flex-col text-center whitespace-nowrap text-white justify-center bg-purple-500"
          />
        </div>
        <div className="flex justify-between items-center">
          <span className="text-xs text-gray-500">0</span>
          <span className="text-lg font-bold text-purple-600">{percentage.toFixed(1)}%</span>
          <span className="text-xs text-gray-500">100</span>
        </div>
      </div>
    </div>
  );

case 'metric':
  return (
    <div key={title} className={`${cardClasses} border-l-orange-500`}>
      <div className="flex justify-between items-center mb-2">
        <h3 className="text-sm font-medium text-gray-600">{title}</h3>
        <Award className="text-orange-500" size={18} />
      </div>
      <p className="text-2xl font-bold text-orange-600 mb-1">{formatValue(value)}</p>
      <p className="text-xs text-gray-500">{subtitle}</p>
    </div>
  );

case 'trend':
  const trendValue = parseFloat(value);
  const isTrendPositive = trendValue >= 0;

  return (
    <div key={title} className={`${cardClasses} border-l-indigo-500`}>
      <div className="flex justify-between items-center mb-2">
        <h3 className="text-sm font-medium text-gray-600">{title}</h3>
        {isTrendPositive ? (
          <TrendingUp className="text-green-500" size={18} />
        ) : (
          <TrendingDown className="text-red-500" size={18} />
        )}
      </div>
      <p className="text-2xl font-bold text-gray-900 mb-1">{formatValue(value)}</p>
      <div className="flex items-center text-xs">
        {isTrendPositive ? (
          <ChevronsUp className="text-green-500 mr-1" size={14} />
        ) : (
          <ChevronsDown className="text-red-500 mr-1" size={14} />
        )}
        <span className={isTrendPositive ? 'text-green-500' : 'text-red-500'}>
          {Math.abs(trendValue).toFixed(2)}%
        </span>
        <span className="text-gray-500 ml-1">vs previous</span>
      </div>
    </div>
  );

case 'status':
  const statusValue = parseFloat(value);
  let statusColor = 'gray';
  let statusIcon = <Activity className="text-gray-500" size={18} />;

  if (statusValue > 0) {
    statusColor = 'green';
    statusIcon = <CheckCircle className="text-green-500" size={18} />;
  } else if (statusValue < 0) {
    statusColor = 'red';
    statusIcon = <AlertTriangle className="text-red-500" size={18} />;
  }

  return (
    <div key={title} className={`${cardClasses} border-l-${statusColor}-500`}>
      <div className="flex justify-between items-center mb-2">
        <h3 className="text-sm font-medium text-gray-600">{title}</h3>
        {statusIcon}
      </div>
      <p className={`text-2xl font-bold text-${statusColor}-600 mb-1`}>{formatValue(value)}</p>
      <p className="text-xs text-gray-500">{subtitle}</p>
    </div>
  );

default:
  return null;

    }
  };
  // Near top of ChartLibrary.jsx (before renderChart or inside the component)

const buildWaterfallData = (rawData, { categoryField = 'category' } = {}) => {
  if (!rawData || rawData.length === 0) return [];

  const categories = ['sales', 'expenses', 'profit'];
  const chartData = [];
  let cumulative = 0;

  chartData.push({
    name: "Start",
    value: 0,
    cumulative: 0,
    fill: "#8884d8",
  });

  rawData.forEach((row) => {
    categories.forEach((field) => {
      const val = field === "expenses" ? -row[field] : row[field];
      cumulative += val;
      const label = `${row[categoryField] || 'Unknown'} ${field}`;

      chartData.push({
        name: label,
        value: val,
        cumulative,
        fill: field === "expenses" ? "#f87171" : "#34d399",
      });
    });
  });

  chartData.push({
    name: "End",
    value: 0,
    cumulative,
    fill: "#8884d8",
  });

  return chartData;
};



  


  // Render the appropriate chart based on selection
  const renderChart = () => {
    const chartData = getChartData();
    if (!chartData.length) return <div className="text-center text-gray-500 py-8">Select columns to display chart</div>;

    const commonProps = {
      data: chartData,
      margin: { top: 20, right: 30, left: 20, bottom: 5 }
    };

    // Card Types
    if (chartType.includes('Card')) {
      const firstYCol = chartConfig.yAxis[0];
      if (!firstYCol || !summaryStats[firstYCol]) {
        return <div className="text-center text-gray-500 py-8">Select a numeric column for card display</div>;
      }
      
      const stats = summaryStats[firstYCol];
      const cards = [];
      
      if (chartType === 'numberCard') {
  cards.push(renderCard('number', chartData, `Total ${firstYCol}`, stats.sum, `${stats.count} records`));
  cards.push(renderCard('number', chartData, `Average ${firstYCol}`, stats.avg, 'Mean value'));
  cards.push(renderCard('number', chartData, `Maximum ${firstYCol}`, stats.max, 'Highest value'));
  cards.push(renderCard('number', chartData, `Minimum ${firstYCol}`, stats.min, 'Lowest value'));
} else if (chartType === 'kpiCard') {
  cards.push(renderCard('kpi', chartData, `${firstYCol} Performance`, stats.sum, 'Total'));
  cards.push(renderCard('kpi', chartData, `${firstYCol} Average`, stats.avg, 'Mean'));
  cards.push(renderCard('kpi', chartData, `${firstYCol} Trend`, stats.trend * 100, 'Change'));
} else if (chartType === 'gaugeCard') {
  cards.push(renderCard('gauge', chartData, `${firstYCol} Progress`, stats.avg, ''));
  cards.push(renderCard('gauge', chartData, `${firstYCol} Completion`, stats.sum, ''));
} else if (chartType === 'metricCard') {
  cards.push(renderCard('metric', chartData, `${firstYCol} Score`, stats.avg, ''));
  cards.push(renderCard('metric', chartData, `${firstYCol} Total`, stats.sum, ''));
} else if (chartType === 'trendCard') {
  cards.push(renderCard('trend', chartData, `${firstYCol} Trend`, stats.trend * 100, 'Change'));
  cards.push(renderCard('trend', chartData, `${firstYCol} Growth`, (stats.last - stats.first) * 100, 'Overall'));
} else if (chartType === 'statusCard') {
  cards.push(renderCard('status', chartData, `${firstYCol} Status`, stats.trend * 100, 'Trend'));
  cards.push(renderCard('status', chartData, `${firstYCol} Alert`, stats.last - stats.avg, 'Deviation'));
}

return (
  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 p-4">
    {cards}
  </div>
);
    }
    // Chart types
    switch (chartType) {
      case 'bar':
      
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Bar 
                  key={col} 
                  dataKey={col} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  name={col}
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'stackedBar':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Bar 
                  key={col} 
                  dataKey={col} 
                  stackId="a" 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  name={col}
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'horizontalBar':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart {...commonProps} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" />
              <YAxis dataKey={chartConfig.xAxis} type="category" />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Bar 
                  key={col} 
                  dataKey={col} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  name={col}
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

        case 'groupedBar':
  return (
    <ResponsiveContainer width="100%" height={400}>
      <BarChart {...commonProps}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey={chartConfig.xAxis} />
        <YAxis />
        <Tooltip />
        <Legend />
        {chartConfig.yAxis.map((col, index) => (
          <Bar 
            key={col} 
            dataKey={col} 
            fill={chartConfig.colors[index % chartConfig.colors.length]} 
            name={col}
            // Add margin between bars in the same group
            radius={[4, 4, 0, 0]}
          />
        ))}
      </BarChart>
    </ResponsiveContainer>
  );
        
        case 'waterfall':
  const waterfallData = buildWaterfallData(data, { categoryField: 'category' });

  return (
    <ResponsiveContainer width="100%" height={400}>
      <BarChart data={waterfallData}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="name" />
        <YAxis />
        <Tooltip />
        <Legend />
        <Bar
          dataKey="value"
          stackId="a"
          fill="#8884d8"
          label={{ position: "top" }}
          isAnimationActive={false}
        >
         {
  waterfallData.map((entry, index) => (
    <Cell key={`cell-${index}`} fill={entry.fill} />
  ))
}

        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );

      case 'line':
      case 'multiLine':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <LineChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Line 
                  key={col} 
                  type="monotone" 
                  dataKey={col} 
                  stroke={chartConfig.colors[index % chartConfig.colors.length]} 
                  strokeWidth={2}
                  dot={{ r: 4 }}
                  activeDot={{ r: 6 }}
                  name={col}
                />
              ))}
            </LineChart>
          </ResponsiveContainer>
        );

      case 'area':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <AreaChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Area 
                  key={col} 
                  type="monotone" 
                  dataKey={col} 
                  stroke={chartConfig.colors[index % chartConfig.colors.length]} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  fillOpacity={0.4}
                  name={col}
                />
              ))}
            </AreaChart>
          </ResponsiveContainer>
        );

      case 'stackedArea':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <AreaChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Area 
                  key={col} 
                  type="monotone" 
                  dataKey={col} 
                  stackId="1" 
                  stroke={chartConfig.colors[index % chartConfig.colors.length]} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  name={col}
                />
              ))}
            </AreaChart>
          </ResponsiveContainer>
        );

      case 'pie':
        const pieData = chartData.map(item => ({
          name: item[chartConfig.xAxis],
          value: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
        }));
        return (
          <ResponsiveContainer width="100%" height={400}>
            <PieChart>
              <Pie
                data={pieData}
                dataKey="value"
                nameKey="name"
                cx="50%"
                cy="50%"
                outerRadius={120}
                label
              >
                {pieData.map((entry, index) => (
  <Cell
    key={`cell-${index}`}
    fill={chartConfig.colors[index % chartConfig.colors.length]}
  />
))}

              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        );

      case 'doughnut':
        const doughnutData = chartData.map(item => ({
          name: item[chartConfig.xAxis],
          value: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
        }));
        return (
          <ResponsiveContainer width="100%" height={400}>
            <PieChart>
              <Pie
                data={doughnutData}
                dataKey="value"
                nameKey="name"
                cx="50%"
                cy="50%"
                innerRadius={60}
                outerRadius={100}
                label
              >
               {doughnutData.map((entry, index) => (
  <Cell
    key={`cell-${index}`}
    fill={chartConfig.colors[index % chartConfig.colors.length]}
  />
))}

              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        );

      case 'radialBar':
        const radialData = chartData.map(item => ({
          name: item[chartConfig.xAxis],
          value: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
        }));
        return (
          <ResponsiveContainer width="100%" height={400}>
            <RadialBarChart 
              innerRadius="10%" 
              outerRadius="80%" 
              data={radialData}
              startAngle={180}
              endAngle={0}
            >
              <RadialBar 
                minAngle={15} 
                label={{ position: 'insideStart', fill: '#fff' }} 
                background 
                dataKey="value"
              >
               {radialData.map((entry, index) => (
  <Cell
    key={`cell-${index}`}
    fill={chartConfig.colors[index % chartConfig.colors.length]}
  />
))}

              </RadialBar>
              <Legend />
              <Tooltip />
            </RadialBarChart>
          </ResponsiveContainer>
        );

      case 'scatter':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <ScatterChart {...commonProps}>
              <CartesianGrid />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Scatter 
                  key={col} 
                  name={col} 
                  dataKey={col} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                />
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        );

      case 'bubble':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <ScatterChart {...commonProps}>
              <CartesianGrid />
              <XAxis dataKey={chartConfig.xAxis} />
              <YAxis />
              <Tooltip />
              <Legend />
              {chartConfig.yAxis.map((col, index) => (
                <Scatter 
                  key={col} 
                  name={col} 
                  dataKey={col} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  shape="circle"
                />
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        );

      case 'radar':
        const radarData = chartData.map(item => ({
          subject: item[chartConfig.xAxis],
          ...chartConfig.yAxis.reduce((acc, col) => {
            acc[col] = item[col] || 0;
            return acc;
          }, {})
        }));
        return (
          <ResponsiveContainer width="100%" height={400}>
            <RadarChart data={radarData}>
              <PolarGrid />
              <PolarAngleAxis dataKey="subject" />
              <PolarRadiusAxis />
              {chartConfig.yAxis.map((col, index) => (
                <Radar 
                  key={col} 
                  name={col} 
                  dataKey={col} 
                  stroke={chartConfig.colors[index % chartConfig.colors.length]} 
                  fill={chartConfig.colors[index % chartConfig.colors.length]} 
                  fillOpacity={0.6} 
                />
              ))}
              <Legend />
              <Tooltip />
            </RadarChart>
          </ResponsiveContainer>
        );

      case 'treemap':
        const treemapData = chartData.map(item => ({
          name: item[chartConfig.xAxis],
          size: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
        }));
        return (
          <ResponsiveContainer width="100%" height={400}>
            <Treemap
              data={treemapData}
              dataKey="size"
              ratio={4/3}
              stroke="#fff"
              fill="#8884d8"
            >
              <Tooltip />
            </Treemap>
          </ResponsiveContainer>
        );

      case 'sunburst':
  const nivoData = {
    name: "Total",
    children: chartData.map(item => ({
      name: item[chartConfig.xAxis],
      loc: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
    }))
  };

  return (
    <div style={{ height: 400 }}>
      <ResponsiveSunburst
        data={nivoData}
        id="name"
        value="loc"
        cornerRadius={2}
        colors={{ scheme: 'nivo' }}
        childColor={{ from: 'color' }}
        borderWidth={1}
        borderColor={{ from: 'color', modifiers: [['darker', 0.1]] }}
        animate={true}
        motionConfig="gentle"
      />
    </div>
  );


  case 'sankey': {
  const scaleFactor = 10;

  // Create dynamic nodes and links from your data
  const productNode = { name: 'Fruits' };
  const monthNodes = data.map((item, index) => ({ name: item.category }));
  const metricNodes = ['Sales', 'Profit', 'Expenses'].map(m => ({ name: m }));

  const nodes = [productNode, ...monthNodes, ...metricNodes];

  const links = [];

  // Link products to months using price
  data.forEach((item, i) => {
    links.push({
      source: 0, // Fruits node
      target: i + 1,
      value: item.price * scaleFactor
    });
  });

  // Link months to metrics
  data.forEach((item, i) => {
    const base = i + 1;
    links.push({ source: base, target: data.length + 1, value: item.sales * scaleFactor });
    links.push({ source: base, target: data.length + 2, value: item.profit * scaleFactor });
    links.push({ source: base, target: data.length + 3, value: item.expenses * scaleFactor });
  });

  return (
    <div style={{ width: '100%', height: '600px' }}>
      <ResponsiveContainer width="100%" height="100%">
        <Sankey
          data={{ nodes, links }}
          nodePadding={40}
          nodeWidth={20}
          margin={{ top: 20, right: 100, bottom: 20, left: 100 }}
        >
          <Tooltip />
        </Sankey>
      </ResponsiveContainer>
    </div>
  );
}





      case 'funnel':
        const funnelData = chartData.map(item => ({
          name: item[chartConfig.xAxis],
          value: chartConfig.yAxis.reduce((sum, col) => sum + (item[col] || 0), 0)
        })).sort((a, b) => b.value - a.value);
        return (
          <ResponsiveContainer width="100%" height={400}>
            <FunnelChart>
              <Tooltip />
              <Funnel
                data={funnelData}
                dataKey="value"
              >
                {funnelData.map((entry, index) => (
  <Cell
    key={`cell-${index}`}
    fill={chartConfig.colors[index % chartConfig.colors.length]}
  />
))}

                <LabelList position="right" fill="#000" stroke="none" dataKey="name" />
              </Funnel>
            </FunnelChart>
          </ResponsiveContainer>
        );

      case 'heatmap':
        return (
          <div
  key={`${rowIndex}-${colIndex}`}
  className="w-8 h-8 rounded flex items-center justify-center text-xs font-medium"
  style={{
    backgroundColor: `rgba(59, 130, 246, ${intensity})`,
    color: intensity > 0.5 ? 'white' : 'black'
  }}
  title={`${col}: ${value}`}
>
  {Math.round(value)}
</div>
);


      case 'gauge':
  // Calculate the gauge value (using the first Y-axis column)
  const firstYCol = chartConfig.yAxis[0];
  if (!firstYCol || !summaryStats[firstYCol]) {
    return <div className="text-center text-gray-500 py-8">Select a numeric column for gauge display</div>;
  }

  // Get values for scaling
  const currentValue = summaryStats[firstYCol].avg;
  const minValue = summaryStats[firstYCol].min;
  const maxValue = summaryStats[firstYCol].max;
  
  // Normalize the value to 0-1 range for the gauge
  const normalizedValue = (currentValue - minValue) / (maxValue - minValue);

  return (
    <div style={{ width: '100%', height: '300px', padding: '20px' }}>
  <GaugeChart
    id={`gauge-${firstYCol}`}
    nrOfLevels={20}
    percent={normalizedValue}
    colors={['#EA4228', '#F5CD19', '#5BE12C']}
    arcWidth={0.3}
    textColor="#000000"
    needleColor="#345243"
    needleBaseColor="#345243"
    formatTextValue={(value) => {
      // Convert back from normalized value to actual value
      const actualValue = minValue + (value * (maxValue - minValue));
      return `${actualValue.toFixed(1)} (${firstYCol})`;
    }}
  />
  <div className="text-center mt-2 text-sm text-gray-600">
    Range: {minValue.toFixed(1)} to {maxValue.toFixed(1)}
  </div>
</div>

  );
      default:
        return <div className="text-center text-gray-500 py-8">Select a chart type to display</div>;
    }
  };

  // Render data table view
  const renderDataTable = () => {
    if (!processedData.length) return <div className="text-center text-gray-500 py-8">No data to display</div>;

    return (
      <div className="overflow-x-auto">
                <table className="min-w-full bg-white rounded-lg overflow-hidden">
          <thead className="bg-gray-100">
            <tr>
              {columns.map(col => (
                <th 
                  key={col} 
                  className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-200"
                  onClick={() => handleSort(col)}
                >
                  <div className="flex items-center justify-between">
                    {col}
                    {sortConfig.key === col && (
                      <span>
                        {sortConfig.direction === 'asc' ? '↑' : '↓'}
                      </span>
                    )}
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-200">
  {processedData.map((row, index) => (
    <tr key={index} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
      {columns.map(col => (
        <td key={`${index}-${col}`} className="px-4 py-2 text-sm text-gray-700">
          {row[col]}
        </td>
      ))}
    </tr>
  ))}
</tbody>

        </table>
      </div>
    );
  };

  // Render column selection controls
  const renderColumnControls = () => {
    if (!columns.length) return null;

    return (
  <div className="bg-white p-4 rounded-lg shadow mb-4">
    <h3 className="text-lg font-medium mb-3 flex items-center">
      <Database className="mr-2" size={18} />
      Data Columns
    </h3>

    <div className="mb-4">
      <h4 className="text-sm font-medium mb-2 flex items-center">
        <Sliders className="mr-1" size={16} />
        X-Axis (Category)
      </h4>
      <div className="flex flex-wrap gap-2">
        {columns.map(col => (
          <button
            key={`x-${col}`}
            onClick={() => handleColumnSelect(col, 'x')}
            className={`px-3 py-1 text-xs rounded-full flex items-center ${
              chartConfig.xAxis === col 
                ? 'bg-blue-100 text-blue-800 border border-blue-300' 
                : 'bg-gray-100 text-gray-800 hover:bg-gray-200'
            }`}
          >
            {col}
            {chartConfig.xAxis === col && <CheckCircle className="ml-1" size={14} />}
          </button>
        ))}
      </div>
    </div>

    <div>
      <h4 className="text-sm font-medium mb-2 flex items-center">
        <BarChart2 className="mr-1" size={16} />
        Y-Axis (Values)
      </h4>
      <div className="flex flex-wrap gap-2">
        {columns.map(col => (
          <button
            key={`y-${col}`}
            onClick={() => handleColumnSelect(col, 'y')}
            className={`px-3 py-1 text-xs rounded-full flex items-center ${
              chartConfig.yAxis.includes(col)
                ? 'bg-green-100 text-green-800 border border-green-300' 
                : 'bg-gray-100 text-gray-800 hover:bg-gray-200'
            }`}
          >
            {col}
            {chartConfig.yAxis.includes(col) && <CheckCircle className="ml-1" size={14} />}
          </button>
        ))}
      </div>
    </div>
  </div>
);
  }

  // Render filters
  const renderFilters = () => {
    if (!columns.length) return null;

    return (
      <div className="bg-white p-4 rounded-lg shadow mb-4">
        <h3 className="text-lg font-medium mb-3 flex items-center">
          <Filter className="mr-2" size={18} />
          Filters
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
  {columns.slice(0, 6).map(col => (
    <div key={`filter-${col}`} className="space-y-1">
      <label className="block text-sm font-medium text-gray-700">{col}</label>
      <input
        type="text"
        value={filters[col] || ''}
        onChange={(e) => handleFilterChange(col, e.target.value)}
        placeholder={`Filter ${col}...`}
        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500 text-sm"
      />
    </div>

          ))}
        </div>
      </div>
    );
  };

  // Group chart types by category
  const groupedChartTypes = chartTypes.reduce((acc, type) => {
    if (!acc[type.category]) {
      acc[type.category] = [];
    }
    acc[type.category].push(type);
    return acc;
  }, {});

  // Render chart type selection
  const renderChartTypeSelector = () => {
    return (
      <div className="bg-white p-4 rounded-lg shadow mb-4">
        <h3 className="text-lg font-medium mb-3 flex items-center">
          <BarChart2 className="mr-2" size={18} />
          Chart Type
        </h3>
        <div className="space-y-4">
  {Object.entries(groupedChartTypes).map(([category, types]) => (
    <div key={category}>
      <h4 className="text-sm font-medium mb-2 text-gray-700">{category}</h4>
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-3 gap-2">
        {types.map(type => (
          <button
            key={type.id}
            onClick={() => setChartType(type.id)}
            className={`flex flex-col items-center p-2 rounded-lg border w-auto min-w-fit ${
              chartType === type.id
                ? 'bg-blue-50 border-blue-300 text-blue-700'
                : 'bg-gray-50 border-gray-200 text-gray-700 hover:bg-gray-100'
            }`}
          >
            <type.icon size={20} className="mb-1" />
            <span className="text-xs text-center whitespace-nowrap">{type.name}</span>
          </button>
        ))}
      </div>
    </div>
  ))}
</div>

      </div>
    );
  };

 const chartRef = useRef(null);

const handleDownloadChart = async () => {
  if (!chartRef.current) {
    alert("Chart not available for export");
    return;
  }

  try {
    const canvas = await html2canvas(chartRef.current, {
      scale: 2, // Higher quality
      logging: false,
      useCORS: true,
      allowTaint: true
    });

    const link = document.createElement('a');
    link.download = `chart-export-${new Date().toISOString().slice(0, 10)}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  } catch (error) {
    console.error("Error exporting chart:", error);
    alert("Failed to export chart");
  }
};
const handleDownloadChartImage = (format = 'png') => {
  const chartElement = document.querySelector('.chart-container');
  
  if (!chartElement) {
    alert("Chart not found");
    return;
  }

  html2canvas(chartElement).then(canvas => {
    const link = document.createElement('a');
    link.download = `chart_${new Date().toISOString().slice(0,10)}.${format}`;
    link.href = canvas.toDataURL(`image/${format}`);
    link.click();
  });
};
  // Render data summary
  const renderDataSummary = () => {
    if (!data.length) return null;

    return (
  <div className="bg-white p-4 rounded-lg shadow mb-4">
    <h3 className="text-lg font-medium mb-3 flex items-center">
      <FileText className="mr-2" size={18} />
      Data Summary
    </h3>
    <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
      
      {/* Total Records */}
      <div className="bg-blue-50 p-3 rounded-lg">
        <p className="text-sm text-blue-800">Total Records</p>
        <p className="text-xl font-bold text-blue-900">{data.length}</p>
      </div>

      {/* Columns */}
      <div className="bg-green-50 p-3 rounded-lg">
        <p className="text-sm text-green-800">Columns</p>
        <p className="text-xl font-bold text-green-900">{columns.length}</p>
      </div>

      {/* File Name */}
      <div className="bg-purple-50 p-3 rounded-lg">
        <p className="text-sm text-purple-800">File</p>
        <p className="text-sm font-medium text-purple-900 truncate">{fileName}</p>
      </div>

      {/* Last Updated */}
      <div className="bg-yellow-50 p-3 rounded-lg">
        <p className="text-sm text-yellow-800">Last Updated</p>
        <p className="text-sm font-medium text-yellow-900">
          {new Date().toLocaleString()}
        </p>
      </div>

    </div>
  </div>
);
  }
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [fullscreenMode, setFullscreenMode] = useState(false);

  const toggleSidebar = () => {
    setSidebarOpen(!sidebarOpen);
  };

  const toggleFullscreen = () => {
    setFullscreenMode(!fullscreenMode);
  };


  return (
  <div className="min-h-screen bg-gray-50">
    {/* Header Section */}
    <header className="bg-white shadow-sm">
      <div className="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8 flex justify-between items-center">
        <div>
          <h1 className="text-2xl font-bold text-gray-900 flex items-center">
            <BarChart2 className="mr-3 text-blue-600" size={28} />
            Strix Charts
          </h1>
          <p className="text-gray-600 mt-1">Transform your data into actionable insights</p>
        </div>
        <div className="flex items-center space-x-4">
          <button
            onClick={handleDownloadChart}
            className="inline-flex items-center px-4 py-2 border border-gray-300 shadow-sm text-sm font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50"
          >
            <Download className="mr-2" size={16} />
            Export
          </button>
        </div>
      </div>
    </header>

    {/* Main Content */}
    <main className="max-w-7xl mx-auto px-4 py-6 sm:px-6 lg:px-8">
      {/* Top Control Section */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
        {/* Data Upload Card */}
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
              {fileName && (
                <p className="text-xs text-gray-500 mt-1 truncate max-w-full">
                  {fileName}
                </p>
              )}
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
            <div>
              <h4 className="text-sm font-medium text-gray-700 mb-2">X-Axis (Category)</h4>
              <div className="flex flex-wrap gap-2">
                {columns.map(col => (
                  <button
                    key={`x-${col}`}
                    onClick={() => handleColumnSelect(col, 'x')}
                    className={`px-3 py-1 text-xs rounded-full ${
                      chartConfig.xAxis === col
                        ? 'bg-blue-100 text-blue-800 border border-blue-300'
                        : 'bg-gray-100 text-gray-800 hover:bg-gray-200'
                    }`}
                  >
                    {col}
                    {chartConfig.xAxis === col && <CheckCircle className="ml-1" size={14} />}
                  </button>
                ))}
              </div>
            </div>
            <div>
              <h4 className="text-sm font-medium text-gray-700 mb-2">Y-Axis (Values)</h4>
              <div className="flex flex-wrap gap-2">
                {columns.map(col => (
                  <button
                    key={`y-${col}`}
                    onClick={() => handleColumnSelect(col, 'y')}
                    className={`px-3 py-1 text-xs rounded-full ${
                      chartConfig.yAxis.includes(col)
                        ? 'bg-green-100 text-green-800 border border-green-300'
                        : 'bg-gray-100 text-gray-800 hover:bg-gray-200'
                    }`}
                  >
                    {col}
                    {chartConfig.yAxis.includes(col) && <CheckCircle className="ml-1" size={14} />}
                  </button>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
        {/* Left Sidebar - Chart Types */}
        <div className="lg:col-span-1">
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200 sticky top-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-3 flex items-center">
              <BarChart2 className="mr-2 text-blue-500" size={20} />
              Chart Type
            </h3>
            <div className="space-y-4">
              {Object.entries(groupedChartTypes).map(([category, types]) => (
                <div key={category}>
                  <h4 className="text-sm font-medium text-gray-700 mb-2">{category}</h4>
                  <div className="grid grid-cols-2 gap-2">
                    {types.map(type => (
                      <button
                        key={type.id}
                        onClick={() => setChartType(type.id)}
                        className={`flex flex-col items-center p-2 rounded-lg border ${
                          chartType === type.id
                            ? 'bg-blue-50 border-blue-200 text-blue-700'
                            : 'bg-gray-50 border-gray-200 text-gray-700 hover:bg-gray-100'
                        }`}
                      >
                        <type.icon size={18} className="mb-1" />
                        <span className="text-xs">{type.name}</span>
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Main Content Area */}
        <div className="lg:col-span-4 space-y-4">
          {/* View Toggle and Filters */}
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
            <div className="flex justify-between items-center mb-4">
              <div className="flex space-x-2">
                <button
                  onClick={() => setDataView("chart")}
                  className={`px-3 py-1 rounded-md flex items-center ${
                    dataView === "chart"
                      ? "bg-blue-100 text-blue-800"
                      : "text-gray-600 hover:bg-gray-100"
                  }`}
                >
                  <Eye className="mr-1" size={16} />
                  Chart View
                </button>
                <button
                  onClick={() => setDataView("table")}
                  className={`px-3 py-1 rounded-md flex items-center ${
                    dataView === "table"
                      ? "bg-blue-100 text-blue-800"
                      : "text-gray-600 hover:bg-gray-100"
                  }`}
                >
                  <Table className="mr-1" size={16} />
                  Data Table
                </button>
              </div>
              <div className="flex items-center space-x-2">
                <button className="px-3 py-1 text-gray-600 hover:bg-gray-100 rounded-md flex items-center">
                  <Settings className="mr-1" size={16} />
                  Settings
                </button>
              </div>
            </div>

            {/* Filters */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {columns.slice(0, 3).map(col => (
                <div key={`filter-${col}`} className="space-y-1">
                  <label className="block text-sm font-medium text-gray-700">{col}</label>
                  <input
                    type="text"
                    value={filters[col] || ''}
                    onChange={(e) => handleFilterChange(col, e.target.value)}
                    placeholder={`Filter ${col}...`}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500 text-sm"
                  />
                </div>
              ))}
            </div>
          </div>

          {/* Chart/Table Display */}
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
            {loading ? (
              <div className="flex justify-center items-center h-full">
                <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-blue-500"></div>
              </div>
            ) : dataView === "chart" ? (
              <div ref={chartRef}>{renderChart()}</div>
            ) : (
              renderDataTable()
            )}
          </div>

          {/* Data Summary - Moved to bottom */}
          <div className="bg-white p-4 rounded-lg shadow border border-gray-200">
            <h3 className="text-lg font-semibold text-gray-900 mb-3 flex items-center">
              <FileText className="mr-2 text-blue-500" size={20} />
              Data Summary
            </h3>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
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
      </div>
    </main>
  </div>
);
};

export default ChartLibrary;