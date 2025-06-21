import React, { useState, useRef, useCallback, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, LineChart, Line, PieChart, Pie, Cell, ResponsiveContainer } from 'recharts';
import { Move, X, Plus, BarChart3, TrendingUp, Activity, CreditCard, Upload, FileSpreadsheet, Database, RefreshCw, Edit2, Check } from 'lucide-react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

// File Upload Component with Real Excel/CSV Processing
const FileUploader = ({ onDataLoaded, onClose }) => {
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef(null);

  const handleFiles = async (files) => {
    const file = files[0];
    if (!file) return;

    setLoading(true);
    try {
      if (file.name.endsWith('.csv')) {
        // Parse CSV file
        Papa.parse(file, {
          header: true,
          dynamicTyping: true,
          skipEmptyLines: true,
          complete: (results) => {
            onDataLoaded({
              fileName: file.name,
              data: results.data,
              columns: Object.keys(results.data[0] || {}),
              rowCount: results.data.length
            });
            onClose();
            setLoading(false);
          },
          error: (error) => {
            console.error('CSV parsing error:', error);
            alert('Error parsing CSV file');
            setLoading(false);
          }
        });
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        // Parse Excel file
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            onDataLoaded({
              fileName: file.name,
              data: jsonData,
              columns: Object.keys(jsonData[0] || {}),
              rowCount: jsonData.length
            });
            onClose();
            setLoading(false);
          } catch (error) {
            console.error('Excel parsing error:', error);
            alert('Error parsing Excel file');
            setLoading(false);
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        alert('Please upload a CSV or Excel file');
        setLoading(false);
      }
    } catch (error) {
      console.error('Error processing file:', error);
      alert('Error reading file. Please check the format.');
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg p-6 w-96 max-w-full mx-4" style={{maxWidth: '90vw'}}>
        <div className="flex justify-between items-center mb-4">
          <h3 className="text-lg font-semibold">Upload Spreadsheet</h3>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700">
            <X size={20} />
          </button>
        </div>
        
        <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center">
          {loading ? (
            <div className="flex items-center justify-center">
              <RefreshCw className="animate-spin mr-2" size={20} />
              Processing file...
            </div>
          ) : (
            <>
              <Upload size={48} className="mx-auto text-gray-400 mb-4" />
              <p className="text-gray-600 mb-2">Click to upload your file</p>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
              >
                Browse Files
              </button>
              <p className="text-sm text-gray-500 mt-2">
                Supports CSV, Excel (.xlsx, .xls)
              </p>
            </>
          )}
        </div>
        
        <input
          ref={fileInputRef}
          type="file"
          accept=".csv,.xlsx,.xls"
          onChange={(e) => handleFiles(e.target.files)}
          className="hidden"
        />
      </div>
    </div>
  );
};

// Column Selector Component
const ColumnSelector = ({ columns, selectedColumns, onSelectionChange, chartType }) => {
  const getRequiredFields = () => {
    switch (chartType) {
      case 'bar':
      case 'line':
        return ['X-Axis (Category)', 'Y-Axis (Value)', 'Y-Axis 2 (Optional)'];
      case 'pie':
        return ['Category', 'Value'];
      case 'card':
        return ['Value Column'];
      default:
        return [];
    }
  };

  const fields = getRequiredFields();

  return (
    <div className="mb-4 p-3 bg-gray-50 rounded-lg">
      <h4 className="text-sm font-medium mb-2">Map Columns</h4>
      {fields.map((field, index) => (
        <div key={field} className="mb-2">
          <label className="block text-xs font-medium text-gray-700 mb-1">
            {field} {index < (chartType === 'pie' ? 2 : chartType === 'card' ? 1 : 1) && '*'}
          </label>
          <select
            value={selectedColumns[index] || ''}
            onChange={(e) => {
              const newSelection = [...selectedColumns];
              newSelection[index] = e.target.value;
              onSelectionChange(newSelection);
            }}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded"
            required={index < (chartType === 'pie' ? 2 : chartType === 'card' ? 1 : 1)}
          >
            <option value="">Select column...</option>
            {columns.map(col => (
              <option key={col} value={col}>{col}</option>
            ))}
          </select>
        </div>
      ))}
    </div>
  );
};

// Data Preview Component
const DataPreview = ({ data, onClose }) => {
  const previewData = (data?.data || []).slice(0, 5);
  
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg p-6 max-w-4xl overflow-auto mx-4" style={{maxHeight: '80vh'}}>
        <div className="flex justify-between items-center mb-4">
          <div>
            <h3 className="text-lg font-semibold">Data Preview</h3>
            <p className="text-sm text-gray-600">
              {data.fileName} - {data.rowCount} rows, {data.columns?.length || 0} columns
            </p>
          </div>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700">
            <X size={20} />
          </button>
        </div>
        
        <div className="overflow-x-auto">
          <table className="min-w-full table-auto border-collapse">
            <thead>
              <tr className="bg-gray-100">
                {(data.columns || []).map(col => (
                  <th key={col} className="border border-gray-300 px-3 py-2 text-left text-sm font-medium">
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewData.map((row, index) => (
                <tr key={index} className="hover:bg-gray-50">
                  {(data.columns || []).map(col => (
                    <td key={col} className="border border-gray-300 px-3 py-2 text-sm">
                      {row[col] || ''}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        
        {data.rowCount > 5 && (
          <p className="text-sm text-gray-500 mt-2">
            Showing first 5 rows of {data.rowCount} total rows
          </p>
        )}
      </div>
    </div>
  );
};

// Chart Name Editor Component
const ChartNameEditor = ({ currentName, onSave, onCancel }) => {
  const [name, setName] = useState(currentName);

  const handleSave = () => {
    if (name.trim()) {
      onSave(name.trim());
    }
  };

  const handleKeyPress = (e) => {
    if (e.key === 'Enter') {
      handleSave();
    } else if (e.key === 'Escape') {
      onCancel();
    }
  };

  return (
    <div className="flex items-center gap-2">
      <input
        type="text"
        value={name}
        onChange={(e) => setName(e.target.value)}
        onKeyDown={handleKeyPress}
        className="text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded px-2 py-1 min-w-0 flex-1"
        autoFocus
      />
      <button
        onClick={handleSave}
        className="p-1 hover:bg-gray-200 rounded text-green-600"
      >
        <Check size={14} />
      </button>
      <button
        onClick={onCancel}
        className="p-1 hover:bg-gray-200 rounded text-gray-500"
      >
        <X size={14} />
      </button>
    </div>
  );
};

// Enhanced Draggable Widget with Manual Naming
const DraggableWidget = ({ 
  id, 
  children, 
  onDelete, 
  initialPosition = { x: 0, y: 0 }, 
  initialSize = { width: 300, height: 200 },
  onConfigureData,
  hasData,
  chartType,
  chartName,
  onNameChange
}) => {
  const [position, setPosition] = useState(initialPosition);
  const [size, setSize] = useState(initialSize);
  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);
  const [isEditingName, setIsEditingName] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const [resizeStart, setResizeStart] = useState({ x: 0, y: 0, width: 0, height: 0 });

  const handleMouseDown = useCallback((e) => {
    if (e.target.closest('.resize-handle') || e.target.closest('.widget-button') || e.target.closest('input')) return;
    setIsDragging(true);
    setDragStart({
      x: e.clientX - position.x,
      y: e.clientY - position.y
    });
  }, [position]);

  const handleResizeMouseDown = useCallback((e) => {
    e.stopPropagation();
    setIsResizing(true);
    setResizeStart({
      x: e.clientX,
      y: e.clientY,
      width: size.width,
      height: size.height
    });
  }, [size]);

  const handleMouseMove = useCallback((e) => {
    if (isDragging) {
      setPosition({
        x: Math.max(0, e.clientX - dragStart.x),
        y: Math.max(0, e.clientY - dragStart.y)
      });
    } else if (isResizing) {
      const newWidth = Math.max(200, resizeStart.width + (e.clientX - resizeStart.x));
      const newHeight = Math.max(150, resizeStart.height + (e.clientY - resizeStart.y));
      setSize({ width: newWidth, height: newHeight });
    }
  }, [isDragging, isResizing, dragStart, resizeStart]);

  const handleMouseUp = useCallback(() => {
    setIsDragging(false);
    setIsResizing(false);
  }, []);

  useEffect(() => {
    if (isDragging || isResizing) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
      return () => {
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
      };
    }
  }, [isDragging, isResizing, handleMouseMove, handleMouseUp]);

  const handleNameSave = (newName) => {
    onNameChange(id, newName);
    setIsEditingName(false);
  };

  return (
    <div
      className="absolute bg-white border border-gray-300 rounded-lg shadow-lg select-none"
      style={{
        left: position.x,
        top: position.y,
        width: size.width,
        height: size.height,
        cursor: isDragging ? 'grabbing' : 'default'
      }}
    >
      <div
        className="flex items-center justify-between p-2 bg-gray-50 border-b border-gray-200 rounded-t-lg cursor-grab hover:bg-gray-100"
        onMouseDown={handleMouseDown}
      >
        <div className="flex items-center gap-2 flex-1 min-w-0">
          <Move size={16} className="text-gray-500 flex-shrink-0" />
          {isEditingName ? (
            <ChartNameEditor
              currentName={chartName}
              onSave={handleNameSave}
              onCancel={() => setIsEditingName(false)}
            />
          ) : (
            <div className="flex items-center gap-2 min-w-0 flex-1">
              <span className="text-sm font-medium text-gray-700 truncate">
                {chartName}
              </span>
              <button
                onClick={() => setIsEditingName(true)}
                className="widget-button p-1 hover:bg-gray-200 rounded transition-colors flex-shrink-0"
                title="Edit chart name"
              >
                <Edit2 size={12} className="text-gray-500" />
              </button>
            </div>
          )}
          {hasData && (
            <div className="w-2 h-2 bg-green-500 rounded-full flex-shrink-0" title="Connected to data" />
          )}
        </div>
        <div className="flex items-center gap-1 flex-shrink-0">
          <button
            onClick={() => onConfigureData(id)}
            className="widget-button p-1 hover:bg-gray-200 rounded transition-colors"
            title="Configure Data"
          >
            <Database size={14} className="text-gray-500" />
          </button>
          <button
            onClick={() => onDelete(id)}
            className="widget-button p-1 hover:bg-gray-200 rounded transition-colors"
          >
            <X size={14} className="text-gray-500" />
          </button>
        </div>
      </div>

      <div className="p-2" style={{ height: size.height - 45 }}>
        {children}
      </div>

      <div
        className="resize-handle absolute bottom-0 right-0 w-4 h-4 cursor-se-resize hover:bg-gray-400 transition-colors opacity-50"
        onMouseDown={handleResizeMouseDown}
        style={{
          background: 'linear-gradient(-45deg, transparent 30%, #9ca3af 30%, #9ca3af 40%, transparent 40%, transparent 60%, #9ca3af 60%, #9ca3af 70%, transparent 70%)'
        }}
      />
    </div>
  );
};

// Chart Components - NO SAMPLE DATA
const BarChartWidget = ({ data, columns }) => {
  if (!data || !Array.isArray(data) || data.length === 0 || !columns || columns.length < 2) {
    return (
      <div className="h-full flex items-center justify-center text-gray-500 bg-gray-50 rounded">
        <div className="text-center">
          <BarChart3 size={48} className="mx-auto mb-2 text-gray-300" />
          <p className="text-sm">No data connected</p>
          <p className="text-xs">Upload data and configure columns</p>
        </div>
      </div>
    );
  }

  // Filter out null/undefined values and convert to numbers where needed
  const chartData = data.map(item => ({
    [columns[0]]: item[columns[0]],
    [columns[1]]: parseFloat(item[columns[1]]) || 0,
    ...(columns[2] && { [columns[2]]: parseFloat(item[columns[2]]) || 0 })
  })).filter(item => item[columns[0]] !== undefined);

  return (
    <ResponsiveContainer width="100%" height="100%">
      <BarChart data={chartData} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey={columns[0]} />
        <YAxis />
        <Tooltip />
        <Bar dataKey={columns[1]} fill="#3b82f6" name={columns[1]} />
        {columns[2] && <Bar dataKey={columns[2]} fill="#10b981" name={columns[2]} />}
      </BarChart>
    </ResponsiveContainer>
  );
};

const LineChartWidget = ({ data, columns }) => {
  if (!data || !Array.isArray(data) || data.length === 0) {
    return (
      <div className="h-full flex items-center justify-center text-gray-500 bg-gray-50 rounded">
        <div className="text-center">
          <TrendingUp size={48} className="mx-auto mb-2 text-gray-300" />
          <p className="text-sm">No data connected</p>
          <p className="text-xs">Upload data and configure columns</p>
        </div>
      </div>
    );
  }

  const columnMap = columns && columns.length > 0 ? columns : [];
  
  return (
    <ResponsiveContainer width="100%" height="100%">
      <LineChart data={data} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey={columnMap[0]} />
        <YAxis />
        <Tooltip />
        {columnMap[1] && <Line type="monotone" dataKey={columnMap[1]} stroke="#3b82f6" strokeWidth={3} name={columnMap[1]} />}
        {columnMap[2] && <Line type="monotone" dataKey={columnMap[2]} stroke="#10b981" strokeWidth={3} name={columnMap[2]} />}
      </LineChart>
    </ResponsiveContainer>
  );
};

const PieChartWidget = ({ data, columns }) => {
  if (!data || !Array.isArray(data) || data.length === 0 || !columns || columns.length < 2) {
    return (
      <div className="h-full flex items-center justify-center text-gray-500 bg-gray-50 rounded">
        <div className="text-center">
          <Activity size={48} className="mx-auto mb-2 text-gray-300" />
          <p className="text-sm">No data connected</p>
          <p className="text-xs">Upload data and configure columns</p>
        </div>
      </div>
    );
  }

  const pieData = data.map((item, index) => ({
    name: item[columns[0]] || `Item ${index + 1}`,
    value: parseFloat(item[columns[1]]) || 0,
    color: `hsl(${index * 137.5}, 70%, 50%)`
  }));

  return (
    <ResponsiveContainer width="100%" height="100%">
      <PieChart>
        <Pie
          data={pieData}
          cx="50%"
          cy="50%"
          labelLine={false}
          label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
          outerRadius={60}
          fill="#8884d8"
          dataKey="value"
        >
          {pieData.map((entry, index) => (
            <Cell key={`cell-${index}`} fill={entry.color} />
          ))}
        </Pie>
        <Tooltip />
      </PieChart>
    </ResponsiveContainer>
  );
};

const MetricCard = ({ data, column, title = "Metric" }) => {
  if (!data || !Array.isArray(data) || data.length === 0 || !column) {
    return (
      <div className="h-full flex items-center justify-center text-gray-500 bg-gray-50 rounded">
        <div className="text-center">
          <CreditCard size={48} className="mx-auto mb-2 text-gray-300" />
          <p className="text-sm">No data connected</p>
          <p className="text-xs">Upload data and configure column</p>
        </div>
      </div>
    );
  }

  const values = data.map(row => parseFloat(row[column])).filter(v => !isNaN(v));
  let value = "N/A";
  let change = null;

  if (values.length > 0) {
    const sum = values.reduce((a, b) => a + b, 0);
    value = sum.toLocaleString();
    
    if (values.length > 1) {
      const recent = values[values.length - 1];
      const previous = values[values.length - 2];
      change = ((recent - previous) / previous * 100).toFixed(1);
    }
  }

  return (
    <div className="h-full flex flex-col justify-center p-4 bg-gradient-to-br from-gray-50 to-white">
      <h3 className="text-sm font-medium text-gray-600 mb-2">{title}</h3>
      <div className="flex items-end gap-2">
        <span className="text-3xl font-bold text-blue-600">{value}</span>
        {change && (
          <span className={`text-sm font-medium ${parseFloat(change) >= 0 ? 'text-green-600' : 'text-red-600'}`}>
            {parseFloat(change) >= 0 ? '↗' : '↘'} {Math.abs(change)}%
          </span>
        )}
      </div>
    </div>
  );
};

// Main Dashboard Component
const PowerBIDashboard = () => {
  const [widgets, setWidgets] = useState([]);
  const [showUploader, setShowUploader] = useState(false);
  const [uploadedData, setUploadedData] = useState(null);
  const [showDataPreview, setShowDataPreview] = useState(false);
  const [configuringWidget, setConfiguringWidget] = useState(null);
  const [selectedColumns, setSelectedColumns] = useState([]);

  const getDefaultChartName = (type, id) => {
    const typeNames = {
      'bar': 'Bar Chart',
      'line': 'Line Chart',
      'pie': 'Pie Chart',
      'card': 'Metric Card'
    };
    return `${typeNames[type]} ${id}`;
  };

  const addWidget = (type) => {
    const newId = Date.now();
    const newWidget = {
      id: newId,
      type,
      name: getDefaultChartName(type, widgets.length + 1),
      position: { 
        x: Math.random() * 300 + 50, 
        y: Math.random() * 200 + 100 
      },
      size: type === 'card' ? { width: 250, height: 180 } : { width: 350, height: 280 },
      data: null,
      columns: []
    };
    setWidgets([...widgets, newWidget]);
  };

  const deleteWidget = (id) => {
    setWidgets(widgets.filter(widget => widget.id !== id));
  };

  const handleNameChange = (id, newName) => {
    setWidgets(widgets.map(widget => 
      widget.id === id ? { ...widget, name: newName } : widget
    ));
  };

  const handleDataLoaded = (data) => {
    setUploadedData(data);
    setShowDataPreview(true);
  };

  const configureWidgetData = (widgetId) => {
    if (!uploadedData) {
      setShowUploader(true);
      return;
    }
    setConfiguringWidget(widgetId);
    setSelectedColumns([]);
  };

  const applyDataToWidget = () => {
  if (!configuringWidget || !uploadedData) return;

  const widgetType = widgets.find(w => w.id === configuringWidget)?.type;
  const requiredColumns = widgetType === 'card' ? 1 : widgetType === 'pie' ? 2 : 1;
  
  if (selectedColumns.filter(col => col && col.trim() !== '').length < requiredColumns) {
    alert(`Please select at least ${requiredColumns} column(s) for this chart type`);
    return;
  }

  const updatedWidgets = widgets.map(widget => {
    if (widget.id === configuringWidget) {
      return {
        ...widget,
        data: uploadedData.data,
        columns: selectedColumns.slice(0, widgetType === 'card' ? 1 : widgetType === 'pie' ? 2 : 3)
      };
    }
    return widget;
  });

  setWidgets(updatedWidgets);
  setConfiguringWidget(null);
  setSelectedColumns([]);
};

  const renderWidget = (widget) => {
    const hasData = widget.data && Array.isArray(widget.data) && widget.columns && widget.columns.length > 0;
    
    switch (widget.type) {
      case 'bar':
        return <BarChartWidget data={widget.data} columns={widget.columns} />;
      case 'line':
        return <LineChartWidget data={widget.data} columns={widget.columns} />;
      case 'pie':
        return <PieChartWidget data={widget.data} columns={widget.columns} />;
      case 'card':
        return <MetricCard data={widget.data} column={widget.columns[0]} title={widget.columns[0] || "Metric"} />;
      default:
        return <div className="p-4 text-gray-500">Unknown widget type</div>;
    }
  };

  return (
    <div className="w-full h-[900px] bg-gray-100 relative overflow-hidden">
      {/* Header */}
      <div className="bg-white border-b border-gray-20 p-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-4">
          <h1 className="text-2xl font-bold text-gray-800">Power BI Dashboard</h1>
          {uploadedData && (
            <div className="flex items-center gap-2 px-3 py-1 bg-green-100 text-green-800 rounded-full text-sm">
              <FileSpreadsheet size={16} />
              {uploadedData.fileName}
              <button
                onClick={() => setShowDataPreview(true)}
                className="hover:underline"
              >
                Preview
              </button>
            </div>
          )}
        </div>
        
        <div className="flex gap-2">
          <button
            onClick={() => setShowUploader(true)}
            className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-all duration-200 shadow-sm hover:shadow-md"
          >
            <Upload size={16} />
            Upload Data
          </button>
          <button
            onClick={() => addWidget('bar')}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-all duration-200 shadow-sm hover:shadow-md"
          >
            <BarChart3 size={16} />
            Bar Chart
          </button>
          <button
            onClick={() => addWidget('line')}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-all duration-200 shadow-sm hover:shadow-md"
          >
            <TrendingUp size={16} />
            Line Chart
          </button>
          <button
            onClick={() => addWidget('pie')}
            className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-all duration-200 shadow-sm hover:shadow-md"
          >
            <Activity size={16} />
            Pie Chart
          </button>
          <button
            onClick={() => addWidget('card')}
            className="flex items-center gap-2 px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition-all duration-200 shadow-sm hover:shadow-md"
          >
            <CreditCard size={16} />
            Metric Card
          </button>
        </div>
      </div>

      {/* Canvas Area */}
      <div className="relative w-full bg-gray-50" style={{ height: 'calc(100vh - 80px)' }}>
        {widgets.map((widget) => (
          <DraggableWidget
            key={widget.id}
            id={widget.id}
            onDelete={deleteWidget}
            onConfigureData={configureWidgetData}
            onNameChange={handleNameChange}
            initialPosition={widget.position}
            initialSize={widget.size}
            hasData={widget.data && Array.isArray(widget.data) && widget.columns && widget.columns.length > 0}
            chartType={widget.type}
            chartName={widget.name}
          >
            {renderWidget(widget)}
          </DraggableWidget>
        ))}

        {widgets.length === 0 && (
          <div className="absolute inset-0 flex items-center justify-center text-gray-500">
            <div className="text-center">
              <Plus size={48} className="mx-auto mb-4 text-gray-400" />
              <p className="text-lg font-medium">Add widgets to get started</p>
              <p className="text-sm">Upload your Excel/CSV data and add charts to create your dashboard</p>
            </div>
          </div>
        )}
      </div>

      {/* File Uploader Modal */}
      {showUploader && (
        <FileUploader
          onDataLoaded={handleDataLoaded}
          onClose={() => setShowUploader(false)}
        />
      )}

      {/* Data Preview Modal */}
      {showDataPreview && uploadedData && (
        <DataPreview
          data={uploadedData}
          onClose={() => setShowDataPreview(false)}
        />
      )}

      {/* Column Configuration Modal */}
      {configuringWidget && uploadedData && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96 max-w-full mx-4">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-lg font-semibold">Configure Chart Data</h3>
              <button 
                onClick={() => setConfiguringWidget(null)} 
                className="text-gray-500 hover:text-gray-700"
              >
                <X size={20} />
              </button>
            </div>
            
            <ColumnSelector
              columns={uploadedData.columns}
              selectedColumns={selectedColumns}
              onSelectionChange={setSelectedColumns}
              chartType={widgets.find(w => w.id === configuringWidget)?.type}
            />
            
            <div className="flex gap-2 justify-end mt-4">
              <button
                onClick={() => setConfiguringWidget(null)}
                className="px-4 py-2 text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                onClick={applyDataToWidget}
                disabled={selectedColumns.filter(col => col && col.trim() !== '').length === 0}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
              >
                Apply
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PowerBIDashboard;














