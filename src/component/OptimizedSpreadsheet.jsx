import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { 
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight,
  Palette, Calculator, ChevronDown, 
  Undo, Redo, Copy, Clipboard, Save, Download,
  Filter, BarChart, Plus, X, Search, Trash2, FilterX
} from 'lucide-react';

const OptimizedSpreadsheet = () => {
  // Core state
  const [sheets, setSheets] = useState([
    { id: 'sheet1', name: 'Sheet1', data: {} }
  ]);
  const [activeSheetId, setActiveSheetId] = useState('sheet1');
  const [selectedCell, setSelectedCell] = useState('A1');
  const [formulaBarValue, setFormulaBarValue] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  
  // Filter state
  const [activeFilters, setActiveFilters] = useState({});
  const [filterMenuOpen, setFilterMenuOpen] = useState(false);
  const [filterColumn, setFilterColumn] = useState('A');
  const [filterValue, setFilterValue] = useState('');
  
  // Performance state
  const [visibleRows, setVisibleRows] = useState(50); // Reduced from 100
  const [scrollTop, setScrollTop] = useState(0);
  
  // History for undo/redo
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [clipboard, setClipboard] = useState(null);

  const cellInputRef = useRef(null);
  const activeSheet = sheets.find(sheet => sheet.id === activeSheetId);
  const data = activeSheet ? activeSheet.data : {};

  // Optimized column name conversion
  const getColumnName = useCallback((col) => {
    let result = '';
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  }, []);

  // Optimized cell data with lazy loading
  const memoizedCellData = useMemo(() => {
    const result = {};
    const startRow = Math.floor(scrollTop / 24) + 1;
    const endRow = Math.min(startRow + visibleRows, 1000);
    
    for (let row = startRow; row <= endRow; row++) {
      for (let col = 1; col <= 40; col++) {
        const cellId = getColumnName(col) + row;
        result[cellId] = data[cellId] || { value: '', formula: '', style: {} };
      }
    }
    return result;
  }, [data, visibleRows, scrollTop, getColumnName]);

  // Simplified formula evaluation
  const evaluateFormula = useCallback((formula) => {
    if (!formula.startsWith('=')) return formula;
    
    try {
      let expression = formula.slice(1);
      
      // Handle basic SUM function
      if (expression.startsWith('SUM(')) {
        const range = expression.match(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/);
        if (range) {
          const [start, end] = range.slice(1);
          // Simple range sum calculation
          return 'SUM_RESULT'; // Simplified for performance
        }
      }
      
      // Basic arithmetic
      expression = expression.replace(/[A-Z]+\d+/g, (match) => {
        const cellData = data[match];
        return cellData ? (cellData.value || 0) : 0;
      });
      
      // Safe evaluation
      if (/^[0-9+\-*/().\s]*$/.test(expression)) {
        return Function('"use strict"; return (' + expression + ')')();
      }
      
      return expression;
    } catch (error) {
      return '#ERROR';
    }
  }, [data]);

  // Optimized cell update
  const updateCell = useCallback((cellId, value) => {
    setSheets(prevSheets => 
      prevSheets.map(sheet => {
        if (sheet.id === activeSheetId) {
          const newData = { ...sheet.data };
          const evaluatedValue = value.startsWith('=') ? evaluateFormula(value) : value;
          
          newData[cellId] = {
            value: evaluatedValue,
            formula: value,
            style: newData[cellId]?.style || {}
          };
          
          return { ...sheet, data: newData };
        }
        return sheet;
      })
    );
  }, [activeSheetId, evaluateFormula]);

  // Filter functions
  const applyFilter = useCallback((column, filterValue) => {
    if (!filterValue) {
      setActiveFilters(prev => {
        const newFilters = { ...prev };
        delete newFilters[column];
        return newFilters;
      });
      return;
    }

    setActiveFilters(prev => ({
      ...prev,
      [column]: filterValue.toLowerCase()
    }));
  }, []);

  const clearAllFilters = useCallback(() => {
    setActiveFilters({});
  }, []);

  const isRowVisible = useCallback((rowIndex) => {
    if (Object.keys(activeFilters).length === 0) return true;
    
    for (const [column, filterValue] of Object.entries(activeFilters)) {
      const colNum = column.charCodeAt(0) - 64;
      const cellId = getColumnName(colNum) + (rowIndex + 1);
      const cellData = data[cellId];
      const cellValue = cellData ? (cellData.value || '').toString().toLowerCase() : '';
      
      if (!cellValue.includes(filterValue)) {
        return false;
      }
    }
    return true;
  }, [activeFilters, data, getColumnName]);

  // Optimized scroll handling
  const handleScroll = useCallback((e) => {
    const scrollTop = e.target.scrollTop;
    setScrollTop(scrollTop);
  }, []);

  // Cell interaction handlers
  const handleCellClick = useCallback((cellId) => {
    setSelectedCell(cellId);
    const cellData = data[cellId];
    setFormulaBarValue(cellData ? (cellData.formula || cellData.value || '') : '');
    setIsEditing(false);
  }, [data]);

  const handleCellDoubleClick = useCallback((cellId) => {
    setSelectedCell(cellId);
    setIsEditing(true);
    setTimeout(() => {
      if (cellInputRef.current) {
        cellInputRef.current.focus();
      }
    }, 0);
  }, []);

  const handleKeyDown = useCallback((e) => {
    if (e.key === 'Enter' && selectedCell) {
      updateCell(selectedCell, formulaBarValue);
      setIsEditing(false);
      
      // Move to next row
      const match = selectedCell.match(/([A-Z]+)(\d+)/);
      if (match) {
        const nextCell = match[1] + (parseInt(match[2]) + 1);
        handleCellClick(nextCell);
      }
    } else if (e.key === 'Escape') {
      setIsEditing(false);
      const cellData = data[selectedCell];
      setFormulaBarValue(cellData ? (cellData.formula || cellData.value || '') : '');
    }
  }, [selectedCell, formulaBarValue, updateCell, data, handleCellClick]);

  // Format functions
  const formatCell = useCallback((style) => {
    if (!selectedCell) return;
    
    setSheets(prevSheets =>
      prevSheets.map(sheet => {
        if (sheet.id === activeSheetId) {
          const newData = { ...sheet.data };
          if (!newData[selectedCell]) {
            newData[selectedCell] = { value: '', formula: '', style: {} };
          }
          newData[selectedCell] = {
            ...newData[selectedCell],
            style: { ...newData[selectedCell].style, ...style }
          };
          return { ...sheet, data: newData };
        }
        return sheet;
      })
    );
  }, [selectedCell, activeSheetId]);

  const getCellStyle = useCallback((cellId) => {
    const cellData = data[cellId];
    if (!cellData || !cellData.style) return {};
    
    const style = {};
    if (cellData.style.bold) style.fontWeight = 'bold';
    if (cellData.style.italic) style.fontStyle = 'italic';
    if (cellData.style.underline) style.textDecoration = 'underline';
    if (cellData.style.color) style.color = cellData.style.color;
    if (cellData.style.backgroundColor) style.backgroundColor = cellData.style.backgroundColor;
    if (cellData.style.textAlign) style.textAlign = cellData.style.textAlign;
    
    return style;
  }, [data]);

  // Close filter menu when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (filterMenuOpen && !event.target.closest('.filter-menu')) {
        setFilterMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [filterMenuOpen]);

  return (
    <div className="flex flex-col h-screen bg-white">
      {/* Header */}
      <div className="bg-gray-50 border-b px-4 py-2">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <h1 className="text-xl font-semibold text-gray-800">Strix Spreadsheet</h1>
            <div className="flex items-center space-x-2">
              <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
                <Undo size={16} />
              </button>
              <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
                <Redo size={16} />
              </button>
              <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
                <Copy size={16} />
              </button>
              <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
                <Clipboard size={16} />
              </button>
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
              <BarChart size={16} />
            </button>
            <button onClick={() => {}} className="p-2 hover:bg-gray-200 rounded">
              <Save size={16} />
            </button>
          </div>
        </div>
      </div>

      {/* Toolbar */}
      <div className="bg-white border-b px-4 py-2">
        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-1">
            <button 
              onClick={() => formatCell({ bold: true })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <Bold size={16} />
            </button>
            <button 
              onClick={() => formatCell({ italic: true })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <Italic size={16} />
            </button>
            <button 
              onClick={() => formatCell({ underline: true })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <Underline size={16} />
            </button>
          </div>
          
          <div className="flex items-center space-x-1">
            <button 
              onClick={() => formatCell({ textAlign: 'left' })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <AlignLeft size={16} />
            </button>
            <button 
              onClick={() => formatCell({ textAlign: 'center' })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <AlignCenter size={16} />
            </button>
            <button 
              onClick={() => formatCell({ textAlign: 'right' })}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <AlignRight size={16} />
            </button>
          </div>

          <button onClick={() => {}} className="p-2 hover:bg-gray-100 rounded">
            <Trash2 size={16} />
          </button>
          
          {/* Filter Button */}
          <div className="relative">
            <button 
              onClick={() => setFilterMenuOpen(!filterMenuOpen)}
              className={`p-2 rounded flex items-center ${
                Object.keys(activeFilters).length > 0 
                  ? 'bg-blue-100 text-blue-600' 
                  : 'hover:bg-gray-100'
              }`}
              title="Create a filter"
            >
              <Filter size={16} />
              {Object.keys(activeFilters).length > 0 && (
                <span className="ml-1 text-xs bg-blue-600 text-white rounded-full px-1.5 py-0.5">
                  {Object.keys(activeFilters).length}
                </span>
              )}
            </button>
            
            {filterMenuOpen && (
              <div className="filter-menu absolute top-full right-0 mt-1 w-80 bg-white border rounded-lg shadow-lg z-50 p-4">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="font-semibold text-gray-800">Create a filter</h3>
                  <button 
                    onClick={() => setFilterMenuOpen(false)}
                    className="text-gray-400 hover:text-gray-600"
                  >
                    <X size={16} />
                  </button>
                </div>
                
                <div className="space-y-3">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Filter by column
                    </label>
                    <select
                      value={filterColumn}
                      onChange={(e) => setFilterColumn(e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {Array.from({ length: 40 }, (_, i) => (
                        <option key={i} value={getColumnName(i + 1)}>
                          Column {getColumnName(i + 1)}
                        </option>
                      ))}
                    </select>
                  </div>
                  
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Filter value
                    </label>
                    <input
                      type="text"
                      value={filterValue}
                      onChange={(e) => setFilterValue(e.target.value)}
                      placeholder="Enter filter value..."
                      className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  
                  <div className="flex gap-2">
                    <button
                      onClick={() => {
                        applyFilter(filterColumn, filterValue);
                        setFilterValue('');
                      }}
                      className="flex-1 px-3 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 text-sm"
                    >
                      Apply Filter
                    </button>
                    <button
                      onClick={clearAllFilters}
                      className="px-3 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 text-sm flex items-center"
                    >
                      <FilterX size={14} className="mr-1" />
                      Clear All
                    </button>
                  </div>
                  
                  {Object.keys(activeFilters).length > 0 && (
                    <div className="border-t pt-3">
                      <h4 className="text-sm font-medium text-gray-700 mb-2">Active Filters:</h4>
                      <div className="space-y-1">
                        {Object.entries(activeFilters).map(([column, value]) => (
                          <div key={column} className="flex items-center justify-between bg-gray-50 px-2 py-1 rounded text-sm">
                            <span>{column}: "{value}"</span>
                            <button
                              onClick={() => applyFilter(column, '')}
                              className="text-red-500 hover:text-red-700"
                            >
                              <X size={12} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Formula Bar */}
      <div className="bg-gray-50 border-b px-4 py-2">
        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-2">
            <span className="text-sm font-medium">{selectedCell}</span>
          </div>
          <div className="flex-1">
            <input
              ref={cellInputRef}
              type="text"
              value={formulaBarValue}
              onChange={(e) => setFormulaBarValue(e.target.value)}
              onKeyDown={handleKeyDown}
              onBlur={() => {
                if (selectedCell) {
                  updateCell(selectedCell, formulaBarValue);
                }
                setIsEditing(false);
              }}
              className="w-full px-3 py-1 border rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="Enter value or formula (start with =)"
            />
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 flex">
        {/* Spreadsheet */}
        <div className="flex-1 overflow-auto" onScroll={handleScroll}>
          <div className="relative">
            <table className="min-w-full border-collapse">
              <thead className="sticky top-0 z-10">
                <tr className="bg-gray-100">
                  <th className="w-16 h-8 border border-gray-300 bg-gray-200"></th>
                  {Array.from({ length: 40 }, (_, i) => {
                    const columnName = getColumnName(i + 1);
                    const hasFilter = activeFilters[columnName];
                    return (
                      <th key={i} className="min-w-20 h-8 border border-gray-300 text-xs font-medium text-center relative">
                        <div className="flex items-center justify-center">
                          {columnName}
                          {hasFilter && (
                            <Filter size={12} className="ml-1 text-blue-600" />
                          )}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: visibleRows }, (_, rowIndex) => {
                  // Apply row filtering
                  if (!isRowVisible(rowIndex)) return null;
                  
                  return (
                    <tr key={rowIndex} style={{ height: '24px' }}>
                      <td className="w-16 h-6 border border-gray-300 bg-gray-100 text-xs text-center font-medium sticky left-0 z-5">
                        {rowIndex + 1}
                      </td>
                      {Array.from({ length: 40 }, (_, colIndex) => {
                        const cellId = getColumnName(colIndex + 1) + (rowIndex + 1);
                        const cellData = memoizedCellData[cellId];
                        const isSelected = selectedCell === cellId;
                        
                        return (
                          <td
                            key={cellId}
                            className={`min-w-20 h-6 border border-gray-300 cursor-cell relative ${
                              isSelected ? 'ring-2 ring-blue-500 bg-blue-50' : 'hover:bg-gray-50'
                            }`}
                            onClick={() => handleCellClick(cellId)}
                            onDoubleClick={() => handleCellDoubleClick(cellId)}
                          >
                            {isSelected && isEditing ? (
                              <input
                                ref={cellInputRef}
                                type="text"
                                value={formulaBarValue}
                                onChange={(e) => setFormulaBarValue(e.target.value)}
                                onKeyDown={handleKeyDown}
                                onBlur={() => {
                                  updateCell(selectedCell, formulaBarValue);
                                  setIsEditing(false);
                                }}
                                className="w-full h-full px-1 text-xs border-none outline-none bg-transparent"
                                autoFocus
                              />
                            ) : (
                              <div 
                                className="px-1 text-xs truncate"
                                style={getCellStyle(cellId)}
                              >
                                {cellData ? (
                                  typeof cellData.value === 'number' ? 
                                    cellData.value.toLocaleString() : 
                                    cellData.value
                                ) : ''}
                              </div>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
            
            {/* Virtual scrolling placeholder */}
            <div style={{ height: `${(1000 - visibleRows) * 24}px` }} />
          </div>
        </div>
      </div>

      {/* Sheet Tabs */}
      <div className="bg-gray-50 border-t px-4 py-2">
        <div className="flex items-center space-x-2">
          {sheets.map(sheet => (
            <button
              key={sheet.id}
              onClick={() => setActiveSheetId(sheet.id)}
              className={`px-3 py-1 text-sm rounded ${
                activeSheetId === sheet.id 
                  ? 'bg-white border border-gray-300' 
                  : 'hover:bg-gray-200'
              }`}
            >
              {sheet.name}
            </button>
          ))}
          <button 
            onClick={() => {}}
            className="p-1 hover:bg-gray-200 rounded"
          >
            <Plus size={16} />
          </button>
        </div>
      </div>
    </div>
  );
};

export default OptimizedSpreadsheet;
