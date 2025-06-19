import React, { useState, useEffect, useCallback,useRef } from 'react';
import {
Download, Upload, Share2, BarChart3, Calculator, Grid, Save, Plus, Trash2,
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, PieChart,
  LineChart, TrendingUp, FunctionSquare, X, Search, DollarSign, Percent, Palette, 
  Calendar, ChevronDown, WrapText, Mail, Link2, Users, Undo, Redo, Copy, 
  FileText, Folder, Printer, Clock, Undo2, Redo2, Scissors, Grid3X3, Eye, 
  ZoomIn, ZoomOut, Maximize, ArrowUpDown, Filter, Shield, Table, Settings, 
  Puzzle, HelpCircle, BookOpen, Clipboard, Image,  Keyboard
} from 'lucide-react';

import { useSpreadsheetData } from '../context/SpreadsheetDataContext';
import { useNavigate,Link } from 'react-router-dom';
const functionCategories = [
  {
    name: 'Math',
    functions: [
      'ABS', 'ACOS', 'ASIN', 'ATAN', 'ATAN2', 'CEILING', 'COS', 'DEGREES',
      'EXP', 'FACT', 'FLOOR', 'GCD', 'LCM', 'LN', 'LOG', 'LOG10', 'MAX',
      'MIN', 'MOD', 'PI', 'POWER', 'PRODUCT', 'RADIANS', 'RAND', 'RANDBETWEEN',
      'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'SIGN', 'SIN', 'SQRT', 'SUM',
      'SUMPRODUCT', 'TAN','TANradian', 'TRUNC'
    ]
  },
  {
    name: 'Financial',
    functions: ['FV', 'PV', 'NPV', 'PMT', 'IPMT', 'PPMT', 'NPER', 'RATE',
      'IRR', 'MIRR', 'SLN', 'SYD', 'DB', 'DDB', 'VDB']
  },
  {
    name: 'Date',
    functions: [
      'NOW', 'TODAY', 'DATE', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360',
      'EDATE', 'EOMONTH', 'HOUR', 'MINUTE', 'MONTH', 'NETWORKDAYS',
      'SECOND', 'TIME', 'TIMEVALUE', 'WEEKDAY', 'WEEKNUM',
      'WORKDAY', 'YEAR', 'YEARFRAC'
    ]
  },
  {
    name: 'Text',
    functions: [
      'CHAR', 'CLEAN', 'CODE', 'CONCATENATE', 'EXACT', 'FIND', 'LEFT',
      'LEN', 'LOWER', 'MID', 'PROPER', 'REPLACE', 'REPT', 'RIGHT',
      'SEARCH', 'SUBSTITUTE', 'T', 'TEXT','TEXTNOW','CLEAN', 'TRIM', 'UPPER', 'VALUE'
    ] 
  },
  {
    name: 'Combinatorial',
    functions: ['COMBIN', 'PERMUT']
  },
  {
    name: 'Distribution',
    functions: [
      'NORMDIST', 'NORMSDIST', 'NORMINV', 'NORMSINV', 'POISSON', 'WEIBULL'
    ]
  },
  {
    name: 'Logical',
    functions: [
      'AND', 'FALSE', 'IF', 'IFERROR', 'IFS', 'NOT',
      'OR', 'SWITCH', 'TRUE', 'XOR'
    ]
  },
  {
    name: 'Statistical',
    functions: [
      'AVERAGE', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'MEDIAN', 'MODE',
      'STDEV', 'STDEVP', 'VAR', 'VARP', 'STDEVA', 'STDEVPA', 'VARA', 'VARPA',
      'CORREL', 'COVAR', 'GEOMEAN', 'HARMEAN', 'KURT', 'SKEW', 'SLOPE',
      'INTERCEPT', 'RSQ', 'FORECAST', 'TREND', 'PERCENTILE', 'PERCENTRANK',
      'QUARTILE', 'RANK', 'LARGE', 'SMALL', 'DEVSQ', 'TRIMMEAN'
    ]
  },
  {
    name: 'Array',
    functions: [
      'ARRAYFORMULA', 'FILTER', 'FLATTEN',
      'SORT', 'SORTN', 'UNIQUE'
    ]
  },
   {
    name: 'Lookup',
    functions: ['CHOOSE', 'HLOOKUP', 'INDEX', 'MATCH', 'VLOOKUP']
  },
  {
    name: 'Information',
    functions: [
      'CELL', 'ERROR_TYPE', 'INFO', 'ISBLANK', 'ISERR', 'ISERROR', 'ISEVEN',
      'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISODD',
      'ISREF', 'ISTEXT', 'N', 'NA', 'SHEET', 'SHEETS', 'TYPE'
    ]
  }
];

export default function SpreadsheetApp(){
  const [activeMenu, setActiveMenu] = useState(null);
const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 });
const menuRef = useRef(null);
  const [history, setHistory] = useState([{}]); // Initial empty state
  const [sheetName, setSheetName] = useState('Untitled spreadsheet');
const [clipboard, setClipboard] = useState(null);
const [findReplace, setFindReplace] = useState({ show: false, find: '', replace: '' });
const [filterActive, setFilterActive] = useState(false);
const [sortDialog, setSortDialog] = useState({ show: false, column: 0, ascending: true });
const [chartDialog, setChartDialog] = useState({ show: false, type: 'line' });
const [chartData, setChartData] = useState([]);
const [zoom, setZoom] = useState(100);
const [showGridlines, setShowGridlines] = useState(true);
const [showFormulas, setShowFormulas] = useState(false);
const [notification, setNotification] = useState(null);
const fileInputRef = useRef(null);
const [historyIndex, setHistoryIndex] = useState(0);
  const [cells, setCells] = useState({});
  const navigate = useNavigate();
 const [showShareDropdown, setShowShareDropdown] = useState(false);
  const [isCopied, setIsCopied] = useState(false);
  const dropdownRef = useRef(null);
  const [isOpen, setIsOpen] = useState(false);
  const [selectedCell, setSelectedCell] = useState('A1');
  const [selectedCategory, setSelectedCategory] = useState(null);
const [isHoveringSub, setIsHoveringSub] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);
  const [formulaBar, setFormulaBar] = useState('');
  const [sheets, setSheets] = useState([{ id: 1, name: 'Sheet1', active: true }]);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [activeSheet, setActiveSheet] = useState(1);
  const [showChart, setShowChart] = useState(false);
  const [chartType, setChartType] = useState('bar');
  const [showFunctions, setShowFunctions] = useState(false);
  const [functionSearch, setFunctionSearch] = useState('');
  const [selectedRange, setSelectedRange] = useState('');
  const formulaBtnRef = useRef(null);
  // Add these state variables at the top of your component
const [colWidths, setColWidths] = useState({});
const [rowHeights, setRowHeights] = useState({});
const [resizing, setResizing] = useState({ active: false, type: null, index: null, startPos: 0 });
const [dragPos, setDragPos] = useState(null);

// Default dimensions
const DEFAULT_COL_WIDTH = 80;
const DEFAULT_ROW_HEIGHT = 24;

// Helper functions for resizing
const getColWidth = (col) => colWidths[col] || DEFAULT_COL_WIDTH;
const getRowHeight = (row) => rowHeights[row] || DEFAULT_ROW_HEIGHT;
const getColumnLabel = (index) => String.fromCharCode(65 + index);
  
  const getCellKey = (row, col) => `${row}-${col}`;
  
  const getCellValue = (row, col) => {
    const key = getCellKey(row, col);
    return cellData[key] || '';
  };

  const getCellStyle = (row, col) => {
    const key = getCellKey(row, col);
    return cellStyles[key] || {};
  };

  const setCellValue = (row, col, value) => {
    const key = getCellKey(row, col);
    const newData = { ...cellData, [key]: value };
    setCellData(newData);
    addToHistory(newData, cellStyles);
  };

  const setCellStyle = (row, col, style) => {
    const key = getCellKey(row, col);
    const newStyles = { ...cellStyles, [key]: { ...cellStyles[key], ...style } };
    setCellStyles(newStyles);
    addToHistory(cellData, newStyles);
  };

  const addToHistory = (data, styles) => {
    const newHistory = history.slice(0, historyIndex + 1);
    newHistory.push({ data, styles });
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
  };

  const showNotification = (message) => {
    setNotification(message);
    setTimeout(() => setNotification(null), 3000);
  };


// Call this whenever cells change to save to history
const saveToHistory = (currentCells) => {
  // Don't save if no changes from current state
  if (JSON.stringify(currentCells) === JSON.stringify(history[historyIndex])) {
    return;
  }

  // If we're not at the end of history, truncate future states
  const newHistory = history.slice(0, historyIndex + 1);
  
  setHistory([...newHistory, JSON.parse(JSON.stringify(currentCells))]);
  setHistoryIndex(newHistory.length);
};
const handleMenuClick = (menuName, event) => {
  if (activeMenu === menuName) {
    setActiveMenu(null);
  } else {
    const rect = event.currentTarget.getBoundingClientRect();
    setMenuPosition({
      top: rect.bottom + window.scrollY,
      left: rect.left + window.scrollX
    });
    setActiveMenu(menuName);
  }
};
useEffect(() => {
  const handleClickOutside = (event) => {
    if (menuRef.current && !menuRef.current.contains(event.target)) {
      setActiveMenu(null);
    }
  };

  document.addEventListener('mousedown', handleClickOutside);
  return () => {
    document.removeEventListener('mousedown', handleClickOutside);
  };
}, []);
const undo = () => {
  if (historyIndex > 0) {
    setHistoryIndex(historyIndex - 1);
    setCells(history[historyIndex - 1]);
  }
};

const redo = () => {
  if (historyIndex < history.length - 1) {
    setHistoryIndex(historyIndex + 1);
    setCells(history[historyIndex + 1]);
  }
};
// Mouse handlers for resizing
const handleResizeMouseDown = (type, index, e) => {
  setResizing({
    active: true,
    type,
    index,
    startPos: type === 'col' ? e.clientX : e.clientY
  });
  setDragPos(type === 'col' ? e.clientX : e.clientY);
  e.preventDefault();
  e.stopPropagation();
};
useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);
useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setShowShareDropdown(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);



const { setSpreadsheetData, setDataSource } = useSpreadsheetData();

  // Your existing state and functions...

 const handleGenerateChart = () => {
  // Track which columns have non-zero values
  const activeColumns = new Set();
  const allHeaders = [];
  
  // First pass: Scan all data to find columns with actual values
  for (let row = 2; row <= 101; row++) {
    for (let col = 0; col < 26; col++) {
      const cellRef = `${String.fromCharCode(65 + col)}${row}`;
      const value = cells[cellRef]?.display;
      
      // Check if value exists and is not zero/empty
      if (value && value !== "0" && value !== 0 && value !== "0.0") {
        activeColumns.add(col);
      }
    }
  }

  // Second pass: Collect only relevant headers
  const headers = [];
  for (let col = 0; col < 26; col++) {
    if (activeColumns.has(col)) {
      const cellRef = `${String.fromCharCode(65 + col)}1`;
      headers.push({
        index: col,
        name: cells[cellRef]?.display || `Column ${col + 1}`
      });
    }
  }

  // Third pass: Build the final dataset
  const chartData = [];
  for (let row = 2; row <= 101; row++) {
    const rowData = {};
    let hasData = false;

    headers.forEach(header => {
      const cellRef = `${String.fromCharCode(65 + header.index)}${row}`;
      const value = cells[cellRef]?.display;
      
      // Only include if value exists and is not zero
      if (value && value !== "0" && value !== 0) {
        rowData[header.name] = isNaN(value) ? value : Number(value);
        hasData = true;
      }
    });

    if (hasData) chartData.push(rowData);
  }

  // Set the final data and navigate
  setSpreadsheetData(chartData);
  setDataSource('spreadsheet');
  navigate('/charts');
};

const handleResizeMouseMove = (e) => {
  if (!resizing.active) return;
  
  const pos = resizing.type === 'col' ? e.clientX : e.clientY;
  setDragPos(pos);
  
  if (Math.abs(pos - resizing.startPos) > 5) {
    const delta = pos - resizing.startPos;
    
    if (resizing.type === 'col') {
      const newWidth = Math.max(20, (colWidths[resizing.index] || DEFAULT_COL_WIDTH) + delta);
      setColWidths(prev => ({ ...prev, [resizing.index]: newWidth }));
    } else {
      const newHeight = Math.max(20, (rowHeights[resizing.index] || DEFAULT_ROW_HEIGHT) + delta);
      setRowHeights(prev => ({ ...prev, [resizing.index]: newHeight }));
    }
    
    setResizing(prev => ({ ...prev, startPos: pos }));
  }
};

const handleResizeMouseUp = () => {
  setResizing({ active: false, type: null, index: null, startPos: 0 });
  setDragPos(null);
};

// Add these event listeners to useEffect
useEffect(() => {
  document.addEventListener('mousemove', handleResizeMouseMove);
  document.addEventListener('mouseup', handleResizeMouseUp);
  return () => {
    document.removeEventListener('mousemove', handleResizeMouseMove);
    document.removeEventListener('mouseup', handleResizeMouseUp);
  };
}, [resizing]);

  const ROWS = 100;
  const COLS = 26;

  const getColumnHeader = (index) => String.fromCharCode(65 + index);

  const cellToCoords = (cellRef) => {
    const col = cellRef.charCodeAt(0) - 65;
    const row = parseInt(cellRef.slice(1)) - 1;
    return { row, col };
  };

  const coordsToCell = (row, col) => {
    return getColumnHeader(col) + (row + 1);
  };

  useEffect(() => {
    const initialCells = {};
    for (let row = 0; row < ROWS; row++) {
      for (let col = 0; col < COLS; col++) {
        const cellRef = coordsToCell(row, col);
        initialCells[cellRef] = { value: '', formula: '', display: '' };
      }
    }
    setCells(initialCells);
  }, []);

  const getCellNumericValue = (cellRef) => {
    const cell = cells[cellRef];
    if (!cell) return 0;
    
    const value = cell.display || cell.value || '';
    const numValue = parseFloat(value);
    return isNaN(numValue) ? 0 : numValue;
  };

 const evaluateFormula = useCallback((formula, cellRef) => {
  if (!formula.startsWith('=')) return formula;
  
  try {
    let expression = formula.slice(1).toUpperCase();
    
    // Helper function to evaluate cell references in functions
    const evaluateCellReference = (ref, currentCell) => {
      if (/^[A-Z]+\d+$/.test(ref)) {
        const cell = cells[ref];
        if (!cell) return '';
        return cell.display || cell.value || '';
      }
      return ref.replace(/^"(.*)"$/, '$1'); // Remove surrounding quotes if present
    };

    // Range-based functions
    const handleRangeFunction = (funcName, callback) => {
      const regex = new RegExp(`${funcName}\\(([A-Z]\\d+):([A-Z]\\d+)\\)`, 'g');
      return expression.replace(regex, (match, startCell, endCell) => {
        const startCoords = cellToCoords(startCell);
        const endCoords = cellToCoords(endCell);
        
        const values = [];
        for (let row = startCoords.row; row <= endCoords.row; row++) {
          for (let col = startCoords.col; col <= endCoords.col; col++) {
            const cellKey = coordsToCell(row, col);
            if (cellKey !== cellRef) {
              const cellValue = getCellNumericValue(cellKey);
              if (cellValue !== 0 || cells[cellKey]?.value === '0') {
                values.push(cellValue);
              }
            }
          }
        }
        return callback(values);
      });
    };

    // Math Functions
    expression = handleRangeFunction('SUM', (values) => values.reduce((sum, val) => sum + val, 0).toString());
    expression = handleRangeFunction('AVERAGE', (values) => values.length > 0 ? (values.reduce((sum, val) => sum + val, 0) / values.length).toFixed(2) : '0');
    expression = handleRangeFunction('COUNT', (values) => values.length.toString());
    expression = handleRangeFunction('MAX', (values) => values.length > 0 ? Math.max(...values).toString() : '0');
    expression = handleRangeFunction('MIN', (values) => values.length > 0 ? Math.min(...values).toString() : '0');
    expression = handleRangeFunction('PRODUCT', (values) => values.reduce((prod, val) => prod * val, 1).toString());

    expression = expression.replace(/ABS\(([^)]+)\)/g, (_, num) => Math.abs(parseFloat(num)).toString());
    expression = expression.replace(/SQRT\(([^)]+)\)/g, (_, num) => {
      const n = parseFloat(num);
      return n >= 0 ? Math.sqrt(n).toString() : '#ERROR!';
    });
    expression = expression.replace(/POWER\(([^,]+),\s*([^)]+)\)/g, (_, base, exp) => Math.pow(parseFloat(base), parseFloat(exp)).toString());
    expression = expression.replace(/ROUND\(([^,]+),\s*(\d+)\)/g, (_, num, digits) => (Math.round(parseFloat(num) * Math.pow(10, parseInt(digits))) / Math.pow(10, parseInt(digits))).toString());
    expression = expression.replace(/ROUNDUP\(([^,]+),\s*(\d+)\)/g, (_, num, digits) => Math.ceil(parseFloat(num) * Math.pow(10, parseInt(digits))) / Math.pow(10, parseInt(digits)).toString());
    expression = expression.replace(/ROUNDDOWN\(([^,]+),\s*(\d+)\)/g, (_, num, digits) => Math.floor(parseFloat(num) * Math.pow(10, parseInt(digits))) / Math.pow(10, parseInt(digits)).toString());
    expression = expression.replace(/CEILING\(([^)]+)\)/g, (_, num) => Math.ceil(parseFloat(num)).toString());
    expression = expression.replace(/FLOOR\(([^)]+)\)/g, (_, num) => Math.floor(parseFloat(num)).toString());
    expression = expression.replace(/TRUNC\(([^)]+)\)/g, (_, num) => Math.trunc(parseFloat(num)).toString());
    expression = expression.replace(/MOD\(([^,]+),\s*([^)]+)\)/g, (_, num, divisor) => (parseFloat(num) % parseFloat(divisor)).toString());
    expression = expression.replace(/SIGN\(([^)]+)\)/g, (_, num) => Math.sign(parseFloat(num)).toString());
    expression = expression.replace(/PI\(\)/g, Math.PI.toString());
    expression = expression.replace(/RAND\(\)/g, Math.random().toString());
    expression = expression.replace(/RANDBETWEEN\(([^,]+),\s*([^)]+)\)/g, (_, low, high) => (Math.floor(Math.random() * (parseInt(high) - parseInt(low) + 1)) + parseInt(low)).toString());

    // Trigonometric Functions
    expression = expression.replace(/SIN\(([^)]+)\)/g, (_, num) => Math.sin(parseFloat(num)).toString());
    expression = expression.replace(/COS\(([^)]+)\)/g, (_, num) => Math.cos(parseFloat(num)).toString());
    expression = expression.replace(/TAN\(([^)]+)\)/g, (_, num) => Math.tan(parseFloat(num)).toString());
    expression = expression.replace(/ASIN\(([^)]+)\)/g, (_, num) => Math.asin(parseFloat(num)).toString());
    expression = expression.replace(/ACOS\(([^)]+)\)/g, (_, num) => Math.acos(parseFloat(num)).toString());
    expression = expression.replace(/ATAN\(([^)]+)\)/g, (_, num) => Math.atan(parseFloat(num)).toString());
    expression = expression.replace(/ATAN2\(([^,]+),\s*([^)]+)\)/g, (_, y, x) => Math.atan2(parseFloat(y), parseFloat(x)).toString());
    expression = expression.replace(/DEGREES\(([^)]+)\)/g, (_, radians) => (parseFloat(radians) * 180 / Math.PI).toString());
    expression = expression.replace(/RADIANS\(([^)]+)\)/g, (_, degrees) => (parseFloat(degrees) * Math.PI / 180).toString());

    // Financial Functions
    if (expression.includes('PV(')) {
      const match = expression.match(/PV\(([^)]+)\)/);
      if (match) {
        const params = match[1].split(',').map(p => parseFloat(p.trim()));
        if (params.length >= 3) {
          const [rate, nper, pmt, fv = 0, type = 0] = params;
          if (rate === 0) return (-pmt * nper - fv).toFixed(2);
          const pv = pmt * ((1 - Math.pow(1 + rate, -nper)) / rate) + fv / Math.pow(1 + rate, nper);
          return (-pv).toFixed(2);
        }
      }
    }

    if (expression.includes('FV(')) {
      const match = expression.match(/FV\(([^)]+)\)/);
      if (match) {
        const params = match[1].split(',').map(p => parseFloat(p.trim()));
        if (params.length >= 3) {
          const [rate, nper, pmt, pv = 0, type = 0] = params;
          if (rate === 0) return (-pmt * nper - pv).toFixed(2);
          const fv = pmt * (Math.pow(1 + rate, nper) - 1) / rate - pv * Math.pow(1 + rate, nper);
          return fv.toFixed(2);
        }
      }
    }

    if (expression.includes('PMT(')) {
      const match = expression.match(/PMT\(([^)]+)\)/);
      if (match) {
        const params = match[1].split(',').map(p => parseFloat(p.trim()));
        if (params.length >= 3) {
          const [rate, nper, pv, fv = 0, type = 0] = params;
          if (rate === 0) return ((-pv - fv) / nper).toFixed(2);
          const pmt = (rate * (pv * Math.pow(1 + rate, nper) + fv)) / (Math.pow(1 + rate, nper) - 1);
          return (-pmt).toFixed(2);
        }
      }
    }

    if (expression.includes('SLN(')) {
      const match = expression.match(/SLN\(([^)]+)\)/);
      if (match) {
        const params = match[1].split(',').map(p => parseFloat(p.trim()));
        if (params.length >= 3) {
          const [cost, salvage, life] = params;
          return ((cost - salvage) / life).toFixed(2);
        }
      }
    }

    // Date Functions
    if (expression.includes('TODAY()')) return new Date().toLocaleDateString();
    if (expression.includes('NOW()')) return new Date().toLocaleString();
    if (expression.includes('DATE(')) {
      const match = expression.match(/DATE\(([^)]+)\)/);
      if (match) {
        const [year, month, day] = match[1].split(',').map(p => parseInt(p.trim()));
        return new Date(year, month - 1, day).toLocaleDateString();
      }
    }

    // Text Functions
    expression = expression.replace(/UPPER\(([^)]+)\)/g, (_, text) => evaluateCellReference(text, cellRef).toUpperCase());
    expression = expression.replace(/LOWER\(([^)]+)\)/g, (_, text) => evaluateCellReference(text, cellRef).toLowerCase());
    expression = expression.replace(/PROPER\(([^)]+)\)/g, (_, text) => 
      evaluateCellReference(text, cellRef).replace(/\w\S*/g, word => word.charAt(0).toUpperCase() + word.substr(1).toLowerCase())
    );
    expression = expression.replace(/TRIM\(([^)]+)\)/g, (_, text) => evaluateCellReference(text, cellRef).trim());
    expression = expression.replace(/CLEAN\(([^)]+)\)/g, (_, text) => evaluateCellReference(text, cellRef).replace(/[\x00-\x1F\x7F]/g, ''));
    expression = expression.replace(/LEN\(([^)]+)\)/g, (_, text) => evaluateCellReference(text, cellRef).length.toString());
    expression = expression.replace(/LEFT\(([^,]+),\s*([^)]+)\)/g, (_, text, num) => 
      evaluateCellReference(text, cellRef).substring(0, parseInt(evaluateCellReference(num, cellRef)) || 1)
    );
    expression = expression.replace(/RIGHT\(([^,]+),\s*([^)]+)\)/g, (_, text, num) => {
      const str = evaluateCellReference(text, cellRef);
      return str.substring(str.length - (parseInt(evaluateCellReference(num, cellRef)) || 1));
    });
    expression = expression.replace(/MID\(([^,]+),\s*([^,]+),\s*([^)]+)\)/g, (_, text, start, num) => {
      const str = evaluateCellReference(text, cellRef);
      const s = (parseInt(evaluateCellReference(start, cellRef)) || 1) - 1;
      const n = parseInt(evaluateCellReference(num, cellRef)) || 1;
      return str.substring(s, s + n);
    });
    expression = expression.replace(/CONCATENATE\(([^)]+)\)/g, (_, args) => 
      args.split(',').map(part => evaluateCellReference(part.trim(), cellRef)).join('')
    );
    expression = expression.replace(/TEXT\(([^,]+),\s*"([^"]+)"\)/g, (_, value, format) => {
      const num = parseFloat(evaluateCellReference(value, cellRef));
      if (isNaN(num)) return evaluateCellReference(value, cellRef);
      if (format.includes('$')) return '$' + num.toFixed(format.split('.')[1]?.length || 0);
      if (format.includes('%')) return (num * 100).toFixed(2) + '%';
      if (format.includes(',')) return num.toLocaleString();
      return num.toString();
    });
    expression = expression.replace(/EXACT\(([^,]+),\s*([^)]+)\)/g, (_, text1, text2) => 
      (evaluateCellReference(text1, cellRef) === evaluateCellReference(text2, cellRef)).toString()
    );
    expression = expression.replace(/FIND\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/g, (_, find, within, start) => {
      const str = evaluateCellReference(within, cellRef);
      const pos = str.indexOf(evaluateCellReference(find, cellRef), start ? parseInt(evaluateCellReference(start, cellRef)) - 1 : 0);
      return (pos === -1 ? '#VALUE!' : (pos + 1).toString());
    });

    // Logical Functions
    expression = expression.replace(/AND\(([^)]+)\)/g, (_, args) => 
      args.split(',').every(arg => evaluateCellReference(arg.trim(), cellRef).toUpperCase() === 'TRUE').toString()
    );
    expression = expression.replace(/OR\(([^)]+)\)/g, (_, args) => 
      args.split(',').some(arg => evaluateCellReference(arg.trim(), cellRef).toUpperCase() === 'TRUE').toString()
    );
    expression = expression.replace(/NOT\(([^)]+)\)/g, (_, arg) => 
      (evaluateCellReference(arg, cellRef).toUpperCase() !== 'TRUE').toString()
    );
    expression = expression.replace(/IF\(([^,]+),\s*([^,]+),\s*([^)]+)\)/g, (_, condition, ifTrue, ifFalse) => 
      evaluateCellReference(condition, cellRef).toUpperCase() === 'TRUE' ? evaluateCellReference(ifTrue, cellRef) : evaluateCellReference(ifFalse, cellRef)
    );
    expression = expression.replace(/IFERROR\(([^,]+),\s*([^)]+)\)/g, (_, value, ifError) => {
      try {
        const result = evaluateFormula('=' + evaluateCellReference(value, cellRef), cellRef);
        return result.startsWith('#') ? evaluateCellReference(ifError, cellRef) : result;
      } catch {
        return evaluateCellReference(ifError, cellRef);
      }
    });

    // Cell references
    expression = expression.replace(/[A-Z]\d+/g, (match) => {
      if (match === cellRef) return '0';
      return getCellNumericValue(match).toString();
    });
    
    // Final evaluation of remaining expressions
    if (/^[-+*/().\d\s]+$/.test(expression)) {
      const result = Function(`"use strict"; return (${expression})`)();
      return isNaN(result) ? '#ERROR!' : result.toString();
    }
    
    return expression;
  } catch (error) {
    return '#ERROR!';
  }
}, [cells]);

  const handleCellChange = (cellRef, value) => {
  setCells(prev => {
    const newCells = { ...prev };
    const isFormula = value.startsWith('=');
    newCells[cellRef] = { 
      value: value, 
      formula: isFormula ? value : '', 
      display: isFormula ? evaluateFormula(value, cellRef) : value 
    };
    
    // Save to history after state update
    setTimeout(() => saveToHistory(newCells), 0);
    return newCells;
  });
};

  const handleCellClick = (cellRef) => {
    setSelectedCell(cellRef);
    setFormulaBar(cells[cellRef]?.value || '');
  };

  const handleFormulaBarChange = (value) => {
    setFormulaBar(value);
    handleCellChange(selectedCell, value);
  };

  const insertFunction = (funcName) => {
    const newFormula = `=${funcName}()`;
    setFormulaBar(newFormula);
    setShowDropdown(false);
    setHoveredCategory(null);
    handleCellChange(selectedCell, newFormula);
  };

  // Add to your formatCell function
const formatCell = (styleUpdates) => {
  // Return early if no cell is selected
  if (!selectedCell) return;

  setCells(prevCells => {
    // Create a new copy of the cells object
    const newCells = { ...prevCells };
    
    // Ensure the selected cell exists
    if (!newCells[selectedCell]) {
      newCells[selectedCell] = { value: '', formula: '', display: '', style: {} };
    }

    // Update the cell's style
    newCells[selectedCell] = {
      ...newCells[selectedCell],
      style: {
        ...newCells[selectedCell].style, // Keep existing styles
        ...styleUpdates                 // Apply new styles
      }
    };

    return newCells;
  });

  // Save to undo history
  saveToHistory(cells);
};

// Add this to your CSS or style object to make the alignment work
// This should be in your stylesheet or at the top of your component
const cellStyle = {
  textAlign: 'left', // default alignment
  // ... other default styles
};


// Add a text wrap button to your toolbar
<button 
  onClick={() => formatCell({ 
    whiteSpace: cells[selectedCell]?.style?.whiteSpace === 'normal' ? 'nowrap' : 'normal' 
  })}
  className={`p-2 rounded transition-colors ${
    cells[selectedCell]?.style?.whiteSpace === 'normal' ? 'bg-gray-300' : 'hover:bg-gray-200'
  }`}
  title="Toggle text wrap"
>
  <WrapText size={16} />
</button>
const renderCell = (cellRef) => {
  const cell = cells[cellRef] || {};
  return (
    <td
      className="border border-gray-200 p-0"
      style={cell.style || {}} // Apply the cell's style object here
      onClick={() => handleCellClick(cellRef)}
    >
      {/* Your cell content */}
    </td>
  );
};

 const generateAdvancedChart = (options = {}) => {
  const {
    minRow = 1,
    maxRow = 10,
    labelCol = 0,
    valueCol = 1,
    includeZeroValues = false,
    skipEmptyLabels = true
  } = options;

  const data = [];
  let seriesCount = 1;

  for (let row = minRow; row <= maxRow; row++) {
    const labelCell = coordsToCell(row, labelCol);
    const valueCell = coordsToCell(row, valueCol);
    
    const label = cells[labelCell]?.display?.toString().trim() || '';
    const rawValue = cells[valueCell]?.display;
    const value = parseFloat(rawValue) || 0;
    
    // Skip conditions
    if (skipEmptyLabels && label === '') continue;
    if (!includeZeroValues && value === 0) continue;
    if (label === '' && rawValue === undefined) continue;
    
    data.push({
      label: label || `Series ${seriesCount++}`,
      value,
      color: `hsl(${row * 36}, 70%, 50%)`
    });
  }

  setChartData(data);
  setShowChart(data.length > 0);
  
};

  const exportToCSV = () => {
    let csvContent = '';
    for (let row = 0; row < 50; row++) {
      const rowData = [];
      for (let col = 0; col < 20; col++) {
        const cellRef = coordsToCell(row, col);
        rowData.push(cells[cellRef]?.display || '');
      }
      csvContent += rowData.join(',') + '\n';
    }
    
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'spreadsheet.csv';
    a.click();
    URL.revokeObjectURL(url);
  };
  // Add this handler for double-click auto-resize
const handleHeaderDoubleClick = (type, index) => {
  if (type === 'col') {
    // Auto-size column based on content
    let maxWidth = DEFAULT_COL_WIDTH;
    
    for (let row = 0; row < ROWS; row++) {
      const cellRef = coordsToCell(row, index);
      const cell = cells[cellRef];
      if (cell?.display) {
        // Approximate text width (10px per character)
        const textWidth = cell.display.length * 10;
        if (textWidth > maxWidth) {
          maxWidth = Math.min(300, textWidth + 10); // Add padding
        }
      }
    }
    
    setColWidths(prev => ({ ...prev, [index]: maxWidth }));
  } else {
    // Auto-size row based on content
    let maxHeight = DEFAULT_ROW_HEIGHT;
    
    for (let col = 0; col < COLS; col++) {
      const cellRef = coordsToCell(index, col);
      const cell = cells[cellRef];
      if (cell?.display) {
        // Count lines if wrapping is enabled
        const lines = cell.style?.whiteSpace === 'normal' 
          ? Math.ceil(cell.display.length / 20) 
          : 1;
        const height = lines * DEFAULT_ROW_HEIGHT;
        if (height > maxHeight) {
          maxHeight = Math.min(200, height);
        }
      }
    }
    
    setRowHeights(prev => ({ ...prev, [index]: maxHeight }));
  }
};

// Update your column and row headers to include onDoubleClick

  const importCSV = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const csv = e.target.result;
      const rows = csv.split('\n');
      const newCells = { ...cells };
      
      rows.forEach((row, rowIndex) => {
        const columns = row.split(',');
        columns.forEach((value, colIndex) => {
          if (rowIndex < ROWS && colIndex < COLS) {
            const cellRef = coordsToCell(rowIndex, colIndex);
            const trimmedValue = value.trim();
            newCells[cellRef] = {
              value: trimmedValue,
              formula: '',
              display: trimmedValue
            };
          }
        });
      });
      
      setCells(newCells);
    };
    reader.readAsText(file);
  };

  const addSheet = () => {
    const newId = Math.max(...sheets.map(s => s.id)) + 1;
    setSheets(prev => [...prev, { id: newId, name: `Sheet${newId}`, active: false }]);
  };

  const switchSheet = (sheetId) => {
  setSheets(prev =>
    prev.map(sheet => ({ ...sheet, active: sheet.id === sheetId }))
  );
  setActiveSheet(sheetId);
};


  const deleteSheet = (sheetId) => {
    if (sheets.length > 1) {
      setSheets(prev => prev.filter(sheet => sheet.id !== sheetId));
      if (activeSheet === sheetId) {
        const remainingSheets = sheets.filter(s => s.id !== sheetId);
        switchSheet(remainingSheets[0].id);
      }
    }
  };

  const downloadChart = () => {
    if (!chartData.length) return;
    
    const svgElement = document.querySelector('.chart-container svg');
    if (!svgElement) return;
    
    const serializer = new XMLSerializer();
    let svgStr = serializer.serializeToString(svgElement);
    svgStr = '<?xml version="1.0" standalone="no"?>\r\n' + svgStr;
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    
    img.onload = () => {
      canvas.width = img.width;
      canvas.height = img.height;
      ctx.drawImage(img, 0, 0);
      
      const pngFile = canvas.toDataURL('image/png');
      const downloadLink = document.createElement('a');
      downloadLink.href = pngFile;
      downloadLink.download = `chart-${new Date().toISOString().slice(0,10)}.png`;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
    };
    
    const svgBlob = new Blob([svgStr], {type: 'image/svg+xml;charset=utf-8'});
    const url = URL.createObjectURL(svgBlob);
    img.src = url;
  };

  const renderChart = () => {
    if (!chartData.length) return null;

    switch (chartType) {
      case 'pie':
        const total = chartData.reduce((sum, item) => sum + item.value, 0);
        if (total === 0) return null;
        
        let currentAngle = 0;
        return (
          <svg width="300" height="300" viewBox="0 0 300 300" xmlns="http://www.w3.org/2000/svg">
            <g transform="translate(150,150)">
              {chartData.map((item, index) => {
                const percentage = (item.value / total) * 100;
                const angle = (item.value / total) * 2 * Math.PI;
                const x1 = Math.cos(currentAngle) * 80;
                const y1 = Math.sin(currentAngle) * 80;
                const x2 = Math.cos(currentAngle + angle) * 80;
                const y2 = Math.sin(currentAngle + angle) * 80;
                
                const largeArcFlag = angle > Math.PI ? 1 : 0;
                const path = `M 0 0 L ${x1} ${y1} A 80 80 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;
                
                currentAngle += angle;
                
                return (
                  <g key={index}>
                    <path
                      d={path}
                      fill={item.color}
                      stroke="white"
                      strokeWidth="2"
                    />
                    <text 
                      x={Math.cos(currentAngle - angle/2) * 120} 
                      y={Math.sin(currentAngle - angle/2) * 120}
                      textAnchor="middle"
                      fontSize="12"
                      fill="#333"
                    >
                      {`${item.label} (${percentage.toFixed(1)}%)`}
                    </text>
                  </g>
                );
              })}
              <text y={-140} textAnchor="middle" fontSize="14" fontWeight="bold">
                Pie Chart
              </text>
            </g>
          </svg>
        );
        
      case 'line':
        const maxValue = Math.max(...chartData.map(d => d.value));
        const minValue = Math.min(...chartData.map(d => d.value));
        const range = maxValue - minValue || 1;
        
        return (
          <svg width="300" height="250" xmlns="http://www.w3.org/2000/svg">
            <rect width="100%" height="100%" fill="white" />
            <g transform="translate(40, 20)">
              {[0, 0.25, 0.5, 0.75, 1].map((ratio, i) => {
                const value = minValue + ratio * range;
                const y = 160 - (ratio * 160);
                return (
                  <g key={`y-label-${i}`}>
                    <line x1="0" y1={y} x2="240" y2={y} stroke="#e5e7eb" strokeWidth="1" />
                    <text x="-10" y={y + 4} textAnchor="end" fontSize="10" fill="#666">
                      {value.toFixed(2)}
                    </text>
                  </g>
                );
              })}
              
              <polyline
                fill="none"
                stroke="#3b82f6"
                strokeWidth="3"
                points={chartData.map((item, index) => {
                  const x = (index / Math.max(1, chartData.length - 1)) * 240;
                  const y = 160 - ((item.value - minValue) / range) * 160;
                  return `${x},${y}`;
                }).join(' ')}
              />
              
              {chartData.map((item, index) => {
                const x = (index / Math.max(1, chartData.length - 1)) * 240;
                const y = 160 - ((item.value - minValue) / range) * 160;
                return (
                  <g key={index}>
                    <circle cx={x} cy={y} r="4" fill="#3b82f6" />
                    <text 
                      x={x}
                      y={y - 10}
                      textAnchor="middle"
                      fontSize="10"
                      fill="#333"
                    >
                      {item.value.toFixed(2)}
                    </text>
                  </g>
                );
              })}
              
              {chartData.map((item, index) => {
                const x = (index / Math.max(1, chartData.length - 1)) * 240;
                return (
                  <text
                    key={`label-${index}`}
                    x={x}
                    y="190"
                    textAnchor="middle"
                    fontSize="10"
                    fill="#333"
                  >
                    {item.label}
                  </text>
                );
              })}
              
              <text x="120" y="220" textAnchor="middle" fontSize="14" fontWeight="bold">
                Line Chart
              </text>
            </g>
          </svg>
        );
        
      default:
        const maxVal = Math.max(...chartData.map(d => d.value));
        return (
          <svg width="300" height="250" xmlns="http://www.w3.org/2000/svg">
            <rect width="100%" height="100%" fill="white" />
            {chartData.map((item, index) => {
              const barHeight = Math.max(10, (item.value / maxVal) * 160);
              const x = index * 30 + 50;
              const y = 200 - barHeight;
              
              return (
                <g key={index} transform={`translate(${x}, ${y})`}>
                  <rect 
                    width="20" 
                    height={barHeight} 
                    fill={item.color}
                    rx="2"
                  />
                  <text 
                    x="10" 
                    y={barHeight + 15} 
                    textAnchor="middle" 
                    fontSize="10"
                    fill="#333"
                  >
                    {item.label}
                  </text>
                  <text 
                    x="10" 
                    y="-5" 
                    textAnchor="middle" 
                    fontSize="10"
                    fill="#333"
                  >
                    {item.value.toFixed(2)}
                  </text>
                </g>
              );
            })}
            <text x="150" y="230" textAnchor="middle" fontSize="14" fontWeight="bold">
              Bar Chart
            </text>
          </svg>
        );
    }
  };
  const handleFileLoad = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = JSON.parse(e.target.result);
          setCellData(data.cellData || {});
          setCellStyles(data.cellStyles || {});
          setSheetName(data.sheetName || 'Untitled spreadsheet');
          showNotification('Spreadsheet loaded');
        } catch (error) {
          showNotification('Error loading file');
        }
      };
      reader.readAsText(file);
    }
  };
  const handleSort = () => {
    const columnData = [];
    for (let row = 0; row < rows; row++) {
      const value = getCellValue(row, sortDialog.column);
      if (value) {
        columnData.push({ row, value, originalRow: row });
      }
    }

    columnData.sort((a, b) => {
      const aVal = isNaN(a.value) ? a.value : parseFloat(a.value);
      const bVal = isNaN(b.value) ? b.value : parseFloat(b.value);
      
      if (sortDialog.ascending) {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });

    const newData = { ...cellData };
    const newStyles = { ...cellStyles };
    
    columnData.forEach((item, index) => {
      for (let col = 0; col < cols; col++) {
        const oldKey = getCellKey(item.originalRow, col);
        const newKey = getCellKey(index, col);
        newData[newKey] = cellData[oldKey] || '';
        newStyles[newKey] = cellStyles[oldKey] || {};
      }
    });

    setCellData(newData);
    setCellStyles(newStyles);
    addToHistory(newData, newStyles);
    setSortDialog({ show: false, column: 0, ascending: true });
    showNotification('Data sorted');
  };
  const handleNew = () => {
    setCellData({});
    setCellStyles({});
    setHistory([{}]);
    setHistoryIndex(0);
    setSheetName('Untitled spreadsheet');
    showNotification('New spreadsheet created');
  };
  const handleOpen = () => {
    fileInputRef.current?.click();
  };
  const handleSave = () => {
    const data = { cellData, cellStyles, sheetName };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sheetName}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showNotification('Spreadsheet saved');
  };
  const handleDownloadCSV = () => {
    let csv = '';
    for (let row = 0; row < rows; row++) {
      const rowData = [];
      for (let col = 0; col < cols; col++) {
        rowData.push(getCellValue(row, col));
      }
      csv += rowData.join(',') + '\n';
    }
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sheetName}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showNotification('CSV downloaded');
  };
   const handleUndo = () => {
    if (historyIndex > 0) {
      const prevState = history[historyIndex - 1];
      setCellData(prevState.data || {});
      setCellStyles(prevState.styles || {});
      setHistoryIndex(historyIndex - 1);
      showNotification('Undone');
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      const nextState = history[historyIndex + 1];
      setCellData(nextState.data || {});
      setCellStyles(nextState.styles || {});
      setHistoryIndex(historyIndex + 1);
      showNotification('Redone');
    }
  };
  const handleCut = () => {
    handleCopy();
    setCellValue(selectedCell.row, selectedCell.col, '');
    showNotification('Cut');
  };

  const handlePaste = () => {
    if (clipboard) {
      setCellValue(selectedCell.row, selectedCell.col, clipboard.value);
      setCellStyle(selectedCell.row, selectedCell.col, clipboard.style);
      showNotification('Pasted');
    }
  };

  const handleDelete = () => {
    setCellValue(selectedCell.row, selectedCell.col, '');
    showNotification('Deleted');
  };

  const handleFindReplace = () => {
    if (!findReplace.find) return;
    
    let replaced = 0;
    const newData = { ...cellData };
    
    Object.keys(newData).forEach(key => {
      if (newData[key].includes(findReplace.find)) {
        newData[key] = newData[key].replace(new RegExp(findReplace.find, 'g'), findReplace.replace);
        replaced++;
      }
    });
    
    setCellData(newData);
    addToHistory(newData, cellStyles);
    showNotification(`Replaced ${replaced} instances`);
    setFindReplace({ show: false, find: '', replace: '' });
  };
  const handleCopy = () => {
    const value = getCellValue(selectedCell.row, selectedCell.col);
    const style = getCellStyle(selectedCell.row, selectedCell.col);
    setClipboard({ value, style, row: selectedCell.row, col: selectedCell.col });
    showNotification('Copied');
  };
  const insertRow = () => {
    const newData = {};
    const newStyles = {};
    
    Object.keys(cellData).forEach(key => {
      const [row, col] = key.split('-').map(Number);
      if (row >= selectedCell.row) {
        const newKey = getCellKey(row + 1, col);
        newData[newKey] = cellData[key];
        newStyles[newKey] = cellStyles[key] || {};
      } else {
        newData[key] = cellData[key];
        newStyles[key] = cellStyles[key] || {};
      }
    });
    
    setCellData(newData);
    setCellStyles(newStyles);
    addToHistory(newData, newStyles);
    showNotification('Row inserted');
  };
  
  const insertColumn = () => {
    const newData = {};
    const newStyles = {};
    
    Object.keys(cellData).forEach(key => {
      const [row, col] = key.split('-').map(Number);
      if (col >= selectedCell.col) {
        const newKey = getCellKey(row, col + 1);
        newData[newKey] = cellData[key];
        newStyles[newKey] = cellStyles[key] || {};
      } else {
        newData[key] = cellData[key];
        newStyles[key] = cellStyles[key] || {};
      }
    });
    
    setCellData(newData);
    setCellStyles(newStyles);
    addToHistory(newData, newStyles);
    showNotification('Column inserted');
  };
  const toggleBold = () => {
    const currentStyle = getCellStyle(selectedCell.row, selectedCell.col);
    setCellStyle(selectedCell.row, selectedCell.col, { 
      fontWeight: currentStyle.fontWeight === 'bold' ? 'normal' : 'bold' 
    });
  };
  const toggleItalic = () => {
    const currentStyle = getCellStyle(selectedCell.row, selectedCell.col);
    setCellStyle(selectedCell.row, selectedCell.col, { 
      fontStyle: currentStyle.fontStyle === 'italic' ? 'normal' : 'italic' 
    });
  };

  const toggleUnderline = () => {
    const currentStyle = getCellStyle(selectedCell.row, selectedCell.col);
    setCellStyle(selectedCell.row, selectedCell.col, { 
      textDecoration: currentStyle.textDecoration === 'underline' ? 'none' : 'underline' 
    });
  };

  const setAlignment = (align) => {
    setCellStyle(selectedCell.row, selectedCell.col, { textAlign: align });
  };

  const setColor = (type, color) => {
    const styleKey = type === 'text' ? 'color' : 'backgroundColor';
    setCellStyle(selectedCell.row, selectedCell.col, { [styleKey]: color });
    setColorPicker({ show: false, type: '', color: '#ffffff' });
  };
  const toggleFilter = () => {
    setFilterActive(!filterActive);
    showNotification(filterActive ? 'Filter removed' : 'Filter applied');
  };
  const MenuDropdown = ({ items, isOpen, onClose }) => {
    if (!isOpen) return null;
    
    return (
      <div className="absolute top-full left-0 mt-1 bg-white border border-gray-200 rounded-md shadow-lg z-50 min-w-48">
        {items.map((item, index) => (
          <button
            key={index}
            className="w-full flex items-center gap-3 px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 text-left"
            onClick={() => {
              item.action();
              onClose();
            }}
          >
            <item.icon size={16} />
            {item.label}
          </button>
        ))}
      </div>
    );
  };
  const handleZoomIn = () => setZoom(Math.min(200, zoom + 25));
  const handleZoomOut = () => setZoom(Math.max(50, zoom - 25));
  const toggleGridlines = () => setShowGridlines(!showGridlines);
  const toggleFormulas = () => setShowFormulas(!showFormulas);

 const menuItems = {
  File: [
    { 
      icon: FileText, 
      label: 'New', 
      action: () => {
        setCells({});
        setSheetName('Untitled spreadsheet');
        setNotification('New spreadsheet created');
        saveToHistory({});
      },
      shortcut: 'Ctrl+N'
    },
    { 
      icon: Folder, 
      label: 'Open', 
      action: () => fileInputRef.current?.click(),
      shortcut: 'Ctrl+O'
    },
    { 
      icon: Save, 
      label: 'Save', 
      action: handleSave,
      shortcut: 'Ctrl+S'
    },
    { 
      icon: Download, 
      label: 'Download as CSV', 
      action: exportToCSV
    },
    { 
      icon: Printer, 
      label: 'Print', 
      action: () => window.print(),
      shortcut: 'Ctrl+P'
    },
    { 
      icon: Share2, 
      label: 'Share', 
      action: () => setShowShareDropdown(!showShareDropdown)
    },
    { 
      icon: Clock, 
      label: 'Version history', 
      action: () => setNotification('Version history would show here')
    }
  ],
  Edit: [
    { 
      icon: Undo2, 
      label: 'Undo', 
      action: undo,
      shortcut: 'Ctrl+Z',
      disabled: historyIndex === 0
    },
    { 
      icon: Redo2, 
      label: 'Redo', 
      action: redo,
      shortcut: 'Ctrl+Y',
      disabled: historyIndex === history.length - 1
    },
    { 
      icon: Scissors, 
      label: 'Cut', 
      action: handleCut,
      shortcut: 'Ctrl+X'
    },
    { 
      icon: Copy, 
      label: 'Copy', 
      action: handleCopy,
      shortcut: 'Ctrl+C'
    },
    { 
      icon: Clipboard, 
      label: 'Paste', 
      action: handlePaste,
      shortcut: 'Ctrl+V'
    },
    { 
      icon: Search, 
      label: 'Find and replace', 
      action: () => setFindReplace({ ...findReplace, show: true }),
      shortcut: 'Ctrl+F'
    },
    { 
      icon: Trash2, 
      label: 'Delete', 
      action: () => handleCellChange(selectedCell, ''),
      shortcut: 'Del'
    }
  ],
  View: [
    { 
      icon: Grid3X3, 
      label: showGridlines ? 'Hide gridlines' : 'Show gridlines', 
      action: toggleGridlines
    },
    { 
      icon: Eye, 
      label: showFormulas ? 'Hide formulas' : 'Show formulas', 
      action: toggleFormulas,
      shortcut: 'Ctrl+`'
    },
    { 
      icon: ZoomIn, 
      label: 'Zoom in', 
      action: handleZoomIn,
      shortcut: 'Ctrl+Plus'
    },
    { 
      icon: ZoomOut, 
      label: 'Zoom out', 
      action: handleZoomOut,
      shortcut: 'Ctrl+-'
    },
    { 
      icon: Maximize, 
      label: 'Full screen', 
      action: () => document.documentElement.requestFullscreen(),
      shortcut: 'F11'
    }
  ],
  Insert: [
    { 
      icon: Plus, 
      label: 'Insert row above', 
      action: () => insertRow('above'),
      shortcut: 'Ctrl+Shift++'
    },
    { 
      icon: Plus, 
      label: 'Insert column left', 
      action: () => insertColumn('left'),
      shortcut: 'Ctrl+Shift++'
    },
    { 
      icon: BarChart3, 
      label: 'Chart', 
      action: () => setChartDialog({ show: true, type: 'line' })
    },
    { 
      icon: Image, 
      label: 'Image', 
      action: () => setNotification('Image upload would be implemented')
    },
    { 
      icon: Link2, 
      label: 'Link', 
      action: () => setNotification('Link insertion would be implemented'),
      shortcut: 'Ctrl+K'
    }
  ],
  Format: [
    { 
      icon: Bold, 
      label: 'Bold', 
      action: () => formatCell({ fontWeight: cells[selectedCell]?.style?.fontWeight === 'bold' ? 'normal' : 'bold' }),
      shortcut: 'Ctrl+B',
      active: cells[selectedCell]?.style?.fontWeight === 'bold'
    },
    { 
      icon: Italic, 
      label: 'Italic', 
      action: () => formatCell({ fontStyle: cells[selectedCell]?.style?.fontStyle === 'italic' ? 'normal' : 'italic' }),
      shortcut: 'Ctrl+I',
      active: cells[selectedCell]?.style?.fontStyle === 'italic'
    },
    { 
      icon: Underline, 
      label: 'Underline', 
      action: () => formatCell({ textDecoration: cells[selectedCell]?.style?.textDecoration === 'underline' ? 'none' : 'underline' }),
      shortcut: 'Ctrl+U',
      active: cells[selectedCell]?.style?.textDecoration === 'underline'
    },
    { 
      icon: AlignLeft, 
      label: 'Align left', 
      action: () => formatCell({ textAlign: 'left' }),
      active: cells[selectedCell]?.style?.textAlign === 'left'
    },
    { 
      icon: AlignCenter, 
      label: 'Align center', 
      action: () => formatCell({ textAlign: 'center' }),
      active: cells[selectedCell]?.style?.textAlign === 'center'
    },
    { 
      icon: AlignRight, 
      label: 'Align right', 
      action: () => formatCell({ textAlign: 'right' }),
      active: cells[selectedCell]?.style?.textAlign === 'right'
    },
    { 
      icon: Palette, 
      label: 'Fill color', 
      action: () => setShowColorPicker(!showColorPicker)
    },
    { 
      icon: WrapText, 
      label: 'Wrap text', 
      action: () => formatCell({ whiteSpace: cells[selectedCell]?.style?.whiteSpace === 'normal' ? 'nowrap' : 'normal' }),
      active: cells[selectedCell]?.style?.whiteSpace === 'normal'
    }
  ],
  Data: [
    { 
      icon: ArrowUpDown, 
      label: 'Sort range', 
      action: () => setSortDialog({ show: true, column: selectedCell?.col || 0, ascending: true })
    },
    { 
      icon: Filter, 
      label: filterActive ? 'Remove filter' : 'Create filter', 
      action: toggleFilter
    },
    { 
      icon: Shield, 
      label: 'Data validation', 
      action: () => setNotification('Data validation would be implemented')
    },
    { 
      icon: Table, 
      label: 'Pivot table', 
      action: () => setNotification('Pivot table would be implemented')
    }
  ],
  Tools: [
    { 
      icon: Settings, 
      label: 'Spell check', 
      action: () => setNotification('Spell check would run')
    },
    { 
      icon: Mail, 
      label: 'Notifications', 
      action: () => setNotification('Notification settings would open')
    },
    { 
      icon: Calculator, 
      label: 'Script editor', 
      action: () => setNotification('Script editor would open'),
      shortcut: 'Alt+Shift+X'
    }
  ],
  Extensions: [
    { 
      icon: Puzzle, 
      label: 'Add-ons', 
      action: () => setNotification('Add-ons store would open')
    },
    { 
      icon: Users, 
      label: 'Apps Script', 
      action: () => setNotification('Apps Script would open')
    }
  ],
  Help: [
    { 
      icon: HelpCircle, 
      label: 'Help', 
      action: () => window.open('https://support.google.com/docs/answer/6000292', '_blank'),
      shortcut: 'F1'
    },
    { 
      icon: Keyboard, 
      label: 'Keyboard shortcuts', 
      action: () => setNotification('Keyboard shortcuts: Ctrl+C (copy), Ctrl+V (paste), Ctrl+Z (undo)'),
      shortcut: 'Ctrl+/'
    },
    { 
      icon: BookOpen, 
      label: 'Training', 
      action: () => window.open('https://support.google.com/a/users/answer/9282664', '_blank')
    }
  ]
};
     
 
 return (
  <div className="w-full h-screen bg-white flex flex-col font-sans text-gray-800">
    {/* Main App Container */}
    <div className="flex flex-col h-full">
      {/* Top Toolbar */}
      <div className="bg-gray-50 border-b border-gray-200 px-4 py-1">
        <div className="flex items-center gap-1">
{Object.keys(menuItems).map((menuName) => (
  <div key={menuName} className="relative">
    <button
      className={`px-3 py-2 rounded hover:bg-gray-200 ${
        activeMenu === menuName ? 'bg-gray-200' : ''
      }`}
      onClick={(e) => handleMenuClick(menuName, e)}
    >
      {menuName}
    </button>
    
    {activeMenu === menuName && (
      <div
        ref={menuRef}
        className="absolute top-full left-0 mt-1 bg-white border border-gray-200 rounded-md shadow-lg z-50 min-w-48"
        style={{ top: `${menuPosition.top}px`, left: `${menuPosition.left}px` }}
      >
        {menuItems[menuName].map((item, index) => (
          <button
            key={index}
            className="w-full flex items-center gap-3 px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 text-left"
            onClick={() => {
              item.action();
              setActiveMenu(null);
            }}
          >
            <item.icon size={16} />
            {item.label}
          </button>
        ))}
      </div>
    )}
  </div>
))}
        </div>
      </div>

      {/* Formatting Toolbar */}
      <div className="bg-white border-b border-gray-200 px-4 py-2 flex items-center gap-1 flex-wrap">
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          <button 
            onClick={undo} 
            disabled={historyIndex === 0}
            className={`p-1.5 rounded-md ${historyIndex === 0 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-gray-100'}`}
            title="Undo (Ctrl+Z)"
          >
            <Undo size={18} className="text-gray-700" />
          </button>
          <button 
            onClick={redo} 
            disabled={historyIndex === history.length - 1}
            className={`p-1.5 rounded-md ${historyIndex === history.length - 1 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-gray-100'}`}
            title="Redo (Ctrl+Y)"
          >
            <Redo size={18} className="text-gray-700" />
          </button>
        </div>

        {/* Text Formatting */}
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          <button
            onClick={() => formatCell({
              fontWeight: cells[selectedCell]?.style?.fontWeight === 'bold' ? 'normal' : 'bold'
            })}
            className={`p-1.5 rounded-md ${
              cells[selectedCell]?.style?.fontWeight === 'bold' 
                ? 'bg-blue-100 text-blue-700' 
                : 'hover:bg-gray-100'
            }`}
            title="Bold (Ctrl+B)"
          >
            <Bold size={18} />
          </button>

          <button
            onClick={() => formatCell({
              fontStyle: cells[selectedCell]?.style?.fontStyle === 'italic' ? 'normal' : 'italic'
            })}
            className={`p-1.5 rounded-md ${
              cells[selectedCell]?.style?.fontStyle === 'italic' 
                ? 'bg-blue-100 text-blue-700' 
                : 'hover:bg-gray-100'
            }`}
            title="Italic (Ctrl+I)"
          >
            <Italic size={18} />
          </button>

          <button
            onClick={() => formatCell({
              textDecoration: cells[selectedCell]?.style?.textDecoration === 'underline' ? 'none' : 'underline'
            })}
            className={`p-1.5 rounded-md ${
              cells[selectedCell]?.style?.textDecoration === 'underline' 
                ? 'bg-blue-100 text-blue-700' 
                : 'hover:bg-gray-100'
            }`}
            title="Underline (Ctrl+U)"
          >
            <Underline size={18} />
          </button>
        </div>

        {/* Alignment */}
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          {['left', 'center', 'right'].map((align) => {
            const Icon = align === 'left' ? AlignLeft : align === 'center' ? AlignCenter : AlignRight;
            return (
              <button
                key={align}
                onClick={() => formatCell({ textAlign: align })}
                className={`p-1.5 rounded-md ${
                  cells[selectedCell]?.style?.textAlign === align 
                    ? 'bg-blue-100 text-blue-700' 
                    : 'hover:bg-gray-100'
                }`}
                title={`Align ${align}`}
              >
                <Icon size={18} />
              </button>
            );
          })}
        </div>

        {/* Number Formatting */}
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          <button 
            className="p-1.5 rounded-md hover:bg-gray-100"
            title="Currency Format"
          >
            <DollarSign size={18} />
          </button>
          <button 
            className="p-1.5 rounded-md hover:bg-gray-100"
            title="Percentage Format"
          >
            <Percent size={18} />
          </button>
        </div>

        {/* Formula Button */}
        <div className="relative">
          <button
            onClick={() => setShowDropdown(!showDropdown)}
            className="p-1.5 rounded-md hover:bg-gray-100 flex items-center gap-1"
            ref={formulaBtnRef}
          >
            <FunctionSquare size={18} />
            <span className="text-sm font-medium">Functions</span>
            <ChevronDown size={16} className="text-gray-500" />
          </button>

          {showDropdown && (
            <div
              ref={dropdownRef}
              className="absolute left-0 mt-1 flex bg-white border border-gray-200 rounded-md shadow-lg z-50"
              onMouseLeave={() => {
                if (!isHoveringSub) setShowDropdown(false);
              }}
            >
              <div className="w-48 border-r">
                {functionCategories.map((cat, idx) => (
                  <div
                    key={idx}
                    onClick={() => setSelectedCategory(cat)}
                    onMouseEnter={() => setSelectedCategory(cat)}
                    className={`flex justify-between items-center px-4 py-2 text-sm cursor-pointer ${
                      selectedCategory?.name === cat.name
                        ? 'bg-blue-50 text-blue-700'
                        : 'hover:bg-gray-100'
                    }`}
                  >
                    {cat.name}
                    <ChevronDown size={16} className="text-gray-400" />
                  </div>
                ))}
              </div>

              {selectedCategory && (
                <div
                  className="w-48 overflow-y-auto max-h-[300px] bg-white"
                  onMouseEnter={() => setIsHoveringSub(true)}
                  onMouseLeave={() => {
                    setIsHoveringSub(false);
                    setSelectedCategory(null);
                  }}
                >
                  {selectedCategory.functions.map((func, idx) => (
                    <div
                      key={idx}
                      onClick={() => {
                        insertFunction(func);
                        setShowDropdown(false);
                        setSelectedCategory(null);
                      }}
                      className="px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 cursor-pointer"
                    >
                      {func}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Color Picker */}
        <div className="relative">
          <button
            onClick={() => setShowColorPicker(!showColorPicker)}
            className="p-1.5 rounded-md hover:bg-gray-100 flex items-center gap-1"
          >
            <Palette size={18} />
            <ChevronDown size={16} className="text-gray-500" />
          </button>
          {showColorPicker && (
            <div className="absolute top-full mt-1 left-0 w-56 bg-white border border-gray-200 rounded-md shadow-lg z-50 p-3">
              <div className="mb-3">
                <div className="text-xs font-medium text-gray-500 mb-1">Text Color</div>
                <div className="grid grid-cols-8 gap-1">
                  {['#000000', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#FFA500'].map(color => (
                    <button
                      key={color}
                      onClick={() => {
                        formatCell({ color });
                        setShowColorPicker(false);
                      }}
                      className="w-6 h-6 rounded-md border border-gray-200 hover:ring-2 hover:ring-blue-200"
                      style={{ backgroundColor: color }}
                      title={color}
                    />
                  ))}
                </div>
              </div>
              <div className="border-t pt-3">
                <div className="text-xs font-medium text-gray-500 mb-1">Fill Color</div>
                <div className="grid grid-cols-8 gap-1">
                  {['#FFFFFF', '#FFEEEE', '#EEFFEE', '#EEEEFF', '#FFFFEE', '#FFEEFF', '#EEFFFF', '#FFE4B5'].map(color => (
                    <button
                      key={color}
                      onClick={() => {
                        formatCell({ backgroundColor: color });
                        setShowColorPicker(false);
                      }}
                      className="w-6 h-6 rounded-md border border-gray-200 hover:ring-2 hover:ring-blue-200"
                      style={{ backgroundColor: color }}
                      title={color}
                    />
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Action Buttons */}
        <div className="ml-auto flex items-center gap-2">
          <button
            onClick={handleGenerateChart}
            className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-md shadow-sm transition-colors"
          >
            <BarChart3 size={16} />
            Create Chart
          </button>
          <label className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md shadow-sm cursor-pointer transition-colors">
            <Upload size={16} />
            Import CSV
            <input type="file" accept=".csv" onChange={importCSV} className="hidden" />
          </label>
        </div>
      </div>

      {/* Formula Bar */}
      <div className="bg-gray-50 border-b border-gray-200 px-4 py-2 flex items-center gap-3">
        <div className="flex items-center gap-2">
          <span className="text-sm font-medium text-gray-500">fx</span>
          <span className="font-mono text-sm bg-gray-200 px-2 py-1 rounded font-medium text-gray-700">
            {selectedCell}
          </span>
        </div>
        <input
          type="text"
          value={formulaBar}
          onChange={(e) => handleFormulaBarChange(e.target.value)}
          className="flex-1 px-3 py-1.5 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white"
          placeholder="Enter formula or value..."
        />
      </div>

      {/* Spreadsheet Grid */}
      <div className="flex-1 flex relative overflow-hidden bg-white">
        <div className="flex-1 overflow-auto">
          <table className="border-collapse w-full" style={{ zoom: `${zoom}%` }}>
            <thead>
              <tr>
                <th className="w-10 h-8 bg-gray-100 border border-gray-300 sticky top-0 left-0 z-10"></th>
                {Array.from({ length: COLS }).map((_, colIndex) => (
                  <th
                    key={colIndex}
                    className="bg-gray-100 border border-gray-300 sticky top-0 z-10"
                    style={{ width: getColWidth(colIndex) }}
                    onDoubleClick={() => handleHeaderDoubleClick('col', colIndex)}
                  >
                    <div className="relative h-full flex items-center justify-center">
                      <span className="text-xs font-medium text-gray-700">
                        {getColumnHeader(colIndex)}
                      </span>
                      <div
                        className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-500"
                        onMouseDown={(e) => handleResizeMouseDown('col', colIndex, e)}
                      />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: ROWS }).map((_, rowIndex) => (
                <tr key={rowIndex}>
                  <th
                    className="w-10 bg-gray-100 border border-gray-300 sticky left-0 z-10"
                    onDoubleClick={() => handleHeaderDoubleClick('row', rowIndex)}
                  >
                    <div className="relative flex items-center justify-center h-full">
                      <span className="text-xs font-medium text-gray-700">
                        {rowIndex + 1}
                      </span>
                      <div
                        className="absolute right-0 bottom-0 left-0 h-1 cursor-row-resize hover:bg-blue-500"
                        onMouseDown={(e) => handleResizeMouseDown('row', rowIndex, e)}
                      />
                    </div>
                  </th>
                  {Array.from({ length: COLS }).map((_, colIndex) => {
                    const cellRef = coordsToCell(rowIndex, colIndex);
                    const cell = cells[cellRef] || {};
                    const isSelected = selectedCell === cellRef;
                    return (
                      <td
                        key={colIndex}
                        className={`border ${showGridlines ? 'border-gray-200' : 'border-transparent'} p-0 ${
                          isSelected ? 'ring-1 ring-blue-500' : ''
                        }`}
                        style={{
                          ...cell.style,
                          height: getRowHeight(rowIndex),
                          backgroundColor: cell.style?.backgroundColor || 'transparent'
                        }}
                        onClick={() => handleCellClick(cellRef)}
                      >
                        <input
                          type="text"
                          className={`w-full h-full px-2 py-1 outline-none text-sm ${
                            cell.style?.fontWeight === 'bold' ? 'font-semibold' : 'font-normal'
                          }`}
                          value={cell.display || ''}
                          onChange={(e) => handleCellChange(cellRef, e.target.value)}
                          style={{
                            textAlign: cell.style?.textAlign || 'left',
                            color: cell.style?.color || 'inherit',
                            fontStyle: cell.style?.fontStyle || 'normal',
                            textDecoration: cell.style?.textDecoration || 'none'
                          }}
                        />
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sheet Tabs */}
      <div className="bg-gray-50 border-t border-gray-200 px-4 py-1 flex items-center gap-2 z-10">
        <div className="flex items-center gap-1 overflow-x-auto">
          {sheets.map(sheet => (
            <div key={sheet.id} className="flex items-center shrink-0">
              <button
                onClick={() => switchSheet(sheet.id)}
                className={`px-3 py-1.5 text-sm rounded-t-md transition-colors flex items-center gap-1 ${
                  sheet.active 
                    ? 'bg-white border-t-2 border-t-blue-500 border-l border-r border-gray-300 -mb-px font-medium text-gray-900' 
                    : 'bg-gray-100 hover:bg-gray-200 text-gray-700'
                }`}
              >
                {sheet.name}
              </button>
              {sheets.length > 1 && (
                <button
                  onClick={() => deleteSheet(sheet.id)}
                  className="ml-1 p-1 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-md"
                >
                  <X size={14} />
                </button>
              )}
            </div>
          ))}
        </div>

        <button
          onClick={addSheet}
          className="flex items-center gap-1 px-2 py-1 text-sm bg-gray-100 hover:bg-gray-200 rounded-md text-gray-700 shrink-0"
        >
          <Plus size={16} />
          Add Sheet
        </button>

        <div className="ml-auto text-xs text-gray-500 shrink-0">
          {Object.keys(cells).filter(key => cells[key]?.value).length} cells with data
        </div>
      </div>
    </div>

    {/* Notification */}
    {notification && (
      <div className="fixed bottom-4 right-4 bg-gray-800 text-white px-4 py-2 rounded-md shadow-lg text-sm animate-fade-in">
        {notification}
      </div>
    )}

    {/* Hidden file input for opening files */}
    <input 
      type="file" 
      ref={fileInputRef} 
      onChange={handleFileLoad} 
      className="hidden" 
      accept=".json,.csv"
    />
  </div>
);


};

