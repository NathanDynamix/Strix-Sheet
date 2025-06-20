import  { useState, useEffect, useCallback,useRef } from 'react';
import {
Download, Upload, Share2, BarChart3, Calculator, Save, Plus, Trash2,
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, FunctionSquare, X, Search, DollarSign, Percent, Palette, 
  ChevronDown, WrapText, Mail, Link2, Users, Undo, Redo, Copy, 
  FileText, Folder, Clock, Undo2, Redo2, Scissors, Grid3X3, Eye, 
  ZoomIn, ZoomOut, Maximize, ArrowUpDown, Filter, Shield, Table, Settings, 
  Puzzle, HelpCircle, BookOpen, Clipboard, Image,  Keyboard
} from 'lucide-react';

import { useSpreadsheetData } from '../context/SpreadsheetDataContext';
import { useNavigate } from 'react-router-dom';
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
  const [history, setHistory] = useState([{}]); 
  const [sheetName, setSheetName] = useState('Untitled spreadsheet');
  const [images, setImages] = useState([]);
 

const [filterActive, setFilterActive] = useState(false);
const [sortDialog, setSortDialog] = useState({ show: false, column: 0, ascending: true });
const [chartDialog, setChartDialog] = useState({ show: false, type: 'line' });
const [chartData, setChartData] = useState([]);
const [zoom, setZoom] = useState(100);
const [showFormulas, setShowFormulas] = useState(false);
const [notification, setNotification] = useState(null);
const fileInputRef = useRef(null);
const [clipboard, setClipboard] = useState({
  value: null,
  formula: null,
  display: null,
  style: null,
  action: null 
});
const [historyIndex, setHistoryIndex] = useState(0);
 const [cells, setCells] = useState({});
const [selectedCell, setSelectedCell] = useState(null);
  const navigate = useNavigate();
 const [showShareDropdown, setShowShareDropdown] = useState(false);
  const dropdownRef = useRef(null);
  const [selectedCategory, setSelectedCategory] = useState(null);
const [isHoveringSub, setIsHoveringSub] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);
  const [formulaBar, setFormulaBar] = useState('');
  const [sheets, setSheets] = useState([{ id: 1, name: 'Sheet1', active: true }]);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [activeSheet, setActiveSheet] = useState(1);
  const [showChart, setShowChart] = useState(false);
  const [showFunctions, setShowFunctions] = useState(false);
  const [functionSearch, setFunctionSearch] = useState('');
  const [selectedRange, setSelectedRange] = useState('');
  const formulaBtnRef = useRef(null);
  const [showHistory, setShowHistory] = useState(false);
const [colWidths, setColWidths] = useState({});
const [rowHeights, setRowHeights] = useState({});
const [resizing, setResizing] = useState({ active: false, type: null, index: null, startPos: 0 });
const [dragPos, setDragPos] = useState(null);
const [uploadedImages, setUploadedImages] = useState([]);
const { setSpreadsheetData, setDataSource } = useSpreadsheetData();


const DEFAULT_COL_WIDTH = 80;
const DEFAULT_ROW_HEIGHT = 24;


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

  const addToHistory = (currentCells) => {
  const newHistory = history.slice(0, historyIndex + 1);
  newHistory.push({
    cells: JSON.parse(JSON.stringify(currentCells)),
    timestamp: new Date().toISOString()
  });
  setHistory(newHistory);
  setHistoryIndex(newHistory.length - 1);
};

  const showNotification = (message) => {
    setNotification(message);
    setTimeout(() => setNotification(null), 3000);
  };



const saveToHistory = (currentCells) => {
  const newHistory = [...history];
  // Only save if there are actual changes
  if (JSON.stringify(currentCells) !== JSON.stringify(newHistory[newHistory.length - 1])) {
    newHistory.push(JSON.parse(JSON.stringify(currentCells)));
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
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

 const handleGenerateChart = () => {
  const activeColumns = new Set();
  const allHeaders = [];
  
  for (let row = 2; row <= 101; row++) {
    for (let col = 0; col < 26; col++) {
      const cellRef = `${String.fromCharCode(65 + col)}${row}`;
      const value = cells[cellRef]?.display;
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



const formatCell = (styleUpdates) => {
  if (!selectedCell) return;
  
  setCells(prev => ({
    ...prev,
    [selectedCell]: {
      ...prev[selectedCell],
      style: {
        ...prev[selectedCell]?.style,
        ...styleUpdates
      }
    }
  }));
};

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
function renderCellContent(cell) {
  if (cell.isLink && cell.linkUrl) {
    return (
      <a 
        href={cell.linkUrl} 
        target="_blank" 
        rel="noopener noreferrer"
        onClick={(e) => e.stopPropagation()} // Prevent cell selection when clicking link
        style={{ color: 'blue', textDecoration: 'underline' }}
      >
        {cell.value}
      </a>
    );
  }
  return cell.value;
}

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

const handleHeaderDoubleClick = (type, index) => {
  if (type === 'col') {
  
    let maxWidth = DEFAULT_COL_WIDTH;
    
    for (let row = 0; row < ROWS; row++) {
      const cellRef = coordsToCell(row, index);
      const cell = cells[cellRef];
      if (cell?.display) {
        
        const textWidth = cell.display.length * 10;
        if (textWidth > maxWidth) {
          maxWidth = Math.min(300, textWidth + 10); 
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

  
function handleSpellCheck(setNotification, documentText) {
  
  setNotification('Starting spell check...');
  
  
  const dictionary = ['apple', 'banana', 'document', 'check', 'example'];
  
 
  const words = documentText.split(/\s+/);
  const errors = [];
  

  words.forEach((word, index) => {
    const cleanWord = word.toLowerCase().replace(/[.,!?]/g, '');
    if (cleanWord.length > 0 && !dictionary.includes(cleanWord)) {
      errors.push({
        word,
        position: index,
        suggestions: getSuggestions(cleanWord, dictionary)
      });
    }
  });
  
  
  if (errors.length === 0) {
    setNotification('Spell check complete - no errors found');
  } else {
    setNotification(`Found ${errors.length} errors. ${errors.length > 10 ? 'First 10 shown' : ''}`);
    console.log('Spelling errors:', errors.slice(0, 10));
    return errors; // Return errors for highlighting
  }
}


function getSuggestions(word, dictionary) {

  return dictionary.filter(dictWord => 
    dictWord.startsWith(word[0]) && 
    Math.abs(dictWord.length - word.length) <= 2
  ).slice(0, 3);
}

  function handleSortRange(selectedRange, setCells) {
  if (!selectedRange) {
    console.error("No range selected");
    return;
  }

  // Get all cells in the selected range
  const rangeCells = Object.entries(cells).filter(([key]) => {
    const [row, col] = key.split('-').map(Number);
    return (
      row >= selectedRange.startRow &&
      row <= selectedRange.endRow &&
      col >= selectedRange.startCol &&
      col <= selectedRange.endCol
    );
  });

  // Sort by the first column in the range
  const sortColumn = selectedRange.startCol;
  
  rangeCells.sort((a, b) => {
    const [keyA] = a;
    const [keyB] = b;
    const [rowA] = keyA.split('-').map(Number);
    const [rowB] = keyB.split('-').map(Number);
    
    const cellA = cells[`${rowA}-${sortColumn}`]?.value;
    const cellB = cells[`${rowB}-${sortColumn}`]?.value;

    // Numeric comparison
    if (!isNaN(cellA) && !isNaN(cellB)) {
      return Number(cellA) - Number(cellB);
    }
    
    // String comparison
    return String(cellA).localeCompare(String(cellB));
  });

  // Rebuild the cells object with new order
  const newCells = {...cells};
  let newRowIndex = selectedRange.startRow;
  
  rangeCells.forEach(([key, cellData]) => {
    const [oldRow, col] = key.split('-').map(Number);
    const newKey = `${newRowIndex}-${col}`;
    
    if (newKey !== key) {
      newCells[newKey] = {...cellData};
      delete newCells[key];
    }
    
    newRowIndex++;
  });

  setCells(newCells);
}

  

  const loadFromStorage = () => {
  try {
    const savedData = localStorage.getItem('spreadsheetData');
    if (savedData) {
      const data = JSON.parse(savedData);
      setCells(data.cells || {});
      setSheetName(data.sheetName || 'Untitled spreadsheet');
      setHistory(data.history || [{}]);
      setHistoryIndex(data.historyIndex || 0);
      showNotification('Spreadsheet loaded from storage');
    }
  } catch (error) {
    console.error('Load failed:', error);
    showNotification('Error: Could not load saved data');
  }
};

// Call this in useEffect when component mounts
useEffect(() => {
  loadFromStorage();
}, []);
 const handleSave = async () => {
  try {
    // Convert cells data to CSV format
    const csvData = convertToCSV(cells);

    // Check if File System Access API is available (Chrome/Edge)
    if ('showSaveFilePicker' in window) {
      const fileHandle = await window.showSaveFilePicker({
        suggestedName: `${sheetName}.csv`,
        types: [{
          description: 'CSV Files',
          accept: { 'text/csv': ['.csv'] },
        }],
      });

      // Write CSV content to the selected file
      const writableStream = await fileHandle.createWritable();
      await writableStream.write(csvData);
      await writableStream.close();

      showNotification(`Saved as "${fileHandle.name}"`);
    } else {
      // Fallback for browsers without File System Access API
      downloadCSV(csvData, `${sheetName}.csv`);
      showNotification(`Downloaded "${sheetName}.csv"`);
    }
  } catch (error) {
    if (error.name !== 'AbortError') {
      console.error('Save failed:', error);
      showNotification('Error: Could not save file');
    }
  }
};
const convertToCSV = (cells) => {
  const ROWS = 100;
  const COLS = 26;
  let csvContent = '';

  // Build header row (A, B, C, ...)
  const headers = Array.from({ length: COLS }, (_, i) => 
    String.fromCharCode(65 + i)
  );
  csvContent += headers.join(',') + '\n';

  // Build data rows
  for (let row = 1; row <= ROWS; row++) {
    const rowData = [];
    for (let col = 0; col < COLS; col++) {
      const cellRef = `${String.fromCharCode(65 + col)}${row}`;
      const cellValue = cells[cellRef]?.display || cells[cellRef]?.value || '';
      // Escape CSV special characters (comma, quotes, newlines)
      rowData.push(`"${cellValue.toString().replace(/"/g, '""')}"`);
    }
    csvContent += rowData.join(',') + '\n';
  }

  return csvContent;
};

const downloadCSV = (csvData, filename) => {
  const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

  const [findReplace, setFindReplace] = useState({
  show: false,
  find: '',
  replace: '',
  matchCase: false,
  results: [],
  currentIndex: -1
});

 const handleFindReplace = (action = 'find') => {
  if (!findReplace.find.trim()) {
    showNotification('Please enter text to find');
    return;
  }

  try {
    const flags = findReplace.matchCase ? 'g' : 'gi';
    const pattern = new RegExp(escapeRegExp(findReplace.find), flags);
    const results = [];
    let replaceCount = 0;

    // First pass: Find all matches and collect results
    {Object.entries(cells).map(([cellId, cell]) => (
  <div key={cellId} className="border p-2 w-20 h-20 flex items-center justify-center">
    {cell.type === 'image' ? (
      <img src={cell.value} alt="Cell" className="max-w-full max-h-full object-contain" />
    ) : (
      cell.value
    )}
  </div>
))}


    // Second pass: Perform replacement if needed
    if (action === 'replace' && results.length > 0) {
      const newCells = { ...cells };
      
      results.forEach(cellRef => {
        const cell = newCells[cellRef];
        const newValue = cell.value.replace(pattern, findReplace.replace);
        
        newCells[cellRef] = {
          ...cell,
          value: newValue,
          display: newValue.startsWith('=') ? evaluateFormula(newValue, cellRef) : newValue
        };
        replaceCount++;
      });

      setCells(newCells);
      saveToHistory(newCells); // Make sure to update history
    }

    // Update state and show results
    setFindReplace(prev => ({
      ...prev,
      results,
      currentIndex: results.length > 0 ? 0 : -1
    }));

    // Show appropriate notification
    const message = action === 'find'
      ? results.length > 0 
        ? `Found ${results.length} matches` 
        : 'No matches found'
      : replaceCount > 0
        ? `Replaced ${replaceCount} occurrences`
        : 'No matches found';
    
    showNotification(message);

    // Highlight first match if found
    if (results.length > 0) {
      setSelectedCell(results[0]);
    }

  } catch (error) {
    console.error('Find/replace error:', error);
    showNotification('Invalid search pattern');
  }
};

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
const handleCut = () => {
  if (!selectedCell || !cells[selectedCell]) {
    showNotification('No cell selected or cell is empty');
    return;
  }

  const cell = cells[selectedCell];
  setClipboard({
    value: cell.value,
    formula: cell.formula,
    display: cell.display,
    style: cell.style || {},
    action: 'cut'
  });

  // Update the cells state
  setCells(prev => {
    const newCells = {...prev};
    newCells[selectedCell] = {
      value: '',
      formula: '',
      display: '',
      style: {}
    };
    return newCells;
  });

  showNotification('Content cut to clipboard');
};

const handleCopy = () => {
  if (!selectedCell || !cells[selectedCell]) {
    showNotification('No cell selected or cell is empty');
    return;
  }

  const cell = cells[selectedCell];
  setClipboard({
    value: cell.value,
    formula: cell.formula,
    display: cell.display,
    style: cell.style || {},
    action: 'copy'
  });

  showNotification('Content copied to clipboard');
};

const handlePaste = () => {
  if (!selectedCell) {
    showNotification('No cell selected');
    return;
  }

  if (!clipboard || clipboard.value === null) {
    showNotification('Clipboard is empty');
    return;
  }

  setCells(prev => ({
    ...prev,
    [selectedCell]: {
      value: clipboard.value,
      formula: clipboard.formula,
      display: clipboard.display,
      style: {...clipboard.style}
    }
  }));

  // If it was a cut operation, clear the clipboard
  if (clipboard.action === 'cut') {
    setClipboard({
      value: null,
      formula: null,
      display: null,
      style: null,
      action: null
    });
  }

  showNotification('Content pasted');
};

  const insertRow = (position = 'above') => {
  if (!selectedCell) {
    showNotification('Please select a cell first');
    return;
  }

  // Get the row number from selected cell (e.g., "B5" → 5)
  const selectedRow = parseInt(selectedCell.match(/\d+$/)[0]);
  
  setCells(prevCells => {
    const newCells = {...prevCells};
    const newRow = position === 'above' ? selectedRow : selectedRow + 1;

    // Shift all cells below down by 1 row
    Object.keys(prevCells).forEach(cellKey => {
      const [col, row] = cellKey.match(/^([A-Z]+)(\d+)$/).slice(1);
      const rowNum = parseInt(row);

      if (rowNum >= newRow) {
        // Move cell down by 1 row
        const newKey = `${col}${rowNum + 1}`;
        newCells[newKey] = {...prevCells[cellKey]};
        
        // Clear original cell if it's being vacated
        if (rowNum === newRow) {
          newCells[cellKey] = { value: '', display: '' };
        }
      }
    });

    return newCells;
  });

  // Update history
  saveToHistory(cells);
  showNotification(`Row inserted ${position}`);
};
  
  const insertColumn = (position = 'left') => {
  if (!selectedCell) {
    showNotification('Please select a cell first');
    return;
  }

  // Get column letter and index (B5 → "B" → 1)
  const colLetter = selectedCell.replace(/\d+/g, '');
  const colIndex = colLetter.charCodeAt(0) - 65; // A=0, B=1, etc.
  const insertAt = position === 'left' ? colIndex : colIndex + 1;

  setCells(prevCells => {
    const newCells = {};
    const colsToShift = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(insertAt);

    // 1. Process all existing cells
    Object.entries(prevCells).forEach(([cellRef, cellData]) => {
      const [currentCol, row] = cellRef.split(/(\d+)/).filter(Boolean);
      const currentIndex = currentCol.charCodeAt(0) - 65;

      if (currentIndex >= insertAt) {
        // Shift right
        const newCol = String.fromCharCode(65 + currentIndex + 1);
        newCells[`${newCol}${row}`] = { ...cellData };
        
        // Clear original if it's the insertion point
        if (currentIndex === insertAt) {
          newCells[cellRef] = { value: '', display: '', formula: '' };
        }
      } else {
        // Copy unchanged
        newCells[cellRef] = { ...cellData };
      }
    });

    return newCells;
  });

  showNotification(`Column inserted ${position} of ${colLetter}`);
};
  
  const toggleFilter = () => {
    setFilterActive(!filterActive);
    showNotification(filterActive ? 'Filter removed' : 'Filter applied');
  };
  const handleZoomIn = () => setZoom(Math.min(200, zoom + 25));
  const handleZoomOut = () => setZoom(Math.max(50, zoom - 25));
  const toggleFormulas = () => setShowFormulas(!showFormulas);
const restoreVersion = (index) => {
  if (index < 0 || index >= history.length) return;
  
  const version = history[index];
  setCells(version.cells || {});
  setHistoryIndex(index);
  setShowHistory(false);
  showNotification(`Restored version ${index + 1}`);
};
const [showGridlines, setShowGridlines] = useState(true); // Default to visible
useEffect(() => {
  console.log('Current gridlines state:', showGridlines);
}, [showGridlines]);
const toggleGridlines = () => {
  console.log('Before toggle:', showGridlines); // Debug
  setShowGridlines(!showGridlines);
  console.log('After toggle:', !showGridlines); // Debug
};

 const menuItems = {
  File: [
    { 
  icon: FileText, 
  label: 'New', 
  action: () => {
    addSheet(); // Call your addSheet function instead of individual state setters
    setNotification('New spreadsheet created');
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
      icon: Share2,
      label: 'Share',
      action: () => {
        if (navigator.share) {
          navigator.share({
            title: sheetName,
            text: 'Check out this spreadsheet',
            url: window.location.href
          }).catch(console.error);
        } else {
          setShowShareDropdown(!showShareDropdown);
        }
      }
    },
    {
  icon: Clock,
  label: 'Version history',
  action: () => {
    setShowHistory(true);
    showNotification(`Loaded ${history.length} versions`);
  },
  shortcut: 'Ctrl+H'
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
    shortcut: 'Ctrl+X',
    disabled: !selectedCell || !cells[selectedCell]
  },
    { 
    icon: Copy, 
    label: 'Copy', 
    action: handleCopy,
    shortcut: 'Ctrl+C',
    disabled: !selectedCell || !cells[selectedCell]
  },
    { 
    icon: Clipboard, 
    label: 'Paste', 
    action: handlePaste,
    shortcut: 'Ctrl+V',
    disabled: !clipboard || clipboard.value === null
  },
   {
  icon: Search, 
  label: 'Find and replace', 
  action: () => setFindReplace({ 
    show: true,
    find: '',
    replace: '',
    matchCase: false,
    results: [],
    currentIndex: -1
  }),
  shortcut: 'Ctrl+F'
},
    {
  icon: Trash2, 
  label: 'Delete', 
  action: () => {
    if (selectedCell) {  // Check if a cell is selected
      handleCut();       // Clear content (or use custom delete logic)
      setNotification('Cell content deleted');
    } else {
      setNotification('No cell selected');
    }
  },
  shortcut: 'Delete'  // Prefer 'Delete' over 'Del' for clarity
}
  ],
  View: [
    {
  icon: Grid3X3,
  label: showGridlines ? 'Hide gridlines' : 'Show gridlines',
  action: toggleGridlines,
  className: showGridlines ? 'bg-blue-100' : '' // Visual feedback
},
    {
  icon: Eye,
  label: showFormulas ? 'Hide formulas' : 'Show formulas',
  action:()=>{
    setShowDropdown(!showDropdown)
  } ,
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
      action: () => {
        if (!document.fullscreenElement) {
          document.documentElement.requestFullscreen();
        } else {
          document.exitFullscreen();
        }
      },
      shortcut: 'F11'
    }
  ],
  Insert: [
    { 
  icon: Plus, 
  label: 'Insert row above', 
  action: () => {
    if (selectedCell) {
      insertRow('above');
    } else {
      showNotification('Please select a cell first');
    }
  },
  shortcut: 'Ctrl+Shift++',
  disabled: !selectedCell // Disable if no cell selected
},
    {
  icon: Plus,
  label: 'Insert column left',
  action: () => {
    if (selectedCell) {
      insertColumn('left');
    } else {
      showNotification('Please select a cell first', { type: 'error' });
    }
  },
  shortcut: 'Ctrl+Shift+]',
  disabled: !selectedCell
},
    {
  icon: BarChart3,
  label: 'Chart',
  action: handleGenerateChart
},
    {
  icon: Image,
  label: 'Image',
  action: async () => {
    const result = await handleImageUpload(setNotification);
    if (result) {
      // Add to your images state
      setImages(prev => [...prev, result]);
      
      // Or insert into your document/editor
      insertImageIntoEditor(result.url); // Your custom function
      
      console.log('Uploaded image:', result);
    }
  }
},
    {
  icon: Link2,
  label: 'Link',
  action: () => handleLinkInsertion(selectedCell, cells, setCells, setNotification),
  shortcut: 'Ctrl+K'
}
  ],
  Format: [
    { 
      icon: Bold, 
      label: 'Bold', 
      action: () => {
        if (selectedCell) {
          formatCell({ 
            fontWeight: cells[selectedCell]?.style?.fontWeight === 'bold' ? 'normal' : 'bold' 
          });
        }
      },
      shortcut: 'Ctrl+B',
      active: selectedCell && cells[selectedCell]?.style?.fontWeight === 'bold'
    },
    { 
      icon: Italic, 
      label: 'Italic', 
      action: () => {
        if (selectedCell) {
          formatCell({ 
            fontStyle: cells[selectedCell]?.style?.fontStyle === 'italic' ? 'normal' : 'italic' 
          });
        }
      },
      shortcut: 'Ctrl+I',
      active: selectedCell && cells[selectedCell]?.style?.fontStyle === 'italic'
    },
    { 
      icon: Underline, 
      label: 'Underline', 
      action: () => {
        if (selectedCell) {
          formatCell({ 
            textDecoration: cells[selectedCell]?.style?.textDecoration === 'underline' ? 'none' : 'underline' 
          });
        }
      },
      shortcut: 'Ctrl+U',
      active: selectedCell && cells[selectedCell]?.style?.textDecoration === 'underline'
    },
    { 
      icon: AlignLeft, 
      label: 'Align left', 
      action: () => {
        if (selectedCell) {
          formatCell({ textAlign: 'left' });
        }
      },
      active: selectedCell && cells[selectedCell]?.style?.textAlign === 'left'
    },
    { 
      icon: AlignCenter, 
      label: 'Align center', 
      action: () => {
        if (selectedCell) {
          formatCell({ textAlign: 'center' });
        }
      },
      active: selectedCell && cells[selectedCell]?.style?.textAlign === 'center'
    },
    { 
      icon: AlignRight, 
      label: 'Align right', 
      action: () => {
        if (selectedCell) {
          formatCell({ textAlign: 'right' });
        }
      },
      active: selectedCell && cells[selectedCell]?.style?.textAlign === 'right'
    },
    { 
      icon: Palette, 
      label: 'Fill color', 
      action: () => {
        if (selectedCell) {
          setShowColorPicker({
            show: true,
            cell: selectedCell,
            type: 'background'
          });
        }
      }
    },
    {
  icon: WrapText,
  label: 'Wrap text',
  action: () => {
    if (selectedCell) {
      const currentCell = cells[selectedCell];
      const isCurrentlyWrapped = currentCell?.style?.whiteSpace === 'normal';
      
      formatCell({
        whiteSpace: isCurrentlyWrapped ? 'nowrap' : 'normal'
      });
      
      setNotification(
        isCurrentlyWrapped 
          ? 'Text unwrapped' 
          : 'Text wrapped'
      );
    } else {
      setNotification('Please select a cell first');
    }
  },
  active: selectedCell && cells[selectedCell]?.style?.whiteSpace === 'normal',
  shortcut: 'Ctrl+Shift+W'
}
  ],
  Data: [
    {
  icon: ArrowUpDown,
  label: 'Sort range',
  action: () => handleSortRange(selectedCell, setSortDialog)
},
    { 
      icon: Filter, 
      label: filterActive ? 'Remove filter' : 'Create filter', 
      action: toggleFilter
    },
    { 
      icon: Shield, 
      label: 'Data validation', 
      action: () => {
        if (selectedCell) {
          const rules = prompt('Enter validation rules (e.g., number>0):');
          if (rules) {
            setNotification(`Validation rules set for ${selectedCell}`);
          }
        }
      }
    },
    { 
      icon: Table, 
      label: 'Pivot table', 
      action: () => {
        setNotification('Pivot table would analyze selected data');
      }
    }
  ],
  Tools: [
    {
  icon: Settings,
  label: 'Spell check',
  action: () => {
    const errors = handleSpellCheck(
      setNotification, 
      getDocumentText() // Your function to get current document text
    );
    
    if (errors) {
      // Highlight errors in your UI
      highlightSpellingErrors(errors);
    }
  }
},
    { 
      icon: Mail, 
      label: 'Notifications', 
      action: () => {
        setNotification('Notification settings would open');
      }
    },
    { 
      icon: Calculator, 
      label: 'Script editor', 
      action: () => {
        setNotification('Script editor would open');
      },
      shortcut: 'Alt+Shift+X'
    }
  ],
  Extensions: [
    { 
      icon: Puzzle, 
      label: 'Add-ons', 
      action: () => {
        setNotification('Add-ons store would open');
      }
    },
    { 
      icon: Users, 
      label: 'Apps Script', 
      action: () => {
        setNotification('Apps Script editor would open');
      }
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
      action: () => {
        setNotification(`
          Common shortcuts:
          Ctrl+C - Copy
          Ctrl+V - Paste
          Ctrl+Z - Undo
          Ctrl+Y - Redo
          Ctrl+F - Find
        `);
      },
      shortcut: 'Ctrl+/'
    },
    { 
      icon: BookOpen, 
      label: 'Training', 
      action: () => window.open('https://support.google.com/a/users/answer/9282664', '_blank')
    }
  ]
};

function handleLinkInsertion(selectedCell, setCells, setNotification) {
  if (!selectedCell) {
    setNotification('Please select a cell first');
    return;
  }

  // Get current cell text
  const currentText = cells[selectedCell]?.value || '';

  // Google Sheets-style dialog
  const url = prompt('Enter URL (e.g., https://example.com):');
  if (!url) return; // User cancelled

  // Get display text (defaults to URL if empty)
  const displayText = prompt('Enter text to display (optional):', currentText) || url;

  // Create Google Sheets-style link
  setCells(prev => ({
    ...prev,
    [selectedCell]: {
      ...prev[selectedCell],
      value: displayText,
      linkUrl: url.startsWith('http') ? url : `https://${url}`,
      isLink: true,
      // Special formatting like Google Sheets
      style: {
        ...prev[selectedCell]?.style,
        color: '#1155cc',
        textDecoration: 'underline',
        cursor: 'pointer'
      }
    }
  }));

  setNotification('Link created! Click to open.');
}

function handleImageUpload(setNotification) {
  return new Promise((resolve) => {
    // Create file input
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.style.display = 'none'; // Hide the input element
    
    // Add to DOM temporarily
    document.body.appendChild(input);
    
    input.onchange = async (e) => {
      const file = e.target.files[0];
      
      if (!file) {
        setNotification('No file selected');
        document.body.removeChild(input);
        return resolve(null);
      }
      
      // Validate file type
      if (!file.type.startsWith('image/')) {
        setNotification('Please select an image file (JPEG, PNG, etc.)');
        document.body.removeChild(input);
        return resolve(null);
      }
      
      try {
        // Create preview URL
        const imageUrl = URL.createObjectURL(file);
        
        setNotification(`Image loaded: ${file.name}`);
        
        // Return both URL and file data
        resolve({
          url: imageUrl,
          file: file,
          name: file.name,
          size: file.size,
          type: file.type
        });
        
      } catch (error) {
        console.error('Upload error:', error);
        setNotification('Failed to load image');
        resolve(null);
      } finally {
        // Clean up
        document.body.removeChild(input);
      }
    };
    
    // Trigger file dialog
    input.click();
  });
}

     useEffect(() => {
  const handleKeyDown = (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

    const ctrlKey = e.ctrlKey || e.metaKey; // Support both Ctrl and Cmd
    
    if (ctrlKey) {
      switch (e.key.toLowerCase()) {
        case 'x':
          e.preventDefault();
          handleCut();
          break;
        case 'c':
          e.preventDefault();
          handleCopy();
          break;
        case 'v':
          e.preventDefault();
          handlePaste();
          break;
      }
    }
  };

  window.addEventListener('keydown', handleKeyDown);
  return () => window.removeEventListener('keydown', handleKeyDown);
}, [selectedCell, clipboard, cells]);

 
 return (
  <div className="w-full h-screen bg-white flex flex-col font-sans text-gray-800">
    {/* Main App Container */}
    <div className="flex flex-col h-full">
      {/* Top Toolbar */}
     <div className="bg-gray-50 border-b border-gray-200 px-4 py-1">
  <div className="flex items-center gap-1">
   {Object.keys(menuItems).map((menuName) => (
  <div key={menuName} className="relative inline-block group">
    <button
      className={`px-3 py-2 rounded hover:bg-gray-200 transition-colors ${
        activeMenu === menuName ? 'bg-gray-200' : ''
      }`}
      onClick={(e) => {
        e.stopPropagation();
        setActiveMenu(activeMenu === menuName ? null : menuName);
      }}
      onMouseEnter={() => {
        if (window.innerWidth > 768) {
          setActiveMenu(menuName);
        }
      }}
    >
      {menuName}
    </button>
    
    {activeMenu === menuName && (
      <div
        ref={menuRef}
        className="absolute left-0 mt-1 bg-white border border-gray-200 rounded-md shadow-lg z-50 min-w-48 py-1"
        onMouseLeave={() => {
          if (window.innerWidth > 768) {
            setActiveMenu(null);
          }
        }}
      >
        {menuItems[menuName].map((item, index) => (
          <button
            key={index}
            className={`w-full flex items-center gap-3 px-4 py-2 text-sm text-left ${
              item.disabled 
                ? 'text-gray-400 cursor-not-allowed' 
                : 'text-gray-700 hover:bg-gray-100'
            } ${
              item.active ? 'bg-blue-50 text-blue-700' : ''
            }`}
            onClick={() => {
              if (!item.disabled) {
                item.action();
                setActiveMenu(null);
              }
            }}
            disabled={item.disabled}
          >
            <item.icon 
              size={16} 
              className={`flex-shrink-0 ${
                item.active ? 'text-blue-600' : 'text-gray-500'
              }`} 
            />
            <span className="flex-grow">{item.label}</span>
            {item.shortcut && (
              <span className="text-xs text-gray-400 ml-2">
                {item.shortcut}
              </span>
            )}
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
                        className={`
  p-0
  ${showGridlines ? 'border border-gray-200' : 'border-none'}
  ${selectedCell === cellRef ? 'ring-1 ring-blue-500' : ''}`}
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
{/* Find/Replace Dialog */}
{findReplace.show && (
  <div className="fixed top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 bg-white p-4 rounded-lg shadow-xl z-50 w-96">
    <div className="flex justify-between items-center mb-4">
      <h3 className="text-lg font-medium">Find and Replace</h3>
      <button 
        onClick={() => setFindReplace(prev => ({...prev, show: false}))}
        className="text-gray-500 hover:text-gray-700"
      >
        <X size={20} />
      </button>
    </div>
    
    <div className="space-y-4">
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">Find</label>
        <input
          type="text"
          value={findReplace.find}
          onChange={(e) => setFindReplace(prev => ({...prev, find: e.target.value}))}
          className="w-full px-3 py-2 border border-gray-300 rounded-md"
          autoFocus
        />
      </div>
      
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">Replace with</label>
        <input
          type="text"
          value={findReplace.replace}
          onChange={(e) => setFindReplace(prev => ({...prev, replace: e.target.value}))}
          className="w-full px-3 py-2 border border-gray-300 rounded-md"
        />
      </div>
      
      <label className="flex items-center">
        <input
          type="checkbox"
          checked={findReplace.matchCase}
          onChange={(e) => setFindReplace(prev => ({...prev, matchCase: e.target.checked}))}
          className="h-4 w-4 text-blue-600 rounded"
        />
        <span className="ml-2 text-sm text-gray-700">Match case</span>
      </label>
      
      <div className="flex justify-between pt-2">
        <div className="flex space-x-2">
          <button
            onClick={() => handleFindReplace('find')}
            className="px-4 py-2 bg-gray-100 hover:bg-gray-200 rounded-md text-sm"
          >
            Find
          </button>
          <button
  onClick={() => {
    if (findReplace.results.length === 0) {
      // If no results, run find first
      handleFindReplace('find');
    } else {
      // If we have results, perform replace
      handleFindReplace('replace');
    }
  }}
  className={`px-4 py-2 rounded-md text-sm ${
    findReplace.results.length === 0
      ? 'bg-gray-200 text-gray-500 cursor-not-allowed'
      : 'bg-blue-100 hover:bg-blue-200 text-blue-800'
  }`}
  disabled={findReplace.results.length === 0}
>
  Replace All
</button>
        </div>
        
        {findReplace.results.length > 0 && (
          <div className="text-sm text-gray-500">
            {findReplace.currentIndex + 1} of {findReplace.results.length}
          </div>
        )}
      </div>
    </div>
  </div>
)}
{showHistory && (
  <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
    <div className="bg-white rounded-lg shadow-xl w-full max-w-2xl max-h-[80vh] flex flex-col">
      <div className="p-4 border-b flex justify-between items-center">
        <h3 className="text-lg font-medium">Version History ({history.length})</h3>
        <button 
          onClick={() => setShowHistory(false)}
          className="text-gray-500 hover:text-gray-700"
        >
          <X size={20} />
        </button>
      </div>
      
      <div className="overflow-y-auto flex-1">
        {history.map((version, index) => (
          <div 
            key={index}
            className={`p-3 border-b hover:bg-gray-50 cursor-pointer ${
              historyIndex === index ? 'bg-blue-50' : ''
            }`}
            onClick={() => restoreVersion(index)}
          >
            <div className="flex justify-between">
              <span className="font-medium">
                Version {index + 1} - {new Date(version.timestamp || Date.now()).toLocaleString()}
              </span>
              {historyIndex === index && (
                <span className="text-sm text-green-600">Current</span>
              )}
            </div>
            <div className="text-sm text-gray-500 mt-1">
              {Object.keys(version.cells || {}).length} cells modified
            </div>
          </div>
        ))}
      </div>
      
      <div className="p-3 border-t bg-gray-50 text-right">
        <button
          onClick={() => setShowHistory(false)}
          className="px-4 py-2 bg-gray-200 hover:bg-gray-300 rounded-md"
        >
          Close
        </button>
      </div>
    </div>
  </div>
)}

    {/* Notification */}
    {notification && (
  <div className="fixed top-4 right-4 bg-gray-800 text-white px-4 py-2 rounded-md shadow-lg text-sm animate-fade-in z-50">
    <div className="flex items-center justify-between gap-4">
      <span>{notification}</span>
      <button 
        onClick={() => setNotification(null)}
        className="text-gray-300 hover:text-white"
      >
        <X size={16} />
      </button>
    </div>
  </div>
)}

    {/* Hidden file input for opening files */}
    <input 
      type="file" 
      ref={fileInputRef} 
      onChange={importCSV} 
      className="hidden" 
      accept=".json,.csv"
    />
  </div>
);


};

