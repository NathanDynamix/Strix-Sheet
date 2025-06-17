import React, { useState, useEffect, useCallback,useRef } from 'react';
import {
  Download, Upload, Share2, BarChart3, Calculator, Grid, Save, Plus, Trash2,
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, PieChart,
  LineChart, TrendingUp, FunctionSquare, X, Search, DollarSign, Percent, Palette, Calendar, ChevronDown,WrapText, Mail, Link2, Users,Undo,Redo, Copy
} from 'lucide-react';
import { useChartData } from './ChartDataContext';
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

const SpreadsheetApp = () => {
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

  const copyToClipboard = async (text) => {
    try {
      await navigator.clipboard.writeText(text);
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 2000);
      return true;
    } catch (err) {
      console.error('Failed to copy: ', err);
      // Fallback for browsers that don't support clipboard API
      const textarea = document.createElement('textarea');
      textarea.value = text;
      document.body.appendChild(textarea);
      textarea.select();
      try {
        document.execCommand('copy');
        setIsCopied(true);
        setTimeout(() => setIsCopied(false), 2000);
        return true;
      } catch (err) {
        console.error('Fallback copy failed: ', err);
        return false;
      } finally {
        document.body.removeChild(textarea);
      }
    }
  };

  const handleCopyLink = () => {
    copyToClipboard(window.location.href);
    setShowShareDropdown(false);
  };

  const handleEmailShare = () => {
    const subject = "Check out this spreadsheet";
    const body = `I wanted to share this spreadsheet with you: ${window.location.href}`;
    window.open(`mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`);
    setShowShareDropdown(false);
  };

  const handleShareWithPeople = () => {
    if (navigator.share) {
      navigator.share({
        title: 'Spreadsheet',
        text: 'Check out this spreadsheet',
        url: window.location.href,
      }).catch(err => console.log('Error sharing:', err));
    } else {
      copyToClipboard(window.location.href);
    }
    setShowShareDropdown(false);
  };

  const handleGetShareableLink = () => {
    copyToClipboard(window.location.href);
    setShowShareDropdown(false);
  };

const { setSpreadsheetData, setDataSource } = useSpreadsheetData();

  // Your existing state and functions...

  const handleGenerateChart = () => {
    // Prepare the data structure
    const headers = [];
    const chartData = [];

    // Get headers from first row (A1, B1, etc.)
    for (let col = 0; col < 26; col++) { // A-Z columns
      const cellRef = `${String.fromCharCode(65 + col)}1`;
      const header = cells[cellRef]?.display || `Column ${col + 1}`;
      headers.push(header);
    }

    // Get data from rows 2-101
    for (let row = 2; row <= 101; row++) {
      const rowData = {};
      let hasData = false;

      for (let col = 0; col < headers.length; col++) {
        const colLetter = String.fromCharCode(65 + col);
        const cellRef = `${colLetter}${row}`;
        const value = cells[cellRef]?.display || '';
        
        // Convert to number if possible
        rowData[headers[col]] = isNaN(value) ? value : Number(value);
        if (value) hasData = true;
      }

      if (hasData) chartData.push(rowData);
    }

    // Save to context and navigate
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

  const generateAdvancedChart = () => {
    const data = [];
    for (let row = 1; row <= 10; row++) {
      const labelCell = coordsToCell(row, 0);
      const valueCell = coordsToCell(row, 1);
      const label = cells[labelCell]?.display || `Series ${row}`;
      const value = getCellNumericValue(valueCell);
      
      if (label && value !== 0) {
        data.push({ label, value, color: `hsl(${row * 36}, 70%, 50%)` });
      }
    }
    setChartData(data);
    setShowChart(true);
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
 
 return (
  <div className="w-full h-screen bg-white flex flex-col">
    {/* Toolbar */}
    <div className=" border-b border-gray-200 p-2 flex items-center gap-2 flex-wrap">
      <div className="relative" ref={dropdownRef}>
      {/* File Button */}
      <button 
        onClick={(e) => {
          e.stopPropagation();
          setIsOpen(!isOpen);
        }}
        className="flex items-center gap-2 px-3 py-1 text-black rounded hover:bg-gray-100 transition-colors"
        aria-expanded={isOpen}
        aria-haspopup="true"
      >
        File
        <ChevronDown 
          size={14} 
          className={`transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`}
        />
      </button>

      {/* Dropdown Menu */}
      {isOpen && (
        <div 
          className="absolute top-full left-0 mt-1 w-40 bg-white border border-gray-200 rounded-md shadow-lg z-50 py-1"
          role="menu"
        >
          {/* Save as Button */}
          <button
            onClick={exportToCSV}
            className="w-full flex items-center gap-2 px-3 py-2 text-sm text-gray-700 hover:bg-gray-100 text-left"
            role="menuitem"
          >
            <Download size={16} className="text-gray-500" />
            Save as CSV
          </button>
        </div>
      )}
    </div>
      <button className="flex items-center gap-2 px-3 py-1  text-black rounded hover:bg-gray-100 transition-colors">
        Edit
      </button>
      <button className="flex items-center gap-2 px-3 py-1  text-black rounded hover:bg-gray-100 transition-colors">View</button>
      <button className="flex items-center gap-2 px-3 py-1  text-black rounded hover:bg-gray-100 transition-colors">Format</button>
      <button className="flex items-center gap-2 px-3 py-1  text-black rounded hover:bg-gray-100 transition-colors">Data</button>
      <button className="flex items-center gap-2 px-3 py-1  text-black rounded hover:bg-gray-100 transition-colors">Tools</button>

      <div className="flex items-center gap-1">
        <label className="flex items-center gap-1 px-2 py-1 text-black rounded hover:bg-gray-100 cursor-pointer transition-colors">
          {/* <Upload size={16} /> */}
          Insert
          <input type="file" accept=".csv" onChange={importCSV} className="hidden" />
        </label>

        {/* <button onClick={exportToCSV} className="flex items-center gap-1 px-2 py-1 text-black rounded hover:bg-gray-100 transition-colors">
          Save as
        </button> */}
      </div>

      <div className="flex items-center gap-1">
        <button onClick={handleGenerateChart} className='flex items-center gap-1 px-2 py-1 text-black rounded hover:bg-gray-100 cursor-pointer transition-colors mr-[20%]'>
        Charts
      </button>
      </div>
         <div className="relative" ref={dropdownRef}>
      <button 
        onClick={(e) => {
          e.stopPropagation();
          setShowShareDropdown(!showShareDropdown);
        }}
        className="bg-blue-200 flex items-center gap-1 ml-[20%] px-4 py-2 text-black rounded hover:bg-blue-300 transition-colors"
      >
        <Share2 size={16} />
        Share
        <div className="w-[1.5px] h-7 bg-white mx-2"></div>
        <ChevronDown size={16}/>
      </button>

      {showShareDropdown && (
        <div className="absolute top-full left-0 mt-1 w-64 bg-white border border-gray-200 rounded-md shadow-lg z-50 py-1">
          <button
            onClick={handleCopyLink}
            className="w-full flex items-center px-4 py-2 text-sm text-gray-700 hover:bg-gray-100"
          >
            <Copy size={16} className="mr-3" />
            {isCopied ? 'Copied!' : 'Copy link'}
          </button>
          
          <button
            onClick={handleGetShareableLink}
            className="w-full flex items-center px-4 py-2 text-sm text-gray-700 hover:bg-gray-100"
          >
            <Link2 size={16} className="mr-3" />
            Get shareable link
          </button>
          
          <button
            onClick={handleEmailShare}
            className="w-full flex items-center px-4 py-2 text-sm text-gray-700 hover:bg-gray-100"
          >
            <Mail size={16} className="mr-3" />
            Email
          </button>
          
          <button
            onClick={handleShareWithPeople}
            className="w-full flex items-center px-4 py-2 text-sm text-gray-700 hover:bg-gray-100"
          >
            <Users size={16} className="mr-3" />
            Share with people
          </button>
        </div>
      )}
    </div>
    </div>

    {/* Format Bar */}
    <div className="bg-white border-b border-gray-200 p-2 flex items-center gap-2 flex-wrap">
      <div className="flex items-center gap-1 ml-4 relative">
        <button  className="p-2 hover:bg-gray-200 rounded">
                        <Undo size={16} />
                      </button>
                      <button  className="p-2 hover:bg-gray-200 rounded">
                        <Redo size={16} />
                        </button>
        <button
          onClick={() => formatCell({
            fontWeight: cells[selectedCell]?.style?.fontWeight === 'bold' ? 'normal' : 'bold'
          })}
          className={`p-2 rounded transition-colors ${cells[selectedCell]?.style?.fontWeight === 'bold' ? 'bg-gray-300' : 'hover:bg-gray-200'}`}
        >
          <Bold size={16} />
        </button>

        <button
          onClick={() => formatCell({
            fontStyle: cells[selectedCell]?.style?.fontStyle === 'italic' ? 'normal' : 'italic'
          })}
          className={`p-2 rounded transition-colors ${cells[selectedCell]?.style?.fontStyle === 'italic' ? 'bg-gray-300' : 'hover:bg-gray-200'}`}
        >
          <Italic size={16} />
        </button>

        <button
          onClick={() => formatCell({
            textDecoration: cells[selectedCell]?.style?.textDecoration === 'underline' ? 'none' : 'underline'
          })}
          className={`p-2 rounded transition-colors ${cells[selectedCell]?.style?.textDecoration === 'underline' ? 'bg-gray-300' : 'hover:bg-gray-200'}`}
        >
          <Underline size={16} />
        </button>

        <div className="w-px h-6 bg-gray-300 mx-2"></div>

        <button
    onClick={() => formatCell({ textAlign: 'left' })}
    className={`p-2 rounded transition-colors ${
      cells[selectedCell]?.style?.textAlign === 'left' ? 'bg-gray-300' : 'hover:bg-gray-200'
    }`}
    title="Align left"
    disabled={!selectedCell}
  >
    <AlignLeft size={16} />
  </button>

  <button
    onClick={() => formatCell({ textAlign: 'center' })}
    className={`p-2 rounded transition-colors ${
      cells[selectedCell]?.style?.textAlign === 'center' ? 'bg-gray-300' : 'hover:bg-gray-200'
    }`}
    title="Align center"
    disabled={!selectedCell}
  >
    <AlignCenter size={16} />
  </button>

  <button
    onClick={() => formatCell({ textAlign: 'right' })}
    className={`p-2 rounded transition-colors ${
      cells[selectedCell]?.style?.textAlign === 'right' ? 'bg-gray-300' : 'hover:bg-gray-200'
    }`}
    title="Align right"
    disabled={!selectedCell}
  >
    <AlignRight size={16} />
  </button>

        <div className="w-px h-6 bg-gray-300 mx-2"></div>

        <button className="p-2 hover:bg-gray-200 rounded transition-colors">
          <DollarSign size={16} />
        </button>

        <button className="p-2 hover:bg-gray-200 rounded transition-colors">
          <Percent size={16} />
        </button>

        <button
          onClick={() => setShowDropdown(!showDropdown)}
          className="p-2 hover:bg-gray-200 rounded flex items-center transition-colors"
        >
          <span>∑</span>
        </button>

        {showDropdown && (
          <div className="absolute top-full left-70 mt-1 flex bg-white border rounded shadow-lg z-10">
            <div className="w-48 border-r">
              {functionCategories.map((cat, idx) => (
                <div
                  key={idx}
                  onClick={() => setSelectedCategory(cat)}
                  onMouseEnter={() => setSelectedCategory(cat)}
                  className={`flex justify-between items-center px-4 py-2 text-sm cursor-pointer ${selectedCategory?.name === cat.name ? 'bg-blue-50' : 'hover:bg-gray-100'}`}
                >
                  {cat.name}
                  <span className="ml-2 text-gray-400">›</span>
                </div>
              ))}
            </div>

            {selectedCategory && (
              <div
                className="w-35 bg-white overflow-y-auto"
                style={{ maxHeight: '400px' }}
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
                      setSelectedCategory(null);
                      setIsHoveringSub(false);
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

        <button
          onClick={() => setShowColorPicker(!showColorPicker)}
          className="p-2 hover:bg-gray-200 rounded transition-colors"
        >
          <Palette size={16} />
        </button>

        {showColorPicker && (
          <div className="absolute top-full right-0 mt-1 w-48 bg-white border rounded-lg shadow-lg z-50 p-3">
            <div className="grid grid-cols-6 gap-2 mb-3">
              {['#000000', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF',
                '#00FFFF', '#FFA500', '#800080', '#008000', '#FFC0CB', '#A52A2A'].map(color => (
                  <button
                    key={color}
                    onClick={() => {
                      formatCell({ color });
                      setShowColorPicker(false);
                    }}
                    className="w-6 h-6 rounded border"
                    style={{ backgroundColor: color }}
                    title={color}
                  />
                ))}
            </div>
            <div className="border-t pt-2">
              <div className="text-sm font-medium mb-2">Background</div>
              <div className="grid grid-cols-6 gap-2">
                {['#FFFFFF', '#F0F0F0', '#FFEEEE', '#EEFFEE', '#EEEEFF', '#FFFFEE',
                  '#FFEEFF', '#EEFFFF', '#FFE4B5', '#E6E6FA', '#F0E68C', '#FFB6C1'].map(color => (
                    <button
                      key={color}
                      onClick={() => {
                        formatCell({ backgroundColor: color });
                        setShowColorPicker(false);
                      }}
                      className="w-6 h-6 rounded border"
                      style={{ backgroundColor: color }}
                      title={color}
                    />
                  ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>

    {/* Formula Bar */}
    <div className="bg-white border-b border-gray-200 p-2 flex items-center gap-2">
      <div className="flex items-center gap-2">
        <Calculator size={16} className="text-gray-500" />
        <span className="font-mono text-sm bg-gray-100 px-2 py-1 rounded min-w-[40px] font-semibold">
          {selectedCell}
        </span>
      </div>
      <input
        type="text"
        value={formulaBar}
        onChange={(e) => handleFormulaBarChange(e.target.value)}
        className="flex-1 px-3 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
        placeholder="Enter formula or value... (e.g., =SUM(A1:A10) or =PMT(0.05/12,360,100000))"
      />
    </div>

    {/* Main Content */}
      <div className="flex-1 flex relative">
        {/* Spreadsheet Grid */}
        <div className="flex-1 overflow-auto">
          <div className="inline-block min-w-full">
{resizing.active && (
  <div 
    className="fixed bg-blue-500 pointer-events-none z-50"
    style={{
      [resizing.type === 'col' ? 'left' : 'top']: dragPos,
      [resizing.type === 'col' ? 'width' : 'height']: '1px',
      [resizing.type === 'col' ? 'height' : 'width']: resizing.type === 'col' ? '100vh' : '100vw',
    }}
  />
)}
<table className="border-collapse">
  <thead>
    <tr>
      <th className="w-12 h-8 bg-gray-100 border border-gray-300 text-xs font-medium"></th>
      {Array.from({ length: COLS }, (_, i) => (
        <th 
          key={i} 
          className="relative bg-gray-100 border border-gray-300 text-xs font-medium"
          style={{ width: getColWidth(i) }}
        >
          {getColumnHeader(i)}
          <div 
            className="absolute top-0 right-0 w-1 h-full cursor-col-resize hover:bg-blue-500 active:bg-blue-700"
            onMouseDown={(e) => handleResizeMouseDown('col', i, e)}
          ></div>
        </th>
      ))}
    </tr>
  </thead>
  <tbody>
    {Array.from({ length: ROWS }, (_, row) => (
      <tr key={row}>
        <td 
          className="relative w-12 bg-gray-100 border border-gray-300 text-xs font-medium text-center"
          style={{ height: getRowHeight(row) }}
        >
          {row + 1}
          <div 
            className="absolute bottom-0 left-0 w-full h-1 cursor-row-resize hover:bg-blue-500 active:bg-blue-700"
            onMouseDown={(e) => handleResizeMouseDown('row', row, e)}
          ></div>
        </td>
        {Array.from({ length: COLS }, (_, col) => {
          const cellRef = coordsToCell(row, col);
          const cell = cells[cellRef];
          const isSelected = selectedCell === cellRef;
          
          return (
            <td 
              key={col} 
              className="border border-gray-300 p-0"
              style={{ 
                width: getColWidth(col),
                height: getRowHeight(row)
              }}
            >
              <input
                type="text"
                value={isSelected ? formulaBar : (cell?.display || '')}
                onChange={(e) => {
                  if (isSelected) {
                    handleFormulaBarChange(e.target.value);
                  }
                }}
                onFocus={() => handleCellClick(cellRef)}
                onClick={() => handleCellClick(cellRef)}
                className={`w-full h-full px-1 text-xs border-none outline-none ${
                  isSelected ? 'bg-blue-100 ring-2 ring-blue-500' : 'bg-transparent hover:bg-gray-50'
                } ${cell?.formula ? 'font-medium' : ''}`}
                style={{
                  textAlign: !isNaN(parseFloat(cell?.display || '')) ? 'right' : 'left',
                  color: cell?.style?.color || 'inherit',
                  backgroundColor: cell?.style?.backgroundColor || 'inherit',
                  fontWeight: cell?.style?.fontWeight || 'inherit',
                  fontStyle: cell?.style?.fontStyle || 'inherit',
                  textDecoration: cell?.style?.textDecoration || 'inherit',
                  whiteSpace: cell?.style?.whiteSpace || 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis'
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
      </div>

      {/* Sheet Tabs */}
      <div className="bg-gray-50 border-t border-gray-200 p-2 flex items-center gap-2">
        <div className="flex items-center gap-1">
          {sheets.map(sheet => (
            <div key={sheet.id} className="flex items-center">
              <button
                onClick={() => switchSheet(sheet.id)}
                className={`px-4 py-2 text-sm rounded-t transition-colors ${
                  sheet.active ? 'bg-white border-t border-l border-r border-gray-300 -mb-px' : 'bg-gray-200 hover:bg-gray-300'
                }`}
              >
                {sheet.name}
              </button>
              {sheets.length > 1 && (
                <button
                  onClick={() => deleteSheet(sheet.id)}
                  className="ml-1 p-1 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded transition-colors"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          ))}
        </div>
        
        <button
          onClick={addSheet}
          className="flex items-center gap-1 px-3 py-1 text-sm bg-gray-200 hover:bg-gray-300 rounded transition-colors"
        >
          <Plus size={16} />
          Add Sheet
        </button>
        
        <div className="ml-auto text-xs text-gray-500">
          Ready • {Object.keys(cells).filter(key => cells[key]?.value).length} cells with data
        </div>
      </div>
  </div>
);

};

export default SpreadsheetApp;