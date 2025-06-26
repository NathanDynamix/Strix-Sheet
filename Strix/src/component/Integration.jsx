import  { useState, useEffect, useCallback,useRef } from 'react';
import { evaluate, parse } from 'mathjs';
import {
Download, Upload, Share2, BarChart3, Calculator, Save, Plus, Trash2,
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, FunctionSquare, X, Search, DollarSign, Percent, Palette, 
  ChevronDown, WrapText, Mail, Link2, Users, Undo, Redo, Copy, 
  FileText, Folder, Clock, Undo2, Redo2, Scissors, Grid3X3, Eye, 
  ZoomIn, ZoomOut, Maximize, ArrowUpDown, Filter, Shield, Table, Settings, 
  Puzzle, HelpCircle, BookOpen, Clipboard, Image,  Keyboard,Link,ArrowUp,ArrowDown,ChevronRight
} from 'lucide-react';
import { formatDistanceToNow } from 'date-fns';
import * as math from 'mathjs';
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
const COLS = 26;
const ROWS = 100;
const getColLetter = (i) => String.fromCharCode(65 + i);
export default function SpreadsheetApp(){
const [selectedCell, setSelectedCell] = useState({ row: 0, col: 0 });
const getColumnName = (colIndex) => {
    return String.fromCharCode(65 + colIndex);
  };
 
const handleChange = (key, value) => {
  const updatedCells = {
    ...cells,
    [key]: {
      ...(cells[key] || {}),
      display: value
    }
  };
  updateCellsAndStyles(updatedCells, cellStyles); // Save value + keep styles
};
  const getSelectedValue = () => {
    return selectedCell ? cells[selectedCell] || '' : '';
  };
const [validations, setValidations] = useState({});
  const [showImageDialog, setShowImageDialog] = useState(false);
const [imageUrl, setImageUrl] = useState('');
const [previewImage, setPreviewImage] = useState(null);
const [activeImageTab, setActiveImageTab] = useState('url'); // 'url' or 'upload'
const fileInputRef = useRef(null);
  const [editingCell, setEditingCell] = useState(null);
  const [showLinkDialog, setShowLinkDialog] = useState(false);
const [submenuTimeout, setSubmenuTimeout] = useState(null);
const timeoutRef = useRef(null);
const [hoveredItem, setHoveredItem] = useState(null);
  const [activeTab, setActiveTab] = useState('url');
  const [currentInput, setCurrentInput] = useState('');
  const [isLoadingLink, setIsLoadingLink] = useState(false);
  const [linkText, setLinkText] = useState('');
  const [linkUrl, setLinkUrl] = useState('');
  const [activeMenu, setActiveMenu] = useState(null);
  const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 });
  const menuRef = useRef(null);
const [future, setFuture] = useState([]);
  const [sheetName, setSheetName] = useState('Untitled spreadsheet');
  const [images, setImages] = useState([]);
  const [chartDialog, setChartDialog] = useState({ show: false, type: 'line' });
  const [chartData, setChartData] = useState([]);
  const [zoom, setZoom] = useState(100);
  const [showFormulas, setShowFormulas] = useState(false);
  const [notification, setNotification] = useState(null);
  const [cellValidationRules, setCellValidationRules] = useState({});
const [validationDialog, setValidationDialog] = useState({
  show: false,
  cell: null,
  type: 'number',
  condition: 'between',
  min: '',
  max: '',
  list: '',
  customFormula: '',
  inputMessage: '',
  errorMessage: 'Invalid input',
});

const [scriptEditor, setScriptEditor] = useState({
  isOpen: false,
  scripts: [
    {
      id: 1,
      name: "Sample Script",
      code: `function helloWorld() {
  const cell = getActiveCell();
  cell.setValue("Hello World!");
}`,
      lastEdited: new Date(),
    },
  ],
  activeScriptId: 1,
});


  const [sortDialog, setSortDialog] = useState({
  show: false,
  range: null,
  sortBy: '',
  order: 'asc',
  hasHeader: true
});
const [selectedRange, setSelectedRange] = useState(null);
const [filterActive, setFilterActive] = useState(false);
const [filterValues, setFilterValues] = useState({}); // e.g., { A: 'apple' }
  const [cells, setCells] = useState({});
  const [clipboard, setClipboard] = useState(null); // holds copied cell data
const [cellStyles, setCellStyles] = useState({});
const [history, setHistory] = useState([]);
const [historyIndex, setHistoryIndex] = useState(-1);

  const [spellCheck, setSpellCheck] = useState({
  active: false,
  errors: [],
  currentErrorIndex: 0,
  dictionary: [
    'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'any', 'can', 'her', 'was', 'one', 'our', 'out', 'day', 'get', 'has', 'him', 'his', 
    'how', 'man', 'new', 'now', 'old', 'see', 'two', 'way', 'who', 'boy', 'did', 'its', 'let', 'put', 'say', 'she', 'too', 'use', 'why', 'ask',
    'big', 'buy', 'far', 'few', 'got', 'hit', 'hot', 'job', 'lot', 'may', 'run', 'set', 'try', 'act', 'add', 'age', 'ago', 'air', 'all', 'and',
    // Add more words as needed
  ],
  userDictionary: [],
  ignoreWords: [],
  suggestions: {}
});
const toggleStyle = (styleKey) => {
  if (!selectedCell) return;
  setCellStyles((prev) => {
    const prevStyle = prev[selectedCell] || {};
    return {
      ...prev,
      [selectedCell]: {
        ...prevStyle,
        [styleKey]: !prevStyle[styleKey]
      }
    };
  });
};

const toggleBold = () => toggleStyle('bold');
const toggleItalic = () => toggleStyle('italic');
const toggleUnderline = () => toggleStyle('underline');
const setAlignment = (alignment) => {
  if (!selectedCell) return;
  setCellStyles((prev) => {
    const prevStyle = prev[selectedCell] || {};
    return {
      ...prev,
      [selectedCell]: {
        ...prevStyle,
        align: alignment
      }
    };
  });
};

const alignLeft = () => setAlignment('left');
const alignCenter = () => setAlignment('center');
const alignRight = () => setAlignment('right');

useEffect(() => {
    const handleClickOutside = (event) => {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        setActiveMenu(null);
        setHoveredItem(null);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Clear timeout on unmount
  useEffect(() => {
    return () => {
      if (timeoutRef.current) clearTimeout(timeoutRef.current);
    };
  }, []);

  const handleMenuEnter = (menuName) => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    setActiveMenu(menuName);
  };





  const handleMenuLeave = () => {
    timeoutRef.current = setTimeout(() => {
      setActiveMenu(null);
      setHoveredItem(null);
    }, 300);
  };
const handleSortSimple = (order) => {
  if (!selectedRange) {
    setNotification('Please select a range using Shift+Click');
    return;
  }

  const [startKey, endKey] = selectedRange;
  const [startRow, startCol] = startKey.split(',').map(Number);
  const [endRow, endCol] = endKey.split(',').map(Number);

  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);

  const sortCol = left; // sort by first column in range

  const rows = [];
  for (let r = top; r <= bottom; r++) {
    rows.push({
      originalRow: r,
      sortValue: cells[`${r},${sortCol}`]?.value?.toString().toLowerCase() || '',
    });
  }

  rows.sort((a, b) =>
    order === 'asc'
      ? a.sortValue.localeCompare(b.sortValue)
      : b.sortValue.localeCompare(a.sortValue)
  );

  const newCells = { ...cells };
  rows.forEach((row, i) => {
    const toRow = top + i;
    for (let c = left; c <= right; c++) {
      const fromKey = `${row.originalRow},${c}`;
      const toKey = `${toRow},${c}`;
      newCells[toKey] = { ...cells[fromKey] };
    }
  });

  setCells(newCells);
  setNotification(`Sorted ${order === 'asc' ? 'A → Z' : 'Z → A'}`);
};
const handleFullSort = (order) => {
  const sortedRows = [];

  for (let r = 0; r < ROWS; r++) {
    const rowCells = {};
    for (let c = 0; c < COLS; c++) {
      const key = `${r},${c}`;
      rowCells[c] = cells[key] || {};
    }
    sortedRows.push({ rowIndex: r, data: rowCells });
  }

  sortedRows.sort((a, b) => {
    const valA = a.data[0]?.value?.toString().toLowerCase() || '';
    const valB = b.data[0]?.value?.toString().toLowerCase() || '';
    return order === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
  });

  const newCells = {};
  sortedRows.forEach((row, newIndex) => {
    for (let c = 0; c < COLS; c++) {
      const key = `${newIndex},${c}`;
      newCells[key] = row.data[c] || {};
    }
  });

  setCells(newCells);
  setNotification(`Sorted all rows by column A (${order === 'asc' ? 'A → Z' : 'Z → A'})`);
};



  const handleItemEnter = (item) => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    if (item.submenu) setHoveredItem(item);
  };
  const cellToCoords = (cellRef) => {
  const matches = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!matches) return { row: 0, col: 0 };
  
  const col = matches[1].split('').reduce((acc, char) => 
    acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const row = parseInt(matches[2]) - 1;
  
  return { row, col };
};

const coordsToCell = (row, col) => {
  let colName = '';
  let c = col + 1;
  while (c > 0) {
    let remainder = (c - 1) % 26;
    colName = String.fromCharCode(65 + remainder) + colName;
    c = Math.floor((c - 1) / 26);
  }
  return `${colName}${row + 1}`;
};

const ROWS = 100;
  const COLS = 26;


  const navigate = useNavigate();
 const [showShareDropdown, setShowShareDropdown] = useState(false);
  const dropdownRef = useRef(null);
  const [notifications, setNotifications] = useState([
  {
    id: 1,
    message: 'Your changes have been saved',
    timestamp: new Date(Date.now() - 3600000), // 1 hour ago
    read: false
  },
  {
    id: 2,
    message: 'New version available',
    timestamp: new Date(Date.now() - 86400000), // 1 day ago
    read: false
  }
]);
const [showNotifications, setShowNotifications] = useState(false);
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

  const formulaBtnRef = useRef(null);
  const [showHistory, setShowHistory] = useState(false);
const [colWidths, setColWidths] = useState({});
const [rowHeights, setRowHeights] = useState({});
const [resizing, setResizing] = useState({ active: false, type: null, index: null, startPos: 0 });
const [dragPos, setDragPos] = useState(null);
const [uploadedImages, setUploadedImages] = useState([]);
const { setSpreadsheetData, setDataSource } = useSpreadsheetData();
const [copiedValue, setCopiedValue] = useState(null);

const spellCheckDictionary = [
  // Add your dictionary words here
  'the', 'quick', 'brown', 'fox', 'jumps', 'over', 'lazy', 'dog',
  'hello', 'world', 'spreadsheet', 'function', 'example', 'correct'
  // Add more words as needed
];

function isWordCorrect(word) {
  // Remove punctuation from the end of words
  const cleanWord = word.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, '').toLowerCase();
  return spellCheckDictionary.includes(cleanWord);
}
// Add to your state



 


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
const SortDialog = () => {
  const availableColumns = Array.from({ length: 26 }, (_, i) => 
    String.fromCharCode(65 + i)
  );

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-md">
        <div className="p-4 border-b flex justify-between items-center">
          <h3 className="text-lg font-medium">Sort Range</h3>
          <button 
            onClick={() => setSortDialog(prev => ({...prev, show: false}))}
            className="text-gray-500"
          >
            <X size={20} />
          </button>
        </div>

        <div className="p-4 space-y-4">
          <div>
            <label className="block text-sm font-medium mb-1">Sort by column</label>
            <select
              value={sortDialog.sortBy}
              onChange={(e) => setSortDialog(prev => ({...prev, sortBy: e.target.value}))}
              className="w-full border rounded p-2"
            >
              {availableColumns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium mb-1">Order</label>
            <select
              value={sortDialog.order}
              onChange={(e) => setSortDialog(prev => ({...prev, order: e.target.value}))}
              className="w-full border rounded p-2"
            >
              <option value="asc">A → Z (Ascending)</option>
              <option value="desc">Z → A (Descending)</option>
            </select>
          </div>

          <label className="flex items-center">
            <input
              type="checkbox"
              checked={sortDialog.hasHeader}
              onChange={(e) => setSortDialog(prev => ({...prev, hasHeader: e.target.checked}))}
              className="h-4 w-4 text-blue-600 rounded"
            />
            <span className="ml-2 text-sm">My data has headers</span>
          </label>
        </div>

        <div className="p-4 border-t flex justify-end">
          <button
            onClick={() => setSortDialog(prev => ({...prev, show: false}))}
            className="px-4 py-2 mr-2 border rounded"
          >
            Cancel
          </button>
          <button
            onClick={applySort}
            className="px-4 py-2 bg-blue-600 text-white rounded"
          >
            Sort Data
          </button>
        </div>
      </div>
    </div>
  );
};

  const getCellStyle = (row, col) => {
    const key = getCellKey(row, col);
    return cellStyles[key] || {};
  };
  const applySort = () => {
  if (!sortDialog.range) {
    showNotification('No range selected');
    return;
  }

  const { start, end } = sortDialog.range;
  const sortColumn = sortDialog.sortBy;
  const sortOrder = sortDialog.order;
  const hasHeader = sortDialog.hasHeader;

  try {
    // Parse range boundaries
    const [startCol, startRow] = parseCellReference(start);
    const [endCol, endRow] = parseCellReference(end);
    const sortColIndex = sortColumn.charCodeAt(0) - 65;

    // Extract the data range
    const rows = [];
    for (let row = startRow; row <= endRow; row++) {
      const rowData = [];
      for (let col = startCol; col <= endCol; col++) {
        const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;
        rowData.push({
          ref: cellRef,
          value: cells[cellRef]?.value || '',
          display: cells[cellRef]?.display || ''
        });
      }
      rows.push(rowData);
    }

    // Separate header if needed
    const header = hasHeader ? rows.shift() : null;

    // Sort the data
    rows.sort((a, b) => {
      const aValue = a[sortColIndex]?.value || '';
      const bValue = b[sortColIndex]?.value || '';

      // Numeric comparison
      const aNum = parseFloat(aValue);
      const bNum = parseFloat(bValue);
      if (!isNaN(aNum) && !isNaN(bNum)) {
        return sortOrder === 'asc' ? aNum - bNum : bNum - aNum;
      }

      // String comparison
      return sortOrder === 'asc' 
        ? aValue.localeCompare(bValue)
        : bValue.localeCompare(aValue);
    });

    // Reattach header if it exists
    if (header) rows.unshift(header);

    // Create new cell data with sorted values
    const newCells = {...cells};
    let currentRow = startRow;
    
    rows.forEach(rowData => {
      rowData.forEach((cell, colIndex) => {
        const col = startCol + colIndex;
        const cellRef = `${String.fromCharCode(65 + col)}${currentRow + 1}`;
        newCells[cellRef] = {
          ...newCells[cellRef],
          value: cell.value,
          display: cell.display
        };
      });
      currentRow++;
    });

    setCells(newCells);
    showNotification(`Range sorted by column ${sortColumn}`);
    setSortDialog(prev => ({...prev, show: false}));

  } catch (error) {
    console.error('Sorting failed:', error);
    showNotification('Failed to sort range');
  }
};
const ScriptSandbox = ({ 
  script, 
  selectedCell, 
  cells, 
  setCells,
  onConsoleOutput,
  onScriptComplete
}) => {
  useEffect(() => {
    if (!script) return;

    const sandbox = {
      // Spreadsheet functions
      getActiveCell: () => {
        return {
          getValue: () => cells[selectedCell]?.value || '',
          setValue: (value) => {
            setCells(prev => ({
              ...prev,
              [selectedCell]: {
                value: value,
                display: value,
                formula: ''
              }
            }));
          }
        };
      },
      getCell: (ref) => {
        return {
          getValue: () => cells[ref]?.value || '',
          setValue: (value) => {
            setCells(prev => ({
              ...prev,
              [ref]: {
                value: value,
                display: value,
                formula: ''
              }
            }));
          }
        };
      },
      refreshSheet: () => {
        // This will trigger a re-render
        setCells(prev => ({ ...prev }));
      },
      // Console functions
      console: {
        log: (...args) => onConsoleOutput('log', args.join(' ')),
        error: (...args) => onConsoleOutput('error', args.join(' '))
      },
      // Limited Math functions
      Math: {
        abs: Math.abs,
        max: Math.max,
        min: Math.min,
        random: Math.random,
        round: Math.round,
        floor: Math.floor,
        ceil: Math.ceil,
        sqrt: Math.sqrt,
        pow: Math.pow,
        sin: Math.sin,
        cos: Math.cos,
        tan: Math.tan,
        log: Math.log,
        PI: Math.PI
      }
    };

    try {
      // Create the function with sandbox context
      const scriptFn = new Function(
        'getActiveCell',
        'getCell',
        'refreshSheet',
        'console',
        'Math',
        `
        try {
          ${script.code}
          
          // Call main if it exists
          if (typeof main === 'function') {
            main();
          }
        } catch (e) {
          console.error('Script error:', e.message);
          throw e;
        }
        `
      );

      // Execute the script with sandboxed functions
      scriptFn(
        sandbox.getActiveCell,
        sandbox.getCell,
        sandbox.refreshSheet,
        sandbox.console,
        sandbox.Math
      );

      onScriptComplete();
    } catch (error) {
      onConsoleOutput('error', `Script execution failed: ${error.message}`);
      onScriptComplete(error);
    }
  }, [script]);

  return null; // This component doesn't render anything
};
const ScriptEditor = ({ scriptEditor, setScriptEditor }) => {
  const activeScript = scriptEditor.scripts.find(s => s.id === scriptEditor.activeScriptId);

  const handleCodeChange = (e) => {
    setScriptEditor(prev => ({
      ...prev,
      scripts: prev.scripts.map(script => 
        script.id === prev.activeScriptId 
          ? {...script, code: e.target.value, lastEdited: new Date()}
          : script
      )
    }));
  };

  const handleNewScript = () => {
    const newId = Math.max(0, ...scriptEditor.scripts.map(s => s.id)) + 1;
    setScriptEditor(prev => ({
      ...prev,
      scripts: [
        ...prev.scripts,
        {
          id: newId,
          name: `Script ${prev.scripts.length + 1}`,
          code: 'function main() {\n  // Your code here\n  const cell = getActiveCell();\n  console.log("Current cell value:", cell.getValue());\n}',
          lastEdited: new Date(),
          isRunning: false
        }
      ],
      activeScriptId: newId,
      consoleOutput: []
    }));
  };

  const handleClearConsole = () => {
    setScriptEditor(prev => ({...prev, consoleOutput: []}));
  };

  const handleClose = () => {
    setScriptEditor(prev => ({...prev, isOpen: false}));
  };
function toggleFilter() {
  setFilterActive((prev) => !prev);

  // Optionally clear filters when turning off
  if (filterActive) {
    setFilters({});
  }
}
  const handleRunScript = () => {
    if (!activeScript || activeScript.isRunning) return;

    // Mark as running
    setScriptEditor(prev => ({
      ...prev,
      scripts: prev.scripts.map(s => 
        s.id === prev.activeScriptId ? {...s, isRunning: true} : s
      ),
      consoleOutput: [
        ...prev.consoleOutput,
        {type: 'log', message: `Running: ${activeScript.name}`}
      ]
    }));

    try {
      // Create a safer execution environment
      const sandbox = {
        // Spreadsheet functions
        getActiveCell: () => {
          return {
            getValue: () => cells[selectedCell]?.value || '',
            setValue: (value) => {
              setCells(prev => ({
                ...prev,
                [selectedCell]: {
                  value: value,
                  display: value,
                  formula: ''
                }
              }));
            }
          };
        },
        getCell: (ref) => {
          return {
            getValue: () => cells[ref]?.value || '',
            setValue: (value) => {
              setCells(prev => ({
                ...prev,
                [ref]: {
                  value: value,
                  display: value,
                  formula: ''
                }
              }));
            }
          };
        },
        getRange: (startRef, endRef) => {
          // Simple range implementation
          return {
            getValues: () => {
              const [startCol, startRow] = [startRef.match(/[A-Z]+/)[0], parseInt(startRef.match(/\d+/)[0])];
              const [endCol, endRow] = [endRef.match(/[A-Z]+/)[0], parseInt(endRef.match(/\d+/)[0])];
              
              const startColCode = startCol.charCodeAt(0);
              const endColCode = endCol.charCodeAt(0);
              
              const values = [];
              for (let row = startRow; row <= endRow; row++) {
                const rowValues = [];
                for (let col = startColCode; col <= endColCode; col++) {
                  const cellRef = `${String.fromCharCode(col)}${row}`;
                  rowValues.push(cells[cellRef]?.value || '');
                }
                values.push(rowValues);
              }
              return values;
            }
          };
        },
        refreshSheet: () => {
          // Force re-render by creating a new object
          setCells({...cells});
        },
        // Console functions
        console: {
          log: (...args) => {
            setScriptEditor(prev => ({
              ...prev,
              consoleOutput: [
                ...prev.consoleOutput,
                {type: 'log', message: args.map(arg => String(arg)).join(' ')}
              ]
            }));
          },
          error: (...args) => {
            setScriptEditor(prev => ({
              ...prev,
              consoleOutput: [
                ...prev.consoleOutput,
                {type: 'error', message: args.map(arg => String(arg)).join(' ')}
              ]
            }));
          }
        },
        // Math functions available in the sandbox
        Math: {
          abs: Math.abs,
          max: Math.max,
          min: Math.min,
          random: Math.random,
          round: Math.round,
          floor: Math.floor,
          ceil: Math.ceil,
          sqrt: Math.sqrt,
          pow: Math.pow,
          sin: Math.sin,
          cos: Math.cos,
          tan: Math.tan,
          log: Math.log,
          PI: Math.PI
        },
        // Date functions
        Date: Date
      };

      // Wrap the script in a try-catch and execute it
      const scriptToRun = `
        try {
          ${activeScript.code}
          
          // Call main if it exists
          if (typeof main === 'function') {
            main();
          }
        } catch (e) {
          console.error('Script error:', e.message);
        }
      `;

      // Create the function with all sandbox properties as parameters
      const paramNames = Object.keys(sandbox);
      const paramValues = paramNames.map(key => sandbox[key]);
      
      const scriptFn = new Function(...paramNames, scriptToRun);
      scriptFn(...paramValues);

    } catch (error) {
      setScriptEditor(prev => ({
        ...prev,
        consoleOutput: [
          ...prev.consoleOutput,
          {type: 'error', message: `Execution failed: ${error.message}`}
        ]
      }));
    } finally {
      setScriptEditor(prev => ({
        ...prev,
        scripts: prev.scripts.map(s => 
          s.id === prev.activeScriptId ? {...s, isRunning: false} : s
        )
      }));
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-5xl h-[80vh] flex flex-col">
        {/* Header */}
        <div className="p-4 border-b flex justify-between items-center">
          <h3 className="text-lg font-medium">Script Editor</h3>
          <div className="flex items-center gap-2">
            <select
              value={scriptEditor.activeScriptId}
              onChange={(e) => setScriptEditor(prev => ({
                ...prev,
                activeScriptId: Number(e.target.value)
              }))}
              className="border rounded px-2 py-1 text-sm"
            >
              {scriptEditor.scripts.map(script => (
                <option key={script.id} value={script.id}>{script.name}</option>
              ))}
            </select>
            <button
              onClick={handleNewScript}
              className="px-3 py-1 bg-blue-600 text-white rounded text-sm"
            >
              New
            </button>
            <button
              onClick={handleClose}
              className="text-gray-500 hover:text-gray-700"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Editor and Console */}
        <div className="flex flex-1 overflow-hidden">
          {/* Editor Panel */}
          <div className="flex-1 border-r flex flex-col">
            <div className="p-2 bg-gray-100 border-b flex justify-between">
              <span className="text-sm font-medium">{activeScript.name}</span>
              <span className="text-xs text-gray-500">
                Last edited: {activeScript.lastEdited.toLocaleString()}
              </span>
            </div>
            <textarea
              value={activeScript.code}
              onChange={handleCodeChange}
              className="flex-1 w-full p-4 font-mono text-sm outline-none resize-none"
              spellCheck="false"
            />
          </div>

          {/* Console Panel */}
          <div className="w-1/3 flex flex-col border-l">
            <div className="p-2 bg-gray-100 border-b">
              <h4 className="text-sm font-medium">Console Output</h4>
            </div>
            <div className="flex-1 overflow-y-auto p-4 bg-black text-green-400 font-mono text-sm">
              {scriptEditor.consoleOutput.map((msg, i) => (
                <div key={i} className={msg.type === 'error' ? 'text-red-400' : ''}>
                  {msg.message}
                </div>
              ))}
            </div>
            <div className="p-2 border-t flex justify-end gap-2">
              <button
                onClick={handleClearConsole}
                className="px-3 py-1 border rounded text-sm"
              >
                Clear
              </button>
              <button
                onClick={handleRunScript}
                disabled={activeScript?.isRunning}
                className={`px-3 py-1 rounded text-sm text-white ${
                  activeScript?.isRunning ? 'bg-gray-400' : 'bg-green-600 hover:bg-green-700'
                }`}
              >
                {activeScript?.isRunning ? 'Running...' : 'Run Script'}
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const startSpellCheck = () => {
  // Perform spell check and find errors
  const errors = [];
  
  // Example implementation - replace with your actual spell check logic
  Object.entries(cells).forEach(([cellRef, cell]) => {
    if (cell.display && typeof cell.display === 'string') {
      const words = cell.display.split(/\s+/);
      
      words.forEach(word => {
        const cleanWord = word.replace(/[^\w]/g, '').toLowerCase();
        if (cleanWord && !spellCheck.dictionary.includes(cleanWord)) {
          errors.push({
            cellRef,
            word,
            suggestions: getSuggestions(cleanWord)
          });
        }
      });
    }
  });
  
  if (errors.length === 0) {
    alert('No spelling errors found!');
    return;
  }
  
  setSpellCheck({
    active: true,
    errors,
    currentErrorIndex: 0,
    dictionary: spellCheck.dictionary // Preserve existing dictionary
  });
};

const getSuggestions = (word) => {
  if (!word) return [];
  
  // Simple suggestion algorithm - find similar words in dictionary
  return spellCheck.dictionary
    .filter(dictWord => dictWord.startsWith(word[0]))
    .slice(0, 4); // Return max 4 suggestions
};
const replaceWord = (cellRef, oldWord, newWord) => {
  setCells(prev => {
    const cell = prev[cellRef];
    if (!cell) return prev;
    
    const newDisplay = cell.display.replace(new RegExp(`\\b${oldWord}\\b`, 'g'), newWord);
    
    return {
      ...prev,
      [cellRef]: {
        ...cell,
        display: newDisplay,
        value: cell.formula ? cell.value : newDisplay
      }
    };
  });
  
  // Move to next error
  if (currentSpellCheckIndex < spellCheckResults.length - 1) {
    setCurrentSpellCheckIndex(prev => prev + 1);
  } else {
    setSpellCheckResults([]); // Close when done
  }
};

const ignoreWord = (word) => {
  // Just move to next error
  if (currentSpellCheckIndex < spellCheckResults.length - 1) {
    setCurrentSpellCheckIndex(prev => prev + 1);
  } else {
    setSpellCheckResults([]);
  }
};

const addToDictionary = (word) => {
  const cleanWord = word.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, '').toLowerCase();
  if (!spellCheckDictionary.includes(cleanWord)) {
    spellCheckDictionary.push(cleanWord);
  }
  
  // Move to next error
  if (currentSpellCheckIndex < spellCheckResults.length - 1) {
    setCurrentSpellCheckIndex(prev => prev + 1);
  } else {
    setSpellCheckResults([]);
  }
};
const [spellCheckResults, setSpellCheckResults] = useState([]);
const [currentSpellCheckIndex, setCurrentSpellCheckIndex] = useState(0);

const showSpellCheckUI = (results) => {
  setSpellCheckResults(results);
  setCurrentSpellCheckIndex(0);
  // You might want to show a modal or panel here
};

const SpellCheckModal = () => {
  if (spellCheckResults.length === 0) return null;
  
  const currentError = spellCheckResults[currentSpellCheckIndex];
  
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white p-6 rounded-lg w-96">
        <h3 className="text-lg font-bold mb-4">Spell Check</h3>
        
        <div className="mb-4">
          <p className="text-red-600">{currentError.word}</p>
          <p className="text-sm text-gray-600">in cell {currentError.cellRef}</p>
        </div>
        
        <div className="mb-4">
          <p className="font-medium mb-2">Suggestions:</p>
          <div className="space-y-2">
            {currentError.suggestions.map((suggestion, i) => (
              <button
                key={i}
                className="block w-full text-left p-2 hover:bg-gray-100 rounded"
                onClick={() => replaceWord(currentError.cellRef, currentError.word, suggestion)}
              >
                {suggestion}
              </button>
            ))}
          </div>
        </div>
        
        <div className="flex justify-between mt-4">
          <button 
            className="px-4 py-2 bg-gray-200 rounded"
            onClick={() => ignoreWord(currentError.word)}
          >
            Ignore
          </button>
          <button 
            className="px-4 py-2 bg-gray-200 rounded"
            onClick={() => addToDictionary(currentError.word)}
          >
            Add to Dictionary
          </button>
          <div className="flex space-x-2">
            <button 
              className="px-4 py-2 bg-blue-500 text-white rounded disabled:opacity-50"
              disabled={currentSpellCheckIndex === 0}
              onClick={() => setCurrentSpellCheckIndex(prev => prev - 1)}
            >
              Previous
            </button>
            <button 
              className="px-4 py-2 bg-blue-500 text-white rounded disabled:opacity-50"
              disabled={currentSpellCheckIndex === spellCheckResults.length - 1}
              onClick={() => setCurrentSpellCheckIndex(prev => prev + 1)}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
const highlightError = (error) => {
  setSelectedCell(error.cellRef);
  setEditingCell(error.cellRef);
};

const handleSpellCheckAction = (action, suggestion = null) => {
  const currentError = spellCheck.errors[spellCheck.currentErrorIndex];
  
  switch(action) {
    case 'replace':
      // Replace the word in the cell
      setCells(prevCells => {
        const newCells = {...prevCells};
        const cell = newCells[currentError.cellRef];
        
        if (cell) {
          const regex = new RegExp(`\\b${currentError.word}\\b`, 'gi');
          const newDisplay = cell.display.replace(regex, suggestion);
          
          newCells[currentError.cellRef] = {
            ...cell,
            display: newDisplay,
            value: cell.formula ? cell.value : newDisplay
          };
        }
        
        return newCells;
      });
      break;
      
    case 'ignore':
      // No action needed, just move to next error
      break;
      
    case 'add':
      // Add to dictionary
      setSpellCheck(prev => ({
        ...prev,
        dictionary: [...prev.dictionary, currentError.word.toLowerCase()]
      }));
      break;
  }
  
  // Move to next error or finish
  moveToNextError();
};

const moveToNextError = () => {
  setSpellCheck(prev => {
    if (prev.currentErrorIndex < prev.errors.length - 1) {
      return {...prev, currentErrorIndex: prev.currentErrorIndex + 1};
    } else {
      return {...prev, active: false};
    }
  });
};

const moveToPrevError = () => {
  setSpellCheck(prev => {
    if (prev.currentErrorIndex > 0) {
      return {...prev, currentErrorIndex: prev.currentErrorIndex - 1};
    }
    return prev;
  });
};

function generatePivotTableAndInsert({ cells, selectedCell, setCells, setNotification }) {
  const [rowStr, colStr] = selectedCell.split(',');
  const row = parseInt(rowStr);
  const col = parseInt(colStr);

  if (col < 2) {
    setNotification("Select a cell at least in column C");
    return;
  }

  const categoryCol = col - 2;
  const valueCol = col - 1;
  const pivotData = {};

  for (let r = 0; r < row; r++) {
    const catKey = `${r},${categoryCol}`;
    const valKey = `${r},${valueCol}`;

    const category = cells[catKey]?.value?.toString().trim();
    const rawValue = cells[valKey]?.value;
    const value = parseFloat(rawValue);

    if (category && !isNaN(value)) {
      if (!pivotData[category]) pivotData[category] = 0;
      pivotData[category] += value;
    }
  }

  const newCells = { ...cells };
  newCells[`${row},${col}`] = { value: "Category", display: "Category" };
  newCells[`${row},${col + 1}`] = { value: "Sum", display: "Sum" };

  let offset = 1;
  for (const [cat, sum] of Object.entries(pivotData)) {
    newCells[`${row + offset},${col}`] = { value: cat, display: cat };
    newCells[`${row + offset},${col + 1}`] = {
      value: sum.toString(),
      display: sum.toString(),
    };
    offset++;
  }

  setCells(newCells);
  setNotification("Pivot table created");
}

// Helper function to parse cell references
const parseCellReference = (ref) => {
  const matches = ref.match(/^([A-Z]+)(\d+)$/);
  if (!matches) return [0, 0];
  
  const col = matches[1].split('').reduce((acc, char) => 
    acc * 26 + char.charCodeAt(0) - 65, 0);
  const row = parseInt(matches[2]) - 1;
  
  return [col, row];
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
function handleImageInsert() {
  if (!selectedCell) {
    setNotification('Please select a cell first');
    return;
  }

  const url = prompt('Enter image URL:');
  if (url && /\.(jpg|jpeg|png|gif|webp|svg)$/i.test(url)) {
    const imageHTML = `<img src="${url}" alt="img" style="max-width:100%; max-height:100%;" />`;

    setCells((prev) => ({
      ...prev,
      [selectedCell]: {
        ...prev[selectedCell],
        display: imageHTML,
        value: url,
      },
    }));

    setNotification('Image inserted');
  } else {
    setNotification('Invalid image URL');
  }
}



const handleFileUpload = (e) => {
  const file = e.target.files[0];
  if (file && file.type.startsWith('image/')) {
    const reader = new FileReader();
    reader.onload = (event) => {
      setPreviewImage(event.target.result);
    };
    reader.readAsDataURL(file);
  }
};

const handleUrlChange = (e) => {
  const url = e.target.value;
  setImageUrl(url);
  if (url && (url.startsWith('http') || url.startsWith('data:'))) {
    setPreviewImage(url);
  } else {
    setPreviewImage(null);
  }
};
function isImageCell(text) {
  return typeof text === 'string' && text.includes('<img');
}


function handleImageInsert() {
  if (!selectedCell) {
    setNotification('Please select a cell first');
    return;
  }

  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';

  input.onchange = () => {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      const base64Image = `<img src="${reader.result}" alt="img" style="max-width:100%; max-height:100%;" />`;

      setCells((prev) => ({
        ...prev,
        [selectedCell]: {
          ...prev[selectedCell],
          display: base64Image,
          value: reader.result, // Save raw data if needed
        },
      }));

      setNotification('Image inserted');
    };

    reader.readAsDataURL(file);
  };

  input.click(); // Trigger file picker
}

 
  const removeImage = (cellRef) => {
  setCells(prev => ({
    ...prev,
    [cellRef]: {
      value: '',
      type: 'text',
      style: prev[cellRef]?.style || {}
    }
  }));
};




const saveToHistory = () => {
  setHistory((prev) => [...prev, { cells: { ...cells }, cellStyles: { ...cellStyles } }]);
  setFuture([]); // Clear redo stack on new change
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

const updateCellsAndStyles = (newCells, newStyles) => {
  const state = {
    cells: newCells,
    styles: newStyles
  };
  const newHistory = history.slice(0, historyIndex + 1);
  newHistory.push(state);
  setHistory(newHistory);
  setHistoryIndex(newHistory.length - 1);
  setCells(newCells);
  setCellStyles(newStyles);
};


const handleUndo = () => {
  if (historyIndex > 0) {
    const prev = history[historyIndex - 1];
    setCells(prev.cells);
    setCellStyles(prev.styles);
    setHistoryIndex(historyIndex - 1);
  }
};

const handleRedo = () => {
  if (historyIndex < history.length - 1) {
    const next = history[historyIndex + 1];
    setCells(next.cells);
    setCellStyles(next.styles);
    setHistoryIndex(historyIndex + 1);
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

  // Pass 1: Identify active columns with non-zero values
  for (let row = 1; row < ROWS; row++) {
    for (let col = 0; col < COLS; col++) {
      const key = `${row},${col}`;
      const cell = cells[key];
      const value = typeof cell === 'object' ? cell.display ?? cell.value : cell;

      if (value && value !== "0" && value !== "0.0") {
        activeColumns.add(col);
      }
    }
  }

  // Pass 2: Get column headers from row 0
  const headers = [];
  activeColumns.forEach((col) => {
    const key = `0,${col}`;
    const headerCell = cells[key];
    const name = typeof headerCell === 'object' ? headerCell.display ?? headerCell.value : headerCell;
    headers.push({
      index: col,
      name: name || `Column ${col + 1}`
    });
  });

  // Pass 3: Construct chart data rows
  const chartData = [];
  for (let row = 1; row < ROWS; row++) {
    const rowData = {};
    let hasData = false;

    headers.forEach((header) => {
      const key = `${row},${header.index}`;
      const cell = cells[key];
      const value = typeof cell === 'object' ? cell.display ?? cell.value : cell;

      if (value && value !== "0") {
        rowData[header.name] = isNaN(value) ? value : Number(value);
        hasData = true;
      }
    });

    if (hasData) chartData.push(rowData);
  }

  // Final step: save and navigate
  setSpreadsheetData(chartData);
  setDataSource("spreadsheet");
  navigate("/charts");
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


  const getColumnHeader = (index) => String.fromCharCode(65 + index);



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
 
const validateCell = (cellRef, value) => {
  const validation = validations[cellRef];
  if (!validation) return true;
  
  const isValid = validation.validator(value);
  
  if (!isValid) {
    switch (validation.errorStyle) {
      case 'stop':
        alert(validation.errorMessage);
        return false;
      case 'warning':
        if (!window.confirm(`${validation.errorMessage}\n\nDo you want to continue anyway?`)) {
          return false;
        }
        return true;
      case 'info':
        alert(validation.errorMessage);
        return true;
      default:
        return false;
    }
  }
  
  return true;
};


const evaluateFormula = (formula, cells) => {
  try {
    if (typeof formula !== 'string' || !formula.startsWith('=')) {
      return formula; // Return as-is if not a formula
    }

    // Remove the leading '='
    const expression = formula.slice(1);

    // Replace cell references with their values (e.g., A1 with its value)
    const parsedExpression = expression.replace(
      /([A-Z]+)([0-9]+)/g, 
      (match, col, row) => {
        const cellRef = `${col}${row}`;
        const cellValue = cells[cellRef]?.value || cells[cellRef] || '';
        
        // If the referenced cell has a formula, evaluate it recursively
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return evaluateFormula(cellValue, cells);
        }
        
        // Convert to number if possible
        const numValue = parseFloat(cellValue);
        return isNaN(numValue) ? `"${cellValue}"` : numValue;
      }
    );

    // Use math.js to evaluate the parsed expression
    return math.evaluate(parsedExpression);
  } catch (error) {
    console.error('Formula error:', error);
    return '#ERROR';
  }
};



const handleCellChange = (cellRef, value) => {
  setCells(prev => {
    const isFormula = value.startsWith('=');
    const newCells = { ...prev };
    
    // Update the current cell
    newCells[cellRef] = {
      ...(newCells[cellRef] || {}),
      value,
      formula: isFormula ? value : '',
      display: isFormula ? evaluateFormula(value, cellRef) : value
    };

    // Find all cells that depend on this cell and update them
    Object.keys(newCells).forEach(ref => {
      if (ref !== cellRef && newCells[ref]?.formula?.includes(cellRef.replace(/(\d+)/, '$1'))) {
        newCells[ref].display = evaluateFormula(newCells[ref].formula, ref);
      }
    });

    return newCells;
  });
};

  // Add to your cell click handler
const handleCellClick = (cellRef, e) => {
  if (e.shiftKey && selectedCell) {
    setSelectedRange({
      start: selectedCell,
      end: cellRef
    });
  } else {
    setSelectedCell(cellRef);
    setSelectedRange(null);
  }
  setEditingCell(null);
};
const handleFormulaBarChange = (newValue) => {
  setFormulaBar(newValue);
  
  if (selectedCell) {
    handleCellChange(selectedCell, newValue);
  }
};

// When a cell is selected, update the formula bar
useEffect(() => {
  if (selectedCell) {
    const cell = cells[selectedCell];
    setFormulaBar(cell?.formula || cell?.value || '');
  }
}, [selectedCell, cells]);
useEffect(() => {
  if (selectedCell && cells[selectedCell]) {
    const content = cells[selectedCell].formula || cells[selectedCell].value || '';
    setFormulaBar(content);
  } else {
    setFormulaBar('');
  }
}, [selectedCell, cells]);




  const insertFunction = (funcName) => {
    const newFormula = `=${funcName}()`;
    setFormulaBar(newFormula);
    setShowDropdown(false);
    setHoveredCategory(null);
    handleCellChange(selectedCell, newFormula);
  };



function formatCell(styleUpdate) {
  setCellStyles((prev) => ({
    ...prev,
    [selectedCell]: {
      ...prev[selectedCell],
      ...styleUpdate,
    },
  }));
}

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
// In your cell rendering logic (where you create the td elements)
const renderCellContent = (cell, cellRef) => {
  const isEditing = editingCell === cellRef;
  const showValue = showFormulas 
    ? (cell?.formula || cell?.value || '')
    : (cell?.display || cell?.value || '');

  if (isEditing) {
    return (
      <input
        type="text"
        value={formulaBar}
        onChange={(e) => handleFormulaBarChange(e.target.value)}
        className="w-full h-full px-2 outline-none bg-white"
        autoFocus
        onBlur={() => setEditingCell(null)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            setEditingCell(null);
          }
        }}
      />
    );
  }

  return (
    <div className="w-full h-full px-2 overflow-hidden">
      {showValue}
    </div>
  );
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
const handleSortSelectedColumn = (order = 'asc') => {
  if (!selectedCell) {
    setNotification('Please select a cell first');
    return;
  }

  const [, colIndex] = selectedCell.split(',').map(Number);

  // Collect values from the selected column
  const columnData = Array.from({ length: ROWS }, (_, r) => {
    const key = `${r},${colIndex}`;
    return {
      row: r,
      key,
      value: cells[key]?.display ?? ''
    };
  });

  // Sort column data
  columnData.sort((a, b) => {
    const aVal = a.value?.toString() ?? '';
    const bVal = b.value?.toString() ?? '';

    const aNum = Number(aVal);
    const bNum = Number(bVal);
    const bothNumbers = !isNaN(aNum) && !isNaN(bNum);

    if (bothNumbers) {
      return order === 'asc' ? aNum - bNum : bNum - aNum;
    } else {
      return order === 'asc'
        ? aVal.localeCompare(bVal)
        : bVal.localeCompare(aVal);
    }
  });

  // Apply sorted values back to cells
  setCells(prev => {
    const updated = { ...prev };
    columnData.forEach((cellData, newRowIdx) => {
      const newKey = `${newRowIdx},${colIndex}`;
      updated[newKey] = {
        ...(updated[newKey] || {}),
        display: cellData.value,
        value: cellData.value
      };
    });
    return updated;
  });

  setNotification(
    `Sorted column ${String.fromCharCode(65 + colIndex)} ${
      order === 'asc' ? 'A → Z' : 'Z → A'
    }`
  );
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
    const rows = csv.trim().split('\n');
    const newCells = {};

    rows.forEach((row, rowIndex) => {
      const columns = row.split(',');
      columns.forEach((value, colIndex) => {
        if (rowIndex < ROWS && colIndex < COLS) {
          const key = `${rowIndex},${colIndex}`;
          const trimmed = value.trim();
          newCells[key] = {
            value: trimmed,
            display: trimmed,
            formula: ''
          };
        }
      });
    });

    setCells((prev) => ({ ...prev, ...newCells }));
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
function validateValue(value, rule) {
  if (rule.type === 'number') {
    const num = parseFloat(value);
    if (isNaN(num)) return false;

    if (rule.condition === 'between') {
      const min = parseFloat(rule.min);
      const max = parseFloat(rule.max);
      return num >= min && num <= max;
    }
    if (rule.condition === 'equal') {
      return num === parseFloat(rule.min);
    }
  }

  if (rule.type === 'list') {
    const options = rule.list.split(',').map(opt => opt.trim());
    return options.includes(value);
  }

  return true;
}


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



 const handleSort = (range, order = 'asc') => {
  const [start, end] = range;
  const [startRow, startCol] = start.split(',').map(Number);
  const [endRow, endCol] = end.split(',').map(Number);

  const col = startCol; // sort by this column
  const rows = [];

  for (let r = startRow; r <= endRow; r++) {
    rows.push({ 
      key: r, 
      value: cells[`${r},${col}`]?.value || '' 
    });
  }

  rows.sort((a, b) => {
    const valA = a.value.toString().toLowerCase();
    const valB = b.value.toString().toLowerCase();
    return order === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
  });

  const newCells = { ...cells };
  rows.forEach((row, i) => {
    for (let c = startCol; c <= endCol; c++) {
      const oldKey = `${row.key},${c}`;
      const newKey = `${startRow + i},${c}`;
      newCells[newKey] = { ...cells[oldKey] };
    }
  });

  setCells(newCells);
};



  

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
const handleSaveValidation = (rule) => {
  setCellValidationRules(prev => ({
    ...prev,
    [rule.cell]: rule
  }));
  setValidationDialog({ show: false });
  setNotification('Validation rule saved');
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
  currentIndex: 0
});

const handleFindReplace = (action) => {
  if (action === 'find') {
    const results = Object.keys(cells).filter((key) => {
      const value = cells[key]?.display || cells[key] || '';
      const search = findReplace.find;
      return findReplace.matchCase
        ? value.includes(search)
        : value.toLowerCase().includes(search.toLowerCase());
    });

    setFindReplace((prev) => ({
      ...prev,
      results,
      currentIndex: 0
    }));

    if (results.length > 0) {
      setSelectedCell(results[0]);
    }
  }

  // Optional: replace logic
};




function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
const handleInputChange = async (value) => {
    setCurrentInput(value);
    const cellKey = getCellKey(selectedCell.row, selectedCell.col);
    
    // Check if the input is a URL
    if (isValidUrl(value)) {
      setIsLoadingLink(true);
      try {
        const title = await fetchPageTitle(value);
        setCells(prev => ({
          ...prev,
          [cellKey]: { 
            value: title, 
            type: 'link', 
            url: value,
            displayText: title
          }
        }));
        setCurrentInput(title);
      } catch (error) {
        setCells(prev => ({
          ...prev,
          [cellKey]: { value, type: 'text' }
        }));
      }
      setIsLoadingLink(false);
    } else {
      setCells(prev => ({
        ...prev,
        [cellKey]: { ...prev[cellKey], value, type: 'text' }
      }));
    }
  };
 const insertLink = () => {
  if (!selectedCell || !linkText) {
    showNotification('Link text cannot be empty');
    return;
  }

  // Add https:// if missing
  let finalUrl = linkUrl;
  if (linkUrl && !/^https?:\/\//i.test(linkUrl)) {
    finalUrl = `https://${linkUrl}`;
  }

  setCells(prev => ({
    ...prev,
    [selectedCell]: {
      ...prev[selectedCell],
      type: 'link',
      value: linkText,
      display: linkText,
      url: finalUrl,
      style: prev[selectedCell]?.style || {}
    }
  }));

  setShowLinkDialog(false);
  showNotification('Link inserted successfully');
};
   const removeLink = (row, col) => {
    const cellKey = getCellKey(row, col);
    const currentCell = cells[cellKey];
    setCells(prev => ({
      ...prev,
      [cellKey]: { value: currentCell.displayText || currentCell.value, type: 'text' }
    }));
  };

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
const applyValidation = () => {
  const { cell, type, condition, min, max, list, customFormula, inputMessage, errorMessage } = validationDialog;
  
  let validator;
  
  switch (type) {
    case 'number':
      validator = (value) => {
        const num = parseFloat(value);
        if (isNaN(num)) return false;
        
        switch (condition) {
          case 'between': 
            return num >= parseFloat(min) && num <= parseFloat(max);
          case 'greater': 
            return num > parseFloat(min);
          case 'less': 
            return num < parseFloat(min);
          case 'equal': 
            return num === parseFloat(min);
          case 'notEqual': 
            return num !== parseFloat(min);
          default: 
            return true;
        }
      };
      break;
      
    case 'list':
      const items = list.split(',').map(item => item.trim());
      validator = (value) => items.includes(value);
      break;
      
    case 'custom':
      validator = (value) => {
        try {
          // This is a simplified implementation - you'd need a proper formula evaluator
          return eval(customFormula.replace(/=/g, '').replace(/value/g, value));
        } catch {
          return false;
        }
      };
      break;
      
    default:
      validator = () => true;
  }
  
  setValidations(prev => ({
    ...prev,
    [cell]: {
      validator,
      inputMessage,
      errorMessage
    }
  }));
  
  setNotification(`Validation rules set for ${cell}`);
  setValidationDialog({...validationDialog, show: false});
};
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


const handleCopy = () => {
  if (!selectedCell || !cells[selectedCell]) return;

  setClipboard({
    key: selectedCell,
    value: cells[selectedCell],
    style: cellStyles[selectedCell] || {}
  });
};


const handleCut = () => {
  if (selectedCell && cells[selectedCell] !== undefined) {
    const val = typeof cells[selectedCell] === 'object'
      ? cells[selectedCell].display || ''
      : cells[selectedCell];

    setCopiedValue(val); // reuse the copy buffer

    const newCells = { ...cells };
    if (typeof newCells[selectedCell] === 'object') {
      newCells[selectedCell] = { ...newCells[selectedCell], display: '' };
    } else {
      newCells[selectedCell] = '';
    }

    setCells(newCells);
  }
};
function getActiveCell() {
  const key = `${selectedCell.row},${selectedCell.col}`;
  return {
    setValue: (val) => {
      setCells(prev => ({
        ...prev,
        [key]: {
          ...(prev[key] || {}),
          value: val,
          display: val
        }
      }));
    },
    getValue: () => cells[key]?.value || '',
    getRow: () => selectedCell.row,
    getCol: () => selectedCell.col
  };
}




function runScript(code) {
  try {
    const func = new Function("getActiveCell", code);
    func(getActiveCell);
    setNotification("Script executed!");
  } catch (err) {
    setNotification("Script error: " + err.message);
  }
}




const handlePaste = () => {
  if (!selectedCell || !clipboard) return;

  const newCells = {
    ...cells,
    [selectedCell]: clipboard.value
  };

  const newStyles = {
    ...cellStyles,
    [selectedCell]: clipboard.style
  };

  updateCellsAndStyles(newCells, newStyles); // or setCells & setCellStyles if you're not using history
};
const insertRow = (position) => {
  if (!selectedCell) return;

  const [rowStr] = selectedCell.split(',');
  const targetRow = parseInt(rowStr, 10);
  const insertAt = position === 'above' ? targetRow : targetRow + 1;

  const newCells = {};
  const newStyles = {};

  for (let r = ROWS - 1; r >= insertAt; r--) {
    for (let c = 0; c < COLS; c++) {
      const oldKey = `${r - 1},${c}`;
      const newKey = `${r},${c}`;
      if (cells[oldKey]) newCells[newKey] = cells[oldKey];
      if (cellStyles[oldKey]) newStyles[newKey] = cellStyles[oldKey];
    }
  }

  // Copy cells above the inserted row
  for (let r = 0; r < insertAt; r++) {
    for (let c = 0; c < COLS; c++) {
      const key = `${r},${c}`;
      if (cells[key]) newCells[key] = cells[key];
      if (cellStyles[key]) newStyles[key] = cellStyles[key];
    }
  }

  // Clear the new row
  for (let c = 0; c < COLS; c++) {
    const key = `${insertAt},${c}`;
    newCells[key] = '';
    newStyles[key] = {};
  }

  setCells(newCells);
  setCellStyles(newStyles);
  showNotification('Row inserted');
};


  
  const insertColumn = (position) => {
  if (!selectedCell) return;

  const [, colStr] = selectedCell.split(',');
  const targetCol = parseInt(colStr, 10);
  const insertAt = position === 'left' ? targetCol : targetCol + 1;

  const newCells = {};
  const newStyles = {};

  for (let r = 0; r < ROWS; r++) {
    for (let c = COLS - 1; c >= insertAt; c--) {
      const oldKey = `${r},${c - 1}`;
      const newKey = `${r},${c}`;
      if (cells[oldKey]) newCells[newKey] = cells[oldKey];
      if (cellStyles[oldKey]) newStyles[newKey] = cellStyles[oldKey];
    }
  }

  // Copy cells left of inserted column
  for (let r = 0; r < ROWS; r++) {
    for (let c = 0; c < insertAt; c++) {
      const key = `${r},${c}`;
      if (cells[key]) newCells[key] = cells[key];
      if (cellStyles[key]) newStyles[key] = cellStyles[key];
    }
  }

  // Clear the inserted column
  for (let r = 0; r < ROWS; r++) {
    const key = `${r},${insertAt}`;
    newCells[key] = '';
    newStyles[key] = {};
  }

  setCells(newCells);
  setCellStyles(newStyles);
  showNotification('Column inserted');
};

  
  const toggleFilter = () => {
    setFilterActive(!filterActive);
    showNotification(filterActive ? 'Filter removed' : 'Filter applied');
  };
  const handleZoomIn = () => setZoom(Math.min(200, zoom + 25));
  const handleZoomOut = () => setZoom(Math.max(50, zoom - 25));
  const toggleFormulas = () => {
  setShowFormulas(!showFormulas);
  setNotification(`Formulas ${!showFormulas ? 'shown' : 'hidden'}`);
};
const restoreVersion = (index) => {
  if (index < 0 || index >= history.length) return;
  
  const version = history[index];
  setCells(version.cells || {});
  setHistoryIndex(index);
  setShowHistory(false);
  showNotification(`Restored version ${index + 1}`);
};
const [showGridlines, setShowGridlines] = useState(true);
useEffect(() => {
  console.log('Current gridlines state:', showGridlines);
}, [showGridlines]);
const toggleGridlines = () => {
  setShowGridlines(prev => !prev);
};

function generatePivotTable(cells, categoryCol, valueCol, maxRow = 1000) {
  const pivot = {};

  for (let row = 1; row <= maxRow; row++) {
    const catKey = `${categoryCol}${row}`;
    const valKey = `${valueCol}${row}`;

    const category = cells[catKey]?.value;
    const value = parseFloat(cells[valKey]?.value);

    if (!category || isNaN(value)) continue;

    if (!pivot[category]) pivot[category] = 0;
    pivot[category] += value;
  }

  return pivot;
}




function handleLinkInsert() {
  if (!selectedCell) {
    setNotification('Please select a cell first');
    return;
  }

  const url = prompt('Enter URL:', 'https://');
  if (url && /^https?:\/\/.+/.test(url)) {
    setCells((prev) => ({
      ...prev,
      [selectedCell]: {
        ...prev[selectedCell],
        display: url,
        value: url
      },
    }));
    setNotification('Link inserted');
  } else {
    setNotification('Invalid URL');
  }
}
function isLinkCell(text) {
  return typeof text === 'string' && /^https?:\/\/\S+$/.test(text);
}

function isImageCell(text) {
  return typeof text === 'string' && text.includes('<img');
}

function isRichCell(text) {
  return isImageCell(text) || isLinkCell(text);
}






const menuItems = {
  File: [
    { 
      icon: FileText, 
      label: 'New', 
      action: () => {
        addSheet();
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
      action: handleUndo,
      shortcut: 'Ctrl+Z',
      disabled: historyIndex === 0
    },
    { 
      icon: Redo2, 
      label: 'Redo', 
      action: handleRedo,
      shortcut: 'Ctrl+Y',
      disabled: historyIndex === history.length - 1
    },
    {
  icon: Scissors, // or another suitable icon
  label: 'Cut',
  action: handleCut,
  shortcut: 'Ctrl+X',
  disabled: !selectedCell || cells[selectedCell] === undefined
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
  disabled: !clipboard
}
,
    {
      icon: Search, 
      label: 'Find and replace', 
      action: () => setFindReplace({ 
        show: true,
        find: '',
        replace: '',
        matchCase: false,
        results: [],
        currentIndex: 0
      }),
      shortcut: 'Ctrl+F'
    },
    {
      icon: Trash2, 
      label: 'Delete', 
      action: () => {
        if (selectedCell) {
          handleCut();
          setNotification('Cell content deleted');
        } else {
          setNotification('No cell selected');
        }
      },
      shortcut: 'Delete'
    }
  ],
  View: [
    {
      icon: Grid3X3,
      label: showGridlines ? 'Hide gridlines' : 'Show gridlines',
      action: toggleGridlines,
      className: showGridlines ? 'bg-blue-100' : ''
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
  disabled: !selectedCell
}
,
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
      action: handleImageInsert
    },
    {
      icon: Link2,
      label: 'Link',
      action: handleLinkInsert,
      shortcut: 'Ctrl+K'
    }
  ],
  Format: [
    { 
      icon: Bold, 
      label: 'Bold', 
      action: toggleBold,
      shortcut: 'Ctrl+B',
      active: selectedCell && cells[selectedCell]?.style?.fontWeight === 'bold'
    },
    { 
      icon: Italic, 
      label: 'Italic', 
      action: toggleItalic,
      shortcut: 'Ctrl+I',
      active: selectedCell && cells[selectedCell]?.style?.fontStyle === 'italic'
    },
    { 
      icon: Underline, 
      label: 'Underline', 
      action: toggleUnderline,
      shortcut: 'Ctrl+U',
      active: selectedCell && cells[selectedCell]?.style?.textDecoration === 'underline'
    },
    { 
      icon: AlignLeft, 
      label: 'Align left', 
      action: alignLeft,
      active: selectedCell && cells[selectedCell]?.style?.textAlign === 'left'
    },
    { 
      icon: AlignCenter, 
      label: 'Align center', 
      action: alignCenter,
      active: selectedCell && cells[selectedCell]?.style?.textAlign === 'center'
    },
    { 
      icon: AlignRight, 
      label: 'Align right', 
      action: alignRight,
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
      const isCurrentlyWrapped = cellStyles[selectedCell]?.whiteSpace === 'normal';

      formatCell({
        whiteSpace: isCurrentlyWrapped ? 'nowrap' : 'normal',
      });

      setNotification(isCurrentlyWrapped ? 'Text unwrapped' : 'Text wrapped');
    } else {
      setNotification('Please select a cell first');
    }
  },
  active: selectedCell && cellStyles[selectedCell]?.whiteSpace === 'normal',
  shortcut: 'Ctrl+Shift+W',
}

  ],
  Data: [
    {
      icon: ArrowUpDown,
      label: 'Sort',
      submenu: [
        {
          icon: ArrowUp,
          label: 'A → Z (Ascending)',
          action: () => {
            if (!selectedRange) {
              showNotification('Please select a range first (Shift+Click)');
              return;
            }
            handleSort(selectedRange, 'asc');
          },
          shortcut: 'Alt+Shift+A'
        },
        {
          icon: ArrowDown,
          label: 'Z → A (Descending)',
          action: () => {
            if (!selectedRange) {
              showNotification('Please select a range first (Shift+Click)');
              return;
            }
            handleSort(selectedRange, 'desc');
          },
          shortcut: 'Alt+Shift+Z'
        },
        {
          icon: Settings,
          label: 'Custom sort...',
          action: () => {
            if (!selectedRange) {
              showNotification('Please select a range first (Shift+Click)');
              return;
            }
            setSortDialog({
              show: true,
              range: selectedRange,
              sortBy: 'A',
              order: 'asc',
              hasHeader: true
            });
          }
        }
      ]
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
    if (!selectedCell) {
      setNotification('Please select a cell first');
      return;
    }
    setValidationDialog({
      show: true,
      cell: selectedCell,
      type: 'number',
      condition: 'between',
      min: '',
      max: '',
      list: '',
      customFormula: '',
      inputMessage: '',
      errorMessage: 'Invalid input',
    });
  }
}
,
    {
  icon: Table,
  label: "Pivot table",
  action: () => {
  if (!selectedCell) {
    setNotification("Please select a cell first");
    return;
  }

  const [rowStr, colStr] = selectedCell.split(',');
  const row = parseInt(rowStr);
  const col = parseInt(colStr);

  if (col < 2) {
    setNotification("Select a cell at least in column C");
    return;
  }

  const categoryCol = col - 2;
  const valueCol = col - 1;
  const pivotData = {};

  for (let r = 0; r < row; r++) {
    const catKey = `${r},${categoryCol}`;
    const valKey = `${r},${valueCol}`;

    const category = cells[catKey]?.value?.toString().trim();
    const rawValue = cells[valKey]?.value;
    const value = parseFloat(rawValue);

    if (category && !isNaN(value)) {
      if (!pivotData[category]) pivotData[category] = 0;
      pivotData[category] += value;
    }
  }

  const newCells = { ...cells };
  newCells[`${row},${col}`] = { value: "Category", display: "Category" };
  newCells[`${row},${col + 1}`] = { value: "Sum", display: "Sum" };

  let offset = 1;
  for (const [cat, sum] of Object.entries(pivotData)) {
    newCells[`${row + offset},${col}`] = { value: cat, display: cat };
    newCells[`${row + offset},${col + 1}`] = {
      value: sum.toString(),
      display: sum.toString(),
    };
    offset++;
  }

  setCells(newCells);
  setNotification("Pivot table created");
}

}
  ],
  Tools: [
    {
      icon: Settings,
      label: 'Spell check',
      action: startSpellCheck,
      shortcut: 'F7'
    },
    {
      icon: Mail,
      label: 'Notifications',
      action: () => setShowNotifications(!showNotifications),
      render: (
        <div className="relative">
          <Mail size={16} />
          {notifications.some(n => !n.read) && (
            <span className="absolute -top-1 -right-1 w-2 h-2 bg-red-500 rounded-full"></span>
          )}
        </div>
      )
    },
    {
      icon: Calculator,
      label: 'Script Editor',
      action: () => setScriptEditor(prev => ({ ...prev, isOpen: true })),
      shortcut: 'Alt+Shift+X'
    }
  ],
  Extensions: [
    { 
      icon: Puzzle, 
      label: 'Add-ons', 
      action: () => setShowAddonsMenu(true)
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
        <div
          key={menuName}
          className="relative"
          onMouseEnter={() => handleMenuEnter(menuName)}
          onMouseLeave={handleMenuLeave}
        >
          <button
            className={`px-3 py-2 rounded hover:bg-gray-200 ${
              activeMenu === menuName ? 'bg-gray-200' : ''
            }`}
          >
            {menuName}
          </button>

          {activeMenu === menuName && (
            <div
              ref={menuRef}
              className="absolute left-0 mt-1 bg-white border border-gray-200 rounded-md shadow-lg z-50 min-w-48 py-1"
              onMouseEnter={() => handleMenuEnter(menuName)}
              onMouseLeave={handleMenuLeave}
            >
              {menuItems[menuName].map((item, index) => (
                <div
                  key={index}
                  className="relative"
                  onMouseEnter={() => handleItemEnter(item)}
                  onMouseLeave={() => {
                    timeoutRef.current = setTimeout(() => {
                      setHoveredItem(null);
                    }, 200);
                  }}
                >
                  <button
                    className={`w-full flex items-center gap-3 px-4 py-2 text-sm text-left ${
                      item.disabled
                        ? 'text-gray-400 cursor-not-allowed'
                        : 'text-gray-700 hover:bg-gray-100'
                    } ${item.active ? 'bg-blue-50 text-blue-700' : ''}`}
                    onClick={() => {
                      if (!item.disabled && !item.submenu) {
                        item.action();
                        setActiveMenu(null);
                      }
                    }}
                    disabled={item.disabled}
                  >
                    {item.icon && (
                      <item.icon
                        size={16}
                        className={`flex-shrink-0 ${
                          item.active ? 'text-blue-600' : 'text-gray-500'
                        }`}
                      />
                    )}
                    <span className="flex-grow">{item.label}</span>
                    {item.shortcut && (
                      <span className="text-xs text-gray-400 ml-2">
                        {item.shortcut}
                      </span>
                    )}
                    {item.submenu && (
                      <ChevronRight size={16} className="text-gray-400" />
                    )}
                  </button>

                  {item.submenu && hoveredItem === item && (
                    <div
                      className="absolute left-full top-0 ml-0.5 bg-white border border-gray-200 rounded-md shadow-lg z-50 min-w-48 py-1"
                      onMouseEnter={() => handleItemEnter(item)}
                      onMouseLeave={() => {
                        timeoutRef.current = setTimeout(() => {
                          setHoveredItem(null);
                        }, 200);
                      }}
                    >
                      {item.submenu.map((subItem, subIndex) => (
                        <button
                          key={subIndex}
                          className={`w-full flex items-center gap-3 px-4 py-2 text-sm text-left ${
                            subItem.disabled
                              ? 'text-gray-400 cursor-not-allowed'
                              : 'text-gray-700 hover:bg-gray-100'
                          }`}
                          onClick={() => {
                            if (!subItem.disabled) {
                              subItem.action();
                              setActiveMenu(null);
                              setHoveredItem(null);
                            }
                          }}
                          disabled={subItem.disabled}
                        >
                          {subItem.icon && (
                            <subItem.icon
                              size={16}
                              className="flex-shrink-0 text-gray-500"
                            />
                          )}
                          <span className="flex-grow">{subItem.label}</span>
                          {subItem.shortcut && (
                            <span className="text-xs text-gray-400 ml-2">
                              {subItem.shortcut}
                            </span>
                          )}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
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
            onClick={handleUndo} 
            disabled={historyIndex === 0}
            className={`p-1.5 rounded-md ${historyIndex === 0 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-gray-100'}`}
            title="Undo (Ctrl+Z)"
          >
            <Undo size={18} className="text-gray-700" />
          </button>
          
                        
          <button 
            onClick={handleRedo} 
            disabled={historyIndex === history.length - 1}
            className={`p-1.5 rounded-md ${historyIndex === history.length - 1 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-gray-100'}`}
            title="Redo (Ctrl+Y)"
          >
            <Redo size={18} className="text-gray-700" />
          </button>
        </div>

        {/* Text Formatting */}
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          <button onClick={toggleBold} className="p-1 hover:bg-gray-200 rounded">
  <Bold size={18} />
</button>
<button onClick={toggleItalic} className="p-1 hover:bg-gray-200 rounded">
  <Italic size={18} />
</button>
<button onClick={toggleUnderline} className="p-1 hover:bg-gray-200 rounded">
  <Underline size={18} />
</button>

        </div>

        {/* Alignment */}
        <div className="flex items-center gap-1 border-r border-gray-200 pr-2 mr-2">
          
              <button onClick={alignLeft} className="p-1 hover:bg-gray-200 rounded">
  <AlignLeft size={18} />
</button>
<button onClick={alignCenter} className="p-1 hover:bg-gray-200 rounded">
  <AlignCenter size={18} />
</button>
<button onClick={alignRight} className="p-1 hover:bg-gray-200 rounded">
  <AlignRight size={18} />
</button>

            
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
         <button onClick={() => handleSortSelectedColumn('asc')} className="px-2 py-1 bg-blue-500 text-white rounded mr-2">Sort A → Z</button>
<button onClick={() => handleSortSelectedColumn('desc')} className="px-2 py-1 bg-blue-500 text-white rounded">Sort Z → A</button>









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
          {/* Category list */}
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

          {/* Function list */}
          {selectedCategory && (
            <div
              className="w-64 overflow-y-auto max-h-[300px] bg-white"
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
                    if (!selectedCell) {
                      setShowDropdown(false);
                      setSelectedCategory(null);
                      return;
                    }

                    const formula = `=${func}()`;
                    const evaluated = evaluateFormula(formula, cells);

                    setCells((prev) => ({
                      ...prev,
                      [selectedCell]: {
                        ...(typeof prev[selectedCell] === 'object'
                          ? prev[selectedCell]
                          : { value: prev[selectedCell] || '' }),
                        formula,
                        display: evaluated,
                      },
                    }));

                    setFormulaBar(formula);
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
      {filterActive && (
  <tr>
    {columnKeys.map((col) => (
      <td key={col}>
        <input
          type="text"
          placeholder={`Filter ${col}`}
          value={filterValues[col] || ''}
          onChange={(e) =>
            setFilterValues((prev) => ({ ...prev, [col]: e.target.value }))
          }
          className="w-full px-1 text-xs border rounded"
        />
      </td>
    ))}
  </tr>
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
            {getColumnName(selectedCell.col)}{selectedCell.row + 1}
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
      <div className="flex-1 overflow-auto">
        <table className="table-fixed border-separate border-spacing-0 min-w-full">
          <thead>
            <tr>
              <th className="w-10 h-8 bg-gray-200 border border-gray-300"></th>
              {Array.from({ length: COLS }).map((_, c) => (
                <th
                  key={c}
                  className="h-8 w-24 text-center text-sm font-medium bg-gray-200 border border-gray-300"
                >
                  {getColLetter(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: ROWS }).map((_, r) => (
              <tr key={r}>
                <th className="w-10 text-center bg-gray-200 border border-gray-300 text-sm font-medium">
                  {r + 1}
                </th>
                {Array.from({ length: COLS }).map((_, c) => {
                  const key = `${r},${c}`;
                  return (
                    <td
  key={key}
  className={`w-24 h-8 p-0 ${showGridlines ? 'border border-gray-300' : ''}`}
  onClick={(e) => {
  const key = `${r},${c}`;
  if (e.shiftKey && selectedCell) {
    setSelectedRange([selectedCell, key]);
  } else {
    setSelectedCell(key);
    setSelectedRange(null);
  }
}}

>
                      <div
  contentEditable={!isRichCell(cells[key]?.display)}  
  suppressContentEditableWarning
  onInput={(e) => {
  const value = e.currentTarget.textContent;
  const rule = cellValidationRules[key];

  if (rule) {
    const isValid = validateValue(value, rule);
    if (!isValid) {
      e.currentTarget.textContent = '';
      setNotification(rule.errorMessage || 'Invalid input');
      return;
    }
  }

  if (!isRichCell(cells[key]?.display)) {
    handleChange(key, value);
  }
}}

  
  className={`w-full h-full px-2 text-sm outline-none cursor-pointer
    flex items-center justify-start overflow-hidden
    ${cellStyles[key]?.bold ? 'font-bold' : ''}
    ${cellStyles[key]?.italic ? 'italic' : ''}
    ${cellStyles[key]?.underline ? 'underline' : ''}
    ${cellStyles[key]?.align === 'center' ? 'justify-center' : ''}
    ${cellStyles[key]?.align === 'right' ? 'justify-end' : ''}
    ${isLinkCell(cells[key]?.display) ? 'text-blue-600 underline' : ''}
  `}
  style={{
    display: 'flex',
    alignItems: 'center', // ✅ vertical center
    justifyContent:
      cellStyles[key]?.align === 'center'
        ? 'center'
        : cellStyles[key]?.align === 'right'
        ? 'flex-end'
        : 'flex-start',
    color: cellStyles[key]?.color || 'inherit',
    backgroundColor: cellStyles[key]?.backgroundColor || 'transparent',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    height: '100%',
    lineHeight: 'normal',
    direction: 'ltr',
  }}
  onClick={(e) => {
    if (isLinkCell(cells[key]?.display)) {
      window.open(cells[key]?.display, '_blank');
      e.preventDefault();
    }
  }}
  ref={(el) => {
    if (
      el &&
      !isRichCell(cells[key]?.display) &&
      el.textContent !== cells[key]?.display
    ) {
      el.textContent = cells[key]?.display || '';
    }
  }}
  dangerouslySetInnerHTML={
    isRichCell(cells[key]?.display)
      ? { __html: cells[key]?.display }
      : undefined
  }
/>

                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
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
        onClick={() => setFindReplace(prev => ({ ...prev, show: false }))}
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
          onChange={(e) => setFindReplace(prev => ({ ...prev, find: e.target.value }))}
          className="w-full px-3 py-2 border border-gray-300 rounded-md"
          autoFocus
        />
      </div>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">Replace with</label>
        <input
          type="text"
          value={findReplace.replace}
          onChange={(e) => setFindReplace(prev => ({ ...prev, replace: e.target.value }))}
          className="w-full px-3 py-2 border border-gray-300 rounded-md"
        />
      </div>

      <label className="flex items-center">
        <input
          type="checkbox"
          checked={findReplace.matchCase}
          onChange={(e) => setFindReplace(prev => ({ ...prev, matchCase: e.target.checked }))}
          className="h-4 w-4 text-blue-600 rounded"
        />
        <span className="ml-2 text-sm text-gray-700">Match case</span>
      </label>

      <div className="flex justify-between pt-2">
        <div className="flex space-x-2">
          <button
            onClick={() => {
              const results = Object.keys(cells).filter(key => {
                const val = typeof cells[key] === 'object' ? cells[key]?.display : cells[key] || '';
                const search = findReplace.find;
                return findReplace.matchCase
                  ? val.includes(search)
                  : val.toLowerCase().includes(search.toLowerCase());
              });

              if (results.length > 0) {
                setSelectedCell(results[0]);
              }

              setFindReplace(prev => ({
                ...prev,
                results,
                currentIndex: 0
              }));
            }}
            className="px-4 py-2 bg-gray-100 hover:bg-gray-200 rounded-md text-sm"
          >
            Find
          </button>

          <button
            onClick={() => {
              if (findReplace.results.length === 0) return;

              const updated = { ...cells };
              findReplace.results.forEach(key => {
                let val = typeof updated[key] === 'object' ? updated[key]?.display : updated[key];
                const search = findReplace.find;
                const replacement = findReplace.replace;

                if (!findReplace.matchCase) {
                  const regex = new RegExp(search, 'gi');
                  val = val.replace(regex, replacement);
                } else {
                  val = val.replaceAll(search, replacement);
                }

                // Handle structured or plain cells
                if (typeof updated[key] === 'object') {
                  updated[key] = { ...updated[key], display: val };
                } else {
                  updated[key] = val;
                }
              });

              setCells(updated);
              setFindReplace(prev => ({
                ...prev,
                results: [],
                currentIndex: 0
              }));
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
{validationDialog.show && (
  <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
    <div className="bg-white rounded-lg shadow-xl p-6 w-96">
      <h3 className="text-lg font-medium mb-4">Data Validation for {validationDialog.cell}</h3>
      
      <div className="space-y-4">
        <div>
          <label className="block text-sm font-medium mb-1">Validation type</label>
          <select
            value={validationDialog.type}
            onChange={(e) => setValidationDialog({...validationDialog, type: e.target.value})}
            className="w-full border rounded p-2"
          >
            <option value="number">Number</option>
            <option value="text">Text</option>
            <option value="date">Date</option>
            <option value="list">List from range</option>
            <option value="custom">Custom formula</option>
          </select>
        </div>

        {validationDialog.type === 'number' && (
          <>
            <div>
              <label className="block text-sm font-medium mb-1">Condition</label>
              <select
                value={validationDialog.condition}
                onChange={(e) => setValidationDialog({...validationDialog, condition: e.target.value})}
                className="w-full border rounded p-2"
              >
                <option value="between">Between</option>
                <option value="greater">Greater than</option>
                <option value="less">Less than</option>
                <option value="equal">Equal to</option>
                <option value="notEqual">Not equal to</option>
              </select>
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium mb-1">
                  {validationDialog.condition === 'between' ? 'Minimum' : 'Value'}
                </label>
                <input
                  type="number"
                  value={validationDialog.min}
                  onChange={(e) => setValidationDialog({...validationDialog, min: e.target.value})}
                  className="w-full border rounded p-2"
                />
              </div>
              
              {validationDialog.condition === 'between' && (
                <div>
                  <label className="block text-sm font-medium mb-1">Maximum</label>
                  <input
                    type="number"
                    value={validationDialog.max}
                    onChange={(e) => setValidationDialog({...validationDialog, max: e.target.value})}
                    className="w-full border rounded p-2"
                  />
                </div>
              )}
            </div>
          </>
        )}

        {validationDialog.type === 'list' && (
          <div>
            <label className="block text-sm font-medium mb-1">List items (comma separated)</label>
            <input
              type="text"
              value={validationDialog.list}
              onChange={(e) => setValidationDialog({...validationDialog, list: e.target.value})}
              className="w-full border rounded p-2"
              placeholder="e.g., Yes,No,Maybe"
            />
          </div>
        )}

        {validationDialog.type === 'custom' && (
          <div>
            <label className="block text-sm font-medium mb-1">Custom formula</label>
            <input
              type="text"
              value={validationDialog.customFormula}
              onChange={(e) => setValidationDialog({...validationDialog, customFormula: e.target.value})}
              className="w-full border rounded p-2"
              placeholder="e.g., =A1>B1"
            />
          </div>
        )}

        <div>
          <label className="block text-sm font-medium mb-1">Input message (optional)</label>
          <input
            type="text"
            value={validationDialog.inputMessage}
            onChange={(e) => setValidationDialog({...validationDialog, inputMessage: e.target.value})}
            className="w-full border rounded p-2"
            placeholder="e.g., Enter a number between 1 and 100"
          />
        </div>

        <div>
          <label className="block text-sm font-medium mb-1">Error message</label>
          <input
            type="text"
            value={validationDialog.errorMessage}
            onChange={(e) => setValidationDialog({...validationDialog, errorMessage: e.target.value})}
            className="w-full border rounded p-2"
          />
        </div>

        <div className="flex justify-end space-x-2 pt-4">
          <button
            onClick={() => setValidationDialog({...validationDialog, show: false})}
            className="px-4 py-2 border rounded"
          >
            Cancel
          </button>
          <button
            onClick={applyValidation}
            className="px-4 py-2 bg-blue-600 text-white rounded"
          >
            Save
          </button>
        </div>
      </div>
    </div>
  </div>
)}
{scriptEditor.isOpen && (
  <div className="fixed inset-0 z-50 bg-black bg-opacity-50 flex items-center justify-center">
    <div className="bg-white p-4 rounded shadow-lg w-[600px]">
      <h2 className="text-lg font-bold mb-2">Script Editor</h2>
      <textarea
        className="w-full h-60 p-2 border rounded font-mono text-sm"
        value={
          scriptEditor.scripts.find((s) => s.id === scriptEditor.activeScriptId)
            ?.code || ''
        }
        onChange={(e) => {
          setScriptEditor((prev) => ({
            ...prev,
            scripts: prev.scripts.map((script) =>
              script.id === prev.activeScriptId
                ? { ...script, code: e.target.value, lastEdited: new Date() }
                : script
            ),
          }));
        }}
      />
      <div className="flex justify-end gap-2 mt-3">
        <button
          className="px-3 py-1 bg-green-600 text-white rounded"
          onClick={() =>
            runScript(
              scriptEditor.scripts.find(
                (s) => s.id === scriptEditor.activeScriptId
              )?.code
            )
          }
        >
          Run
        </button>
        <button
          className="px-3 py-1 bg-gray-300 text-black rounded"
          onClick={() =>
            setScriptEditor((prev) => ({ ...prev, isOpen: false }))
          }
        >
          Close
        </button>
      </div>
    </div>
  </div>
)}


{showLinkDialog && (
  <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
    <div className="bg-white rounded-lg shadow-2xl w-full max-w-md mx-4">
      <div className="flex items-center justify-between p-6 border-b border-gray-200">
        <h2 className="text-xl font-medium text-gray-900">Insert link</h2>
        <button
          onClick={() => setShowLinkDialog(false)}
          className="text-gray-400 hover:text-gray-600 p-1"
        >
          <X size={24} />
        </button>
      </div>

      <div className="p-6 space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Text
          </label>
          <input
            type="text"
            value={linkText}
            onChange={(e) => setLinkText(e.target.value)}
            placeholder="Enter text to display"
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Link
          </label>
          <input
            type="url"
            value={linkUrl}
            onChange={(e) => setLinkUrl(e.target.value)}
            placeholder="Paste or type a link"
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
          />
        </div>
      </div>

      <div className="flex items-center justify-end gap-3 p-6 border-t border-gray-200 bg-gray-50">
        <button
          onClick={() => setShowLinkDialog(false)}
          className="px-4 py-2 text-sm font-medium text-gray-700 border border-gray-300 rounded-md hover:bg-gray-50"
        >
          Cancel
        </button>
        <button
          onClick={() => {
            insertLink();
            setShowLinkDialog(false);
          }}
          disabled={!linkText.trim() || !linkUrl.trim()}
          className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
        >
          Apply
        </button>
      </div>
    </div>
  </div>
)}
{showImageDialog && (
  <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 overflow-y-auto">
    <div className="bg-white rounded-lg shadow-xl w-full max-w-2xl ">
      <div className="flex justify-between items-center p-4 border-b ">
        <h3 className="text-lg font-medium">Insert Image</h3>
        <button onClick={() => setShowImageDialog(false)} className="text-gray-500">
          <X size={24} />
        </button>
      </div>

      <div className="p-4">
        <div className="flex border-b ">
          <button
            className={`px-4 py-2 ${activeImageTab === 'url' ? 'border-b-2 border-blue-500' : ''}`}
            onClick={() => setActiveImageTab('url')}
          >
            From URL
          </button>
          <button
            className={`px-4 py-2 ${activeImageTab === 'upload' ? 'border-b-2 border-blue-500' : ''}`}
            onClick={() => setActiveImageTab('upload')}
          >
            Upload
          </button>
        </div>

        <div className="mt-4">
          {activeImageTab === 'url' ? (
            <div className="space-y-4">
              <input
                type="url"
                value={imageUrl}
                onChange={handleUrlChange}
                placeholder="Enter image URL"
                className="w-full p-2 border rounded"
              />
              {previewImage && (
                <div className="mt-2">
                  <p className="text-sm font-medium mb-1">Preview:</p>
                  <img 
                    src={previewImage} 
                    alt="Preview" 
                    className="max-w-full max-h-48 object-contain border"
                  />
                </div>
              )}
            </div>
          ) : (
            <div className="space-y-4">
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept="image/*"
                className="hidden"
              />
              <button
                onClick={() => fileInputRef.current.click()}
                className="w-full p-8 border-2 border-dashed rounded-lg flex flex-col items-center justify-center hover:bg-gray-50"
              >
                <Upload size={32} className="text-gray-400 mb-2" />
                <p>Click to upload or drag and drop</p>
                <p className="text-sm text-gray-500 mt-1">PNG, JPG, GIF up to 5MB</p>
              </button>
              {previewImage && (
                <div className="mt-2">
                  <p className="text-sm font-medium mb-1">Preview:</p>
                  <img 
                    src={previewImage} 
                    alt="Preview" 
                    className="max-w-full max-h-48 object-contain border"
                  />
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <div className="flex justify-end p-4 border-t">
        <button
          onClick={() => setShowImageDialog(false)}
          className="px-4 py-2 mr-2 border rounded"
        >
          Cancel
        </button>
        <button
          onClick={insertImage}
          disabled={!previewImage}
          className="px-4 py-2 bg-blue-600 text-white rounded disabled:bg-gray-300"
        >
          Insert Image
        </button>
      </div>
    </div>
  </div>
)}
{spellCheck.active && spellCheck.errors.length > 0 && (
  <div className="fixed bottom-4 right-4 bg-white border border-gray-300 rounded-lg shadow-xl z-50 w-96">
    <div className="p-4 border-b flex justify-between items-center bg-gray-50">
      <h3 className="font-medium text-gray-800">
        Spell Check ({spellCheck.currentErrorIndex + 1}/{spellCheck.errors.length})
      </h3>
      <button 
        onClick={() => setSpellCheck(prev => ({...prev, active: false}))}
        className="text-gray-500 hover:text-gray-700 p-1 rounded-full hover:bg-gray-200"
      >
        <X size={18} />
      </button>
    </div>
    
    <div className="p-4">
      <div className="mb-4">
        <p className="text-sm text-gray-600 mb-1">Misspelled word in cell {spellCheck.errors[spellCheck.currentErrorIndex].cellRef}:</p>
        <div className="font-medium bg-red-50 px-3 py-2 rounded border border-red-100 text-red-800">
          {spellCheck.errors[spellCheck.currentErrorIndex].word}
        </div>
      </div>
      
      <div className="mb-4">
        <p className="text-sm text-gray-600 mb-1">Suggestions:</p>
        <div className="grid grid-cols-2 gap-2">
          {spellCheck.errors[spellCheck.currentErrorIndex].suggestions.map((suggestion, i) => (
            <button
              key={i}
              onClick={() => handleSpellCheckAction('replace', suggestion)}
              className="px-3 py-1.5 bg-blue-50 text-blue-700 rounded hover:bg-blue-100 text-sm text-left truncate"
            >
              {suggestion}
            </button>
          ))}
        </div>
      </div>
      
      <div className="flex justify-between items-center pt-2 border-t">
        <div className="flex gap-2">
          <button
            onClick={() => handleSpellCheckAction('ignore')}
            className="px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 text-gray-700"
          >
            Ignore
          </button>
          <button
            onClick={() => handleSpellCheckAction('add')}
            className="px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 text-gray-700"
          >
            Add to Dictionary
          </button>
        </div>
        
        <div className="flex gap-2">
          <button
            onClick={moveToPrevError}
            disabled={spellCheck.currentErrorIndex === 0}
            className="px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 text-gray-700 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            Previous
          </button>
          <button
            onClick={moveToNextError}
            disabled={spellCheck.currentErrorIndex === spellCheck.errors.length - 1}
            className="px-3 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {spellCheck.currentErrorIndex === spellCheck.errors.length - 1 ? 'Finish' : 'Next'}
          </button>
        </div>
      </div>
    </div>
  </div>
)}
{showNotifications && (
  <div className="fixed top-16 right-4 bg-white border border-gray-200 rounded-lg shadow-xl z-50 w-80">
    <div className="p-3 border-b flex justify-between items-center">
      <h3 className="font-medium">Notifications</h3>
      <button 
        onClick={() => {
          setNotifications(prev => prev.map(n => ({...n, read: true})));
          setShowNotifications(false);
        }}
        className="text-sm text-blue-600 hover:text-blue-800"
      >
        Mark all as read
      </button>
    </div>
    
    <div className="max-h-96 overflow-y-auto">
      {notifications.length === 0 ? (
        <div className="p-4 text-center text-gray-500">No notifications</div>
      ) : (
        notifications.map(notification => (
          <div 
            key={notification.id}
            className={`p-3 border-b hover:bg-gray-50 cursor-pointer ${
              !notification.read ? 'bg-blue-50' : ''
            }`}
            onClick={() => {
              setNotifications(prev => 
                prev.map(n => 
                  n.id === notification.id ? {...n, read: true} : n
                )
              );
              // Handle notification click action
            }}
          >
            <div className="flex justify-between">
              <span className={!notification.read ? 'font-medium' : ''}>
                {notification.message}
              </span>
              {!notification.read && (
                <span className="w-2 h-2 bg-blue-500 rounded-full"></span>
              )}
            </div>
            <div className="text-xs text-gray-500 mt-1">
              {formatDistanceToNow(notification.timestamp, { addSuffix: true })}
            </div>
          </div>
        ))
      )}
    </div>
    
    <div className="p-2 border-t text-center">
      <button 
        onClick={() => setShowNotifications(false)}
        className="text-sm text-blue-600 hover:text-blue-800"
      >
        Close
      </button>
    </div>
  </div>
)}
{sortDialog.show && (
  <div className="fixed top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 bg-white p-5 rounded shadow-lg w-96 z-50">
    <h3 className="text-lg font-medium mb-4">Custom Sort</h3>

    <div className="space-y-3">
      <label className="block">
        Sort by column:
        <select
          className="w-full mt-1 border rounded px-2 py-1"
          value={sortDialog.sortBy}
          onChange={(e) => setSortDialog(prev => ({ ...prev, sortBy: e.target.value }))}
        >
          {Array.from({ length: COLS }).map((_, idx) => (
            <option key={idx} value={String.fromCharCode(65 + idx)}>
              {String.fromCharCode(65 + idx)}
            </option>
          ))}
        </select>
      </label>

      <label className="block">
        Order:
        <select
          className="w-full mt-1 border rounded px-2 py-1"
          value={sortDialog.order}
          onChange={(e) => setSortDialog(prev => ({ ...prev, order: e.target.value }))}
        >
          <option value="asc">Ascending</option>
          <option value="desc">Descending</option>
        </select>
      </label>

      <label className="flex items-center">
        <input
          type="checkbox"
          checked={sortDialog.hasHeader}
          onChange={(e) => setSortDialog(prev => ({ ...prev, hasHeader: e.target.checked }))}
          className="mr-2"
        />
        My data has headers
      </label>
    </div>

    <div className="flex justify-end mt-4 space-x-2">
      <button
        onClick={() => setSortDialog(prev => ({ ...prev, show: false }))}
        className="px-4 py-2 bg-gray-200 rounded"
      >
        Cancel
      </button>
      <button
        onClick={handleCustomSort}
        className="px-4 py-2 bg-blue-500 text-white rounded"
      >
        Sort
      </button>
    </div>
  </div>
)}

{validationDialog.show && (
  <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
    <div className="bg-white rounded-lg shadow-xl w-full max-w-md">
      <div className="p-4 border-b flex justify-between items-center">
        <h3 className="text-lg font-medium">Data Validation</h3>
        <button 
          onClick={() => setValidationDialog({...validationDialog, show: false})}
          className="text-gray-500"
        >
          <X size={20} />
        </button>
      </div>

      <div className="p-4 space-y-4">
        <div>
          <label className="block text-sm font-medium mb-1">Validation type</label>
          <select
            value={validationDialog.type}
            onChange={(e) => setValidationDialog({...validationDialog, type: e.target.value})}
            className="w-full border rounded p-2"
          >
            <option value="number">Number</option>
            <option value="text">Text</option>
            <option value="date">Date</option>
            <option value="list">List from range</option>
            <option value="custom">Custom formula</option>
          </select>
        </div>

        {validationDialog.type === 'number' && (
          <>
            <div>
              <label className="block text-sm font-medium mb-1">Condition</label>
              <select
                value={validationDialog.condition}
                onChange={(e) => setValidationDialog({...validationDialog, condition: e.target.value})}
                className="w-full border rounded p-2"
              >
                <option value="between">Between</option>
                <option value="greater">Greater than</option>
                <option value="less">Less than</option>
                <option value="equal">Equal to</option>
                <option value="notEqual">Not equal to</option>
              </select>
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium mb-1">
                  {validationDialog.condition === 'between' ? 'Minimum' : 'Value'}
                </label>
                <input
                  type="number"
                  value={validationDialog.min}
                  onChange={(e) => setValidationDialog({...validationDialog, min: e.target.value})}
                  className="w-full border rounded p-2"
                />
              </div>
              
              {validationDialog.condition === 'between' && (
                <div>
                  <label className="block text-sm font-medium mb-1">Maximum</label>
                  <input
                    type="number"
                    value={validationDialog.max}
                    onChange={(e) => setValidationDialog({...validationDialog, max: e.target.value})}
                    className="w-full border rounded p-2"
                  />
                </div>
              )}
            </div>
          </>
        )}

        {validationDialog.type === 'list' && (
          <div>
            <label className="block text-sm font-medium mb-1">List items (comma separated)</label>
            <input
              type="text"
              value={validationDialog.list}
              onChange={(e) => setValidationDialog({...validationDialog, list: e.target.value})}
              className="w-full border rounded p-2"
              placeholder="e.g., Yes,No,Maybe"
            />
          </div>
        )}

        {validationDialog.type === 'custom' && (
          <div>
            <label className="block text-sm font-medium mb-1">Custom formula</label>
            <input
              type="text"
              value={validationDialog.customFormula}
              onChange={(e) => setValidationDialog({...validationDialog, customFormula: e.target.value})}
              className="w-full border rounded p-2"
              placeholder="e.g., =A1>B1"
            />
          </div>
        )}

        <div>
          <label className="block text-sm font-medium mb-1">Input message (optional)</label>
          <input
            type="text"
            value={validationDialog.inputMessage}
            onChange={(e) => setValidationDialog({...validationDialog, inputMessage: e.target.value})}
            className="w-full border rounded p-2"
            placeholder="e.g., Enter a number between 1 and 100"
          />
        </div>

        <div>
          <label className="block text-sm font-medium mb-1">Error message</label>
          <input
            type="text"
            value={validationDialog.errorMessage}
            onChange={(e) => setValidationDialog({...validationDialog, errorMessage: e.target.value})}
            className="w-full border rounded p-2"
          />
        </div>

        <div>
          <label className="block text-sm font-medium mb-1">Error style</label>
          <select
            value={validationDialog.errorStyle}
            onChange={(e) => setValidationDialog({...validationDialog, errorStyle: e.target.value})}
            className="w-full border rounded p-2"
          >
            <option value="stop">Stop (prevent invalid data)</option>
            <option value="warning">Warning (warn but allow)</option>
            <option value="info">Information (show message)</option>
          </select>
        </div>
      </div>

      <div className="p-4 border-t flex justify-end space-x-2">
        <button
          onClick={() => setValidationDialog({...validationDialog, show: false})}
          className="px-4 py-2 border rounded"
        >
          Cancel
        </button>
        <button
          onClick={applyValidation}
          className="px-4 py-2 bg-blue-600 text-white rounded"
        >
          Save
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
    <SpellCheckModal />
  </div>
);
};

