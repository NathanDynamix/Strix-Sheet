import React, {
  useState,
  useRef,
  useCallback,
  useMemo,
  useEffect,
} from "react";
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useSpreadsheet } from '../context/SpreadsheetContext';
import { useToast } from '../context/ToastContext';
import {
  AlignLeft,
  AlignCenter,
  AlignRight,
  Palette,
  Calculator,
  ChevronDown,
  Undo,
  Redo,
  Filter,
  BarChart,
  MessageSquare,
  Plus,
  X,
  Trash2,
  DollarSign,
  FilterX,
  FileSpreadsheet,
} from "lucide-react";
import {
  LineChart as RechartsLineChart,
  BarChart as RechartsBarChart,
  PieChart as RechartsPieChart,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  Line,
  Bar,
  Cell,
  AreaChart ,
} from "recharts";

const GoogleSheetsClone = () => {
  const { currentUser, logout } = useAuth();
  const { 
    currentSpreadsheet, 
    updateCell, 
    autoSave, 
    getCurrentSheetData,
    getCellData 
  } = useSpreadsheet();
  const { showSuccess, showError } = useToast();
  const navigate = useNavigate();

  // Initialize spreadsheet data for 40,000 cells (40 columns x 1000 rows)
  const initializeData = () => {
    const data = {};
    for (let row = 1; row <= 1000; row++) {
      for (let col = 1; col <= 40; col++) {
        const cellId = getColumnName(col) + row;
        data[cellId] = { value: "", formula: "", style: {} };
      }
    }
    return data;
  };

  // Convert column number to letter(s)
  const getColumnName = (col) => {
    let result = "";
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  };

  // Get data from context or initialize
  const getSheetData = () => {
    if (currentSpreadsheet && currentSpreadsheet.sheets) {
      const activeSheet = currentSpreadsheet.sheets.find(
        sheet => sheet.id === currentSpreadsheet.activeSheetId
      );
      if (activeSheet && activeSheet.data) {
        return Object.fromEntries(activeSheet.data);
      }
    }
    return initializeData();
  };

  const [sheets, setSheets] = useState([
    { id: "sheet1", name: "Sheet1", data: getSheetData() },
  ]);
  const [activeSheetId, setActiveSheetId] = useState("sheet1");
  const [selectedCell, setSelectedCell] = useState("A1");
  const [formulaBarValue, setFormulaBarValue] = useState("");
  const [cellEditValue, setCellEditValue] = useState("");
  const [isEditing, setIsEditing] = useState(false);
  const [isProcessingInput, setIsProcessingInput] = useState(false);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [showFunctionMenu, setShowFunctionMenu] = useState(false);
  const [showChartModal, setShowChartModal] = useState(false);
  const [chartData, setChartData] = useState(null);
  const [chartType, setChartType] = useState("line");
  const [showFilterModal, setShowFilterModal] = useState(false);
  const [filterColumn, setFilterColumn] = useState("A");
  const [filterValue, setFilterValue] = useState("");
  const [activeFilters, setActiveFilters] = useState({});
  const [filterMenuOpen, setFilterMenuOpen] = useState(false);
  
  // Zoom functionality
  const [zoomLevel, setZoomLevel] = useState(100);
  const [showZoomDropdown, setShowZoomDropdown] = useState(false);
  
  // Font functionality
  const [showFontDropdown, setShowFontDropdown] = useState(false);
  const [showFontSizeDropdown, setShowFontSizeDropdown] = useState(false);
  
  // Number formatting functionality
  const [showNumberFormatDropdown, setShowNumberFormatDropdown] = useState(false);
  const [formatUpdateTrigger, setFormatUpdateTrigger] = useState(0);
  
  // Google Sheets-like filter states
  const [columnFilters, setColumnFilters] = useState({});
  const [filterDropdowns, setFilterDropdowns] = useState({});
  const [showFilterDropdown, setShowFilterDropdown] = useState(null);
  const [filterOptions, setFilterOptions] = useState({});
  
  // Cell resizing states
  const [columnWidths, setColumnWidths] = useState({});
  const [rowHeights, setRowHeights] = useState({});
  const [isResizing, setIsResizing] = useState(false);
  const [resizeType, setResizeType] = useState(null); // 'column' or 'row'
  const [resizeTarget, setResizeTarget] = useState(null);
  const [resizeStartPos, setResizeStartPos] = useState({ x: 0, y: 0 });
  const [resizeStartSize, setResizeStartSize] = useState({ width: 0, height: 0 });
  const [showResizeIndicator, setShowResizeIndicator] = useState(false);
  const [visibleRows, setVisibleRows] = useState(100); // Virtualization
  const [scrollTop, setScrollTop] = useState(0);
  const [showFormulaHelper, setShowFormulaHelper] = useState(false);
  const [formulaSearch, setFormulaSearch] = useState("");
  const [showFormulaPrompt, setShowFormulaPrompt] = useState(false);
  const [selectedFunction, setSelectedFunction] = useState(null);

  // Menu states
  const [activeMenu, setActiveMenu] = useState(null);
  const [showFileMenu, setShowFileMenu] = useState(false);
  const [showEditMenu, setShowEditMenu] = useState(false);
  const [showViewMenu, setShowViewMenu] = useState(false);
  const [showInsertMenu, setShowInsertMenu] = useState(false);
  const [showFormatMenu, setShowFormatMenu] = useState(false);
  const [showDataMenu, setShowDataMenu] = useState(false);
  
  // Clipboard functionality
  const [clipboard, setClipboard] = useState(null);
  const [clipboardType, setClipboardType] = useState(null); // 'copy' or 'cut'
  
  // Undo/Redo functionality
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  const cellInputRef = useRef(null);
  const formulaBarInputRef = useRef(null);
  const isNavigatingRef = useRef(false);
  const csvFileInputRef = useRef(null);
  const activeSheet = sheets.find((sheet) => sheet.id === activeSheetId);
  
  // Use useMemo to ensure data updates when sheets change
  const data = useMemo(() => {
    return activeSheet ? (activeSheet.data instanceof Map ? Object.fromEntries(activeSheet.data) : activeSheet.data) : {};
  }, [activeSheet, formatUpdateTrigger]);
  
  // Debug logging (minimal)
  // console.log('Active sheet:', activeSheet);

  // Update sheets when currentSpreadsheet changes
  useEffect(() => {
    if (currentSpreadsheet && currentSpreadsheet.sheets) {
      setSheets(currentSpreadsheet.sheets);
      setActiveSheetId(currentSpreadsheet.activeSheetId);
    }
  }, [currentSpreadsheet]);

  // Performance optimizations - only render visible rows
  const memoizedCellData = useMemo(() => {
    const result = {};
    const startRow = Math.floor(scrollTop / 24) + 1;
    const endRow = Math.min(startRow + visibleRows, 1000);

    for (let row = startRow; row <= endRow; row++) {
      for (let col = 1; col <= 40; col++) {
        const cellId = getColumnName(col) + row;
        const cellData = data[cellId];
        
        // Debug logging for specific cells
        if (cellId === 'A1') {
          console.log(`MemoizedCellData for ${cellId}:`, cellData);
          console.log(`Data[${cellId}]:`, data[cellId]);
        }
        
        result[cellId] = cellData || { value: "", formula: "", style: {} };
      }
    }
    return result;
  }, [data, visibleRows, scrollTop]);

  // Enhanced filter functions for Google Sheets-like behavior
  const applyColumnFilter = useCallback((column, filterType, filterValue, filterValues = []) => {
    if (!filterValue && filterValues.length === 0) {
      setColumnFilters((prev) => {
        const newFilters = { ...prev };
        delete newFilters[column];
        return newFilters;
      });
      return;
    }

    setColumnFilters((prev) => ({
      ...prev,
      [column]: {
        type: filterType,
        value: filterValue,
        values: filterValues,
        enabled: true
      }
    }));
  }, []);

  const isRowVisible = useCallback(
    (rowIndex) => {
      if (Object.keys(columnFilters).length === 0) return true;

      for (const [column, filter] of Object.entries(columnFilters)) {
        if (!filter.enabled) continue;

        const colNum = getColumnNumber(column);
        const cellId = getColumnName(colNum) + (rowIndex + 1);
        const cellData = data[cellId];
        const cellValue = cellData ? (cellData.value || "").toString() : "";

        let isVisible = false;

        switch (filter.type) {
          case 'contains':
            isVisible = cellValue.toLowerCase().includes((filter.value || "").toLowerCase());
            break;
          case 'equals':
            isVisible = cellValue.toLowerCase() === (filter.value || "").toLowerCase();
            break;
          case 'starts_with':
            isVisible = cellValue.toLowerCase().startsWith((filter.value || "").toLowerCase());
            break;
          case 'ends_with':
            isVisible = cellValue.toLowerCase().endsWith((filter.value || "").toLowerCase());
            break;
          case 'greater_than':
            isVisible = parseFloat(cellValue) > parseFloat(filter.value || 0);
            break;
          case 'less_than':
            isVisible = parseFloat(cellValue) < parseFloat(filter.value || 0);
            break;
          case 'greater_equal':
            isVisible = parseFloat(cellValue) >= parseFloat(filter.value || 0);
            break;
          case 'less_equal':
            isVisible = parseFloat(cellValue) <= parseFloat(filter.value || 0);
            break;
          case 'is_empty':
            isVisible = cellValue === "";
            break;
          case 'is_not_empty':
            isVisible = cellValue !== "";
            break;
          case 'is_one_of':
            isVisible = filter.values.some(val => cellValue.toLowerCase() === val.toLowerCase());
            break;
          default:
            isVisible = true;
        }

        if (!isVisible) return false;
      }
      return true;
    },
    [columnFilters, data]
  );

  // Legacy filter function for backward compatibility
  const applyFilter = useCallback((column, filterValue) => {
    applyColumnFilter(column, 'contains', filterValue);
  }, [applyColumnFilter]);

  const clearAllFilters = () => {
    setColumnFilters({});
    setShowFilterDropdown(null);
    showSuccess('All filters cleared');
  };

  // Zoom functions
  const zoomLevels = [50, 75, 90, 100, 125, 150, 175, 200, 250, 300];
  
  const handleZoomIn = () => {
    const currentIndex = zoomLevels.indexOf(zoomLevel);
    if (currentIndex < zoomLevels.length - 1) {
      setZoomLevel(zoomLevels[currentIndex + 1]);
      showSuccess(`Zoomed to ${zoomLevels[currentIndex + 1]}%`);
    }
  };

  const handleZoomOut = () => {
    const currentIndex = zoomLevels.indexOf(zoomLevel);
    if (currentIndex > 0) {
      setZoomLevel(zoomLevels[currentIndex - 1]);
      showSuccess(`Zoomed to ${zoomLevels[currentIndex - 1]}%`);
    }
  };

  const handleZoomToFit = () => {
    setZoomLevel(100);
    showSuccess('Zoomed to fit');
  };

  const handleZoomChange = (newZoom) => {
    setZoomLevel(newZoom);
    showSuccess(`Zoomed to ${newZoom}%`);
    setShowZoomDropdown(false);
  };

  // Font functions
  const fontFamilies = [
    'Arial', 'Arial Black', 'Calibri', 'Cambria', 'Candara', 'Comic Sans MS',
    'Consolas', 'Courier New', 'Georgia', 'Helvetica', 'Impact', 'Lucida Console',
    'Lucida Sans Unicode', 'Microsoft Sans Serif', 'Palatino Linotype',
    'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'
  ];

  const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

  const getCurrentFontFamily = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.fontFamily || 'Arial';
    }
    return 'Arial';
  };

  const getCurrentFontSize = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.fontSize || 10;
    }
    return 10;
  };

  const handleFontFamilyChange = (fontFamily) => {
    if (selectedCell) {
      formatCell({ fontFamily });
      showSuccess(`Font changed to ${fontFamily}`);
      setShowFontDropdown(false);
    }
  };

  const handleFontSizeChange = (fontSize) => {
    if (selectedCell) {
      formatCell({ fontSize });
      showSuccess(`Font size changed to ${fontSize}px`);
      setShowFontSizeDropdown(false);
    }
  };

  // Number formatting functions
  const numberFormats = [
    { label: 'Automatic', value: 'automatic', example: '1,000.12' },
    { label: 'Plain text', value: 'text', example: 'Plain text' },
    { label: 'Number', value: 'number', example: '1,000.12' },
    { label: 'Percent', value: 'percent', example: '10.12%' },
    { label: 'Scientific', value: 'scientific', example: '1.01E+03' },
    { label: 'Accounting', value: 'accounting', example: '$ (1,000.12)' },
    { label: 'Financial', value: 'financial', example: '(1,000.12)' },
    { label: 'Currency', value: 'currency', example: '$1,000.12' },
    { label: 'Currency rounded', value: 'currency_rounded', example: '$1,000' },
    { label: 'Date', value: 'date', example: '9/26/2008' },
    { label: 'Time', value: 'time', example: '3:59:00 PM' },
    { label: 'Date time', value: 'datetime', example: '9/26/2008 15:59:00' },
    { label: 'Duration', value: 'duration', example: '24:01:00' },
  ];

  const getCurrentNumberFormat = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.numberFormat || 'automatic';
    }
    return 'automatic';
  };

  const getCurrentDecimalPlaces = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.decimalPlaces || 2;
    }
    return 2;
  };

  const handleNumberFormatChange = (format) => {
    if (selectedCell) {
      formatCell({ numberFormat: format });
      showSuccess(`Number format changed to ${format}`);
      setShowNumberFormatDropdown(false);
    }
  };

  const handleCurrencyFormat = () => {
    if (selectedCell) {
      formatCell({ numberFormat: 'currency' });
      showSuccess('Applied currency format');
    }
  };

  const handlePercentageFormat = () => {
    if (selectedCell) {
      formatCell({ numberFormat: 'percent' });
      showSuccess('Applied percentage format');
    }
  };

  const handleIncreaseDecimalPlaces = () => {
    if (selectedCell) {
      const currentPlaces = getCurrentDecimalPlaces();
      formatCell({ decimalPlaces: Math.min(currentPlaces + 1, 10) });
      showSuccess(`Increased decimal places to ${Math.min(currentPlaces + 1, 10)}`);
    }
  };

  const handleDecreaseDecimalPlaces = () => {
    if (selectedCell) {
      const currentPlaces = getCurrentDecimalPlaces();
      formatCell({ decimalPlaces: Math.max(currentPlaces - 1, 0) });
      showSuccess(`Decreased decimal places to ${Math.max(currentPlaces - 1, 0)}`);
    }
  };

  // Virtual scrolling
  const handleScroll = useCallback(
    (e) => {
      const scrollTop = e.target.scrollTop;
      setScrollTop(scrollTop);
      const startRow = Math.floor(scrollTop / 24); // 24px per row
      const endRow = Math.min(startRow + visibleRows, 1000);

      // Update visible range for virtualization
      setVisibleRows(endRow - startRow);
    },
    [visibleRows]
  );

  // Logout handler
  const handleLogout = async () => {
    try {
      await logout();
      showSuccess('You have been logged out successfully.');
      navigate('/');
    } catch (error) {
      console.error('Logout error:', error);
      showError('Failed to logout. Please try again.');
    }
  };

  // Menu functionality
  const handleMenuClick = (menuName) => {
    // Close all other menus
    setShowFileMenu(false);
    setShowEditMenu(false);
    setShowViewMenu(false);
    setShowInsertMenu(false);
    setShowFormatMenu(false);
    setShowDataMenu(false);
    
    // Toggle the clicked menu
    switch (menuName) {
      case 'file':
        setShowFileMenu(!showFileMenu);
        break;
      case 'edit':
        setShowEditMenu(!showEditMenu);
        break;
      case 'view':
        setShowViewMenu(!showViewMenu);
        break;
      case 'insert':
        setShowInsertMenu(!showInsertMenu);
        break;
      case 'format':
        setShowFormatMenu(!showFormatMenu);
        break;
      case 'data':
        setShowDataMenu(!showDataMenu);
        break;
    }
    setActiveMenu(menuName);
  };

  // File menu functions
  const handleNewSpreadsheet = () => {
    setShowFileMenu(false);
    showSuccess('Creating new spreadsheet...');
    // Navigate to create new spreadsheet
    navigate('/spreadsheet-dashboard');
  };

  const handleSaveSpreadsheet = async () => {
    setShowFileMenu(false);
    try {
      await autoSave();
      showSuccess('Spreadsheet saved successfully!');
    } catch (error) {
      showError('Failed to save spreadsheet. Please try again.');
    }
  };

  const handleExportCSV = () => {
    setShowFileMenu(false);
    const csvData = convertToCSV();
    const blob = new Blob([csvData], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'spreadsheet.csv';
    a.click();
    window.URL.revokeObjectURL(url);
    showSuccess('CSV exported successfully!');
  };

  const handleImportCSV = () => {
    csvFileInputRef.current?.click();
  };

  const handleCSVFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.csv')) {
      showError('Please select a CSV file');
      return;
    }

    try {
      const text = await file.text();
      const csvData = parseCSV(text);
      importCSVData(csvData);
      showSuccess(`CSV imported successfully! ${csvData.length} rows imported.`);
    } catch (error) {
      console.error('Error importing CSV:', error);
      showError('Failed to import CSV file. Please check the file format.');
    }

    // Reset the file input
    event.target.value = '';
  };

  const parseCSV = (csvText) => {
    const lines = csvText.split('\n');
    const data = [];
    
    for (const line of lines) {
      if (line.trim()) {
        // Simple CSV parsing - handles basic cases
        const row = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          
          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === ',' && !inQuotes) {
            row.push(current.trim());
            current = '';
          } else {
            current += char;
          }
        }
        
        // Add the last field
        row.push(current.trim());
        
        // Remove quotes from fields
        const cleanedRow = row.map(field => {
          if (field.startsWith('"') && field.endsWith('"')) {
            return field.slice(1, -1);
          }
          return field;
        });
        
        data.push(cleanedRow);
      }
    }
    
    return data;
  };

  const importCSVData = (csvData) => {
    if (!csvData || csvData.length === 0) return;

    const newSheets = sheets.map((sheet) => {
      if (sheet.id === activeSheetId) {
        const newData = { ...sheet.data };
        
        // Clear existing data in the import area (first 100 rows, first 26 columns)
        for (let row = 1; row <= Math.min(csvData.length, 100); row++) {
          for (let col = 1; col <= 26; col++) {
            const cellId = getColumnName(col) + row;
            delete newData[cellId];
          }
        }
        
        // Import CSV data
        csvData.forEach((row, rowIndex) => {
          if (rowIndex >= 100) return; // Limit to 100 rows
          
          row.forEach((cellValue, colIndex) => {
            if (colIndex >= 26) return; // Limit to 26 columns (A-Z)
            
            const cellId = getColumnName(colIndex + 1) + (rowIndex + 1);
            newData[cellId] = {
              value: cellValue,
              formula: cellValue,
              style: {}
            };
          });
        });
        
        return { ...sheet, data: newData };
      }
      return sheet;
    });

    setSheets(newSheets);
    
    // Update backend if we have a current spreadsheet
    if (currentSpreadsheet) {
      try {
        const sheetData = newSheets.find(s => s.id === activeSheetId)?.data || {};
        autoSave(activeSheetId, sheetData);
      } catch (error) {
        console.error('Error auto-saving imported data:', error);
      }
    }
  };

  // Edit menu functions
  const handleUndo = () => {
    if (historyIndex > 0) {
      setHistoryIndex(historyIndex - 1);
      const previousState = history[historyIndex - 1];
      // Restore previous state
      showSuccess('Undo completed');
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      setHistoryIndex(historyIndex + 1);
      const nextState = history[historyIndex + 1];
      // Restore next state
      showSuccess('Redo completed');
    }
  };

  // Local function to get cell data from local state
  const getLocalCellData = (cellId) => {
    const cellData = data[cellId];
    return cellData || { value: "", formula: "", style: {} };
  };

  // Function to save current state to history
  const saveHistory = () => {
    const newHistory = history.slice(0, historyIndex + 1);
    newHistory.push(JSON.parse(JSON.stringify(data)));
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
  };

  // Function to update cell data locally and sync with backend
  const updateCellLocal = async (cellId, value, formula = null, style = null) => {
    console.log(`Updating cell ${cellId} with value:`, value, 'formula:', formula, 'style:', style);
    saveHistory();

    const newSheets = sheets.map((sheet) => {
      if (sheet.id === activeSheetId) {
        const newData = { ...sheet.data };

        if (!newData[cellId]) {
          newData[cellId] = { value: "", formula: "", style: {} };
        }

        const formulaValue = formula !== null ? formula : value;
        const evaluatedValue = formulaValue.startsWith("=")
          ? evaluateFormula(formulaValue)
          : formulaValue;

        const cellData = {
          ...newData[cellId],
          formula: formulaValue,
          value: evaluatedValue,
          style: style ? { ...newData[cellId].style, ...style } : newData[cellId].style,
        };

        newData[cellId] = cellData;
        console.log(`Updated cell ${cellId} data:`, cellData);

        // Recalculate dependent cells
        Object.keys(newData).forEach((id) => {
          if (
            newData[id].formula &&
            newData[id].formula.startsWith("=") &&
            newData[id].formula.includes(cellId)
          ) {
            newData[id].value = evaluateFormula(newData[id].formula);
          }
        });

        return { ...sheet, data: newData };
      }
      return sheet;
    });

    setSheets(newSheets);
    console.log('Updated sheets:', newSheets);

    // Update backend if we have a current spreadsheet
    if (currentSpreadsheet) {
      try {
        const cellData = {
          value: formulaValue.startsWith("=") ? evaluatedValue : formulaValue,
          formula: formulaValue,
          style: newSheets.find(s => s.id === activeSheetId)?.data[cellId]?.style || {}
        };
        
        await updateCell(activeSheetId, cellId, cellData);
        
        // Auto-save the entire sheet data
        const sheetData = newSheets.find(s => s.id === activeSheetId)?.data || {};
        await autoSave(activeSheetId, sheetData);
      } catch (error) {
        console.error('Error updating cell in backend:', error);
      }
    }
  };

  const handleCut = useCallback(() => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      console.log('Cutting cell data:', cellData);
      
      // Store the cell data in clipboard
      setClipboard(cellData);
      setClipboardType('cut');
      
      // Clear the cell and update formula bar
      updateCellLocal(selectedCell, '');
      setFormulaBarValue('');
      
      showSuccess('Cell cut to clipboard');
    } else {
      console.warn('No cell selected for cut operation');
      showError('Please select a cell to cut');
    }
  }, [selectedCell, data, updateCellLocal, setFormulaBarValue, showSuccess, showError]);

  const handleCopy = useCallback(() => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      console.log('Copying cell data:', cellData);
      
      setClipboard(cellData);
      setClipboardType('copy');
      showSuccess('Cell copied to clipboard');
    } else {
      console.warn('No cell selected for copy operation');
      showError('Please select a cell to copy');
    }
  }, [selectedCell, data, showSuccess, showError]);

  const handlePaste = useCallback(() => {
    if (clipboard && selectedCell) {
      console.log('Pasting data:', clipboard);
      
      // Update both the cell data and formula bar with formula and styles
      const formulaToPaste = clipboard.formula || clipboard.value || '';
      updateCellLocal(selectedCell, clipboard.value || '', formulaToPaste, clipboard.style);
      
      // Update formula bar to show the pasted content
      setFormulaBarValue(formulaToPaste);
      
      if (clipboardType === 'cut') {
        setClipboard(null);
        setClipboardType(null);
      }
      showSuccess('Cell pasted');
    } else if (!clipboard) {
      console.warn('No data in clipboard for paste operation');
      showError('No data to paste. Please copy a cell first.');
    } else if (!selectedCell) {
      console.warn('No cell selected for paste operation');
      showError('Please select a cell to paste into.');
    }
  }, [clipboard, selectedCell, clipboardType, updateCellLocal, setFormulaBarValue, showSuccess, showError]);

  // Insert menu functions
  const handleInsertRow = () => {
    setShowInsertMenu(false);
    showSuccess('Row inserted');
    // Implementation for inserting row
  };

  const handleInsertColumn = () => {
    setShowInsertMenu(false);
    showSuccess('Column inserted');
    // Implementation for inserting column
  };

  // Format menu functions
  const handleBold = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      const newStyle = { ...cellData.style, fontWeight: cellData.style?.fontWeight === 'bold' ? 'normal' : 'bold' };
      // Apply the style directly to the cell
      formatCell({ fontWeight: newStyle.fontWeight });
      showSuccess('Text formatting updated');
    }
  };

  const handleItalic = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      const newStyle = { ...cellData.style, fontStyle: cellData.style?.fontStyle === 'italic' ? 'normal' : 'italic' };
      // Apply the style directly to the cell
      formatCell({ fontStyle: newStyle.fontStyle });
      showSuccess('Text formatting updated');
    }
  };

  // Data menu functions
  const handleSortAscending = () => {
    setShowDataMenu(false);
    showSuccess('Data sorted in ascending order');
    // Implementation for sorting
  };

  const handleSortDescending = () => {
    setShowDataMenu(false);
    showSuccess('Data sorted in descending order');
    // Implementation for sorting
  };

  const handleFilter = () => {
    setShowDataMenu(false);
    setShowFilterModal(true);
    showSuccess('Filter applied');
  };

  // Get unique values for a column (for filter options)
  const getColumnFilterOptions = (column) => {
    const colNum = getColumnNumber(column);
    const values = new Set();
    
    for (let row = 1; row <= 1000; row++) {
      const cellId = getColumnName(colNum) + row;
      const cellData = data[cellId];
      if (cellData && cellData.value) {
        values.add(cellData.value.toString());
      }
    }
    
    return Array.from(values).sort();
  };

  // Cell resizing functions
  const getColumnWidth = (column) => {
    return columnWidths[column] || 100; // Default width
  };

  const getRowHeight = (row) => {
    return rowHeights[row] || 24; // Default height
  };

  const handleColumnResize = (column, newWidth) => {
    setColumnWidths({
      ...columnWidths,
      [column]: Math.max(50, newWidth) // Minimum width of 50px
    });
  };

  const handleRowResize = (row, newHeight) => {
    setRowHeights({
      ...rowHeights,
      [row]: Math.max(20, newHeight) // Minimum height of 20px
    });
  };

  const startResize = (type, target, event) => {
    setIsResizing(true);
    setResizeType(type);
    setResizeTarget(target);
    setShowResizeIndicator(true);
    
    // Capture initial position and size
    setResizeStartPos({ x: event.clientX, y: event.clientY });
    
    if (type === 'column') {
      setResizeStartSize({ width: getColumnWidth(target), height: 0 });
    } else if (type === 'row') {
      setResizeStartSize({ width: 0, height: getRowHeight(target) });
    }
  };

  const handleMouseMove = (e) => {
    if (!isResizing || !resizeTarget) return;
    
    if (resizeType === 'column') {
      // Calculate delta from start position
      const deltaX = e.clientX - resizeStartPos.x;
      const newWidth = resizeStartSize.width + deltaX;
      handleColumnResize(resizeTarget, Math.max(50, newWidth));
    } else if (resizeType === 'row') {
      // Calculate delta from start position
      const deltaY = e.clientY - resizeStartPos.y;
      const newHeight = resizeStartSize.height + deltaY;
      handleRowResize(resizeTarget, Math.max(20, newHeight));
    }
  };

  const stopResize = () => {
    setIsResizing(false);
    setResizeType(null);
    setResizeTarget(null);
    setShowResizeIndicator(false);
  };

  // Helper function to convert data to CSV
  const convertToCSV = () => {
    const rows = [];
    for (let row = 1; row <= 100; row++) {
      const rowData = [];
      for (let col = 1; col <= 26; col++) {
        const cellId = getColumnName(col) + row;
        const cellData = data[cellId];
        const value = cellData ? (cellData.value || '') : '';
        rowData.push(`"${value}"`);
      }
      rows.push(rowData.join(','));
    }
    return rows.join('\n');
  };

  // Close filter menu, zoom dropdown, and font dropdowns when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (filterMenuOpen && !event.target.closest(".filter-menu")) {
        setFilterMenuOpen(false);
      }
      if (showZoomDropdown && !event.target.closest(".zoom-dropdown")) {
        setShowZoomDropdown(false);
      }
      if (showFontDropdown && !event.target.closest(".font-dropdown")) {
        setShowFontDropdown(false);
      }
      if (showFontSizeDropdown && !event.target.closest(".font-size-dropdown")) {
        setShowFontSizeDropdown(false);
      }
      if (showNumberFormatDropdown && !event.target.closest(".number-format-dropdown")) {
        setShowNumberFormatDropdown(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [filterMenuOpen, showZoomDropdown, showFontDropdown, showFontSizeDropdown, showNumberFormatDropdown]);

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (event) => {
      // Close all menus when pressing Escape
      if (event.key === 'Escape') {
        setShowFileMenu(false);
        setShowEditMenu(false);
        setShowViewMenu(false);
        setShowInsertMenu(false);
        setShowFormatMenu(false);
        setShowDataMenu(false);
        setShowZoomDropdown(false);
        setShowFontDropdown(false);
        setShowFontSizeDropdown(false);
        setShowNumberFormatDropdown(false);
        setActiveMenu(null);
        return;
      }

      // Ctrl/Cmd + key combinations
      if (event.ctrlKey || event.metaKey) {
        switch (event.key.toLowerCase()) {
          case 's':
            event.preventDefault();
            handleSaveSpreadsheet();
            break;
          case 'c':
            event.preventDefault();
            console.log('Cmd+C pressed, calling handleCopy');
            handleCopy();
            break;
          case 'x':
            event.preventDefault();
            handleCut();
            break;
          case 'v':
            event.preventDefault();
            console.log('Cmd+V pressed, calling handlePaste');
            // Use setTimeout to ensure the paste happens after any other processing
            setTimeout(() => {
            handlePaste();
            }, 0);
            break;
          case 'z':
            event.preventDefault();
            if (event.shiftKey) {
              handleRedo();
            } else {
              handleUndo();
            }
            break;
          case 'y':
            event.preventDefault();
            handleRedo();
            break;
          case 'n':
            event.preventDefault();
            handleNewSpreadsheet();
            break;
          case 'b':
            event.preventDefault();
            handleBold();
            break;
          case 'i':
            event.preventDefault();
            handleItalic();
            break;
          case '=':
          case '+':
            event.preventDefault();
            handleZoomIn();
            break;
          case '-':
            event.preventDefault();
            handleZoomOut();
            break;
          case '0':
            event.preventDefault();
            handleZoomToFit();
            break;
        }
      }

      // Plus/Minus keys for zoom (without Ctrl/Cmd)
      if (!event.ctrlKey && !event.metaKey && !event.altKey) {
        if (event.key === '+' || event.key === '=') {
          event.preventDefault();
          handleZoomIn();
        } else if (event.key === '-') {
          event.preventDefault();
          handleZoomOut();
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, []); // Remove dependencies to prevent frequent recreation

  // Mouse events for resizing
  useEffect(() => {
    if (isResizing) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', stopResize);
      
      return () => {
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', stopResize);
      };
    }
  }, [isResizing, resizeType, resizeTarget]);

  // Enhanced spreadsheet functions with comprehensive banking formulas
  const functions = {
    // Math functions
    SUM: (range) => {
      try {
        const values = getRangeValues(range);
        const numValues = values
          .filter((v) => !isNaN(parseFloat(v)))
          .map((v) => parseFloat(v));
        return numValues.reduce((sum, val) => sum + val, 0);
      } catch (e) {
        return "#ERROR";
      }
    },
    AVERAGE: (range) => {
      try {
        const values = getRangeValues(range).filter(
          (v) => !isNaN(parseFloat(v))
        );
        return values.length
          ? values.reduce((sum, val) => sum + parseFloat(val), 0) /
              values.length
          : 0;
      } catch (e) {
        return "#ERROR";
      }
    },
    COUNT: (range) => {
      try {
        return getRangeValues(range).filter((v) => !isNaN(parseFloat(v)))
          .length;
      } catch (e) {
        return "#ERROR";
      }
    },
    COUNTA: (range) => {
      try {
        return getRangeValues(range).filter((v) => v !== "").length;
      } catch (e) {
        return "#ERROR";
      }
    },
    MAX: (range) => {
      try {
        const values = getRangeValues(range).filter(
          (v) => !isNaN(parseFloat(v))
        );
        return values.length
          ? Math.max(...values.map((v) => parseFloat(v)))
          : 0;
      } catch (e) {
        return "#ERROR";
      }
    },
    MIN: (range) => {
      try {
        const values = getRangeValues(range).filter(
          (v) => !isNaN(parseFloat(v))
        );
        return values.length
          ? Math.min(...values.map((v) => parseFloat(v)))
          : 0;
      } catch (e) {
        return "#ERROR";
      }
    },
    PRODUCT: (range) => {
      try {
        const values = getRangeValues(range).filter(
          (v) => !isNaN(parseFloat(v))
        );
        return values.reduce((product, val) => product * parseFloat(val), 1);
      } catch (e) {
        return "#ERROR";
      }
    },
    POWER: (base, exponent) => {
      try {
        return Math.pow(parseFloat(base) || 0, parseFloat(exponent) || 0);
      } catch (e) {
        return "#ERROR";
      }
    },
    SQRT: (number) => {
      try {
        const num = parseFloat(number) || 0;
        return num >= 0 ? Math.sqrt(num) : "#ERROR";
      } catch (e) {
        return "#ERROR";
      }
    },
    ABS: (number) => {
      try {
        return Math.abs(parseFloat(number) || 0);
      } catch (e) {
        return "#ERROR";
      }
    },
    ROUND: (number, digits = 0) => {
      try {
        const num = parseFloat(number) || 0;
        const dig = parseInt(digits) || 0;
        return Math.round(num * Math.pow(10, dig)) / Math.pow(10, dig);
      } catch (e) {
        return "#ERROR";
      }
    },

    // Banking & Financial Functions
    PMT: (rate, nper, pv, fv = 0, type = 0) => {
      try {
        const r = parseFloat(rate) || 0;
        const n = parseFloat(nper) || 0;
        const present = parseFloat(pv) || 0;
        const future = parseFloat(fv) || 0;
        const t = parseInt(type) || 0;

        if (r === 0) return -(present + future) / n;

        const payment =
          (r * (present * Math.pow(1 + r, n) + future)) /
          ((t ? 1 + r : 1) * (Math.pow(1 + r, n) - 1));
        return -payment;
      } catch (e) {
        return "#ERROR";
      }
    },

    PV: (rate, nper, pmt, fv = 0, type = 0) => {
      try {
        const r = parseFloat(rate) || 0;
        const n = parseFloat(nper) || 0;
        const payment = parseFloat(pmt) || 0;
        const future = parseFloat(fv) || 0;
        const t = parseInt(type) || 0;

        if (r === 0) return -(payment * n + future);

        const pv =
          ((payment * (t ? 1 + r : 1) * (1 - Math.pow(1 + r, -n))) / r -
            future) /
          Math.pow(1 + r, n);
        return pv;
      } catch (e) {
        return "#ERROR";
      }
    },

    FV: (rate, nper, pmt, pv = 0, type = 0) => {
      try {
        const r = parseFloat(rate) || 0;
        const n = parseFloat(nper) || 0;
        const payment = parseFloat(pmt) || 0;
        const present = parseFloat(pv) || 0;
        const t = parseInt(type) || 0;

        if (r === 0) return -(present + payment * n);

        const fv = -(
          present * Math.pow(1 + r, n) +
          (payment * (t ? 1 + r : 1) * (Math.pow(1 + r, n) - 1)) / r
        );
        return fv;
      } catch (e) {
        return "#ERROR";
      }
    },

    NPER: (rate, pmt, pv, fv = 0, type = 0) => {
      try {
        const r = parseFloat(rate) || 0;
        const payment = parseFloat(pmt) || 0;
        const present = parseFloat(pv) || 0;
        const future = parseFloat(fv) || 0;
        const t = parseInt(type) || 0;

        if (r === 0) return -(present + future) / payment;

        const temp = (payment * (t ? 1 + r : 1)) / r;
        return Math.log((temp - future) / (temp + present)) / Math.log(1 + r);
      } catch (e) {
        return "#ERROR";
      }
    },

    NPV: (rate, ...values) => {
      try {
        const r = parseFloat(rate) || 0;
        return values.reduce((npv, value, index) => {
          return npv + (parseFloat(value) || 0) / Math.pow(1 + r, index + 1);
        }, 0);
      } catch (e) {
        return "#ERROR";
      }
    },

    IRR: (...values) => {
      try {
        const cashFlows = values.map((v) => parseFloat(v) || 0);
        let rate = 0.1;

        for (let i = 0; i < 100; i++) {
          let npv = 0;
          let dnpv = 0;

          for (let j = 0; j < cashFlows.length; j++) {
            npv += cashFlows[j] / Math.pow(1 + rate, j);
            dnpv -= (j * cashFlows[j]) / Math.pow(1 + rate, j + 1);
          }

          const newRate = rate - npv / dnpv;
          if (Math.abs(newRate - rate) < 0.000001) return newRate;
          rate = newRate;
        }
        return rate;
      } catch (e) {
        return "#ERROR";
      }
    },

    // Logical functions
    IF: (condition, trueValue, falseValue = "") => {
      try {
        return condition ? trueValue : falseValue;
      } catch (e) {
        return "#ERROR";
      }
    },

    // Text functions
    CONCATENATE: (...args) => {
      try {
        return args.map((arg) => String(arg || "")).join("");
      } catch (e) {
        return "#ERROR";
      }
    },

    // Date functions
    TODAY: () => {
      try {
        return new Date().toLocaleDateString();
      } catch (e) {
        return "#ERROR";
      }
    },
    NOW: () => {
      try {
        return new Date().toLocaleString();
      } catch (e) {
        return "#ERROR";
      }
    },
  };

  const getRangeValues = (range) => {
    if (!range) return [];

    try {
      if (range.includes(":")) {
        const [start, end] = range.split(":");
        const startMatch = start.match(/([A-Z]+)(\d+)/);
        const endMatch = end.match(/([A-Z]+)(\d+)/);

        if (!startMatch || !endMatch) return [];

        const startCol = getColumnNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);
        const endCol = getColumnNumber(endMatch[1]);
        const endRow = parseInt(endMatch[2]);

        const values = [];
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            const cellId = getColumnName(col) + row;
            const cellData = data[cellId];
            if (cellData) {
              values.push(cellData.value || "");
            } else {
              values.push("");
            }
          }
        }
        return values;
      } else {
        const cellData = data[range];
        return cellData ? [cellData.value || ""] : [""];
      }
    } catch (e) {
      return [];
    }
  };

  const getColumnNumber = (columnName) => {
    let result = 0;
    for (let i = 0; i < columnName.length; i++) {
      result = result * 26 + (columnName.charCodeAt(i) - 64);
    }
    return result;
  };

  const evaluateFormula = (formula) => {
    if (!formula || !formula.startsWith("=")) return formula;

    try {
      let expression = formula.slice(1).trim();

      // Replace cell references with values
      expression = expression.replace(/[A-Z]+\d+/g, (match) => {
        const cellData = data[match];
        const value = cellData ? cellData.value : "";
        if (value === "" || value === null || value === undefined) return "0";
        const numValue = parseFloat(value);
        return isNaN(numValue) ? `"${value}"` : numValue.toString();
      });

      // Handle function calls
      Object.keys(functions).forEach((funcName) => {
        const regex = new RegExp(`\\b${funcName}\\s*\\(([^()]*)\\)`, "gi");
        expression = expression.replace(regex, (match, args) => {
          try {
            let argList = [];
            if (args && args.trim()) {
              argList = args.split(",").map((arg) => {
                arg = arg.trim();
                if (arg.startsWith('"') && arg.endsWith('"')) {
                  return arg.slice(1, -1);
                }
                if (arg.includes(":")) {
                  return arg;
                }
                const numArg = parseFloat(arg);
                return isNaN(numArg) ? arg : numArg;
              });
            }

            const result = functions[funcName](...argList);
            if (result === "#ERROR") return "#ERROR";
            return typeof result === "string"
              ? `"${result}"`
              : (result || 0).toString();
          } catch (error) {
            return "#ERROR";
          }
        });
      });

      if (expression.includes("#ERROR")) {
        return "#ERROR";
      }

      // Safe evaluation for basic math
      const allowedChars = /^[0-9+\-*/().\s"]*$/;
      const cleanExpression = expression.replace(/"[^"]*"/g, '""');

      if (allowedChars.test(cleanExpression)) {
        try {
          const result = Function(
            '"use strict"; return (' + expression + ")"
          )();
          if (
            typeof result === "number" &&
            (isNaN(result) || !isFinite(result))
          ) {
            return "#ERROR";
          }
          return typeof result === "number"
            ? result
            : typeof result === "string"
            ? result
            : "#ERROR";
        } catch (evalError) {
          return "#ERROR";
        }
      }

      return expression;
    } catch (error) {
      return "#ERROR";
    }
  };



  const undo = () => {
    if (historyIndex > 0) {
      const newSheets = sheets.map((sheet) => {
        if (sheet.id === activeSheetId) {
          return { ...sheet, data: history[historyIndex - 1] };
        }
        return sheet;
      });
      setSheets(newSheets);
      setHistoryIndex(historyIndex - 1);
    }
  };

  const redo = () => {
    if (historyIndex < history.length - 1) {
      const newSheets = sheets.map((sheet) => {
        if (sheet.id === activeSheetId) {
          return { ...sheet, data: history[historyIndex + 1] };
        }
        return sheet;
      });
      setSheets(newSheets);
      setHistoryIndex(historyIndex + 1);
    }
  };


  const deleteCell = () => {
    if (selectedCell) {
      updateCellLocal(selectedCell, "");
    }
  };

  const addSheet = () => {
    const newSheetId = `sheet${sheets.length + 1}`;
    setSheets([
      ...sheets,
      {
        id: newSheetId,
        name: `Sheet${sheets.length + 1}`,
        data: initializeData(),
      },
    ]);
  };

  const formatCell = (style) => {
    if (!selectedCell) return;

    const newSheets = sheets.map((sheet) => {
      if (sheet.id === activeSheetId) {
        const newData = { ...sheet.data };
        if (!newData[selectedCell]) {
          newData[selectedCell] = { value: "", formula: "", style: {} };
        }
        const currentStyle = newData[selectedCell].style || {};
        const updatedStyle = { ...currentStyle, ...style };
        
        newData[selectedCell] = {
          ...newData[selectedCell],
          style: updatedStyle,
        };
        
        return { ...sheet, data: newData };
      }
      return sheet;
    });

    setSheets(newSheets);
    
    // Force re-render by updating trigger
    setFormatUpdateTrigger(prev => prev + 1);
  };

  const getCellStyle = (cellId) => {
    const cellData = data[cellId];
    if (!cellData || !cellData.style) return {};

    const style = {};
    if (cellData.style.fontFamily) style.fontFamily = cellData.style.fontFamily;
    if (cellData.style.fontSize) style.fontSize = `${cellData.style.fontSize}px`;
    if (cellData.style.fontWeight) style.fontWeight = cellData.style.fontWeight;
    if (cellData.style.fontStyle) style.fontStyle = cellData.style.fontStyle;
    if (cellData.style.underline) style.textDecoration = "underline";
    if (cellData.style.textDecoration) style.textDecoration = cellData.style.textDecoration;
    if (cellData.style.color) style.color = cellData.style.color;
    if (cellData.style.backgroundColor)
      style.backgroundColor = cellData.style.backgroundColor;
    if (cellData.style.textAlign) style.textAlign = cellData.style.textAlign;

    return style;
  };

  const formatCellValue = (cellData) => {
    if (!cellData || !cellData.style?.numberFormat || cellData.style.numberFormat === 'automatic') {
      return cellData?.value || '';
    }

    const value = cellData.value;
    const numberFormat = cellData.style.numberFormat;
    const decimalPlaces = cellData.style.decimalPlaces || 2;

    // Convert string to number if it's a valid number
    let numericValue = value;
    if (typeof value === 'string' && !isNaN(value) && value.trim() !== '') {
      numericValue = parseFloat(value);
    }

    if (typeof numericValue !== 'number' || isNaN(numericValue)) {
      return value;
    }

    switch (numberFormat) {
      case 'currency':
        return `$${numericValue.toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces })}`;
      case 'currency_rounded':
        return `$${numericValue.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
      case 'percent':
        return `${(numericValue * 100).toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces })}%`;
      case 'number':
        return numericValue.toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces });
      case 'scientific':
        return numericValue.toExponential(decimalPlaces);
      case 'accounting':
        return numericValue < 0 ? `($${Math.abs(numericValue).toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces })})` : `$${numericValue.toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces })}`;
      case 'financial':
        return numericValue < 0 ? `(${Math.abs(numericValue).toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces })})` : numericValue.toLocaleString(undefined, { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces });
      case 'date':
        return new Date(numericValue).toLocaleDateString();
      case 'time':
        return new Date(numericValue).toLocaleTimeString();
      case 'datetime':
        return new Date(numericValue).toLocaleString();
      default:
        return numericValue;
    }
  };

  const handleCellClick = (cellId) => {
    // Set navigation flag to prevent interference
    isNavigatingRef.current = true;
    
    // If we're currently editing a different cell, save it first
    if (isEditing && selectedCell && selectedCell !== cellId) {
      updateCellLocal(selectedCell, cellEditValue);
      setIsEditing(false);
    }
    
    setSelectedCell(cellId);
    const cellData = data[cellId];
    const cellValue = cellData ? cellData.formula || cellData.value || "" : "";
    setFormulaBarValue(cellValue);
    setCellEditValue(cellValue);
    setIsProcessingInput(false);
    
    // Set editing state and focus immediately
    setIsEditing(true);
    
    // Use requestAnimationFrame for better timing
    requestAnimationFrame(() => {
      if (cellInputRef.current) {
        cellInputRef.current.focus();
        // Position cursor at end instead of selecting all text
        const length = cellInputRef.current.value.length;
        cellInputRef.current.setSelectionRange(length, length);
      }
      // Clear navigation flag after focus is set
      setTimeout(() => {
        isNavigatingRef.current = false;
      }, 50);
    });
  };

  const handleCellDoubleClick = (cellId) => {
    // If we're currently editing a different cell, save it first
    if (isEditing && selectedCell && selectedCell !== cellId) {
      updateCellLocal(selectedCell, cellEditValue);
      setIsEditing(false);
    }
    
    // Double-click selects all text for replacement
    setSelectedCell(cellId);
    const cellData = data[cellId];
    const cellValue = cellData ? cellData.formula || cellData.value || "" : "";
    setCellEditValue(cellValue);
    setFormulaBarValue(cellValue);
    setIsProcessingInput(false);
    
    // Set editing state immediately
    setIsEditing(true);
    
    // Use requestAnimationFrame for better timing
    requestAnimationFrame(() => {
      if (cellInputRef.current) {
        cellInputRef.current.focus();
        cellInputRef.current.select(); // Select all text for replacement
      }
    });
  };

  const handleFormulaBarChange = (value) => {
    setFormulaBarValue(value);
    if (selectedCell) {
      updateCellLocal(selectedCell, value);
    }
  };

  // Handle keyboard input when cell is selected but not editing
  const handleKeyPress = (e) => {
    // Don't interfere with copy/paste shortcuts
    if (e.ctrlKey || e.metaKey) {
      return; // Let the handleKeyDown function handle these
    }
    
    // Only handle printable characters, not special keys
    const isPrintableKey = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
    
    // Only handle if we're not already editing, not processing, and not in formula bar
    if (!isEditing && !isProcessingInput && selectedCell && isPrintableKey && document.activeElement !== formulaBarInputRef.current) {
      // Prevent the default behavior to avoid duplication
      e.preventDefault();
      setIsProcessingInput(true);
      // Start editing when user types a character
      setIsEditing(true);
      const cellData = data[selectedCell];
      const existingValue = cellData ? cellData.formula || cellData.value || "" : "";
      setCellEditValue(existingValue + e.key);
      setFormulaBarValue(existingValue + e.key);
      setTimeout(() => {
        if (cellInputRef.current) {
          cellInputRef.current.focus();
          cellInputRef.current.setSelectionRange(existingValue.length + 1, existingValue.length + 1);
        }
        setIsProcessingInput(false);
      }, 0);
    }
  };

  // Ensure the main container can receive focus for keyboard events
  useEffect(() => {
    const mainContainer = document.querySelector('.main-spreadsheet-container');
    if (mainContainer) {
      mainContainer.focus();
    }
  }, []);

  // Handle clicks outside the spreadsheet to save current cell
  useEffect(() => {
    const handleClickOutside = (e) => {
      const isInsideSpreadsheet = e.target.closest('.main-spreadsheet-container');
      const isCellInput = e.target.tagName === 'INPUT' && e.target.className.includes('border-none');
      
      if (!isInsideSpreadsheet && isEditing && selectedCell && !isCellInput) {
        updateCellLocal(selectedCell, cellEditValue);
        setIsEditing(false);
      }
    };

    // Use a longer delay to prevent interference with cell navigation
    const timeoutId = setTimeout(() => {
      document.addEventListener('mousedown', handleClickOutside);
    }, 200);

    return () => {
      clearTimeout(timeoutId);
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isEditing, selectedCell, cellEditValue]);


  const generateChart = () => {
    const selectedData = [];
    for (let i = 1; i <= 10; i++) {
      const cellA = data[`A${i}`];
      const cellB = data[`B${i}`];
      if (cellA && cellB && cellA.value && cellB.value) {
        selectedData.push({
          name: cellA.value,
          value: parseFloat(cellB.value) || 0,
        });
      }
    }

    if (selectedData.length > 0) {
      setChartData(selectedData);
      setShowChartModal(true);
    }
  };

  const colors = [
    "#8884d8",
    "#82ca9d",
    "#ffc658",
    "#ff7300",
    "#8dd1e1",
    "#d084d0",
  ];

  const renderChart = () => {
    if (!chartData) return null;

    switch (chartType) {
      case "line":
        return (
          <RechartsLineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Tooltip />
            <Legend />
            <Line type="monotone" dataKey="value" stroke="#8884d8" />
          </RechartsLineChart>
        );
      case "bar":
        return (
          <RechartsBarChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Tooltip />
            <Legend />
            <Bar dataKey="value" fill="#8884d8" />
          </RechartsBarChart>
        );
      case "pie":
        return (
          <RechartsPieChart>
            <Tooltip />
            <Legend />
            <Cell fill="#8884d8" />
          </RechartsPieChart>
        );
      default:
        return null;
    }
  };

  const insertFunction = (funcName) => {
    const currentFormula = formulaBarValue;
    const newFormula = currentFormula.startsWith("=")
      ? `${currentFormula}${funcName}()`
      : `=${funcName}()`;
    setFormulaBarValue(newFormula);
    setShowFunctionMenu(false);
  };

  const functionCategories = {
    Math: [
      "SUM",
      "AVERAGE",
      "COUNT",
      "MAX",
      "MIN",
      "ROUND",
      "ABS",
      "POWER",
      "SQRT",
    ],
    Financial: ["PMT", "PV", "FV", "NPER", "NPV", "IRR"],
    Text: ["CONCATENATE"],
    Date: ["TODAY", "NOW"],
    Logical: ["IF"],
  };

  return (
    <div 
      className="flex flex-col h-screen bg-white main-spreadsheet-container" 
      onKeyDown={handleKeyPress}
      tabIndex={0}
    >
      {/* Strix Header */}
      <div className="bg-white border-b border-gray-200">
        {/* Top Header Bar */}
        <div className="flex items-center justify-between px-4 py-1 bg-white border-b border-gray-200">
          <div className="flex items-center">
            <h1 className="text-2xl font-bold mt-2 bg-gradient-to-r from-indigo-600 to-purple-600 bg-clip-text text-transparent">
              Strix Sheets
            </h1>
          </div>

          {/* Logout Button */}
          <button
            onClick={handleLogout}
            className="px-4 py-2 bg-red-600 text-white rounded-md text-sm font-medium hover:bg-red-700 flex items-center space-x-2"
          >
            <svg
              className="w-4 h-4"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M10 19l-7-7m0 0l7-7m-7 7h18"
              />
            </svg>
            <span>Logout</span>
          </button>
        </div>

        {/* Menu Bar */}
        <div className="flex items-center px-4 py-1 bg-white border-b border-gray-200 relative">
          <div className="flex items-center space-x-2">
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('file')}
                className={`text-md px-2 py-1 rounded ${showFileMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              File
            </button>
              {showFileMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button onClick={handleNewSpreadsheet} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    New Spreadsheet
                  </button>
                  <button onClick={handleSaveSpreadsheet} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Save
                  </button>
                  <button onClick={handleExportCSV} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Export as CSV
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('edit')}
                className={`text-md px-2 py-1 rounded ${showEditMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              Edit
            </button>
              {showEditMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button onClick={handleUndo} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Undo
                  </button>
                  <button onClick={handleRedo} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Redo
                  </button>
                  <div className="border-t border-gray-200"></div>
                  <button onClick={handleCut} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Cut
                  </button>
                  <button onClick={handleCopy} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Copy
                  </button>
                  <button onClick={handlePaste} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Paste
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('view')}
                className={`text-md px-2 py-1 rounded ${showViewMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              View
            </button>
              {showViewMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button 
                    onClick={() => {
                      handleZoomIn();
                      setShowViewMenu(false);
                    }}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Zoom In
                  </button>
                  <button 
                    onClick={() => {
                      handleZoomOut();
                      setShowViewMenu(false);
                    }}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Zoom Out
                  </button>
                  <button 
                    onClick={() => {
                      handleZoomToFit();
                      setShowViewMenu(false);
                    }}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Zoom to Fit
                  </button>
                  <div className="border-t my-1"></div>
                  <button className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Show Grid Lines
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('insert')}
                className={`text-md px-2 py-1 rounded ${showInsertMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              Insert
            </button>
              {showInsertMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button onClick={handleInsertRow} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Insert Row
                  </button>
                  <button onClick={handleInsertColumn} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Insert Column
                  </button>
                  <button className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Insert Chart
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('format')}
                className={`text-md px-2 py-1 rounded ${showFormatMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              Format
            </button>
              {showFormatMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button onClick={handleBold} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Bold
                  </button>
                  <button onClick={handleItalic} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Italic
                  </button>
                  <button className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Text Color
                  </button>
                  <button className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Background Color
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick('data')}
                className={`text-md px-2 py-1 rounded ${showDataMenu ? 'bg-gray-200 text-gray-900' : 'text-gray-700 hover:text-gray-900 hover:bg-gray-100'}`}
              >
              Data
            </button>
              {showDataMenu && (
                <div className="absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48">
                  <button onClick={handleSortAscending} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Sort A to Z
                  </button>
                  <button onClick={handleSortDescending} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Sort Z to A
                  </button>
                  <button onClick={handleFilter} className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm">
                    Filter
                  </button>
                </div>
              )}
            </div>
            
            <button className="text-md text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-2 py-1 rounded">
              Tools
            </button>
            <button className="text-md text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-2 py-1 rounded">
              Extensions
            </button>
            <button className="text-md text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-2 py-1 rounded">
              Help
            </button>
          </div>
        </div>
      </div>

      {/* Google Sheets Style Toolbar */}
      <div className="bg-white border-b border-gray-200 px-3 py-2">
        <div className="flex items-center justify-between">
          {/* Left side toolbar */}
          <div className="flex items-center space-x-1">
            {/* Undo/Redo */}
            <button onClick={handleUndo} className="p-2 hover:bg-gray-100 rounded">
              <Undo size={16} className="text-gray-600" />
            </button>
            <button onClick={handleRedo} className="p-2 hover:bg-gray-100 rounded">
              <Redo size={16} className="text-gray-600" />
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Zoom Controls */}
            <div className="flex items-center space-x-1">
              <button 
                onClick={handleZoomOut}
                className="p-1 text-gray-600 hover:bg-gray-100 rounded"
                title="Zoom Out"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 12H4" />
              </svg>
            </button>

              <div className="relative">
                <button 
                  onClick={() => setShowZoomDropdown(!showZoomDropdown)}
                  className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-12"
                >
                  {zoomLevel}%
                </button>
                
                {showZoomDropdown && (
                  <div className="zoom-dropdown absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-32">
                    {zoomLevels.map((level) => (
                      <button
                        key={level}
                        onClick={() => handleZoomChange(level)}
                        className={`block w-full text-left px-3 py-2 text-sm hover:bg-gray-100 ${
                          level === zoomLevel ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        {level}%
              </button>
                    ))}
                  </div>
                )}
            </div>

              <button 
                onClick={handleZoomIn}
                className="p-1 text-gray-600 hover:bg-gray-100 rounded"
                title="Zoom In"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6v6m0 0v6m0-6h6m-6 0H6" />
                </svg>
              </button>
              
              <div className="w-px h-4 bg-gray-300"></div>
            </div>

            {/* Number Format Dropdown */}
            <div className="relative">
              <button 
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  setShowNumberFormatDropdown(!showNumberFormatDropdown);
                }}
                className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-16 text-left"
              >
                {numberFormats.find(f => f.value === getCurrentNumberFormat())?.label || '123'}
              </button>
              
              {showNumberFormatDropdown && (
                <div className="number-format-dropdown absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-64 max-h-80 overflow-y-auto">
                  <div className="p-2">
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2">General</div>
                    {numberFormats.filter(f => ['automatic', 'text'].includes(f.value)).map((format) => (
                      <button
                        key={format.value}
                        onClick={() => handleNumberFormatChange(format.value)}
                        className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                          format.value === getCurrentNumberFormat() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        <div className="flex justify-between items-center">
                          <span>{format.label}</span>
                          <span className="text-xs text-gray-400">{format.example}</span>
                        </div>
                      </button>
                    ))}
                    
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">Number</div>
                    {numberFormats.filter(f => ['number', 'percent', 'scientific'].includes(f.value)).map((format) => (
                      <button
                        key={format.value}
                        onClick={() => handleNumberFormatChange(format.value)}
                        className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                          format.value === getCurrentNumberFormat() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        <div className="flex justify-between items-center">
                          <span>{format.label}</span>
                          <span className="text-xs text-gray-400">{format.example}</span>
                        </div>
                      </button>
                    ))}
                    
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">Currency</div>
                    {numberFormats.filter(f => ['accounting', 'financial', 'currency', 'currency_rounded'].includes(f.value)).map((format) => (
                      <button
                        key={format.value}
                        onClick={() => handleNumberFormatChange(format.value)}
                        className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                          format.value === getCurrentNumberFormat() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        <div className="flex justify-between items-center">
                          <span>{format.label}</span>
                          <span className="text-xs text-gray-400">{format.example}</span>
                        </div>
                      </button>
                    ))}
                    
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">Date & time</div>
                    {numberFormats.filter(f => ['date', 'time', 'datetime', 'duration'].includes(f.value)).map((format) => (
                      <button
                        key={format.value}
                        onClick={() => handleNumberFormatChange(format.value)}
                        className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                          format.value === getCurrentNumberFormat() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        <div className="flex justify-between items-center">
                          <span>{format.label}</span>
                          <span className="text-xs text-gray-400">{format.example}</span>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Currency Button */}
            <button 
              onClick={handleCurrencyFormat}
              className={`p-2 rounded ${
                getCurrentNumberFormat() === 'currency' ? 'bg-blue-100 text-blue-600' : 'hover:bg-gray-100'
              }`}
              title="Format as currency"
            >
              <DollarSign size={16} className="text-gray-600" />
            </button>
            
            {/* Percentage Button */}
            <button 
              onClick={handlePercentageFormat}
              className={`p-2 rounded ${
                getCurrentNumberFormat() === 'percent' ? 'bg-blue-100 text-blue-600' : 'hover:bg-gray-100'
              }`}
              title="Format as percent"
            >
              <span className="text-sm text-gray-600">%</span>
            </button>

            {/* Decimal places - Decrease */}
            <button 
              onClick={handleDecreaseDecimalPlaces}
              className="p-2 hover:bg-gray-100 rounded"
              title="Decrease decimal places"
            >
              <div className="flex items-center">
                <span className="text-xs text-gray-600">.00</span>
                <svg className="w-3 h-3 text-gray-600 ml-1" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clipRule="evenodd" />
              </svg>
              </div>
            </button>
            
            {/* Decimal places - Increase */}
            <button 
              onClick={handleIncreaseDecimalPlaces}
              className="p-2 hover:bg-gray-100 rounded"
              title="Increase decimal places"
            >
              <div className="flex items-center">
                <span className="text-xs text-gray-600">.0</span>
                <svg className="w-3 h-3 text-gray-600 ml-1" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M14.707 12.707a1 1 0 01-1.414 0L10 9.414l-3.293 3.293a1 1 0 01-1.414-1.414l4-4a1 1 0 011.414 0l4 4a1 1 0 010 1.414z" clipRule="evenodd" />
              </svg>
              </div>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Font controls */}
            <div className="flex items-center space-x-1">
              {/* Font Family Dropdown */}
              <div className="relative">
                <button 
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    console.log('Font dropdown clicked, current state:', showFontDropdown);
                    setShowFontDropdown(!showFontDropdown);
                  }}
                  className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-20 text-left"
                  style={{ fontFamily: getCurrentFontFamily() }}
                >
                  {getCurrentFontFamily()}
              </button>
                
                {showFontDropdown && (
                  <div className="font-dropdown absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-48 max-h-64 overflow-y-auto">
                    {console.log('Rendering font dropdown with fonts:', fontFamilies)}
                    {fontFamilies.map((font) => (
                      <button
                        key={font}
                        onClick={() => handleFontFamilyChange(font)}
                        className={`block w-full text-left px-3 py-2 text-sm hover:bg-gray-100 ${
                          font === getCurrentFontFamily() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                        style={{ fontFamily: font }}
                      >
                        {font}
                </button>
                    ))}
                  </div>
                )}
              </div>
              
              {/* Font Size Dropdown */}
              <div className="relative">
                <button 
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    console.log('Font size dropdown clicked, current state:', showFontSizeDropdown);
                    setShowFontSizeDropdown(!showFontSizeDropdown);
                  }}
                  className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-8 text-center"
                >
                  {getCurrentFontSize()}
                  </button>
                
                {showFontSizeDropdown && (
                  <div className="font-size-dropdown absolute top-full left-0 bg-white border border-gray-200 rounded shadow-lg z-50 min-w-16 max-h-48 overflow-y-auto">
                    {fontSizes.map((size) => (
                      <button
                        key={size}
                        onClick={() => handleFontSizeChange(size)}
                        className={`block w-full text-center px-3 py-2 text-sm hover:bg-gray-100 ${
                          size === getCurrentFontSize() ? 'bg-blue-50 text-blue-600' : ''
                        }`}
                      >
                        {size}
                  </button>
                    ))}
                </div>
                )}
              </div>
            </div>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Text formatting */}
            <button
              onClick={handleBold}
              className="p-2 hover:bg-gray-100 rounded font-bold text-gray-600"
            >
              B
            </button>
            <button
              onClick={handleItalic}
              className="p-2 hover:bg-gray-100 rounded italic text-gray-600"
            >
              I
            </button>
            <button
              onClick={() => {
                if (selectedCell) {
                  const cellData = getLocalCellData(selectedCell);
                  const newStyle = { ...cellData.style, textDecoration: cellData.style?.textDecoration === 'line-through' ? 'none' : 'line-through' };
                  formatCell({ textDecoration: newStyle.textDecoration });
                  showSuccess('Text formatting updated');
                }
              }}
              className="p-2 hover:bg-gray-100 rounded underline text-gray-600"
            >
              <span className="text-sm">S</span>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Colors */}
            <button className="p-2 hover:bg-gray-100 rounded">
              <svg
                className="w-4 h-4 text-gray-600"
                fill="currentColor"
                viewBox="0 0 24 24"
              >
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z" />
              </svg>
            </button>
            <button className="p-2 hover:bg-gray-100 rounded">
              <svg
                className="w-4 h-4 text-gray-600"
                fill="currentColor"
                viewBox="0 0 24 24"
              >
                <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM5 19V5h14v14H5z" />
              </svg>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Borders */}
            <button className="p-2 hover:bg-gray-100 rounded">
              <svg
                className="w-4 h-4 text-gray-600"
                fill="currentColor"
                viewBox="0 0 24 24"
              >
                <path d="M3 3h18v2H3V3zm0 4h18v2H3V7zm0 4h18v2H3v-2zm0 4h18v2H3v-2zm0 4h18v2H3v-2z" />
              </svg>
            </button>

            {/* Merge cells */}
            <button className="p-2 hover:bg-gray-100 rounded">
              <svg
                className="w-4 h-4 text-gray-600"
                fill="currentColor"
                viewBox="0 0 24 24"
              >
                <path d="M3 3v18h18V3H3zm16 16H5V5h14v14zM7 7h6v2H7V7zm0 4h6v2H7v-2zm0 4h6v2H7v-2z" />
              </svg>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Alignment */}
            <div className="flex items-center space-x-1">
              <button
                onClick={() => formatCell({ textAlign: "left" })}
                className="p-2 hover:bg-gray-100 rounded"
              >
                <AlignLeft size={16} className="text-gray-600" />
              </button>
              <button
                onClick={() => formatCell({ textAlign: "center" })}
                className="p-2 hover:bg-gray-100 rounded"
              >
                <AlignCenter size={16} className="text-gray-600" />
              </button>
              <button
                onClick={() => formatCell({ textAlign: "right" })}
                className="p-2 hover:bg-gray-100 rounded"
              >
                <AlignRight size={16} className="text-gray-600" />
              </button>
            </div>

            {/* More alignment options */}
            <div className="flex items-center space-x-1">
              <button className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded">
                Top
              </button>
              <button className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded">
                Middle
              </button>
              <button className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded">
                Bottom
              </button>
            </div>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* More tools */}
            <button className="p-2 hover:bg-gray-100 rounded">
              <svg
                className="w-4 h-4 text-gray-600"
                fill="currentColor"
                viewBox="0 0 24 24"
              >
                <path d="M3.9 12c0-1.71 1.39-3.1 3.1-3.1h4V7H7c-2.76 0-5 2.24-5 5s2.24 5 5 5h4v-1.9H7c-1.71 0-3.1-1.39-3.1-3.1zM8 13h8v-2H8v2zm9-6h-4v1.9h4c1.71 0 3.1 1.39 3.1 3.1s-1.39 3.1-3.1 3.1h-4V17h4c2.76 0 5-2.24 5-5s-2.24-5-5-5z" />
              </svg>
            </button>

            <button className="p-2 hover:bg-gray-100 rounded">
              <MessageSquare size={16} className="text-gray-600" />
            </button>

            <button className="p-2 hover:bg-gray-100 rounded">
              <BarChart size={16} className="text-gray-600" />
            </button>

            {/* Filter */}
            <div className="relative">
              <button
                onClick={() => setFilterMenuOpen(!filterMenuOpen)}
                className={`p-2 rounded flex items-center ${
                  Object.keys(columnFilters).length > 0
                    ? "bg-blue-100 text-blue-600"
                    : "hover:bg-gray-100"
                }`}
                title="Create a filter"
              >
                <Filter size={16} />
                {Object.keys(columnFilters).length > 0 && (
                  <span className="ml-1 text-xs bg-blue-600 text-white rounded-full px-1.5 py-0.5">
                    {Object.keys(columnFilters).length}
                  </span>
                )}
              </button>
            </div>
          </div>

          <div className="relative">
            <button
              onClick={() => setShowFunctionMenu(!showFunctionMenu)}
              className="p-2 hover:bg-gray-100 rounded flex items-center"
            >
              <Calculator size={16} />
              <ChevronDown size={12} className="ml-1" />
            </button>

            {showFunctionMenu && (
              <div className="absolute top-full left-0 mt-1 w-64 bg-white border rounded-lg shadow-lg z-50">
                <div className="p-2 border-b">
                  <input
                    type="text"
                    placeholder="Search functions..."
                    className="w-full px-2 py-1 border rounded text-sm"
                    value={formulaSearch}
                    onChange={(e) => setFormulaSearch(e.target.value)}
                  />
                </div>
                <div className="max-h-64 overflow-y-auto">
                  {Object.entries(functionCategories).map(
                    ([category, funcs]) => (
                      <div key={category} className="p-2">
                        <div className="font-semibold text-sm text-gray-600 mb-1">
                          {category}
                        </div>
                        {funcs
                          .filter(
                            (func) =>
                              !formulaSearch ||
                              func
                                .toLowerCase()
                                .includes(formulaSearch.toLowerCase())
                          )
                          .map((func) => (
                            <button
                              key={func}
                              onClick={() => insertFunction(func)}
                              className="block w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded"
                            >
                              {func}
                            </button>
                          ))}
                      </div>
                    )
                  )}
                </div>
              </div>
            )}
          </div>
          <div className="relative">
            <button
              onClick={() => setShowColorPicker(!showColorPicker)}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <Palette size={16} />
            </button>

            {showColorPicker && (
              <div className="absolute top-full left-0 mt-1 w-48 bg-white border rounded-lg shadow-lg z-50 p-3">
                <div className="grid grid-cols-6 gap-2 mb-3">
                  {[
                    "#000000",
                    "#FF0000",
                    "#00FF00",
                    "#0000FF",
                    "#FFFF00",
                    "#FF00FF",
                    "#00FFFF",
                    "#FFA500",
                    "#800080",
                    "#008000",
                    "#FFC0CB",
                    "#A52A2A",
                  ].map((color) => (
                    <button
                      key={color}
                      onClick={() => {
                        formatCell({ color });
                        setShowColorPicker(false);
                      }}
                      className="w-6 h-6 rounded border"
                      style={{ backgroundColor: color }}
                    />
                  ))}
                </div>
                <div className="border-t pt-2">
                  <div className="text-sm font-medium mb-2">Background</div>
                  <div className="grid grid-cols-6 gap-2">
                    {[
                      "#FFFFFF",
                      "#F0F0F0",
                      "#FFEEEE",
                      "#EEFFEE",
                      "#EEEEFF",
                      "#FFFFEE",
                      "#FFEEFF",
                      "#EEFFFF",
                      "#FFE4B5",
                      "#E6E6FA",
                      "#F0E68C",
                      "#FFB6C1",
                    ].map((color) => (
                      <button
                        key={color}
                        onClick={() => {
                          formatCell({ backgroundColor: color });
                          setShowColorPicker(false);
                        }}
                        className="w-6 h-6 rounded border"
                        style={{ backgroundColor: color }}
                      />
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          <button
            onClick={deleteCell}
            className="p-2 hover:bg-gray-100 rounded"
          >
            <Trash2 size={16} />
          </button>

          {/* Filter Button */}
          <div className="relative">
            <button
              onClick={() => setFilterMenuOpen(!filterMenuOpen)}
              className={`p-2 rounded flex items-center ${
                Object.keys(activeFilters).length > 0
                  ? "bg-blue-100 text-blue-600"
                  : "hover:bg-gray-100"
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
                  <h3 className="font-semibold text-gray-800">
                    Create a filter
                  </h3>
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
                        setFilterValue("");
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

                  {Object.keys(columnFilters).length > 0 && (
                    <div className="border-t pt-3">
                      <h4 className="text-sm font-medium text-gray-700 mb-2">
                        Active Filters:
                      </h4>
                      <div className="space-y-1">
                        {Object.entries(columnFilters).map(
                          ([column, filter]) => (
                            <div
                              key={column}
                              className="flex items-center justify-between bg-gray-50 px-2 py-1 rounded text-sm"
                            >
                              <span>
                                {column}: {filter.type.replace('_', ' ')} {filter.value && `"${filter.value}"`}
                                {filter.values && filter.values.length > 0 && `(${filter.values.length} values)`}
                              </span>
                              <button
                                onClick={() => applyColumnFilter(column, '', '', [])}
                                className="text-red-500 hover:text-red-700"
                              >
                                <X size={12} />
                              </button>
                            </div>
                          )
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

           {/* Right side - Action buttons */}
           <div className="flex items-center space-x-2">
             <button 
               onClick={() => navigate('/dashboard')}
               className="px-3 py-1 bg-blue-600 text-white rounded text-sm font-medium hover:bg-blue-700 flex items-center space-x-1"
             >
               <BarChart size={14} />
               <span>Dashboard</span>
             </button>
             <button 
               onClick={() => navigate('/spreadsheet-dashboard')}
               className="px-3 py-1 bg-green-600 text-white rounded text-sm font-medium hover:bg-green-700 flex items-center space-x-1"
             >
               <FileSpreadsheet size={14} />
               <span>My Sheets</span>
             </button>
            <button 
              onClick={() => navigate('/create-chart')}
              className="px-3 py-1 bg-blue-600 text-white rounded text-sm font-medium hover:bg-blue-700 flex items-center space-x-1"
            >
              <BarChart size={14} />
              <span>Create Chart</span>
            </button>
            <button 
              onClick={handleImportCSV}
              className="px-3 py-1 bg-white text-gray-700 border border-gray-300 rounded text-sm font-medium hover:bg-gray-50 flex items-center space-x-1"
            >
              <svg
                className="w-4 h-4"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                />
              </svg>
              <span>Import CSV</span>
            </button>
          </div>
        </div>
      </div>

      {/* Google Sheets Style Formula Bar */}
      <div className="bg-gray-50 border-b border-gray-200 px-4 py-2">
        <div className="flex items-center space-x-3">
          <div className="flex items-center space-x-2">
            <span className="text-sm font-medium text-gray-700 bg-white px-2 py-1 border border-gray-300 rounded min-w-12 text-center">
              {selectedCell}
            </span>
            <button
              onClick={() => setShowFormulaHelper(!showFormulaHelper)}
              className="p-1 hover:bg-gray-200 rounded text-gray-600"
            >
              <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 24 24">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.94-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z" />
              </svg>
            </button>
          </div>
          <div className="flex-1">
            <div 
              className="flex items-center border border-gray-300 rounded bg-white cursor-text"
              onClick={() => {
                setTimeout(() => {
                  if (formulaBarInputRef.current) {
                    formulaBarInputRef.current.focus();
                    // Don't select all text - just position cursor at end for editing
                    const length = formulaBarInputRef.current.value.length;
                    formulaBarInputRef.current.setSelectionRange(length, length);
                  }
                }, 0);
              }}
            >
                <input
                  ref={formulaBarInputRef}
                  type="text"
                  value={formulaBarValue}
                  onChange={(e) => setFormulaBarValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && selectedCell) {
                      updateCellLocal(selectedCell, formulaBarValue);
                      // Move to next row
                      const match = selectedCell.match(/([A-Z]+)(\d+)/);
                      if (match) {
                        const nextCell = match[1] + (parseInt(match[2]) + 1);
                        handleCellClick(nextCell);
                      }
                    } else if (e.key === "Tab" && selectedCell) {
                      e.preventDefault();
                      updateCellLocal(selectedCell, formulaBarValue);
                      // Move to next column
                      const match = selectedCell.match(/([A-Z]+)(\d+)/);
                      if (match) {
                        const colNum = getColumnNumber(match[1]);
                        const nextCell = getColumnName(colNum + 1) + match[2];
                        handleCellClick(nextCell);
                      }
                    }
                  }}
                  onFocus={() => setIsEditing(true)}
                  onBlur={() => {
                    setTimeout(() => {
                      if (selectedCell) {
                        updateCellLocal(selectedCell, formulaBarValue);
                      }
                      setIsEditing(false);
                    }, 150);
                  }}
                  onMouseUp={(e) => {
                    // Allow normal text selection behavior
                    e.stopPropagation();
                  }}
                  className="flex-1 px-3 py-1 focus:outline-none"
                  placeholder="Enter value or formula (start with =)"
                  autoComplete="off"
                />
              <button className="px-2 py-1 text-gray-400 hover:text-gray-600 border-l border-gray-300">
                <svg
                  className="w-4 h-4"
                  fill="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path d="M9 11H7v6h2v-6zm4 0h-2v6h2v-6zm4 0h-2v6h2v-6zm2-7H5c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 14H5V6h14v12z" />
                </svg>
              </button>
            </div>
          </div>
        </div>

        {showFormulaHelper && (
          <div className="mt-2 p-3 bg-blue-50 border border-blue-200 rounded">
            <div className="text-sm">
              <strong>Formula Examples:</strong>
              <ul className="mt-1 space-y-1">
                <li>• =SUM(A1:A10) - Sum range of cells</li>
                <li>• =PMT(5%/12,60,10000) - Monthly payment calculation</li>
                <li>• =IF(A1 &gt 100, "High", "Low") - Conditional logic</li>
                <li>• =AVERAGE(A1:A5) - Average of range</li>
              </ul>
            </div>
          </div>
        )}
      </div>

      {/* Main Content */}
      <div className="flex-1 flex flex-col">
        {/* Resize Indicator */}
        {showResizeIndicator && isResizing && (
          <div className="fixed top-4 right-4 bg-blue-600 text-white px-3 py-1 rounded shadow-lg z-50 text-sm">
            {resizeType === 'column' 
              ? `Width: ${Math.round(getColumnWidth(resizeTarget))}px`
              : `Height: ${Math.round(getRowHeight(resizeTarget))}px`
            }
          </div>
        )}
        
        {/* Google Sheets Style Spreadsheet */}
        <div className="flex-1 overflow-auto bg-white mb-12" style={{ maxHeight: 'calc(100vh - 200px)' }}>
          <div 
            className="relative"
            style={{ 
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: 'top left',
              width: `${100 / (zoomLevel / 100)}%`,
              height: `${100 / (zoomLevel / 100)}%`
            }}
          >
             <table className="min-w-full border-collapse">
               <thead className="sticky top-0 z-20">
                 <tr className="bg-gray-50">
                   <th className="w-16 h-8 border border-gray-300 bg-gray-100 text-xs font-medium text-gray-600 sticky left-0 z-30" style={{ minWidth: '64px', maxWidth: '64px' }}></th>
                   {Array.from({ length: 40 }, (_, i) => {
                     const columnName = getColumnName(i + 1);
                     const hasFilter = activeFilters[columnName];
                     const hasColumnFilter = columnFilters[columnName];
                     const columnWidth = getColumnWidth(columnName);
                     
                     return (
                       <th
                         key={i}
                         className="border border-gray-300 bg-gray-50 text-xs font-medium text-center text-gray-700 relative hover:bg-gray-100 group"
                         style={{ 
                           minWidth: `${columnWidth}px`,
                           width: `${columnWidth}px`
                         }}
                       >
                         <div className="flex items-center justify-between h-full px-1">
                           <div className="flex items-center justify-center flex-1">
                           {columnName}
                           {hasFilter && (
                             <Filter size={12} className="ml-1 text-blue-600" />
                           )}
                             {hasColumnFilter && (
                               <div className="ml-1 w-2 h-2 bg-blue-500 rounded-full"></div>
                           )}
                         </div>
                           
                           {/* Filter button */}
                           <button
                             onClick={(e) => {
                               e.stopPropagation();
                               setShowFilterDropdown(showFilterDropdown === columnName ? null : columnName);
                             }}
                             className={`opacity-0 group-hover:opacity-100 hover:bg-gray-200 rounded p-1 transition-opacity ${
                               hasColumnFilter ? 'opacity-100' : ''
                             }`}
                             title="Filter"
                           >
                             <Filter size={12} className={hasColumnFilter ? 'text-blue-600' : 'text-gray-600'} />
                           </button>
                           
                           {/* Resize handle */}
                           <div
                             className="absolute right-0 top-0 w-2 h-full cursor-col-resize hover:bg-blue-500 opacity-0 group-hover:opacity-100 transition-opacity"
                             onMouseDown={(e) => {
                               e.preventDefault();
                               startResize('column', columnName, e);
                             }}
                           />
                         </div>
                         
                         {/* Enhanced Filter dropdown */}
                         {showFilterDropdown === columnName && (
                           <div className="absolute top-full left-0 bg-white border border-gray-300 rounded shadow-lg z-50 min-w-64 max-h-80 overflow-y-auto">
                             <div className="p-3">
                               <div className="flex items-center justify-between mb-3">
                                 <h4 className="font-medium text-sm">Filter by {columnName}</h4>
                                 <button
                                   onClick={() => setShowFilterDropdown(null)}
                                   className="text-gray-400 hover:text-gray-600"
                                 >
                                   <X size={14} />
                                 </button>
                               </div>
                               
                               {/* Filter Type Selection */}
                               <div className="mb-3">
                                 <label className="block text-xs font-medium text-gray-700 mb-1">Filter type:</label>
                                 <select
                                   value={columnFilters[columnName]?.type || 'contains'}
                                   onChange={(e) => {
                                     const currentFilter = columnFilters[columnName];
                                     applyColumnFilter(columnName, e.target.value, currentFilter?.value || '', currentFilter?.values || []);
                                   }}
                                   className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                                 >
                                   <option value="contains">Contains</option>
                                   <option value="equals">Equals</option>
                                   <option value="starts_with">Starts with</option>
                                   <option value="ends_with">Ends with</option>
                                   <option value="greater_than">Greater than</option>
                                   <option value="less_than">Less than</option>
                                   <option value="greater_equal">Greater than or equal</option>
                                   <option value="less_equal">Less than or equal</option>
                                   <option value="is_empty">Is empty</option>
                                   <option value="is_not_empty">Is not empty</option>
                                   <option value="is_one_of">Is one of</option>
                                 </select>
                                 </div>
                                 
                               {/* Filter Value Input (for text/number filters) */}
                               {!['is_empty', 'is_not_empty', 'is_one_of'].includes(columnFilters[columnName]?.type || 'contains') && (
                                 <div className="mb-3">
                                   <label className="block text-xs font-medium text-gray-700 mb-1">Filter value:</label>
                                   <input
                                     type="text"
                                     value={columnFilters[columnName]?.value || ''}
                                     onChange={(e) => {
                                       const currentFilter = columnFilters[columnName];
                                       applyColumnFilter(columnName, currentFilter?.type || 'contains', e.target.value, currentFilter?.values || []);
                                     }}
                                     placeholder="Enter filter value..."
                                     className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                                   />
                                 </div>
                               )}

                               {/* Value Selection (for is_one_of filter) */}
                               {columnFilters[columnName]?.type === 'is_one_of' && (
                                 <div className="mb-3">
                                   <div className="text-xs text-gray-600 mb-2">
                                     Select values to include:
                                   </div>
                                   <div className="max-h-32 overflow-y-auto border border-gray-200 rounded p-2">
                                 {getColumnFilterOptions(columnName).map((value, index) => (
                                       <label key={index} className="flex items-center space-x-2 text-sm mb-1">
                                     <input
                                       type="checkbox"
                                       checked={columnFilters[columnName]?.values?.includes(value) || false}
                                       onChange={(e) => {
                                         const currentValues = columnFilters[columnName]?.values || [];
                                         const newValues = e.target.checked
                                           ? [...currentValues, value]
                                           : currentValues.filter(v => v !== value);
                                             applyColumnFilter(columnName, 'is_one_of', '', newValues);
                                       }}
                                       className="rounded"
                                     />
                                         <span className="truncate">{value || '(empty)'}</span>
                                   </label>
                                 ))}
                                   </div>
                                   <div className="flex gap-2 mt-2">
                                   <button
                                     onClick={() => {
                                       const allValues = getColumnFilterOptions(columnName);
                                         applyColumnFilter(columnName, 'is_one_of', '', allValues);
                                     }}
                                     className="text-xs text-blue-600 hover:text-blue-800"
                                   >
                                     Select all
                                   </button>
                                   <button
                                       onClick={() => applyColumnFilter(columnName, 'is_one_of', '', [])}
                                       className="text-xs text-blue-600 hover:text-blue-800"
                                   >
                                     Clear all
                                   </button>
                                 </div>
                                 </div>
                               )}
                               
                               {/* Action Buttons */}
                               <div className="flex gap-2 pt-2 border-t">
                                 <button
                                   onClick={() => {
                                     applyColumnFilter(columnName, columnFilters[columnName]?.type || 'contains', columnFilters[columnName]?.value || '', columnFilters[columnName]?.values || []);
                                     setShowFilterDropdown(null);
                                   }}
                                   className="flex-1 px-3 py-1 bg-blue-600 text-white text-xs rounded hover:bg-blue-700"
                                 >
                                   Apply
                                 </button>
                                 <button
                                   onClick={() => {
                                     applyColumnFilter(columnName, '', '', []);
                                     setShowFilterDropdown(null);
                                   }}
                                   className="px-3 py-1 bg-gray-200 text-gray-700 text-xs rounded hover:bg-gray-300"
                                 >
                                   Clear
                                 </button>
                               </div>
                             </div>
                           </div>
                         )}
                       </th>
                     );
                   })}
                 </tr>
               </thead>
               <tbody>
                 {Array.from({ length: 100 }, (_, rowIndex) => {
                   // Apply row filtering
                   if (!isRowVisible(rowIndex)) return null;

                   return (
                     <tr key={rowIndex} style={{ height: `${getRowHeight(rowIndex + 1)}px` }}>
                       <td className="w-16 border border-gray-300 bg-gray-100 text-xs text-center font-medium sticky left-0 z-10 group" style={{ minWidth: '64px', maxWidth: '64px' }}>
                         <div className="flex items-center justify-center h-full">
                         {rowIndex + 1}
                         </div>
                         
                         {/* Row resize handle */}
                         <div
                           className="absolute bottom-0 left-0 w-full h-2 cursor-row-resize hover:bg-blue-500 opacity-0 group-hover:opacity-100 transition-opacity"
                           onMouseDown={(e) => {
                             e.preventDefault();
                             startResize('row', rowIndex + 1, e);
                           }}
                         />
                       </td>
                      {Array.from({ length: 40 }, (_, colIndex) => {
                        const cellId =
                          getColumnName(colIndex + 1) + (rowIndex + 1);
                        const cellData = data[cellId];
                        const isSelected = selectedCell === cellId;

                        const columnName = getColumnName(colIndex + 1);
                        const columnWidth = getColumnWidth(columnName);
                        const rowHeight = getRowHeight(rowIndex + 1);

                        return (
                          <td
                            key={cellId}
                            className={`border border-gray-200 cursor-cell relative bg-white ${
                              isSelected
                                ? "ring-2 ring-blue-500 bg-blue-50"
                                : "hover:bg-gray-50"
                            }`}
                            style={{
                              width: `${columnWidth}px`,
                              minWidth: `${columnWidth}px`,
                              height: `${rowHeight}px`
                            }}
                            onClick={(e) => {
                              e.preventDefault();
                              e.stopPropagation();
                              handleCellClick(cellId);
                            }}
                            onDoubleClick={(e) => {
                              e.preventDefault();
                              e.stopPropagation();
                              handleCellDoubleClick(cellId);
                            }}
                          >
                            {isSelected && isEditing ? (
                              <input
                                ref={cellInputRef}
                                type="text"
                                value={cellEditValue}
                                onChange={(e) => {
                                  setCellEditValue(e.target.value);
                                  setFormulaBarValue(e.target.value);
                                }}
                                onKeyDown={(e) => {
                                  if (e.key === "Enter") {
                                    e.preventDefault();
                                    updateCellLocal(selectedCell, cellEditValue);
                                    setIsEditing(false);
                                    // Move to next row
                                    const match = selectedCell.match(/([A-Z]+)(\d+)/);
                                    if (match) {
                                      const nextCell = match[1] + (parseInt(match[2]) + 1);
                                      handleCellClick(nextCell);
                                    }
                                  } else if (e.key === "Escape") {
                                    e.preventDefault();
                                    setIsEditing(false);
                                    setCellEditValue(cellData ? cellData.formula || cellData.value || "" : "");
                                  }
                                }}
                                onFocus={() => {
                                  // Ensure we're in editing mode when focused
                                  setIsEditing(true);
                                }}
                                onMouseDown={(e) => {
                                  // Prevent the cell click from interfering with input focus
                                  e.stopPropagation();
                                }}
                                onBlur={(e) => {
                                  // Don't blur if we're navigating between cells
                                  if (isNavigatingRef.current) {
                                    return;
                                  }
                                  
                                  // Check if focus is moving to another cell input
                                  const relatedTarget = e.relatedTarget;
                                  const isMovingToAnotherCell = relatedTarget && relatedTarget.tagName === 'INPUT' && relatedTarget.className.includes('border-none');
                                  
                                  if (!isMovingToAnotherCell) {
                                    // Save immediately without delay to prevent focus issues
                                    updateCellLocal(selectedCell, cellEditValue);
                                    setIsEditing(false);
                                  }
                                }}
                                className="w-full h-full px-1 text-xs border-none outline-none bg-transparent"
                                autoFocus
                              />
                            ) : (
                              <div
                                className="px-1 text-xs truncate"
                                style={getCellStyle(cellId)}
                              >
                                {formatCellValue(cellData)}
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
          </div>
        </div>
        
        {/* Sheet Tabs - Fixed at bottom */}
        <div className="bg-gray-50 border-t px-4 py-2 w-full z-50 fixed bottom-0 right-0">
          <div className="flex items-center space-x-2">
            {sheets.map((sheet) => (
              <button
                key={sheet.id}
                onClick={() => setActiveSheetId(sheet.id)}
                className={`px-3 py-1 text-sm rounded ${
                  activeSheetId === sheet.id
                    ? "bg-white border border-gray-300"
                    : "hover:bg-gray-200"
                }`}
              >
                {sheet.name}
              </button>
            ))}
            <button onClick={addSheet} className="p-1 hover:bg-gray-200 rounded">
              <Plus size={16} />
            </button>
          </div>
        </div>
      </div>

      {/* Chart Modal */}
      {showChartModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-3/4 h-3/4 max-w-4xl">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-semibold">Chart Visualization</h2>
              <div className="flex items-center space-x-4">
                <select
                  value={chartType}
                  onChange={(e) => setChartType(e.target.value)}
                  className="px-3 py-1 border rounded"
                >
                  <option value="line">Line Chart</option>
                  <option value="bar">Bar Chart</option>
                  <option value="pie">Pie Chart</option>
                  <option value="area">Area Chart</option>
                </select>
                <button
                  onClick={() => setShowChartModal(false)}
                  className="p-2 hover:bg-gray-100 rounded"
                >
                  <X size={20} />
                </button>
              </div>
            </div>
            <div className="h-full">
              <ResponsiveContainer width="100%" height="90%">
                {renderChart()}
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}

      {/* Function Prompt Modal */}
      {showFormulaPrompt && selectedFunction && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-lg font-semibold">
                {selectedFunction} Function
              </h3>
              <button
                onClick={() => setShowFormulaPrompt(false)}
                className="p-1 hover:bg-gray-100 rounded"
              >
                <X size={16} />
              </button>
            </div>
            <div className="space-y-3">
              <div className="text-sm text-gray-600">
                {selectedFunction === "PMT" &&
                  "Calculate loan payments: PMT(rate, nper, pv, [fv], [type])"}
                {selectedFunction === "SUM" &&
                  "Sum a range of cells: SUM(range)"}
                {selectedFunction === "IF" &&
                  "Conditional logic: IF(condition, true_value, false_value)"}
              </div>
              <div className="flex space-x-2">
                <button
                  onClick={() => {
                    insertFunction(selectedFunction);
                    setShowFormulaPrompt(false);
                  }}
                  className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
                >
                  Insert
                </button>
                <button
                  onClick={() => setShowFormulaPrompt(false)}
                  className="px-4 py-2 border rounded hover:bg-gray-50"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Hidden CSV File Input */}
      <input
        ref={csvFileInputRef}
        type="file"
        accept=".csv"
        onChange={handleCSVFileUpload}
        className="hidden"
      />
    </div>
  );
};

export default GoogleSheetsClone;
