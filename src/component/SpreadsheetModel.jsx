import React, {
  useState,
  useRef,
  useCallback,
  useMemo,
  useEffect,
} from "react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "../context/AuthContext";
import { useSpreadsheet } from "../context/SpreadsheetContext";
import { useToast } from "../context/ToastContext";
import {
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  WrapText,
  RotateCcw,
  Palette,
  Calculator,
  ChevronDown,
  ChevronRight,
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
  AreaChart,
} from "recharts";

const GoogleSheetsClone = () => {
  const { currentUser, logout } = useAuth();
  const { 
    currentSpreadsheet, 
    updateCell, 
    autoSave, 
    getCurrentSheetData,
    getCellData,
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
        (sheet) => sheet.id === currentSpreadsheet.activeSheetId
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
  const [showFormulaMenu, setShowFormulaMenu] = useState(false);
  const [formulaSearch, setFormulaSearch] = useState("");
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [linkUrl, setLinkUrl] = useState("");
  const [linkText, setLinkText] = useState("");
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
  const [showNumberFormatDropdown, setShowNumberFormatDropdown] =
    useState(false);
  const [formatUpdateTrigger, setFormatUpdateTrigger] = useState(0);
  const [showTextRotationDropdown, setShowTextRotationDropdown] =
    useState(false);

  // Auto-save functionality
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [lastSavedTime, setLastSavedTime] = useState(null);
  const [autoSaveInterval, setAutoSaveInterval] = useState(null);

  // Text rotation options
  const textRotationOptions = [
    { label: "0°", value: 0 },
    { label: "45°", value: 45 },
    { label: "90°", value: 90 },
    { label: "135°", value: 135 },
    { label: "180°", value: 180 },
    { label: "Vertical", value: "vertical" },
  ];
  
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
  const [resizeStartSize, setResizeStartSize] = useState({
    width: 0,
    height: 0,
  });
  const [showResizeIndicator, setShowResizeIndicator] = useState(false);
  const [visibleRows, setVisibleRows] = useState(100); // Virtualization
  const [scrollTop, setScrollTop] = useState(0);
  const [showFormulaHelper, setShowFormulaHelper] = useState(false);
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
    return activeSheet
      ? activeSheet.data instanceof Map
        ? Object.fromEntries(activeSheet.data)
        : activeSheet.data
      : {};
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
        if (cellId === "A1") {
          console.log(`MemoizedCellData for ${cellId}:`, cellData);
          console.log(`Data[${cellId}]:`, data[cellId]);
        }

        result[cellId] = cellData || { value: "", formula: "", style: {} };
      }
    }
    return result;
  }, [data, visibleRows, scrollTop]);

  // Enhanced filter functions for Google Sheets-like behavior
  const applyColumnFilter = useCallback(
    (column, filterType, filterValue, filterValues = []) => {
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
          enabled: true,
        },
      }));
    },
    []
  );

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
          case "contains":
            isVisible = cellValue
              .toLowerCase()
              .includes((filter.value || "").toLowerCase());
            break;
          case "equals":
            isVisible =
              cellValue.toLowerCase() === (filter.value || "").toLowerCase();
            break;
          case "starts_with":
            isVisible = cellValue
              .toLowerCase()
              .startsWith((filter.value || "").toLowerCase());
            break;
          case "ends_with":
            isVisible = cellValue
              .toLowerCase()
              .endsWith((filter.value || "").toLowerCase());
            break;
          case "greater_than":
            isVisible = parseFloat(cellValue) > parseFloat(filter.value || 0);
            break;
          case "less_than":
            isVisible = parseFloat(cellValue) < parseFloat(filter.value || 0);
            break;
          case "greater_equal":
            isVisible = parseFloat(cellValue) >= parseFloat(filter.value || 0);
            break;
          case "less_equal":
            isVisible = parseFloat(cellValue) <= parseFloat(filter.value || 0);
            break;
          case "is_empty":
            isVisible = cellValue === "";
            break;
          case "is_not_empty":
            isVisible = cellValue !== "";
            break;
          case "is_one_of":
            isVisible = filter.values.some(
              (val) => cellValue.toLowerCase() === val.toLowerCase()
            );
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
  const applyFilter = useCallback(
    (column, filterValue) => {
      applyColumnFilter(column, "contains", filterValue);
    },
    [applyColumnFilter]
  );

  const clearAllFilters = () => {
    setColumnFilters({});
    setShowFilterDropdown(null);
    showSuccess("All filters cleared");
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
    showSuccess("Zoomed to fit");
  };

  const handleZoomChange = (newZoom) => {
    setZoomLevel(newZoom);
    showSuccess(`Zoomed to ${newZoom}%`);
    setShowZoomDropdown(false);
  };

  // Font functions
  const fontFamilies = [
    "Arial",
    "Arial Black",
    "Calibri",
    "Cambria",
    "Candara",
    "Comic Sans MS",
    "Consolas",
    "Courier New",
    "Georgia",
    "Helvetica",
    "Impact",
    "Lucida Console",
    "Lucida Sans Unicode",
    "Microsoft Sans Serif",
    "Palatino Linotype",
    "Segoe UI",
    "Tahoma",
    "Times New Roman",
    "Trebuchet MS",
    "Verdana",
  ];

  const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

  const getCurrentFontFamily = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.fontFamily || "Arial";
    }
    return "Arial";
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
    { label: "Automatic", value: "automatic", example: "1,000.12" },
    { label: "Plain text", value: "text", example: "Plain text" },
    { label: "Number", value: "number", example: "1,000.12" },
    { label: "Percent", value: "percent", example: "10.12%" },
    { label: "Scientific", value: "scientific", example: "1.01E+03" },
    { label: "Accounting", value: "accounting", example: "$ (1,000.12)" },
    { label: "Financial", value: "financial", example: "(1,000.12)" },
    { label: "Currency", value: "currency", example: "$1,000.12" },
    { label: "Currency rounded", value: "currency_rounded", example: "$1,000" },
    { label: "Date", value: "date", example: "9/26/2008" },
    { label: "Time", value: "time", example: "3:59:00 PM" },
    { label: "Date time", value: "datetime", example: "9/26/2008 15:59:00" },
    { label: "Duration", value: "duration", example: "24:01:00" },
  ];

  const getCurrentNumberFormat = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      return cellData.style?.numberFormat || "automatic";
    }
    return "automatic";
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
      formatCell({ numberFormat: "currency" });
      showSuccess("Applied currency format");
    }
  };

  const handlePercentageFormat = () => {
    if (selectedCell) {
      formatCell({ numberFormat: "percent" });
      showSuccess("Applied percentage format");
    }
  };

  const handleIncreaseDecimalPlaces = () => {
    if (selectedCell) {
      const currentPlaces = getCurrentDecimalPlaces();
      formatCell({ decimalPlaces: Math.min(currentPlaces + 1, 10) });
      showSuccess(
        `Increased decimal places to ${Math.min(currentPlaces + 1, 10)}`
      );
    }
  };

  // Text formatting handlers
  const handleTextWrap = () => {
    if (selectedCell) {
      const currentCell = getLocalCellData(selectedCell);
      const currentWrap = currentCell.style?.textWrap || "nowrap";
      const newWrap = currentWrap === "wrap" ? "nowrap" : "wrap";
      formatCell({ textWrap: newWrap });
      showSuccess(
        newWrap === "wrap" ? "Text wrapping enabled" : "Text wrapping disabled"
      );

      // Force re-render to update row heights
      setFormatUpdateTrigger((prev) => prev + 1);
    }
  };

  const handleTextRotation = (rotation) => {
    if (selectedCell) {
      formatCell({ textRotation: rotation });
      showSuccess(
        `Text rotated to ${
          rotation === "vertical" ? "vertical" : rotation + "°"
        }`
      );
      setShowTextRotationDropdown(false);
    }
  };

  const handleDecreaseDecimalPlaces = () => {
    if (selectedCell) {
      const currentPlaces = getCurrentDecimalPlaces();
      formatCell({ decimalPlaces: Math.max(currentPlaces - 1, 0) });
      showSuccess(
        `Decreased decimal places to ${Math.max(currentPlaces - 1, 0)}`
      );
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
      showSuccess("You have been logged out successfully.");
      navigate("/");
    } catch (error) {
      console.error("Logout error:", error);
      showError("Failed to logout. Please try again.");
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
      case "file":
        setShowFileMenu(!showFileMenu);
        break;
      case "edit":
        setShowEditMenu(!showEditMenu);
        break;
      case "view":
        setShowViewMenu(!showViewMenu);
        break;
      case "insert":
        setShowInsertMenu(!showInsertMenu);
        break;
      case "format":
        setShowFormatMenu(!showFormatMenu);
        break;
      case "data":
        setShowDataMenu(!showDataMenu);
        break;
    }
    setActiveMenu(menuName);
  };

  // File menu functions
  const handleNewSpreadsheet = () => {
    setShowFileMenu(false);
    showSuccess("Creating new spreadsheet...");
    // Navigate to create new spreadsheet
    navigate("/spreadsheet-dashboard");
  };

  const handleSaveSpreadsheet = async () => {
    setShowFileMenu(false);
    setIsSaving(true);
    try {
      await autoSave();
      setHasUnsavedChanges(false);
      setLastSavedTime(new Date());
      showSuccess("Spreadsheet saved successfully!");
    } catch (error) {
      showError("Failed to save spreadsheet. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleExportCSV = () => {
    setShowFileMenu(false);
    const csvData = convertToCSV();
    const blob = new Blob([csvData], { type: "text/csv" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "spreadsheet.csv";
    a.click();
    window.URL.revokeObjectURL(url);
    showSuccess("CSV exported successfully!");
  };

  const handleImportCSV = () => {
    csvFileInputRef.current?.click();
  };

  const handleCSVFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".csv")) {
      showError("Please select a CSV file");
      return;
    }

    try {
      const text = await file.text();
      const csvData = parseCSV(text);
      importCSVData(csvData);
      showSuccess(
        `CSV imported successfully! ${csvData.length} rows imported.`
      );
    } catch (error) {
      console.error("Error importing CSV:", error);
      showError("Failed to import CSV file. Please check the file format.");
    }

    // Reset the file input
    event.target.value = "";
  };

  const parseCSV = (csvText) => {
    const lines = csvText.split("\n");
    const data = [];

    for (const line of lines) {
      if (line.trim()) {
        // Simple CSV parsing - handles basic cases
        const row = [];
        let current = "";
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
          const char = line[i];

          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === "," && !inQuotes) {
            row.push(current.trim());
            current = "";
          } else {
            current += char;
          }
        }

        // Add the last field
        row.push(current.trim());

        // Remove quotes from fields
        const cleanedRow = row.map((field) => {
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
              style: {},
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
        const sheetData =
          newSheets.find((s) => s.id === activeSheetId)?.data || {};
        autoSave(activeSheetId, sheetData);
      } catch (error) {
        console.error("Error auto-saving imported data:", error);
      }
    }
  };

  // Edit menu functions
  const handleUndo = () => {
    if (historyIndex > 0) {
      setHistoryIndex(historyIndex - 1);
      const previousState = history[historyIndex - 1];
      // Restore previous state
      showSuccess("Undo completed");
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      setHistoryIndex(historyIndex + 1);
      const nextState = history[historyIndex + 1];
      // Restore next state
      showSuccess("Redo completed");
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
  const updateCellLocal = async (
    cellId,
    value,
    formula = null,
    style = null
  ) => {
    console.log(
      `Updating cell ${cellId} with value:`,
      value,
      "formula:",
      formula,
      "style:",
      style
    );
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
          style: style
            ? { ...newData[cellId].style, ...style }
            : newData[cellId].style,
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
    console.log("Updated sheets:", newSheets);
    
    // Mark as having unsaved changes
    setHasUnsavedChanges(true);

    // Update backend if we have a current spreadsheet
    if (currentSpreadsheet) {
      try {
        const cellData = {
          value: formulaValue.startsWith("=") ? evaluatedValue : formulaValue,
          formula: formulaValue,
          style:
            newSheets.find((s) => s.id === activeSheetId)?.data[cellId]
              ?.style || {},
        };

        await updateCell(activeSheetId, cellId, cellData);

        // Auto-save the entire sheet data
        const sheetData =
          newSheets.find((s) => s.id === activeSheetId)?.data || {};
        await autoSave(activeSheetId, sheetData);
      } catch (error) {
        console.error("Error updating cell in backend:", error);
      }
    }
  };

  const handleCut = useCallback(() => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      console.log("Cutting cell data:", cellData);
      
      // Store the cell data in clipboard
      setClipboard(cellData);
      setClipboardType("cut");
      
      // Clear the cell and update formula bar
      updateCellLocal(selectedCell, "");
      setFormulaBarValue("");

      showSuccess("Cell cut to clipboard");
    } else {
      console.warn("No cell selected for cut operation");
      showError("Please select a cell to cut");
    }
  }, [
    selectedCell,
    data,
    updateCellLocal,
    setFormulaBarValue,
    showSuccess,
    showError,
  ]);

  const handleCopy = useCallback(() => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      console.log("Copying cell data:", cellData);
      
      setClipboard(cellData);
      setClipboardType("copy");
      showSuccess("Cell copied to clipboard");
    } else {
      console.warn("No cell selected for copy operation");
      showError("Please select a cell to copy");
    }
  }, [selectedCell, data, showSuccess, showError]);

  const handlePaste = useCallback(() => {
    if (clipboard && selectedCell) {
      console.log("Pasting data:", clipboard);

      // Update both the cell data and formula bar with formula and styles
      const formulaToPaste = clipboard.formula || clipboard.value || "";
      updateCellLocal(
        selectedCell,
        clipboard.value || "",
        formulaToPaste,
        clipboard.style
      );
      
      // Update formula bar to show the pasted content
      setFormulaBarValue(formulaToPaste);
      
      if (clipboardType === "cut") {
        setClipboard(null);
        setClipboardType(null);
      }
      showSuccess("Cell pasted");
    } else if (!clipboard) {
      console.warn("No data in clipboard for paste operation");
      showError("No data to paste. Please copy a cell first.");
    } else if (!selectedCell) {
      console.warn("No cell selected for paste operation");
      showError("Please select a cell to paste into.");
    }
  }, [
    clipboard,
    selectedCell,
    clipboardType,
    updateCellLocal,
    setFormulaBarValue,
    showSuccess,
    showError,
  ]);

  // Insert menu functions
  const handleInsertRow = () => {
    setShowInsertMenu(false);
    showSuccess("Row inserted");
    // Implementation for inserting row
  };

  const handleInsertColumn = () => {
    setShowInsertMenu(false);
    showSuccess("Column inserted");
    // Implementation for inserting column
  };

  // Format menu functions
  const handleBold = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      const newStyle = {
        ...cellData.style,
        fontWeight: cellData.style?.fontWeight === "bold" ? "normal" : "bold",
      };
      // Apply the style directly to the cell
      formatCell({ fontWeight: newStyle.fontWeight });
      showSuccess("Text formatting updated");
    }
  };

  const handleItalic = () => {
    if (selectedCell) {
      const cellData = getLocalCellData(selectedCell);
      const newStyle = {
        ...cellData.style,
        fontStyle: cellData.style?.fontStyle === "italic" ? "normal" : "italic",
      };
      // Apply the style directly to the cell
      formatCell({ fontStyle: newStyle.fontStyle });
      showSuccess("Text formatting updated");
    }
  };

  // Data menu functions
  const handleSortAscending = () => {
    setShowDataMenu(false);
    showSuccess("Data sorted in ascending order");
    // Implementation for sorting
  };

  const handleSortDescending = () => {
    setShowDataMenu(false);
    showSuccess("Data sorted in descending order");
    // Implementation for sorting
  };

  const handleFilter = () => {
    setShowDataMenu(false);
    setShowFilterModal(true);
    showSuccess("Filter applied");
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
    const baseHeight = rowHeights[row] || 24; // Default height

    // Check if any cell in this row has text wrapping enabled
    let maxRequiredHeight = baseHeight;
    for (let col = 1; col <= 40; col++) {
      const cellId = getColumnName(col) + row;
      const cellData = data[cellId];
      if (cellData?.style?.textWrap === "wrap" && cellData?.value) {
        const columnWidth = getColumnWidth(getColumnName(col));
        const text = cellData.value.toString();
        // More accurate height calculation based on character width
        const charsPerLine = Math.floor((columnWidth - 8) / 7); // 7px per char, 8px padding
        const lines = Math.ceil(text.length / charsPerLine);
        const requiredHeight = Math.max(24, lines * 16 + 8); // 16px line height, 8px padding
        maxRequiredHeight = Math.max(maxRequiredHeight, requiredHeight);
      }
    }

    return Math.max(baseHeight, maxRequiredHeight);
  };

  const handleColumnResize = (column, newWidth) => {
    setColumnWidths({
      ...columnWidths,
      [column]: Math.max(50, newWidth), // Minimum width of 50px
    });
  };

  const handleRowResize = (row, newHeight) => {
    setRowHeights({
      ...rowHeights,
      [row]: Math.max(20, newHeight), // Minimum height of 20px
    });
  };

  const startResize = (type, target, event) => {
    setIsResizing(true);
    setResizeType(type);
    setResizeTarget(target);
    setShowResizeIndicator(true);
    
    // Capture initial position and size
    setResizeStartPos({ x: event.clientX, y: event.clientY });
    
    if (type === "column") {
      setResizeStartSize({ width: getColumnWidth(target), height: 0 });
    } else if (type === "row") {
      setResizeStartSize({ width: 0, height: getRowHeight(target) });
    }
  };

  const handleMouseMove = (e) => {
    if (!isResizing || !resizeTarget) return;
    
    if (resizeType === "column") {
      // Calculate delta from start position
      const deltaX = e.clientX - resizeStartPos.x;
      const newWidth = resizeStartSize.width + deltaX;
      handleColumnResize(resizeTarget, Math.max(50, newWidth));
    } else if (resizeType === "row") {
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
        const value = cellData ? cellData.value || "" : "";
        rowData.push(`"${value}"`);
      }
      rows.push(rowData.join(","));
    }
    return rows.join("\n");
  };

  // Handle beforeunload to prompt user to save before refresh
  useEffect(() => {
    const handleBeforeUnload = (event) => {
      if (hasUnsavedChanges) {
        event.preventDefault();
        event.returnValue = 'You have unsaved changes. Are you sure you want to leave?';
        return 'You have unsaved changes. Are you sure you want to leave?';
      }
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [hasUnsavedChanges]);

  // Auto-save functionality
  useEffect(() => {
    if (autoSaveInterval) {
      clearInterval(autoSaveInterval);
    }

    const interval = setInterval(() => {
      if (hasUnsavedChanges && !isSaving) {
        autoSave();
        setHasUnsavedChanges(false);
        setLastSavedTime(new Date());
      }
    }, 30000); // Auto-save every 30 seconds

    setAutoSaveInterval(interval);

    return () => {
      if (interval) {
        clearInterval(interval);
      }
    };
  }, [hasUnsavedChanges, isSaving]);

  // Close all dropdowns and modals when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      // Close filter menu
      if (filterMenuOpen && !event.target.closest(".filter-menu")) {
        setFilterMenuOpen(false);
      }
      // Close zoom dropdown
      if (showZoomDropdown && !event.target.closest(".zoom-dropdown")) {
        setShowZoomDropdown(false);
      }
      // Close font dropdown
      if (showFontDropdown && !event.target.closest(".font-dropdown")) {
        setShowFontDropdown(false);
      }
      // Close font size dropdown
      if (
        showFontSizeDropdown &&
        !event.target.closest(".font-size-dropdown")
      ) {
        setShowFontSizeDropdown(false);
      }
      // Close number format dropdown
      if (
        showNumberFormatDropdown &&
        !event.target.closest(".number-format-dropdown")
      ) {
        setShowNumberFormatDropdown(false);
      }
      // Close text rotation dropdown
      if (
        showTextRotationDropdown &&
        !event.target.closest(".text-rotation-dropdown")
      ) {
        setShowTextRotationDropdown(false);
      }
      // Close color picker
      if (showColorPicker && !event.target.closest(".color-picker-container")) {
        setShowColorPicker(false);
      }
      // Close function menu
      if (showFunctionMenu && !event.target.closest(".function-menu-container")) {
        setShowFunctionMenu(false);
      }
      // Close formula menu
      if (showFormulaMenu && !event.target.closest(".formula-menu-container")) {
        setShowFormulaMenu(false);
      }
      // Close filter dropdowns in column headers
      if (showFilterDropdown && !event.target.closest(".filter-dropdown-container")) {
        setShowFilterDropdown(null);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [
    filterMenuOpen,
    showZoomDropdown,
    showFontDropdown,
    showFontSizeDropdown,
    showNumberFormatDropdown,
    showTextRotationDropdown,
    showColorPicker,
    showFunctionMenu,
    showFormulaMenu,
    showFilterDropdown,
  ]);

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (event) => {
      // Close all menus when pressing Escape
      if (event.key === "Escape") {
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
          case "s":
            event.preventDefault();
            handleSaveSpreadsheet();
            break;
          case "c":
            event.preventDefault();
            console.log("Cmd+C pressed, calling handleCopy");
            handleCopy();
            break;
          case "x":
            event.preventDefault();
            handleCut();
            break;
          case "v":
            event.preventDefault();
            console.log("Cmd+V pressed, calling handlePaste");
            // Use setTimeout to ensure the paste happens after any other processing
            setTimeout(() => {
            handlePaste();
            }, 0);
            break;
          case "z":
            event.preventDefault();
            if (event.shiftKey) {
              handleRedo();
            } else {
              handleUndo();
            }
            break;
          case "y":
            event.preventDefault();
            handleRedo();
            break;
          case "n":
            event.preventDefault();
            handleNewSpreadsheet();
            break;
          case "b":
            event.preventDefault();
            handleBold();
            break;
          case "i":
            event.preventDefault();
            handleItalic();
            break;
          case "=":
          case "+":
            event.preventDefault();
            handleZoomIn();
            break;
          case "-":
            event.preventDefault();
            handleZoomOut();
            break;
          case "0":
            event.preventDefault();
            handleZoomToFit();
            break;
        }
      }

      // Plus/Minus keys for zoom (only with Ctrl/Cmd)
      if (event.ctrlKey || event.metaKey) {
        if (event.key === "+" || event.key === "=") {
          event.preventDefault();
          handleZoomIn();
        } else if (event.key === "-") {
          event.preventDefault();
          handleZoomOut();
        }
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, []); // Remove dependencies to prevent frequent recreation

  // Mouse events for resizing
  useEffect(() => {
    if (isResizing) {
      document.addEventListener("mousemove", handleMouseMove);
      document.addEventListener("mouseup", stopResize);
      
      return () => {
        document.removeEventListener("mousemove", handleMouseMove);
        document.removeEventListener("mouseup", stopResize);
      };
    }
  }, [isResizing, resizeType, resizeTarget]);

  // Enhanced spreadsheet functions with comprehensive banking formulas
  const functions = {
    // Math functions
    SUM: (range) => {
      try {
        console.log("SUM function called with range:", range);
        const values = getRangeValues(range);
        console.log("Range values:", values);
        const numValues = values
          .filter((v) => !isNaN(parseFloat(v)))
          .map((v) => parseFloat(v));
        const result = numValues.reduce((sum, val) => sum + val, 0);
        console.log("SUM result:", result);
        return result;
      } catch (e) {
        console.error("SUM error:", e);
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

    console.log("Evaluating formula:", formula);
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

    // Mark as having unsaved changes
    setHasUnsavedChanges(true);

    // Force re-render by updating trigger
    setFormatUpdateTrigger((prev) => prev + 1);
  };

  const getCellStyle = (cellId) => {
    const cellData = data[cellId];
    if (!cellData || !cellData.style) return {};

    const style = {};
    if (cellData.style.fontFamily) style.fontFamily = cellData.style.fontFamily;
    if (cellData.style.fontSize)
      style.fontSize = `${cellData.style.fontSize}px`;
    if (cellData.style.fontWeight) style.fontWeight = cellData.style.fontWeight;
    if (cellData.style.fontStyle) style.fontStyle = cellData.style.fontStyle;
    if (cellData.style.underline) style.textDecoration = "underline";
    if (cellData.style.textDecoration)
      style.textDecoration = cellData.style.textDecoration;
    if (cellData.style.color) style.color = cellData.style.color;
    if (cellData.style.backgroundColor)
      style.backgroundColor = cellData.style.backgroundColor;
    if (cellData.style.textAlign) style.textAlign = cellData.style.textAlign;
    if (cellData.style.verticalAlign)
      style.verticalAlign = cellData.style.verticalAlign;
    if (cellData.style.textWrap) {
      if (cellData.style.textWrap === "wrap") {
        style.whiteSpace = "pre-wrap";
        style.wordWrap = "break-word";
        style.overflowWrap = "break-word";
        style.maxWidth = "100%";
        style.width = "100%";
      } else {
        style.whiteSpace = "nowrap";
      }
    }
    if (cellData.style.textRotation) {
      const rotation = cellData.style.textRotation;
      if (rotation === "vertical") {
        style.writingMode = "vertical-rl";
        style.textOrientation = "mixed";
      } else if (typeof rotation === "number") {
        style.transform = `rotate(${rotation}deg)`;
        style.transformOrigin = "center";
      }
    }

    return style;
  };

  const formatCellValue = (cellData) => {
    if (
      !cellData ||
      !cellData.style?.numberFormat ||
      cellData.style.numberFormat === "automatic"
    ) {
      return cellData?.value || "";
    }

    const value = cellData.value;
    const numberFormat = cellData.style.numberFormat;
    const decimalPlaces = cellData.style.decimalPlaces || 2;

    // Convert string to number if it's a valid number
    let numericValue = value;
    if (typeof value === "string" && !isNaN(value) && value.trim() !== "") {
      numericValue = parseFloat(value);
    }

    if (typeof numericValue !== "number" || isNaN(numericValue)) {
      return value;
    }

    switch (numberFormat) {
      case "currency":
        return `$${numericValue.toLocaleString(undefined, {
          minimumFractionDigits: decimalPlaces,
          maximumFractionDigits: decimalPlaces,
        })}`;
      case "currency_rounded":
        return `$${numericValue.toLocaleString(undefined, {
          minimumFractionDigits: 0,
          maximumFractionDigits: 0,
        })}`;
      case "percent":
        return `${(numericValue * 100).toLocaleString(undefined, {
          minimumFractionDigits: decimalPlaces,
          maximumFractionDigits: decimalPlaces,
        })}%`;
      case "number":
        return numericValue.toLocaleString(undefined, {
          minimumFractionDigits: decimalPlaces,
          maximumFractionDigits: decimalPlaces,
        });
      case "scientific":
        return numericValue.toExponential(decimalPlaces);
      case "accounting":
        return numericValue < 0
          ? `($${Math.abs(numericValue).toLocaleString(undefined, {
              minimumFractionDigits: decimalPlaces,
              maximumFractionDigits: decimalPlaces,
            })})`
          : `$${numericValue.toLocaleString(undefined, {
              minimumFractionDigits: decimalPlaces,
              maximumFractionDigits: decimalPlaces,
            })}`;
      case "financial":
        return numericValue < 0
          ? `(${Math.abs(numericValue).toLocaleString(undefined, {
              minimumFractionDigits: decimalPlaces,
              maximumFractionDigits: decimalPlaces,
            })})`
          : numericValue.toLocaleString(undefined, {
              minimumFractionDigits: decimalPlaces,
              maximumFractionDigits: decimalPlaces,
            });
      case "date":
        return new Date(numericValue).toLocaleDateString();
      case "time":
        return new Date(numericValue).toLocaleTimeString();
      case "datetime":
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
    const isPrintableKey =
      e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
    
    // Only handle if we're not already editing, not processing, and not in formula bar
    if (
      !isEditing &&
      !isProcessingInput &&
      selectedCell &&
      isPrintableKey &&
      document.activeElement !== formulaBarInputRef.current
    ) {
      // Prevent the default behavior to avoid duplication
      e.preventDefault();
      setIsProcessingInput(true);
      // Start editing when user types a character
      setIsEditing(true);
      const cellData = data[selectedCell];
      const existingValue = cellData
        ? cellData.formula || cellData.value || ""
        : "";
      setCellEditValue(existingValue + e.key);
      setFormulaBarValue(existingValue + e.key);
      setTimeout(() => {
        if (cellInputRef.current) {
          cellInputRef.current.focus();
          cellInputRef.current.setSelectionRange(
            existingValue.length + 1,
            existingValue.length + 1
          );
        }
        setIsProcessingInput(false);
      }, 0);
    }
  };

  // Ensure the main container can receive focus for keyboard events
  useEffect(() => {
    const mainContainer = document.querySelector(".main-spreadsheet-container");
    if (mainContainer) {
      mainContainer.focus();
    }
  }, []);

  // Handle clicks outside the spreadsheet to save current cell
  useEffect(() => {
    const handleClickOutside = (e) => {
      const isInsideSpreadsheet = e.target.closest(
        ".main-spreadsheet-container"
      );
      const isCellInput =
        e.target.tagName === "INPUT" &&
        e.target.className.includes("border-none");
      
      if (!isInsideSpreadsheet && isEditing && selectedCell && !isCellInput) {
        updateCellLocal(selectedCell, cellEditValue);
        setIsEditing(false);
      }
    };

    // Use a longer delay to prevent interference with cell navigation
    const timeoutId = setTimeout(() => {
      document.addEventListener("mousedown", handleClickOutside);
    }, 200);

    return () => {
      clearTimeout(timeoutId);
      document.removeEventListener("mousedown", handleClickOutside);
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
    "Most Used": [
      "SUM", "AVERAGE", "COUNT", "MAX", "MIN"
    ],
    "All": [
      "ABS", "ACCRINT", "ACCRINTM", "ACOS", "ACOSH", "ACOT", "ACOTH", "ADDRESS", "AMORLINC", "AND", "ARABIC", "ARRAY_CONSTRAIN", "ARRAYFORMULA", "ASC", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "AVEDEV", "AVERAGE", "AVERAGE.WEIGHTED", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BAHTTEXT", "BASE", "BETA.DIST", "BETA.INV", "BETADIST", "BETAINV", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BINOM.DIST", "BINOM.INV", "BINOMDIST", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "CHAR", "CHIDIST", "CHIINV", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSE", "CLEAN", "CODE", "COLUMN", "COLUMNS", "COMBIN", "COMBINA", "COMPLEX", "CONCAT", "CONCATENATE", "CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T", "CONVERT", "CORREL", "COS", "COSH", "COT", "COTH", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COUNTUNIQUE", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "COVAR", "COVARIANCE.P", "COVARIANCE.S", "CRITBINOM", "CSC", "CSCH", "CUMIPMT", "CUMPRINC", "DATE", "DATEDIF", "DATEVALUE", "DAVERAGE", "DAY", "DAYS", "DAYS360", "DB", "DCOUNT", "DCOUNTA", "DDB", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DECIMAL", "DEGREES", "DELTA", "DEVSQ", "DGET", "DMAX", "DMIN", "DOLLAR", "DOLLARDE", "DOLLARFR", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DURATION", "DVAR", "DVARP", "EDATE", "EFFECT", "EOMONTH", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE", "ERROR.TYPE", "EUROCONVERT", "EVEN", "EXACT", "EXP", "EXPON.DIST", "EXPONDIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FACT", "FACTDOUBLE", "FALSE", "FDIST", "FILTER", "FIND", "FINV", "FISHER", "FISHERINV", "FIXED", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST", "FORECAST.LINEAR", "FORMULATEXT", "FREQUENCY", "FTEST", "FV", "FVSCHEDULE", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMADIST", "GAMMAINV", "GAMMALN", "GAMMALN.PRECISE", "GAUSS", "GCD", "GEOMEAN", "GESTEP", "GROWTH", "HARMEAN", "HEX2BIN", "HEX2DEC", "HEX2OCT", "HLOOKUP", "HOUR", "HYPERLINK", "HYPGEOM.DIST", "HYPGEOMDIST", "IF", "IFERROR", "IFNA", "IFS", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT", "IMSUB", "IMSUM", "IMTAN", "INDEX", "INDIRECT", "INT", "INTERCEPT", "INTRATE", "IPMT", "IRR", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "ISURL", "KURT", "LARGE", "LCM", "LEFT", "LEN", "LINEST", "LN", "LOG", "LOG10", "LOGEST", "LOGNORM.DIST", "LOGNORM.INV", "LOGNORMDIST", "LOGINV", "LOOKUP", "LOWER", "MATCH", "MAX", "MAXA", "MDETERM", "MDURATION", "MEDIAN", "MID", "MIN", "MINA", "MINUTE", "MINVERSE", "MIRR", "MMULT", "MOD", "MODE", "MODE.MULT", "MODE.SNGL", "MODEMULT", "MODESNGL", "MONTH", "MROUND", "MULTINOMIAL", "MUNIT", "N", "NA", "NEGBINOM.DIST", "NEGBINOMDIST", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOMINAL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NORMDIST", "NORMINV", "NORMSDIST", "NORMSINV", "NOT", "NOW", "NPER", "NPV", "OCT2BIN", "OCT2DEC", "OCT2HEX", "ODD", "OFFSET", "OR", "PDURATION", "PEARSON", "PERCENTILE", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUT", "PERMUTATIONA", "PHI", "PI", "PMT", "POISSON", "POISSON.DIST", "POWER", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PROB", "PRODUCT", "PROPER", "PV", "QUARTILE", "QUARTILE.EXC", "QUARTILE.INC", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN", "RANK", "RANK.AVG", "RANK.EQ", "RATE", "RECEIVED", "REGEXEXTRACT", "REGEXMATCH", "REGEXREPLACE", "REPLACE", "REPT", "RIGHT", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "RRI", "RSQ", "RTD", "SEARCH", "SEC", "SECH", "SECOND", "SERIESSUM", "SHEET", "SHEETS", "SIGN", "SIN", "SINH", "SKEW", "SKEW.P", "SLN", "SLOPE", "SMALL", "SQRT", "SQRTPI", "STANDARDIZE", "STDEV", "STDEV.P", "STDEV.S", "STDEVA", "STDEVP", "STDEVPA", "STEYX", "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "SWITCH", "SYD", "T", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TAN", "TANH", "TDIST", "TEXT", "TIME", "TIMEVALUE", "TINV", "TODAY", "TRANSPOSE", "TREND", "TRIM", "TRIMMEAN", "TRUE", "TRUNC", "TTEST", "TYPE", "UNICHAR", "UNICODE", "UNIQUE", "UPPER", "VALUE", "VAR", "VAR.P", "VAR.S", "VARA", "VARP", "VARPA", "VDB", "VLOOKUP", "WEEKDAY", "WEEKNUM", "WEIBULL", "WEIBULL.DIST", "WORKDAY", "WORKDAY.INTL", "XIRR", "XNPV", "XOR", "YEAR", "YEARFRAC", "YIELD", "YIELDDISC", "YIELDMAT", "Z.TEST", "ZTEST"
    ],
    "Array": [
      "ARRAY_CONSTRAIN", "ARRAYFORMULA", "FILTER", "FLATTEN", "FREQUENCY", "GROWTH", "LINEST", "LOGEST", "MMULT", "MUNIT", "SEQUENCE", "SORT", "SORTN", "SPLIT", "TRANSPOSE", "UNIQUE"
    ],
    "Database": [
      "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP"
    ],
    "Date": [
      "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DAYS360", "EDATE", "EOMONTH", "HOUR", "MINUTE", "MONTH", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOW", "SECOND", "TIME", "TIMEVALUE", "TODAY", "WEEKDAY", "WEEKNUM", "WORKDAY", "WORKDAY.INTL", "YEAR", "YEARFRAC"
    ],
    "Engineering": [
      "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "COMPLEX", "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE", "GESTEP", "HEX2BIN", "HEX2DEC", "HEX2OCT", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT", "IMSUB", "IMSUM", "IMTAN", "OCT2BIN", "OCT2DEC", "OCT2HEX"
    ],
    "Filter": [
      "FILTER", "SORT", "SORTN", "UNIQUE"
    ],
    "Financial": [
      "ACCRINT", "ACCRINTM", "AMORLINC", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "CUMIPMT", "CUMPRINC", "DB", "DDB", "DISC", "DOLLAR", "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT", "FV", "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT", "MDURATION", "MIRR", "NOMINAL", "NPER", "NPV", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "PDURATION", "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PV", "RATE", "RECEIVED", "RRI", "SLN", "SYD", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "VDB", "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT"
    ],
    "Google": [
      "GOOGLEFINANCE", "GOOGLETRANSLATE", "IMAGE"
    ],
    "Info": [
      "CELL", "ERROR.TYPE", "INFO", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "N", "NA", "TYPE"
    ],
    "Logical": [
      "AND", "FALSE", "IF", "IFERROR", "IFNA", "IFS", "NOT", "OR", "SWITCH", "TRUE", "XOR"
    ],
    "Lookup": [
      "ADDRESS", "CHOOSE", "COLUMN", "COLUMNS", "FORMULATEXT", "GETPIVOTDATA", "HLOOKUP", "INDEX", "INDIRECT", "LOOKUP", "MATCH", "OFFSET", "ROW", "ROWS", "RTD", "TRANSPOSE", "VLOOKUP", "XLOOKUP", "XMATCH"
    ],
    "Math": [
      "ABS", "ACOS", "ACOSH", "ACOT", "ACOTH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "COMBIN", "COMBINA", "COS", "COSH", "COT", "COTH", "CSC", "CSCH", "DEGREES", "EVEN", "EXP", "FACT", "FACTDOUBLE", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "GCD", "INT", "LCM", "LN", "LOG", "LOG10", "MOD", "MROUND", "MULTINOMIAL", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEC", "SECH", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC"
    ],
    "Operator": [
      "ADD", "DIVIDE", "EQ", "GT", "GTE", "LT", "LTE", "MINUS", "MULTIPLY", "NE", "POW", "UMINUS", "UPLUS", "UNARY_PERCENT"
    ],
    "Parser": [
      "REGEXEXTRACT", "REGEXMATCH", "REGEXREPLACE"
    ],
    "Statistical": [
      "AVEDEV", "AVERAGE", "AVERAGE.WEIGHTED", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BETA.DIST", "BETA.INV", "BETADIST", "BETAINV", "BINOM.DIST", "BINOM.INV", "BINOMDIST", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "CHAR", "CHIDIST", "CHIINV", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSE", "CLEAN", "CODE", "COLUMN", "COLUMNS", "COMBIN", "COMBINA", "COMPLEX", "CONCAT", "CONCATENATE", "CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T", "CONVERT", "CORREL", "COS", "COSH", "COT", "COTH", "CSC", "CSCH", "DEGREES", "EVEN", "EXP", "FACT", "FACTDOUBLE", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "GCD", "INT", "LCM", "LN", "LOG", "LOG10", "MOD", "MROUND", "MULTINOMIAL", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEC", "SECH", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC"
    ],
    "Text": [
      "ARABIC", "ASC", "BAHTTEXT", "CHAR", "CLEAN", "CODE", "CONCAT", "CONCATENATE", "DOLLAR", "EXACT", "FIND", "FIXED", "LEFT", "LEN", "LOWER", "MID", "NUMBERVALUE", "PROPER", "REGEXEXTRACT", "REGEXMATCH", "REGEXREPLACE", "REPLACE", "REPT", "RIGHT", "ROMAN", "SEARCH", "SPLIT", "SUBSTITUTE", "T", "TEXT", "TRIM", "UNICHAR", "UNICODE", "UPPER", "VALUE"
    ]
  };

  return (
    <div 
      className="flex flex-col h-screen bg-white main-spreadsheet-container" 
      onKeyDown={handleKeyPress}
      tabIndex={0}
    >
      {/* Strix Header */}
      <div className="bg-white border-b border-gray-200 shadow-sm">
        {/* Top Header Bar */}
        <div className="flex items-center justify-between px-4 sm:px-6 py-3 bg-white">
          <div className="flex items-center space-x-4 sm:space-x-6">
            <div className="flex items-center space-x-2 sm:space-x-3">
              <div className="w-8 h-8 bg-gradient-to-r from-indigo-600 to-purple-600 rounded-lg flex items-center justify-center">
                <span className="text-white font-bold text-sm">S</span>
              </div>
              <h1 className="text-lg sm:text-xl font-semibold text-gray-900">
                Strix Sheets
              </h1>
            </div>
            
            {/* Save Status Indicator */}
            <div className="flex items-center space-x-2">
              {isSaving ? (
                <div className="flex items-center space-x-2 text-sm text-blue-600">
                  <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-blue-600"></div>
                  <span>Saving...</span>
                </div>
              ) : hasUnsavedChanges ? (
                <div className="flex items-center space-x-2 text-sm text-orange-600">
                  <div className="w-2 h-2 bg-orange-600 rounded-full"></div>
                  <span>Unsaved changes</span>
                </div>
              ) : (
                <div className="flex items-center space-x-2 text-sm text-green-600">
                  <div className="w-2 h-2 bg-green-600 rounded-full"></div>
                  <span>Saved</span>
                  {lastSavedTime && (
                    <span className="text-xs text-gray-500 ml-1">
                      ({lastSavedTime.toLocaleTimeString()})
                    </span>
                  )}
                </div>
              )}
            </div>
          </div>

          {/* Right side actions */}
          <div className="flex items-center space-x-2 sm:space-x-4">
            {/* Action buttons */}
            <div className="hidden sm:flex items-center space-x-2">
              <button
                onClick={() => navigate('/dashboard')}
                className="flex items-center space-x-2 px-3 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors"
              >
                <BarChart size={16} />
                <span>Dashboard</span>
              </button>
              
              <button
                onClick={() => navigate('/spreadsheet-dashboard')}
                className="flex items-center space-x-2 px-3 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition-colors"
              >
                <FileSpreadsheet size={16} />
                <span>My Sheets</span>
              </button>
              
              <button
                onClick={() => navigate('/create-chart')}
                className="flex items-center space-x-2 px-3 py-2 bg-purple-600 text-white rounded-lg text-sm font-medium hover:bg-purple-700 transition-colors"
              >
                <BarChart size={16} />
                <span>Create Chart</span>
              </button>
              
              <button
                onClick={handleExportCSV}
                className="flex items-center space-x-2 px-3 py-2 bg-gray-600 text-white rounded-lg text-sm font-medium hover:bg-gray-700 transition-colors"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10" />
                </svg>
                <span>Import CSV</span>
              </button>
            </div>
            
            {/* Mobile action buttons */}
            <div className="flex sm:hidden items-center space-x-1">
              <button
                onClick={() => navigate('/dashboard')}
                className="p-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                title="Dashboard"
              >
                <BarChart size={16} />
              </button>
              
              <button
                onClick={() => navigate('/spreadsheet-dashboard')}
                className="p-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
                title="My Sheets"
              >
                <FileSpreadsheet size={16} />
              </button>
              
              <button
                onClick={() => navigate('/create-chart')}
                className="p-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors"
                title="Create Chart"
              >
                <BarChart size={16} />
              </button>
            </div>

            {/* Logout Button */}
            <button
              onClick={handleLogout}
              className="flex items-center space-x-1 sm:space-x-2 px-2 sm:px-4 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700 transition-colors"
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
                  d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1"
                />
              </svg>
              <span className="hidden sm:inline">Logout</span>
            </button>
          </div>
        </div>

        {/* Menu Bar */}
        <div className="flex items-center px-4 sm:px-6 py-2 bg-gray-50 border-b border-gray-200 relative overflow-x-auto">
          <div className="flex items-center space-x-1 min-w-max">
            <div className="relative">
              <button 
                onClick={() => handleMenuClick("file")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showFileMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              File
            </button>
              {showFileMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '20px' }}>
                  <button
                    onClick={handleNewSpreadsheet}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    New Spreadsheet
                  </button>
                  <button
                    onClick={handleSaveSpreadsheet}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Save
                  </button>
                  <button
                    onClick={handleExportCSV}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Export as CSV
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick("edit")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showEditMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              Edit
            </button>
              {showEditMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '80px' }}>
                  <button
                    onClick={handleUndo}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Undo
                  </button>
                  <button
                    onClick={handleRedo}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Redo
                  </button>
                  <div className="border-t border-gray-200"></div>
                  <button
                    onClick={handleCut}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Cut
                  </button>
                  <button
                    onClick={handleCopy}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Copy
                  </button>
                  <button
                    onClick={handlePaste}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Paste
                  </button>
                </div>
              )}
            </div>
            
            <div className="relative">
              <button 
                onClick={() => handleMenuClick("view")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showViewMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              View
            </button>
              {showViewMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '140px' }}>
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
                onClick={() => handleMenuClick("insert")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showInsertMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              Insert
            </button>
              {showInsertMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '200px' }}>
                  <button
                    onClick={handleInsertRow}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Insert Row
                  </button>
                  <button
                    onClick={handleInsertColumn}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
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
                onClick={() => handleMenuClick("format")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showFormatMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              Format
            </button>
              {showFormatMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '260px' }}>
                  <button
                    onClick={handleBold}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Bold
                  </button>
                  <button
                    onClick={handleItalic}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
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
                onClick={() => handleMenuClick("data")}
                className={`text-sm px-3 py-2 rounded-md font-medium transition-colors ${
                  showDataMenu
                    ? "bg-gray-200 text-gray-900"
                    : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"
                }`}
              >
              Data
            </button>
              {showDataMenu && (
                <div className="fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '120px', left: '320px' }}>
                  <button
                    onClick={handleSortAscending}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Sort A to Z
                  </button>
                  <button
                    onClick={handleSortDescending}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Sort Z to A
                  </button>
                  <button
                    onClick={handleFilter}
                    className="block w-full text-left px-4 py-2 hover:bg-gray-100 text-sm"
                  >
                    Filter
                  </button>
                </div>
              )}
            </div>
            
            <button className="text-sm text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-3 py-2 rounded-md font-medium transition-colors">
              Tools
            </button>
            <button className="text-sm text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-3 py-2 rounded-md font-medium transition-colors">
              Extensions
            </button>
            <button className="text-sm text-gray-700 hover:text-gray-900 hover:bg-gray-100 px-3 py-2 rounded-md font-medium transition-colors">
              Help
            </button>
          </div>
        </div>
      </div>

      {/* Google Sheets Style Toolbar */}
      <div className="bg-white border-b border-gray-200 px-4 sm:px-6 py-3 shadow-sm">
        <div className="flex items-center justify-between overflow-x-auto">
          {/* Left side toolbar */}
          <div className="flex items-center space-x-2 min-w-max">
            {/* Undo/Redo */}
            <button
              onClick={handleUndo}
              className="p-2 hover:bg-gray-100 rounded-md transition-colors"
              title="Undo (Ctrl+Z)"
            >
              <Undo size={16} className="text-gray-600" />
            </button>
            <button
              onClick={handleRedo}
              className="p-2 hover:bg-gray-100 rounded-md transition-colors"
              title="Redo (Ctrl+Y)"
            >
              <Redo size={16} className="text-gray-600" />
            </button>

            {/* Save Button */}
            <button 
              onClick={handleSaveSpreadsheet} 
              className={`p-2 hover:bg-gray-100 rounded-md transition-colors ${hasUnsavedChanges ? 'bg-blue-50 text-blue-600' : ''}`}
              title="Save (Ctrl+S)"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4" />
              </svg>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Zoom Controls */}
            <div className="flex items-center space-x-1 bg-gray-50 rounded-lg p-1">
              <button
                onClick={handleZoomOut}
                className="p-1.5 text-gray-600 hover:bg-gray-200 rounded-md transition-colors"
                title="Zoom Out"
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
                    d="M20 12H4"
                  />
                </svg>
              </button>

              <div className="relative">
                <button
                  onClick={() => setShowZoomDropdown(!showZoomDropdown)}
                  className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded-md min-w-16 font-medium transition-colors"
                >
                  {zoomLevel}%
                </button>

                {showZoomDropdown && (
                  <div className="zoom-dropdown fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-32" style={{ top: '180px', left: '20px' }}>
                    {zoomLevels.map((level) => (
                      <button
                        key={level}
                        onClick={() => handleZoomChange(level)}
                        className={`block w-full text-left px-3 py-2 text-sm hover:bg-gray-100 ${
                          level === zoomLevel ? "bg-blue-50 text-blue-600" : ""
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
                className="p-1.5 text-gray-600 hover:bg-gray-200 rounded-md transition-colors"
                title="Zoom In"
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
                    d="M12 6v6m0 0v6m0-6h6m-6 0H6"
                  />
                </svg>
              </button>
            </div>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Number Format Dropdown */}
            <div className="relative">
              <button
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  setShowNumberFormatDropdown(!showNumberFormatDropdown);
                }}
                className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-100 rounded-md min-w-20 text-left font-medium transition-colors"
              >
                {numberFormats.find((f) => f.value === getCurrentNumberFormat())
                  ?.label || "123"}
              </button>

              {showNumberFormatDropdown && (
                <div className="number-format-dropdown fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-64 max-h-80 overflow-y-auto" style={{ top: '180px', left: '200px' }}>
                  <div className="p-2">
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2">
                      General
                    </div>
                    {numberFormats
                      .filter((f) => ["automatic", "text"].includes(f.value))
                      .map((format) => (
                        <button
                          key={format.value}
                          onClick={() => handleNumberFormatChange(format.value)}
                          className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                            format.value === getCurrentNumberFormat()
                              ? "bg-blue-50 text-blue-600"
                              : ""
                          }`}
                        >
                          <div className="flex justify-between items-center">
                            <span>{format.label}</span>
                            <span className="text-xs text-gray-400">
                              {format.example}
                            </span>
                          </div>
                        </button>
                      ))}

                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">
                      Number
                    </div>
                    {numberFormats
                      .filter((f) =>
                        ["number", "percent", "scientific"].includes(f.value)
                      )
                      .map((format) => (
                        <button
                          key={format.value}
                          onClick={() => handleNumberFormatChange(format.value)}
                          className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                            format.value === getCurrentNumberFormat()
                              ? "bg-blue-50 text-blue-600"
                              : ""
                          }`}
                        >
                          <div className="flex justify-between items-center">
                            <span>{format.label}</span>
                            <span className="text-xs text-gray-400">
                              {format.example}
                            </span>
                          </div>
                        </button>
                      ))}

                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">
                      Currency
                    </div>
                    {numberFormats
                      .filter((f) =>
                        [
                          "accounting",
                          "financial",
                          "currency",
                          "currency_rounded",
                        ].includes(f.value)
                      )
                      .map((format) => (
                        <button
                          key={format.value}
                          onClick={() => handleNumberFormatChange(format.value)}
                          className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                            format.value === getCurrentNumberFormat()
                              ? "bg-blue-50 text-blue-600"
                              : ""
                          }`}
                        >
                          <div className="flex justify-between items-center">
                            <span>{format.label}</span>
                            <span className="text-xs text-gray-400">
                              {format.example}
                            </span>
                          </div>
                        </button>
                      ))}

                    <div className="text-xs font-medium text-gray-500 mb-2 px-2 mt-3">
                      Date & time
                    </div>
                    {numberFormats
                      .filter((f) =>
                        ["date", "time", "datetime", "duration"].includes(
                          f.value
                        )
                      )
                      .map((format) => (
                        <button
                          key={format.value}
                          onClick={() => handleNumberFormatChange(format.value)}
                          className={`w-full text-left px-2 py-1 text-sm hover:bg-gray-100 rounded ${
                            format.value === getCurrentNumberFormat()
                              ? "bg-blue-50 text-blue-600"
                              : ""
                          }`}
                        >
                          <div className="flex justify-between items-center">
                            <span>{format.label}</span>
                            <span className="text-xs text-gray-400">
                              {format.example}
                            </span>
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
                getCurrentNumberFormat() === "currency"
                  ? "bg-blue-100 text-blue-600"
                  : "hover:bg-gray-100"
              }`}
              title="Format as currency"
            >
              <DollarSign size={16} className="text-gray-600" />
            </button>

            {/* Percentage Button */}
            <button
              onClick={handlePercentageFormat}
              className={`p-2 rounded ${
                getCurrentNumberFormat() === "percent"
                  ? "bg-blue-100 text-blue-600"
                  : "hover:bg-gray-100"
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
                <svg
                  className="w-3 h-3 text-gray-600 ml-1"
                fill="currentColor"
                  viewBox="0 0 20 20"
                >
                  <path
                    fillRule="evenodd"
                    d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
                    clipRule="evenodd"
                  />
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
                <svg
                  className="w-3 h-3 text-gray-600 ml-1"
                fill="currentColor"
                  viewBox="0 0 20 20"
                >
                  <path
                    fillRule="evenodd"
                    d="M14.707 12.707a1 1 0 01-1.414 0L10 9.414l-3.293 3.293a1 1 0 01-1.414-1.414l4-4a1 1 0 011.414 0l4 4a1 1 0 010 1.414z"
                    clipRule="evenodd"
                  />
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
                    console.log(
                      "Font dropdown clicked, current state:",
                      showFontDropdown
                    );
                    setShowFontDropdown(!showFontDropdown);
                  }}
                  className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-20 text-left"
                  style={{ fontFamily: getCurrentFontFamily() }}
                >
                  {getCurrentFontFamily()}
              </button>

                {showFontDropdown && (
                  <div className="font-dropdown fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48 max-h-64 overflow-y-auto" style={{ top: '180px', left: '400px' }}>
                    {fontFamilies.map((font) => (
                      <button
                        key={font}
                        onClick={() => handleFontFamilyChange(font)}
                        className={`block w-full text-left px-3 py-2 text-sm hover:bg-gray-100 ${
                          font === getCurrentFontFamily()
                            ? "bg-blue-50 text-blue-600"
                            : ""
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
                    console.log(
                      "Font size dropdown clicked, current state:",
                      showFontSizeDropdown
                    );
                    setShowFontSizeDropdown(!showFontSizeDropdown);
                  }}
                  className="px-2 py-1 text-sm text-gray-600 hover:bg-gray-100 rounded min-w-8 text-center"
                >
                  {getCurrentFontSize()}
                  </button>

                {showFontSizeDropdown && (
                  <div className="font-size-dropdown fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-16 max-h-48 overflow-y-auto" style={{ top: '180px', left: '500px' }}>
                    {fontSizes.map((size) => (
                      <button
                        key={size}
                        onClick={() => handleFontSizeChange(size)}
                        className={`block w-full text-center px-3 py-2 text-sm hover:bg-gray-100 ${
                          size === getCurrentFontSize()
                            ? "bg-blue-50 text-blue-600"
                            : ""
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
                  const newStyle = {
                    ...cellData.style,
                    textDecoration:
                      cellData.style?.textDecoration === "line-through"
                        ? "none"
                        : "line-through",
                  };
                  formatCell({ textDecoration: newStyle.textDecoration });
                  showSuccess("Text formatting updated");
                }
              }}
              className="p-2 hover:bg-gray-100 rounded underline text-gray-600"
            >
              <span className="text-sm">S</span>
            </button>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            {/* Alignment */}
            <div className="flex items-center space-x-1">
              <button
                onClick={() => formatCell({ textAlign: "left" })}
                className="p-2 hover:bg-gray-100 rounded"
                title="Align Left"
              >
                <AlignLeft size={16} className="text-gray-600" />
              </button>
              <button
                onClick={() => formatCell({ textAlign: "center" })}
                className="p-2 hover:bg-gray-100 rounded"
                title="Align Center"
              >
                <AlignCenter size={16} className="text-gray-600" />
              </button>
              <button
                onClick={() => formatCell({ textAlign: "right" })}
                className="p-2 hover:bg-gray-100 rounded"
                title="Align Right"
              >
                <AlignRight size={16} className="text-gray-600" />
              </button>
            </div>

            {/* Text Wrapping */}
            <button
              onClick={handleTextWrap}
              className={`p-2 hover:bg-gray-100 rounded ${
                selectedCell &&
                getLocalCellData(selectedCell)?.style?.textWrap === "wrap"
                  ? "bg-blue-100 text-blue-600"
                  : ""
              }`}
              title="Wrap Text"
            >
              <WrapText size={16} className="text-gray-600" />
              </button>

            {/* Text Rotation */}
            <div className="relative">
              <button
                onClick={() =>
                  setShowTextRotationDropdown(!showTextRotationDropdown)
                }
                className="p-2 hover:bg-gray-100 rounded flex items-center"
                title="Text Rotation"
              >
                <RotateCcw size={16} className="text-gray-600" />
                <ChevronDown size={12} className="ml-1" />
            </button>

              {showTextRotationDropdown && (
                <div className="text-rotation-dropdown fixed bg-white border border-gray-200 rounded shadow-lg z-[9999] min-w-48" style={{ top: '180px', left: '600px' }}>
                  <div className="p-2">
                    <div className="text-xs font-medium text-gray-500 mb-2 px-2">
                      Text Rotation
                    </div>
                    {textRotationOptions.map((option) => (
              <button
                        key={option.value}
                        onClick={() => handleTextRotation(option.value)}
                        className="block w-full text-left px-2 py-1 hover:bg-gray-100 text-sm"
                      >
                        {option.label}
              </button>
                    ))}
            </div>
                </div>
              )}
          </div>

          {/* Formula Icon (Σ) */}
          <div className="relative formula-menu-container">
            <button
              onClick={() => setShowFormulaMenu(!showFormulaMenu)}
              className="p-2 hover:bg-gray-100 rounded flex items-center"
              title="Functions"
            >
              <span className="text-lg font-bold">Σ</span>
            </button>

            {showFormulaMenu && (
              <div className="fixed w-80 bg-white border rounded-lg shadow-lg z-[9999] max-h-96 overflow-y-auto" style={{ top: '180px', left: '700px' }}>
                <div className="p-2 border-b">
                  <input
                    type="text"
                    placeholder="Search functions..."
                    className="w-full px-2 py-1 border rounded text-sm"
                    value={formulaSearch}
                    onChange={(e) => setFormulaSearch(e.target.value)}
                  />
                </div>
                <div className="max-h-80 overflow-y-auto">
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

          {/* Link Icon */}
          <div className="relative">
            <button
              onClick={() => setShowLinkModal(true)}
              className="p-2 hover:bg-gray-100 rounded flex items-center"
              title="Insert link"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
              </svg>
            </button>
          </div>

            {/* Separator */}
            <div className="w-px h-6 bg-gray-300 mx-2"></div>

            <div className="flex flex-wrap items-start space-x-2 justify-start">
              <div className="relative function-menu-container">
            <button
              onClick={() => setShowFunctionMenu(!showFunctionMenu)}
              className="p-2 hover:bg-gray-100 rounded flex items-center"
            >
              <Calculator size={16} />
              <ChevronDown size={12} className="ml-1" />
            </button>

            {showFunctionMenu && (
              <div className="fixed w-48 bg-white border rounded-lg shadow-lg z-[9999]" style={{ top: '180px', left: '900px' }}>
                <div className="py-1">
                            <button
                    onClick={() => {
                      setShowFunctionMenu(false);
                      showSuccess("Create group by view");
                    }}
                    className="flex items-center w-full px-3 py-2 text-sm hover:bg-gray-100 text-left"
                  >
                    <span className="mr-2">+</span>
                    Create group by view
                    <ChevronRight size={12} className="ml-auto" />
                            </button>
                  <button
                    onClick={() => {
                      setShowFunctionMenu(false);
                      showSuccess("Create filter view");
                    }}
                    className="flex items-center w-full px-3 py-2 text-sm hover:bg-gray-100 text-left"
                  >
                    <span className="mr-2">+</span>
                    Create filter view
                  </button>
                </div>
              </div>
            )}
          </div>
              <div className="relative color-picker-container">
            <button
              onClick={() => setShowColorPicker(!showColorPicker)}
              className="p-2 hover:bg-gray-100 rounded"
            >
              <Palette size={16} />
            </button>

            {showColorPicker && (
              <div className="fixed w-48 bg-white border rounded-lg shadow-lg z-[9999] p-3" style={{ top: '180px', left: '1000px' }}>
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
              <div className="relative filter-menu-container">
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
              <div className="filter-menu fixed w-80 bg-white border rounded-lg shadow-lg z-[9999] p-4" style={{ top: '180px', right: '20px' }}>
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
                                    {column}: {filter.type.replace("_", " ")}{" "}
                                    {filter.value && `"${filter.value}"`}
                                    {filter.values &&
                                      filter.values.length > 0 &&
                                      `(${filter.values.length} values)`}
                              </span>
                              <button
                                    onClick={() =>
                                      applyColumnFilter(column, "", "", [])
                                    }
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
            </div>
          </div>
        </div>
      </div>

      {/* Google Sheets Style Formula Bar */}
      <div className="bg-white border-b border-gray-200 flex-shrink-0 shadow-sm">
        <div className="flex items-center h-12 px-4 sm:px-6">
          {/* Cell reference */}
          <div className="flex items-center border border-gray-300 bg-gray-50 rounded-l-md">
            <span className="text-sm font-medium text-gray-700 px-4 py-2 min-w-20 text-center border-r border-gray-300">
              {selectedCell}
            </span>
            <button
              onClick={() => setShowFormulaHelper(!showFormulaHelper)}
              className="p-2 hover:bg-gray-200 text-gray-600 border-r border-gray-300 transition-colors"
              title="Functions"
            >
              <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 24 24">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.94-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z" />
              </svg>
            </button>
          </div>
          
          {/* Formula input area */}
          <div className="flex-1 flex items-center">
            <div 
              className="flex-1 flex items-center border border-gray-300 border-l-0 bg-white cursor-text h-full rounded-r-md"
              onClick={() => {
                setTimeout(() => {
                  if (formulaBarInputRef.current) {
                    formulaBarInputRef.current.focus();
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
                    console.log("Enter pressed, updating cell with:", formulaBarValue);
                      updateCellLocal(selectedCell, formulaBarValue);
                      const match = selectedCell.match(/([A-Z]+)(\d+)/);
                      if (match) {
                        const nextCell = match[1] + (parseInt(match[2]) + 1);
                        handleCellClick(nextCell);
                      }
                    } else if (e.key === "Tab" && selectedCell) {
                      e.preventDefault();
                    console.log("Tab pressed, updating cell with:", formulaBarValue);
                      updateCellLocal(selectedCell, formulaBarValue);
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
                      console.log("Formula bar blur, updating cell with:", formulaBarValue);
                        updateCellLocal(selectedCell, formulaBarValue);
                      }
                      setIsEditing(false);
                    }, 150);
                  }}
                onMouseUp={(e) => e.stopPropagation()}
                className="flex-1 px-4 py-2 focus:outline-none text-sm h-full border-none placeholder-gray-500"
                  placeholder="Enter value or formula (start with =)"
                  autoComplete="off"
                />
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
      <div className="flex-1 flex flex-col bg-gray-50">
        {/* Resize Indicator */}
        {showResizeIndicator && isResizing && (
          <div className="fixed top-4 right-4 bg-blue-600 text-white px-3 py-1 rounded-lg shadow-lg z-[9999] text-sm font-medium">
            {resizeType === "column"
              ? `Width: ${Math.round(getColumnWidth(resizeTarget))}px`
              : `Height: ${Math.round(getRowHeight(resizeTarget))}px`}
          </div>
        )}
        
        {/* Google Sheets Style Spreadsheet */}
        <div
          className="flex-1 overflow-auto bg-white shadow-inner mb-10"
          style={{ maxHeight: "calc(100vh - 200px)" }}
        >
          <div
            className="relative"
            style={{
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: "top left",
              width: `${100 / (zoomLevel / 100)}%`,
              height: `${100 / (zoomLevel / 100)}%`,
            }}
          >
             <table className="min-w-full border-collapse">
               <thead className="sticky top-0 z-10">
                 <tr className="bg-gray-50">
                  <th
                    className="w-16 h-8 border border-gray-300 bg-gray-100 text-xs font-medium text-gray-600 sticky left-0 z-10"
                    style={{ minWidth: "64px", maxWidth: "64px" }}
                  ></th>
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
                          width: `${columnWidth}px`,
                         }}
                       >
                         <div className="flex items-center justify-between h-full px-1">
                           <div className="flex items-center justify-center flex-1">
                           {columnName}
                           {hasFilter && (
                              <Filter
                                size={12}
                                className="ml-1 text-blue-600"
                              />
                           )}
                             {hasColumnFilter && (
                               <div className="ml-1 w-2 h-2 bg-blue-500 rounded-full"></div>
                           )}
                         </div>
                           
                           {/* Filter button */}
                           <button
                             onClick={(e) => {
                               e.stopPropagation();
                              setShowFilterDropdown(
                                showFilterDropdown === columnName
                                  ? null
                                  : columnName
                              );
                            }}
                            className={`opacity-0 group-hover:opacity-100 hover:bg-gray-200 rounded p-1 transition-opacity ${
                              hasColumnFilter ? "opacity-100" : ""
                            }`}
                             title="Filter"
                           >
                            <Filter
                              size={12}
                              className={
                                hasColumnFilter
                                  ? "text-blue-600"
                                  : "text-gray-600"
                              }
                            />
                           </button>
                           
                           {/* Resize handle */}
                           <div
                            className="absolute right-0 top-0 w-2 h-full cursor-col-resize opacity-0 group-hover:opacity-100 transition-opacity z-10 border-l border-gray-300 bg-gray-200 hover:bg-blue-500"
                             onMouseDown={(e) => {
                               e.preventDefault();
                              e.stopPropagation();
                              startResize("column", columnName, e);
                             }}
                           />
                         </div>
                         
                        {/* Enhanced Filter dropdown */}
                         {showFilterDropdown === columnName && (
                          <div className="filter-dropdown-container fixed bg-white border border-gray-300 rounded shadow-lg z-[9999] min-w-64 max-h-80 overflow-y-auto" style={{ top: '250px', left: '20px' }}>
                             <div className="p-3">
                              <div className="flex items-center justify-between mb-3">
                                <h4 className="font-medium text-sm">
                                  Filter by {columnName}
                                </h4>
                                 <button
                                   onClick={() => setShowFilterDropdown(null)}
                                   className="text-gray-400 hover:text-gray-600"
                                 >
                                   <X size={14} />
                                 </button>
                               </div>
                               
                              {/* Filter Type Selection */}
                              <div className="mb-3">
                                <label className="block text-xs font-medium text-gray-700 mb-1">
                                  Filter type:
                                </label>
                                <select
                                  value={
                                    columnFilters[columnName]?.type ||
                                    "contains"
                                  }
                                  onChange={(e) => {
                                    const currentFilter =
                                      columnFilters[columnName];
                                    applyColumnFilter(
                                      columnName,
                                      e.target.value,
                                      currentFilter?.value || "",
                                      currentFilter?.values || []
                                    );
                                  }}
                                  className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                                >
                                  <option value="contains">Contains</option>
                                  <option value="equals">Equals</option>
                                  <option value="starts_with">
                                    Starts with
                                  </option>
                                  <option value="ends_with">Ends with</option>
                                  <option value="greater_than">
                                    Greater than
                                  </option>
                                  <option value="less_than">Less than</option>
                                  <option value="greater_equal">
                                    Greater than or equal
                                  </option>
                                  <option value="less_equal">
                                    Less than or equal
                                  </option>
                                  <option value="is_empty">Is empty</option>
                                  <option value="is_not_empty">
                                    Is not empty
                                  </option>
                                  <option value="is_one_of">Is one of</option>
                                </select>
                                 </div>
                                 
                              {/* Filter Value Input (for text/number filters) */}
                              {![
                                "is_empty",
                                "is_not_empty",
                                "is_one_of",
                              ].includes(
                                columnFilters[columnName]?.type || "contains"
                              ) && (
                                <div className="mb-3">
                                  <label className="block text-xs font-medium text-gray-700 mb-1">
                                    Filter value:
                                  </label>
                                  <input
                                    type="text"
                                    value={
                                      columnFilters[columnName]?.value || ""
                                    }
                                    onChange={(e) => {
                                      const currentFilter =
                                        columnFilters[columnName];
                                      applyColumnFilter(
                                        columnName,
                                        currentFilter?.type || "contains",
                                        e.target.value,
                                        currentFilter?.values || []
                                      );
                                    }}
                                    placeholder="Enter filter value..."
                                    className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                                  />
                                </div>
                              )}

                              {/* Value Selection (for is_one_of filter) */}
                              {columnFilters[columnName]?.type ===
                                "is_one_of" && (
                                <div className="mb-3">
                                  <div className="text-xs text-gray-600 mb-2">
                                    Select values to include:
                                  </div>
                                  <div className="max-h-32 overflow-y-auto border border-gray-200 rounded p-2">
                                    {getColumnFilterOptions(columnName).map(
                                      (value, index) => (
                                        <label
                                          key={index}
                                          className="flex items-center space-x-2 text-sm mb-1"
                                        >
                                     <input
                                       type="checkbox"
                                            checked={
                                              columnFilters[
                                                columnName
                                              ]?.values?.includes(value) ||
                                              false
                                            }
                                       onChange={(e) => {
                                              const currentValues =
                                                columnFilters[columnName]
                                                  ?.values || [];
                                         const newValues = e.target.checked
                                           ? [...currentValues, value]
                                                : currentValues.filter(
                                                    (v) => v !== value
                                                  );
                                              applyColumnFilter(
                                                columnName,
                                                "is_one_of",
                                                "",
                                                newValues
                                              );
                                       }}
                                       className="rounded"
                                     />
                                          <span className="truncate">
                                            {value || "(empty)"}
                                          </span>
                                   </label>
                                      )
                                    )}
                                  </div>
                                  <div className="flex gap-2 mt-2">
                                   <button
                                     onClick={() => {
                                        const allValues =
                                          getColumnFilterOptions(columnName);
                                        applyColumnFilter(
                                          columnName,
                                          "is_one_of",
                                          "",
                                          allValues
                                        );
                                     }}
                                     className="text-xs text-blue-600 hover:text-blue-800"
                                   >
                                     Select all
                                   </button>
                                   <button
                                      onClick={() =>
                                        applyColumnFilter(
                                          columnName,
                                          "is_one_of",
                                          "",
                                          []
                                        )
                                      }
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
                                    applyColumnFilter(
                                      columnName,
                                      columnFilters[columnName]?.type ||
                                        "contains",
                                      columnFilters[columnName]?.value || "",
                                      columnFilters[columnName]?.values || []
                                    );
                                    setShowFilterDropdown(null);
                                  }}
                                  className="flex-1 px-3 py-1 bg-blue-600 text-white text-xs rounded hover:bg-blue-700"
                                >
                                  Apply
                                </button>
                                <button
                                  onClick={() => {
                                    applyColumnFilter(columnName, "", "", []);
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
                    <tr
                      key={rowIndex}
                      style={{ height: `${getRowHeight(rowIndex + 1)}px` }}
                    >
                      <td
                        className="w-16 border border-gray-300 bg-gray-100 text-xs text-center font-medium sticky left-0 z-5 group"
                        style={{ minWidth: "64px", maxWidth: "64px" }}
                      >
                         <div className="flex items-center justify-center h-full">
                         {rowIndex + 1}
                         </div>
                         
                         {/* Row resize handle */}
                         <div
                          className="absolute bottom-0 left-0 w-full h-2 cursor-row-resize opacity-0 group-hover:opacity-100 transition-opacity z-10 border-t border-gray-300 bg-gray-200 hover:bg-blue-500"
                           onMouseDown={(e) => {
                             e.preventDefault();
                            e.stopPropagation();
                            startResize("row", rowIndex + 1, e);
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
                              maxWidth: `${columnWidth}px`,
                              height: `${rowHeight}px`,
                              tableLayout: "fixed",
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
                                    updateCellLocal(
                                      selectedCell,
                                      cellEditValue
                                    );
                                    setIsEditing(false);
                                    // Move to next row
                                    const match =
                                      selectedCell.match(/([A-Z]+)(\d+)/);
                                    if (match) {
                                      const nextCell =
                                        match[1] + (parseInt(match[2]) + 1);
                                      handleCellClick(nextCell);
                                    }
                                  } else if (e.key === "Escape") {
                                    e.preventDefault();
                                    setIsEditing(false);
                                    setCellEditValue(
                                      cellData
                                        ? cellData.formula ||
                                            cellData.value ||
                                            ""
                                        : ""
                                    );
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
                                  const isMovingToAnotherCell =
                                    relatedTarget &&
                                    relatedTarget.tagName === "INPUT" &&
                                    relatedTarget.className.includes(
                                      "border-none"
                                    );
                                  
                                  if (!isMovingToAnotherCell) {
                                    // Save immediately without delay to prevent focus issues
                                    updateCellLocal(
                                      selectedCell,
                                      cellEditValue
                                    );
                                    setIsEditing(false);
                                  }
                                }}
                                className="w-full h-full px-1 text-xs border-none outline-none bg-transparent"
                                autoFocus
                              />
                            ) : (
                              <div
                                className={`px-1 text-xs ${
                                  cellData?.style?.textWrap === "wrap"
                                    ? "break-words"
                                    : "truncate"
                                }`}
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
        <div className="bg-gray-100 border-t border-gray-200 px-4 sm:px-6 py-1 w-full z-[9999] fixed bottom-0 right-0 shadow-lg">
          <div className="flex items-center space-x-1 overflow-x-auto">
            {sheets.map((sheet) => (
              <button
                key={sheet.id}
                onClick={() => setActiveSheetId(sheet.id)}
                className={`px-4 py-1 text-sm rounded-lg font-medium transition-colors ${
                  activeSheetId === sheet.id
                    ? "bg-white border border-gray-300 shadow-sm text-gray-900"
                    : "hover:bg-gray-200 text-gray-600"
                }`}
              >
                {sheet.name}
              </button>
            ))}
            <button
              onClick={addSheet}
              className="p-2 hover:bg-gray-200 rounded-lg transition-colors text-gray-600"
              title="Add new sheet"
            >
              <Plus size={16} />
            </button>
          </div>
        </div>
      </div>

      {/* Chart Modal */}
      {showChartModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[9999]">
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
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[9999]">
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

      {/* Link Modal */}
      {showLinkModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[9999]">
          <div className="bg-white rounded-lg p-6 w-96">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-semibold">Insert Link</h2>
              <button
                onClick={() => setShowLinkModal(false)}
                className="p-2 hover:bg-gray-100 rounded"
              >
                <X size={20} />
              </button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Link URL
                </label>
                <input
                  type="url"
                  value={linkUrl}
                  onChange={(e) => setLinkUrl(e.target.value)}
                  placeholder="https://example.com"
                  className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Link Text (optional)
                </label>
                <input
                  type="text"
                  value={linkText}
                  onChange={(e) => setLinkText(e.target.value)}
                  placeholder="Display text"
                  className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
              <div className="flex justify-end space-x-2">
                <button
                  onClick={() => setShowLinkModal(false)}
                  className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded"
                >
                  Cancel
                </button>
                <button
                  onClick={() => {
                    if (linkUrl && selectedCell) {
                      const linkValue = linkText ? `=HYPERLINK("${linkUrl}", "${linkText}")` : `=HYPERLINK("${linkUrl}")`;
                      updateCellLocal(selectedCell, linkValue);
                      setShowLinkModal(false);
                      setLinkUrl("");
                      setLinkText("");
                      showSuccess("Link inserted");
                    }
                  }}
                  className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
                >
                  Insert Link
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
