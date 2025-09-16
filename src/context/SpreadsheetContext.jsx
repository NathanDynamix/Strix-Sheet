import React, { createContext, useContext, useState, useEffect, useCallback } from 'react';
import { useAuth } from './AuthContext';
import spreadsheetService from '../services/spreadsheetService';

const SpreadsheetContext = createContext();

export function useSpreadsheet() {
  return useContext(SpreadsheetContext);
}

export function SpreadsheetProvider({ children }) {
  const { currentUser } = useAuth();
  const [spreadsheets, setSpreadsheets] = useState([]);
  const [currentSpreadsheet, setCurrentSpreadsheet] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [autoSaveTimeout, setAutoSaveTimeout] = useState(null);

  // Load all spreadsheets for the current user
  const loadSpreadsheets = useCallback(async () => {
    if (!currentUser) return;

    try {
      setLoading(true);
      setError(null);
      const response = await spreadsheetService.getSpreadsheets();
      setSpreadsheets(response.data || []);
    } catch (err) {
      console.error('Error loading spreadsheets:', err);
      setError('Failed to load spreadsheets');
    } finally {
      setLoading(false);
    }
  }, [currentUser]);

  // Create a new spreadsheet
  const createSpreadsheet = async (title, description = '') => {
    if (!currentUser) {
      throw new Error('User must be logged in to create spreadsheets');
    }

    try {
      setLoading(true);
      setError(null);
      const response = await spreadsheetService.createSpreadsheet(title, description);
      const newSpreadsheet = response.data;
      
      setSpreadsheets(prev => [newSpreadsheet, ...prev]);
      setCurrentSpreadsheet(newSpreadsheet);
      
      return newSpreadsheet;
    } catch (err) {
      console.error('Error creating spreadsheet:', err);
      setError('Failed to create spreadsheet');
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // Load a specific spreadsheet
  const loadSpreadsheet = async (id) => {
    if (!currentUser) {
      throw new Error('User must be logged in to load spreadsheets');
    }

    try {
      setLoading(true);
      setError(null);
      const response = await spreadsheetService.getSpreadsheet(id);
      const spreadsheet = response.data;
      
      setCurrentSpreadsheet(spreadsheet);
      return spreadsheet;
    } catch (err) {
      console.error('Error loading spreadsheet:', err);
      setError('Failed to load spreadsheet');
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // Update spreadsheet metadata
  const updateSpreadsheet = async (id, updates) => {
    try {
      setError(null);
      const response = await spreadsheetService.updateSpreadsheet(id, updates);
      const updatedSpreadsheet = response.data;
      
      setSpreadsheets(prev => 
        prev.map(s => s._id === id ? updatedSpreadsheet : s)
      );
      
      if (currentSpreadsheet && currentSpreadsheet._id === id) {
        setCurrentSpreadsheet(updatedSpreadsheet);
      }
      
      return updatedSpreadsheet;
    } catch (err) {
      console.error('Error updating spreadsheet:', err);
      setError('Failed to update spreadsheet');
      throw err;
    }
  };

  // Delete a spreadsheet
  const deleteSpreadsheet = async (id) => {
    try {
      setError(null);
      await spreadsheetService.deleteSpreadsheet(id);
      
      setSpreadsheets(prev => prev.filter(s => s._id !== id));
      
      if (currentSpreadsheet && currentSpreadsheet._id === id) {
        setCurrentSpreadsheet(null);
      }
    } catch (err) {
      console.error('Error deleting spreadsheet:', err);
      setError('Failed to delete spreadsheet');
      throw err;
    }
  };

  // Update a cell
  const updateCell = async (sheetId, cellId, cellData) => {
    if (!currentSpreadsheet) return;

    try {
      setError(null);
      await spreadsheetService.updateCell(
        currentSpreadsheet._id, 
        sheetId, 
        cellId, 
        cellData
      );

      // Update local state
      setCurrentSpreadsheet(prev => {
        if (!prev) return prev;
        
        const updatedSheets = prev.sheets.map(sheet => {
          if (sheet.id === sheetId) {
            const updatedData = new Map(sheet.data);
            updatedData.set(cellId, cellData);
            return { ...sheet, data: updatedData };
          }
          return sheet;
        });

        return {
          ...prev,
          sheets: updatedSheets,
          lastModified: new Date(),
          version: prev.version + 1
        };
      });
    } catch (err) {
      console.error('Error updating cell:', err);
      setError('Failed to update cell');
      throw err;
    }
  };

  // Auto-save functionality
  const autoSave = useCallback(async (sheetId, data) => {
    if (!currentSpreadsheet || !data || Object.keys(data).length === 0) {
      return;
    }

    // Clear existing timeout
    if (autoSaveTimeout) {
      clearTimeout(autoSaveTimeout);
    }

    // Set new timeout for auto-save
    const timeout = setTimeout(async () => {
      try {
        await spreadsheetService.autoSave(currentSpreadsheet._id, sheetId, data);
      } catch (err) {
        console.error('Auto-save failed:', err);
      }
    }, 2000); // Auto-save after 2 seconds of inactivity

    setAutoSaveTimeout(timeout);
  }, [currentSpreadsheet, autoSaveTimeout]);

  // Add a new sheet
  const addSheet = async (name) => {
    if (!currentSpreadsheet) return;

    try {
      setError(null);
      const response = await spreadsheetService.addSheet(currentSpreadsheet._id, name);
      const newSheet = response.data;

      setCurrentSpreadsheet(prev => {
        if (!prev) return prev;
        return {
          ...prev,
          sheets: [...prev.sheets, newSheet],
          activeSheetId: newSheet.id
        };
      });

      return newSheet;
    } catch (err) {
      console.error('Error adding sheet:', err);
      setError('Failed to add sheet');
      throw err;
    }
  };

  // Update sheet data in bulk
  const updateSheetData = async (sheetId, data) => {
    if (!currentSpreadsheet) return;

    try {
      setError(null);
      await spreadsheetService.updateSheetData(
        currentSpreadsheet._id, 
        sheetId, 
        data
      );

      // Update local state
      setCurrentSpreadsheet(prev => {
        if (!prev) return prev;
        
        const updatedSheets = prev.sheets.map(sheet => {
          if (sheet.id === sheetId) {
            return { ...sheet, data: new Map(Object.entries(data)) };
          }
          return sheet;
        });

        return {
          ...prev,
          sheets: updatedSheets,
          lastModified: new Date(),
          version: prev.version + 1
        };
      });
    } catch (err) {
      console.error('Error updating sheet data:', err);
      setError('Failed to update sheet data');
      throw err;
    }
  };

  // Share spreadsheet
  const shareSpreadsheet = async (collaborators) => {
    if (!currentSpreadsheet) return;

    try {
      setError(null);
      const response = await spreadsheetService.shareSpreadsheet(
        currentSpreadsheet._id, 
        collaborators
      );

      const updatedCollaborators = response.data;
      
      setCurrentSpreadsheet(prev => {
        if (!prev) return prev;
        return {
          ...prev,
          collaborators: updatedCollaborators
        };
      });

      return updatedCollaborators;
    } catch (err) {
      console.error('Error sharing spreadsheet:', err);
      setError('Failed to share spreadsheet');
      throw err;
    }
  };

  // Get current sheet data
  const getCurrentSheetData = () => {
    if (!currentSpreadsheet) return {};
    
    const activeSheet = currentSpreadsheet.sheets.find(
      sheet => sheet.id === currentSpreadsheet.activeSheetId
    );
    
    if (!activeSheet || !activeSheet.data) return {};
    
    // Convert Map to object for easier handling
    return Object.fromEntries(activeSheet.data);
  };

  // Get cell data
  const getCellData = (sheetId, cellId) => {
    if (!currentSpreadsheet) return { value: "", formula: "", style: {} };
    
    const sheet = currentSpreadsheet.sheets.find(s => s.id === sheetId);
    if (!sheet || !sheet.data) return { value: "", formula: "", style: {} };
    
    return sheet.data.get(cellId) || { value: "", formula: "", style: {} };
  };

  // Load spreadsheets when user changes
  useEffect(() => {
    if (currentUser) {
      loadSpreadsheets();
    } else {
      setSpreadsheets([]);
      setCurrentSpreadsheet(null);
    }
  }, [currentUser, loadSpreadsheets]);

  // Cleanup timeout on unmount
  useEffect(() => {
    return () => {
      if (autoSaveTimeout) {
        clearTimeout(autoSaveTimeout);
      }
    };
  }, [autoSaveTimeout]);

  const value = {
    spreadsheets,
    currentSpreadsheet,
    loading,
    error,
    createSpreadsheet,
    loadSpreadsheet,
    updateSpreadsheet,
    deleteSpreadsheet,
    updateCell,
    autoSave,
    addSheet,
    updateSheetData,
    shareSpreadsheet,
    getCurrentSheetData,
    getCellData,
    setCurrentSpreadsheet,
    setError
  };

  return (
    <SpreadsheetContext.Provider value={value}>
      {children}
    </SpreadsheetContext.Provider>
  );
}
