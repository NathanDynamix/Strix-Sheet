const API_BASE_URL = 'http://localhost:2000/api/v1/spreadsheets';

class SpreadsheetService {
  constructor() {
    this.baseURL = API_BASE_URL;
  }

  // Helper method to get auth headers
  getAuthHeaders() {
    const user = JSON.parse(localStorage.getItem('user') || '{}');
    return {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${user.accessToken || ''}`,
    };
  }

  // Helper method to get user ID
  getUserId() {
    const user = JSON.parse(localStorage.getItem('user') || '{}');
    return user.uid || user.id;
  }

  // Create a new spreadsheet
  async createSpreadsheet(title, description = '') {
    try {
      const response = await fetch(`${this.baseURL}`, {
        method: 'POST',
        headers: this.getAuthHeaders(),
        body: JSON.stringify({
          title,
          description,
          owner: this.getUserId()
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
      throw error;
    }
  }

  // Get all spreadsheets for the current user
  async getSpreadsheets() {
    try {
      const response = await fetch(`${this.baseURL}?userId=${this.getUserId()}`, {
        method: 'GET',
        headers: this.getAuthHeaders()
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error fetching spreadsheets:', error);
      throw error;
    }
  }

  // Get a specific spreadsheet by ID
  async getSpreadsheet(id) {
    try {
      const response = await fetch(`${this.baseURL}/${id}?userId=${this.getUserId()}`, {
        method: 'GET',
        headers: this.getAuthHeaders()
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error fetching spreadsheet:', error);
      throw error;
    }
  }

  // Update spreadsheet metadata
  async updateSpreadsheet(id, updates) {
    try {
      const response = await fetch(`${this.baseURL}/${id}?userId=${this.getUserId()}`, {
        method: 'PUT',
        headers: this.getAuthHeaders(),
        body: JSON.stringify(updates)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error updating spreadsheet:', error);
      throw error;
    }
  }

  // Delete a spreadsheet
  async deleteSpreadsheet(id) {
    try {
      const response = await fetch(`${this.baseURL}/${id}?userId=${this.getUserId()}`, {
        method: 'DELETE',
        headers: this.getAuthHeaders()
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error deleting spreadsheet:', error);
      throw error;
    }
  }

  // Update a specific cell
  async updateCell(spreadsheetId, sheetId, cellId, cellData) {
    try {
      const response = await fetch(`${this.baseURL}/${spreadsheetId}/cells/${sheetId}/${cellId}?userId=${this.getUserId()}`, {
        method: 'PUT',
        headers: this.getAuthHeaders(),
        body: JSON.stringify({ cellData })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error updating cell:', error);
      throw error;
    }
  }

  // Get a specific cell
  async getCell(spreadsheetId, sheetId, cellId) {
    try {
      const response = await fetch(`${this.baseURL}/${spreadsheetId}/cells/${sheetId}/${cellId}?userId=${this.getUserId()}`, {
        method: 'GET',
        headers: this.getAuthHeaders()
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error fetching cell:', error);
      throw error;
    }
  }

  // Add a new sheet to a spreadsheet
  async addSheet(spreadsheetId, name) {
    try {
      const response = await fetch(`${this.baseURL}/${spreadsheetId}/sheets?userId=${this.getUserId()}`, {
        method: 'POST',
        headers: this.getAuthHeaders(),
        body: JSON.stringify({ name })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error adding sheet:', error);
      throw error;
    }
  }

  // Update sheet data in bulk
  async updateSheetData(spreadsheetId, sheetId, data) {
    try {
      const response = await fetch(`${this.baseURL}/${spreadsheetId}/sheets/${sheetId}/data?userId=${this.getUserId()}`, {
        method: 'PUT',
        headers: this.getAuthHeaders(),
        body: JSON.stringify({ data })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error updating sheet data:', error);
      throw error;
    }
  }

  // Share spreadsheet with collaborators
  async shareSpreadsheet(spreadsheetId, collaborators) {
    try {
      const response = await fetch(`${this.baseURL}/${spreadsheetId}/share?userId=${this.getUserId()}`, {
        method: 'POST',
        headers: this.getAuthHeaders(),
        body: JSON.stringify({ collaborators })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error sharing spreadsheet:', error);
      throw error;
    }
  }

  // Auto-save functionality
  async autoSave(spreadsheetId, sheetId, data) {
    try {
      // Only save if there are changes
      if (!data || Object.keys(data).length === 0) {
        return;
      }

      await this.updateSheetData(spreadsheetId, sheetId, data);
      console.log('Auto-saved successfully');
    } catch (error) {
      console.error('Auto-save failed:', error);
      // Don't throw error for auto-save failures
    }
  }
}

export default new SpreadsheetService();
