import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useSpreadsheet } from '../context/SpreadsheetContext';
import { 
  Plus, 
  FileSpreadsheet, 
  Search, 
  MoreVertical, 
  Edit, 
  Trash2, 
  Share2, 
  Calendar,
  User,
  Clock
} from 'lucide-react';

const SpreadsheetDashboard = () => {
  const navigate = useNavigate();
  const { currentUser, logout } = useAuth();
  const { 
    spreadsheets, 
    loading, 
    error, 
    createSpreadsheet, 
    deleteSpreadsheet 
  } = useSpreadsheet();
  
  const [searchTerm, setSearchTerm] = useState('');
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [newSpreadsheetTitle, setNewSpreadsheetTitle] = useState('');
  const [newSpreadsheetDescription, setNewSpreadsheetDescription] = useState('');
  const [isCreating, setIsCreating] = useState(false);

  // Filter spreadsheets based on search term
  const filteredSpreadsheets = spreadsheets.filter(spreadsheet =>
    spreadsheet.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
    spreadsheet.description.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const handleCreateSpreadsheet = async (e) => {
    e.preventDefault();
    if (!newSpreadsheetTitle.trim()) return;

    try {
      setIsCreating(true);
      const newSpreadsheet = await createSpreadsheet(
        newSpreadsheetTitle.trim(),
        newSpreadsheetDescription.trim()
      );
      
      // Navigate to the new spreadsheet
      navigate('/integration');
      
      // Reset form
      setNewSpreadsheetTitle('');
      setNewSpreadsheetDescription('');
      setShowCreateModal(false);
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
    } finally {
      setIsCreating(false);
    }
  };

  const handleDeleteSpreadsheet = async (id, title) => {
    if (window.confirm(`Are you sure you want to delete "${title}"? This action cannot be undone.`)) {
      try {
        await deleteSpreadsheet(id);
      } catch (error) {
        console.error('Error deleting spreadsheet:', error);
      }
    }
  };

  const handleLogout = () => {
    logout();
    navigate('/');
  };

  const formatDate = (dateString) => {
    const date = new Date(dateString);
    const now = new Date();
    const diffTime = Math.abs(now - date);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays} days ago`;
    return date.toLocaleDateString();
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center py-4">
            <div className="flex items-center">
              <div className="bg-gradient-to-r from-indigo-600 to-purple-600 w-10 h-10 rounded-lg flex items-center justify-center mr-3">
                <span className="text-xl font-bold text-white">S</span>
              </div>
              <div>
                <h1 className="text-2xl font-bold text-gray-900">Strix Sheets</h1>
                <p className="text-sm text-gray-500">Welcome back, {currentUser?.displayName || currentUser?.email}</p>
              </div>
            </div>
            
            <div className="flex items-center space-x-4">
              <button
                onClick={() => setShowCreateModal(true)}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700 flex items-center space-x-2"
              >
                <Plus size={20} />
                <span>New Spreadsheet</span>
              </button>
              
              <button
                onClick={handleLogout}
                className="text-gray-600 hover:text-gray-900 px-3 py-2 rounded-lg hover:bg-gray-100"
              >
                Logout
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Search Bar */}
        <div className="mb-8">
          <div className="relative max-w-md">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-5 h-5" />
            <input
              type="text"
              placeholder="Search spreadsheets..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
            />
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div className="mb-6 p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg">
            {error}
          </div>
        )}

        {/* Loading State */}
        {loading && (
          <div className="flex justify-center items-center py-12">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div>
          </div>
        )}

        {/* Spreadsheets Grid */}
        {!loading && (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
            {filteredSpreadsheets.map((spreadsheet) => (
              <div
                key={spreadsheet._id}
                className="bg-white rounded-lg border border-gray-200 hover:shadow-lg transition-shadow cursor-pointer group"
                onClick={() => navigate('/integration')}
              >
                <div className="p-4">
                  <div className="flex items-start justify-between mb-3">
                    <div className="flex items-center">
                      <FileSpreadsheet className="w-8 h-8 text-green-600 mr-3" />
                      <div>
                        <h3 className="font-semibold text-gray-900 truncate">
                          {spreadsheet.title}
                        </h3>
                        <p className="text-sm text-gray-500 truncate">
                          {spreadsheet.description || 'No description'}
                        </p>
                      </div>
                    </div>
                    
                    <div className="relative">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          // Handle menu toggle
                        }}
                        className="opacity-0 group-hover:opacity-100 p-1 hover:bg-gray-100 rounded"
                      >
                        <MoreVertical size={16} className="text-gray-400" />
                      </button>
                    </div>
                  </div>
                  
                  <div className="flex items-center justify-between text-sm text-gray-500">
                    <div className="flex items-center space-x-4">
                      <div className="flex items-center">
                        <Calendar size={14} className="mr-1" />
                        {formatDate(spreadsheet.lastModified)}
                      </div>
                      <div className="flex items-center">
                        <User size={14} className="mr-1" />
                        {spreadsheet.collaborators?.length + 1 || 1}
                      </div>
                    </div>
                    
                    <div className="flex items-center space-x-1">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          // Handle edit
                        }}
                        className="p-1 hover:bg-gray-100 rounded opacity-0 group-hover:opacity-100"
                      >
                        <Edit size={14} className="text-gray-400" />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          // Handle share
                        }}
                        className="p-1 hover:bg-gray-100 rounded opacity-0 group-hover:opacity-100"
                      >
                        <Share2 size={14} className="text-gray-400" />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDeleteSpreadsheet(spreadsheet._id, spreadsheet.title);
                        }}
                        className="p-1 hover:bg-gray-100 rounded opacity-0 group-hover:opacity-100"
                      >
                        <Trash2 size={14} className="text-red-400" />
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            ))}
            
            {/* Empty State */}
            {filteredSpreadsheets.length === 0 && !loading && (
              <div className="col-span-full flex flex-col items-center justify-center py-12">
                <FileSpreadsheet className="w-16 h-16 text-gray-300 mb-4" />
                <h3 className="text-lg font-medium text-gray-900 mb-2">
                  {searchTerm ? 'No spreadsheets found' : 'No spreadsheets yet'}
                </h3>
                <p className="text-gray-500 text-center mb-6">
                  {searchTerm 
                    ? 'Try adjusting your search terms'
                    : 'Create your first spreadsheet to get started'
                  }
                </p>
                {!searchTerm && (
                  <button
                    onClick={() => setShowCreateModal(true)}
                    className="bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700 flex items-center space-x-2"
                  >
                    <Plus size={20} />
                    <span>Create Spreadsheet</span>
                  </button>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Create Spreadsheet Modal */}
      {showCreateModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-full max-w-md mx-4">
            <h2 className="text-xl font-semibold text-gray-900 mb-4">
              Create New Spreadsheet
            </h2>
            
            <form onSubmit={handleCreateSpreadsheet}>
              <div className="mb-4">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Title *
                </label>
                <input
                  type="text"
                  value={newSpreadsheetTitle}
                  onChange={(e) => setNewSpreadsheetTitle(e.target.value)}
                  placeholder="Enter spreadsheet title"
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                  required
                />
              </div>
              
              <div className="mb-6">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Description
                </label>
                <textarea
                  value={newSpreadsheetDescription}
                  onChange={(e) => setNewSpreadsheetDescription(e.target.value)}
                  placeholder="Enter description (optional)"
                  rows={3}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                />
              </div>
              
              <div className="flex justify-end space-x-3">
                <button
                  type="button"
                  onClick={() => setShowCreateModal(false)}
                  className="px-4 py-2 text-gray-700 border border-gray-300 rounded-lg hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={isCreating || !newSpreadsheetTitle.trim()}
                  className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isCreating ? 'Creating...' : 'Create'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default SpreadsheetDashboard;
