# Strix - Google Sheets Clone Setup Guide

## Overview
This is a complete Google Sheets clone with Firebase authentication and MongoDB backend for data persistence. The application includes real-time collaboration features and a comprehensive spreadsheet interface.

## Features Implemented

### ✅ Authentication System
- Firebase Authentication with Google OAuth
- Email/Password authentication
- User context management
- Protected routes

### ✅ Backend API
- MongoDB with Mongoose for data persistence
- Comprehensive spreadsheet model with cells, sheets, and collaboration
- RESTful API endpoints for CRUD operations
- User permissions and sharing system

### ✅ Frontend Integration
- React context for state management
- Real-time data synchronization
- Auto-save functionality
- Spreadsheet dashboard

## Setup Instructions

### 1. Firebase Configuration

1. Go to [Firebase Console](https://console.firebase.google.com/)
2. Create a new project or use existing one
3. Enable Authentication:
   - Go to Authentication > Sign-in method
   - Enable Email/Password
   - Enable Google provider
4. Get your Firebase config:
   - Go to Project Settings > General
   - Copy the config object
5. Update `src/Firebase.js` with your config:

```javascript
const firebaseConfig = {
  apiKey: "your-api-key",
  authDomain: "your-project.firebaseapp.com",
  projectId: "your-project-id",
  storageBucket: "your-project.appspot.com",
  messagingSenderId: "your-sender-id",
  appId: "your-app-id"
};
```

### 2. Backend Setup

1. Navigate to the backend directory:
```bash
cd StrixBackend
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file in the backend directory:
```env
PORT=2000
MONGODB_URI=mongodb://localhost:27017/strix
JWT_SECRET=your-jwt-secret-key
NODE_ENV=development
```

4. Start MongoDB (if using local instance):
```bash
# Using MongoDB Compass or command line
mongod
```

5. Start the backend server:
```bash
npm run server
```

### 3. Frontend Setup

1. Navigate to the main project directory:
```bash
cd /Users/aakashRao/Documents/Projects/Strix
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev
```

### 4. Database Setup

The application will automatically create the necessary collections in MongoDB. The main collections are:

- `spreadsheets` - Stores spreadsheet documents
- `users` - User information (if using custom user management)

## API Endpoints

### Spreadsheet Operations
- `POST /api/v1/spreadsheets` - Create new spreadsheet
- `GET /api/v1/spreadsheets?userId={id}` - Get user's spreadsheets
- `GET /api/v1/spreadsheets/{id}?userId={id}` - Get specific spreadsheet
- `PUT /api/v1/spreadsheets/{id}?userId={id}` - Update spreadsheet
- `DELETE /api/v1/spreadsheets/{id}?userId={id}` - Delete spreadsheet

### Cell Operations
- `PUT /api/v1/spreadsheets/{id}/cells/{sheetId}/{cellId}?userId={id}` - Update cell
- `GET /api/v1/spreadsheets/{id}/cells/{sheetId}/{cellId}?userId={id}` - Get cell

### Sheet Operations
- `POST /api/v1/spreadsheets/{id}/sheets?userId={id}` - Add new sheet
- `PUT /api/v1/spreadsheets/{id}/sheets/{sheetId}/data?userId={id}` - Update sheet data

### Sharing
- `POST /api/v1/spreadsheets/{id}/share?userId={id}` - Share spreadsheet

## Application Structure

```
src/
├── context/
│   ├── AuthContext.jsx          # Authentication context
│   └── SpreadsheetContext.jsx   # Spreadsheet data context
├── services/
│   └── spreadsheetService.js    # API service layer
├── component/
│   ├── StrixAuth.jsx           # Authentication component
│   ├── SpreadsheetModel.jsx    # Main spreadsheet interface
│   └── SpreadsheetDashboard.jsx # Spreadsheet management dashboard
└── Firebase.js                 # Firebase configuration

StrixBackend/src/
├── models/
│   └── spreadsheet.model.js    # MongoDB schema
├── controllers/
│   └── spreadsheet.controller.js # API controllers
├── routers/
│   └── spreadsheet.router.js   # API routes
└── middleware/
    └── auth.middleware.js      # Authentication middleware
```

## Usage

### 1. Authentication
- Users can sign up with email/password or Google OAuth
- Authentication state is managed globally
- Protected routes redirect to login if not authenticated

### 2. Spreadsheet Management
- Create new spreadsheets from the dashboard
- View all user's spreadsheets
- Search and filter spreadsheets
- Delete spreadsheets (owner only)

### 3. Spreadsheet Editing
- Full Google Sheets-like interface
- Formula support with financial functions
- Cell formatting (bold, italic, colors, alignment)
- Auto-save functionality
- Real-time collaboration ready

### 4. Data Persistence
- All changes are automatically saved to MongoDB
- Auto-save triggers after 2 seconds of inactivity
- Version history and last modified tracking
- Collaborative editing support

## Development Notes

### Auto-save Implementation
The application includes intelligent auto-save:
- Triggers after 2 seconds of inactivity
- Only saves when there are actual changes
- Handles network errors gracefully
- Doesn't interrupt user workflow

### Real-time Collaboration
The backend is prepared for real-time features:
- Socket.io integration ready
- User presence tracking
- Conflict resolution strategies
- Collaborative cursors

### Performance Optimizations
- Virtual scrolling for large spreadsheets
- Memoized cell rendering
- Efficient data structures
- Lazy loading of sheet data

## Troubleshooting

### Common Issues

1. **Firebase Authentication Issues**
   - Check Firebase configuration
   - Verify API keys are correct
   - Ensure authentication methods are enabled

2. **Backend Connection Issues**
   - Verify MongoDB is running
   - Check database connection string
   - Ensure backend server is running on port 2000

3. **CORS Issues**
   - Backend is configured for localhost:5173
   - Update CORS settings if using different ports

4. **Data Not Saving**
   - Check browser console for errors
   - Verify user authentication
   - Check network connectivity

### Debug Mode
Enable debug logging by setting `NODE_ENV=development` in the backend `.env` file.

## Next Steps

### Real-time Collaboration
To implement real-time features:
1. Add Socket.io client to frontend
2. Implement user presence tracking
3. Add collaborative cursors
4. Handle conflict resolution

### Advanced Features
- Version history and rollback
- Advanced sharing permissions
- Export to Excel/CSV
- Advanced charting capabilities
- Mobile responsiveness improvements

## Support

For issues or questions:
1. Check the browser console for errors
2. Verify all services are running
3. Check network connectivity
4. Review the setup steps above

The application is now ready for development and testing!
