import React, { useEffect, useRef } from 'react';
import { Navigate, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useToast } from '../context/ToastContext';

const ProtectedRoute = ({ children }) => {
  const { currentUser, loading } = useAuth();
  const location = useLocation();
  const { showError } = useToast();
  const hasShownError = useRef(false);

  // Show error message after component mounts, not during render
  useEffect(() => {
    if (!loading && !currentUser && !hasShownError.current) {
      hasShownError.current = true;
      showError('Please log in to access this page');
    }
  }, [loading, currentUser, showError]);

  // Reset the error flag when user changes
  useEffect(() => {
    if (currentUser) {
      hasShownError.current = false;
    }
  }, [currentUser]);

  // Show loading spinner while checking authentication
  if (loading) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-indigo-50 via-white to-purple-50 flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-indigo-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Loading...</p>
        </div>
      </div>
    );
  }

  // If user is not authenticated, redirect to login
  if (!currentUser) {
    return <Navigate to="/" state={{ from: location }} replace />;
  }

  // If user is authenticated, render the protected component
  return children;
};

export default ProtectedRoute;
