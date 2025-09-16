import { Spreadsheet } from "../models/spreadsheet.model.js";
import { ApiError } from "../utils/apiError.js";
import { ApiResponse } from "../utils/apiResponse.js";
import { asyncHandler } from "../utils/asyncHandler.js";

// Create a new spreadsheet
const createSpreadsheet = asyncHandler(async (req, res) => {
  console.log("📝 Creating new spreadsheet...");
  const { title, description, owner } = req.body;
  console.log("   Title:", title);
  console.log("   Description:", description);
  console.log("   Owner:", owner);

  if (!title) {
    throw new ApiError(400, "Spreadsheet title is required");
  }

  if (!owner) {
    throw new ApiError(400, "Owner ID is required");
  }

  // Create initial sheet
  const initialSheet = {
    id: "sheet1",
    name: "Sheet1",
    data: new Map(),
    activeCell: "A1",
    selectedRange: { start: "A1", end: "A1" }
  };

  const spreadsheet = await Spreadsheet.create({
    title,
    description: description || "",
    owner,
    sheets: [initialSheet],
    activeSheetId: "sheet1"
  });

  return res
    .status(201)
    .json(
      new ApiResponse(201, spreadsheet, "Spreadsheet created successfully")
    );
});

// Get all spreadsheets for a user
const getSpreadsheets = asyncHandler(async (req, res) => {
  console.log("📋 Getting spreadsheets for user...");
  const { userId } = req.query;
  console.log("   User ID:", userId);
  
  if (!userId) {
    throw new ApiError(400, "User ID is required");
  }

  const spreadsheets = await Spreadsheet.find({
    $or: [
      { owner: userId },
      { "collaborators.userId": userId }
    ]
  }).sort({ lastModified: -1 });

  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheets, "Spreadsheets retrieved successfully")
    );
});

// Get a specific spreadsheet
const getSpreadsheetById = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check if user has access
  const hasAccess = spreadsheet.owner === userId || 
    spreadsheet.collaborators.some(collab => collab.userId === userId);

  if (!hasAccess) {
    throw new ApiError(403, "Access denied");
  }

  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheet, "Spreadsheet retrieved successfully")
    );
});

// Update spreadsheet metadata
const updateSpreadsheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { title, description, settings, tags } = req.body;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check if user is owner or has edit permissions
  const isOwner = spreadsheet.owner === userId;
  const collaborator = spreadsheet.collaborators.find(collab => collab.userId === userId);
  const canEdit = isOwner || (collaborator && collaborator.permissions.canEdit);

  if (!canEdit) {
    throw new ApiError(403, "You don't have permission to edit this spreadsheet");
  }

  const updateData = {};
  if (title) updateData.title = title;
  if (description !== undefined) updateData.description = description;
  if (settings) updateData.settings = { ...spreadsheet.settings, ...settings };
  if (tags) updateData.tags = tags;
  updateData.lastModified = new Date();

  const updatedSpreadsheet = await Spreadsheet.findByIdAndUpdate(
    id,
    { $set: updateData },
    { new: true, runValidators: true }
  );

  return res
    .status(200)
    .json(
      new ApiResponse(200, updatedSpreadsheet, "Spreadsheet updated successfully")
    );
});

// Delete spreadsheet
const deleteSpreadsheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Only owner can delete
  if (spreadsheet.owner !== userId) {
    throw new ApiError(403, "Only the owner can delete this spreadsheet");
  }

  await Spreadsheet.findByIdAndDelete(id);

  return res
    .status(200)
    .json(new ApiResponse(200, {}, "Spreadsheet deleted successfully"));
});

// Update a specific cell
const updateCell = asyncHandler(async (req, res) => {
  console.log("✏️ Updating cell...");
  const { id } = req.params;
  const { sheetId, cellId, cellData } = req.body;
  const { userId } = req.query;
  console.log("   Spreadsheet ID:", id);
  console.log("   Sheet ID:", sheetId);
  console.log("   Cell ID:", cellId);
  console.log("   Cell Data:", cellData);
  console.log("   User ID:", userId);

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check permissions
  const isOwner = spreadsheet.owner === userId;
  const collaborator = spreadsheet.collaborators.find(collab => collab.userId === userId);
  const canEdit = isOwner || (collaborator && collaborator.permissions.canEdit);

  if (!canEdit) {
    throw new ApiError(403, "You don't have permission to edit this spreadsheet");
  }

  spreadsheet.updateCell(sheetId, cellId, cellData);
  await spreadsheet.save();

  return res
    .status(200)
    .json(
      new ApiResponse(200, { cellId, cellData }, "Cell updated successfully")
    );
});

// Get a specific cell
const getCell = asyncHandler(async (req, res) => {
  const { id, sheetId, cellId } = req.params;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check access
  const hasAccess = spreadsheet.owner === userId || 
    spreadsheet.collaborators.some(collab => collab.userId === userId);

  if (!hasAccess) {
    throw new ApiError(403, "Access denied");
  }

  const cellData = spreadsheet.getCell(sheetId, cellId);

  return res
    .status(200)
    .json(
      new ApiResponse(200, { cellId, cellData }, "Cell retrieved successfully")
    );
});

// Add a new sheet
const addSheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { name } = req.body;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check permissions
  const isOwner = spreadsheet.owner === userId;
  const collaborator = spreadsheet.collaborators.find(collab => collab.userId === userId);
  const canEdit = isOwner || (collaborator && collaborator.permissions.canEdit);

  if (!canEdit) {
    throw new ApiError(403, "You don't have permission to edit this spreadsheet");
  }

  const newSheet = spreadsheet.addSheet(name);
  await spreadsheet.save();

  return res
    .status(200)
    .json(
      new ApiResponse(200, newSheet, "Sheet added successfully")
    );
});

// Update sheet data in bulk
const updateSheetData = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { sheetId, data } = req.body;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Check permissions
  const isOwner = spreadsheet.owner === userId;
  const collaborator = spreadsheet.collaborators.find(collab => collab.userId === userId);
  const canEdit = isOwner || (collaborator && collaborator.permissions.canEdit);

  if (!canEdit) {
    throw new ApiError(403, "You don't have permission to edit this spreadsheet");
  }

  const sheet = spreadsheet.sheets.find(s => s.id === sheetId);
  if (!sheet) {
    throw new ApiError(404, "Sheet not found");
  }

  // Update sheet data
  sheet.data = new Map(Object.entries(data));
  spreadsheet.lastModified = new Date();
  spreadsheet.version += 1;

  await spreadsheet.save();

  return res
    .status(200)
    .json(
      new ApiResponse(200, { sheetId, data }, "Sheet data updated successfully")
    );
});

// Share spreadsheet with collaborators
const shareSpreadsheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { collaborators } = req.body;
  const { userId } = req.query;

  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  // Only owner can share
  if (spreadsheet.owner !== userId) {
    throw new ApiError(403, "Only the owner can share this spreadsheet");
  }

  // Add new collaborators
  collaborators.forEach(collab => {
    const existingIndex = spreadsheet.collaborators.findIndex(
      c => c.userId === collab.userId
    );
    
    if (existingIndex >= 0) {
      spreadsheet.collaborators[existingIndex] = collab;
    } else {
      spreadsheet.collaborators.push(collab);
    }
  });

  await spreadsheet.save();

  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheet.collaborators, "Spreadsheet shared successfully")
    );
});

export {
  createSpreadsheet,
  getSpreadsheets,
  getSpreadsheetById,
  updateSpreadsheet,
  deleteSpreadsheet,
  updateCell,
  getCell,
  addSheet,
  updateSheetData,
  shareSpreadsheet,
}; 