import { Spreadsheet } from "../models/spreadsheet.model.js";
import { ApiError } from "../utils/apiError.js";
import { ApiResponse } from "../utils/apiResponse.js";
import { asyncHandler } from "../utils/asyncHandler.js";

const createSpreadsheet = asyncHandler(async (req, res) => {
  const { name, data } = req.body;

  if (!name) {
    throw new ApiError(400, "Spreadsheet name is required");
  }

  const spreadsheet = await Spreadsheet.create({
    name,
    data: data || [],
  });

  return res
    .status(201)
    .json(
      new ApiResponse(201, spreadsheet, "Spreadsheet created successfully")
    );
});

const getSpreadsheets = asyncHandler(async (req, res) => {
  const spreadsheets = await Spreadsheet.find();
  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheets, "Spreadsheets retrieved successfully")
    );
});

const getSpreadsheetById = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const spreadsheet = await Spreadsheet.findById(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheet, "Spreadsheet retrieved successfully")
    );
});

const updateSpreadsheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { name, data } = req.body;

  const spreadsheet = await Spreadsheet.findByIdAndUpdate(
    id,
    { $set: { name, data } },
    { new: true, runValidators: true }
  );

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  return res
    .status(200)
    .json(
      new ApiResponse(200, spreadsheet, "Spreadsheet updated successfully")
    );
});

const deleteSpreadsheet = asyncHandler(async (req, res) => {
  const { id } = req.params;
  const spreadsheet = await Spreadsheet.findByIdAndDelete(id);

  if (!spreadsheet) {
    throw new ApiError(404, "Spreadsheet not found");
  }

  return res
    .status(200)
    .json(new ApiResponse(200, {}, "Spreadsheet deleted successfully"));
});

export {
  createSpreadsheet,
  getSpreadsheets,
  getSpreadsheetById,
  updateSpreadsheet,
  deleteSpreadsheet,
}; 