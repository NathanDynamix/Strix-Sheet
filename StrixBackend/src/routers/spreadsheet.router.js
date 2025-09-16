import { Router } from "express";
import {
  createSpreadsheet,
  deleteSpreadsheet,
  getSpreadsheetById,
  getSpreadsheets,
  updateSpreadsheet,
  updateCell,
  getCell,
  addSheet,
  updateSheetData,
  shareSpreadsheet,
} from "../controllers/spreadsheet.controller.js";

const router = Router();

// Basic CRUD operations
router.route("/").post(createSpreadsheet).get(getSpreadsheets);

router
  .route("/:id")
  .get(getSpreadsheetById)
  .put(updateSpreadsheet)
  .delete(deleteSpreadsheet);

// Cell operations
router.route("/:id/cells/:sheetId/:cellId")
  .get(getCell)
  .put(updateCell);

// Sheet operations
router.route("/:id/sheets")
  .post(addSheet);

router.route("/:id/sheets/:sheetId/data")
  .put(updateSheetData);

// Sharing
router.route("/:id/share")
  .post(shareSpreadsheet);

export default router; 