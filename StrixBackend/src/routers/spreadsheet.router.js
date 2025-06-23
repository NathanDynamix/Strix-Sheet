import { Router } from "express";
import {
  createSpreadsheet,
  deleteSpreadsheet,
  getSpreadsheetById,
  getSpreadsheets,
  updateSpreadsheet,
} from "../controllers/spreadsheet.controller.js";

const router = Router();

router.route("/").post(createSpreadsheet).get(getSpreadsheets);

router
  .route("/:id")
  .get(getSpreadsheetById)
  .put(updateSpreadsheet)
  .delete(deleteSpreadsheet);

export default router; 