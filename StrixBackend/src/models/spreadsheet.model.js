import mongoose, { Schema } from "mongoose";

// Cell schema for individual cells
const cellSchema = new Schema({
  value: {
    type: String,
    default: ""
  },
  formula: {
    type: String,
    default: ""
  },
  style: {
    bold: { type: Boolean, default: false },
    italic: { type: Boolean, default: false },
    underline: { type: Boolean, default: false },
    color: { type: String, default: "#000000" },
    backgroundColor: { type: String, default: "#FFFFFF" },
    textAlign: { type: String, enum: ["left", "center", "right"], default: "left" },
    fontSize: { type: Number, default: 10 },
    fontFamily: { type: String, default: "Arial" }
  }
}, { _id: false });

// Sheet schema for individual sheets within a spreadsheet
const sheetSchema = new Schema({
  id: {
    type: String,
    required: true
  },
  name: {
    type: String,
    required: true,
    default: "Sheet1"
  },
  data: {
    type: Map,
    of: cellSchema,
    default: new Map()
  },
  activeCell: {
    type: String,
    default: "A1"
  },
  selectedRange: {
    start: { type: String, default: "A1" },
    end: { type: String, default: "A1" }
  }
}, { _id: false });

// Main spreadsheet schema
const spreadsheetSchema = new Schema(
  {
    title: {
      type: String,
      required: true,
      trim: true,
      default: "Untitled Spreadsheet"
    },
    description: {
      type: String,
      trim: true,
      default: ""
    },
    owner: {
      type: String,
      required: true,
      index: true
    },
    collaborators: [{
      userId: {
        type: String,
        required: true
      },
      email: {
        type: String,
        required: true
      },
      role: {
        type: String,
        enum: ["viewer", "editor", "owner"],
        default: "viewer"
      },
      permissions: {
        canEdit: { type: Boolean, default: false },
        canComment: { type: Boolean, default: true },
        canShare: { type: Boolean, default: false }
      }
    }],
    sheets: [sheetSchema],
    activeSheetId: {
      type: String,
      default: "sheet1"
    },
    settings: {
      isPublic: { type: Boolean, default: false },
      allowComments: { type: Boolean, default: true },
      allowSuggestions: { type: Boolean, default: true },
      versionHistory: { type: Boolean, default: true }
    },
    version: {
      type: Number,
      default: 1
    },
    lastModified: {
      type: Date,
      default: Date.now
    },
    tags: [{
      type: String,
      trim: true
    }],
    size: {
      rows: { type: Number, default: 1000 },
      columns: { type: Number, default: 40 }
    }
  },
  { 
    timestamps: true,
    toJSON: { virtuals: true },
    toObject: { virtuals: true }
  }
);

// Indexes for better performance
spreadsheetSchema.index({ owner: 1, createdAt: -1 });
spreadsheetSchema.index({ "collaborators.userId": 1 });
spreadsheetSchema.index({ title: "text", description: "text" });

// Virtual for getting the active sheet
spreadsheetSchema.virtual('activeSheet').get(function() {
  return this.sheets.find(sheet => sheet.id === this.activeSheetId);
});

// Method to add a new sheet
spreadsheetSchema.methods.addSheet = function(name) {
  const newSheetId = `sheet${this.sheets.length + 1}`;
  const newSheet = {
    id: newSheetId,
    name: name || `Sheet${this.sheets.length + 1}`,
    data: new Map(),
    activeCell: "A1",
    selectedRange: { start: "A1", end: "A1" }
  };
  this.sheets.push(newSheet);
  this.activeSheetId = newSheetId;
  return newSheet;
};

// Method to update a cell
spreadsheetSchema.methods.updateCell = function(sheetId, cellId, cellData) {
  const sheet = this.sheets.find(s => s.id === sheetId);
  if (!sheet) {
    throw new Error('Sheet not found');
  }
  
  if (!sheet.data) {
    sheet.data = new Map();
  }
  
  sheet.data.set(cellId, cellData);
  this.lastModified = new Date();
  this.version += 1;
};

// Method to get a cell
spreadsheetSchema.methods.getCell = function(sheetId, cellId) {
  const sheet = this.sheets.find(s => s.id === sheetId);
  if (!sheet || !sheet.data) {
    return { value: "", formula: "", style: {} };
  }
  return sheet.data.get(cellId) || { value: "", formula: "", style: {} };
};

export const Spreadsheet = mongoose.model("Spreadsheet", spreadsheetSchema); 