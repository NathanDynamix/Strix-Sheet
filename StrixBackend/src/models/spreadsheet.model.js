import mongoose, { Schema } from "mongoose";

const spreadsheetSchema = new Schema(
  {
    name: {
      type: String,
      required: true,
      trim: true,
    },
    data: [
      {
        label: {
          type: String,
          required: true,
        },
        value: {
          type: Number,
          required: true,
        },
      },
    ],
  },
  { timestamps: true }
);

export const Spreadsheet = mongoose.model("Spreadsheet", spreadsheetSchema); 