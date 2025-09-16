import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import { createServer } from "http";
import { Server } from "socket.io";


const app = express();
app.use(cookieParser());
app.use(
  cors({
    origin: [
      process.env.FRONTEND_URL || "http://localhost:5173",
    ],
    methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
    allowedHeaders: [
      "Content-Type",
      "Authorization",
      "Access-Control-Allow-Headers",
      "Access-Control-Allow-Credentials",
      "Access-Control-Allow-Origin",
    ],
    credentials: true,
  })
);


const server = createServer(app);
const io = new Server(server, {
  cors: {
    origin: [
      process.env.FRONTEND_URL || "http://localhost:5173",
    ],
    methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
    credentials: true,
  },
});

app.use(express.json({ limit: process.env.MAX_FILE_SIZE || "50mb" }));
app.use(express.urlencoded({ extended: true, limit: process.env.MAX_FILE_SIZE || "16kbs" }));
app.use(express.static("public"));

// import routes
import spreadsheetRouter from "./routers/spreadsheet.router.js";

// routes declaration
app.use("/api/v1/spreadsheets", spreadsheetRouter);

const activeUsers = new Map();

io.on("connection", (socket) => {
  socket.on("authenticate", (userId) => {
    activeUsers.set(userId, socket.id);
    console.log(`User ${userId} connected with socket ID: ${socket.id}`);
  });

  // Handle disconnection
  socket.on("disconnect", () => {
    for (const [userId, socketId] of activeUsers.entries()) {
      if (socketId === socket.id) {
        activeUsers.delete(userId);
        console.log(`Disconnected with this ID ->${userId} disconnected`);
        break;
      }
    }
  });
});

export { app, server };

// http://localhost:5173
