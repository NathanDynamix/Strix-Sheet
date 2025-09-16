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

// Request logging middleware
app.use((req, res, next) => {
  const timestamp = new Date().toISOString();
  const method = req.method;
  const url = req.url;
  const userAgent = req.get('User-Agent') || 'Unknown';
  const ip = req.ip || req.connection.remoteAddress;
  
  console.log(`\n🚀 [${timestamp}] API Request:`);
  console.log(`   Method: ${method}`);
  console.log(`   URL: ${url}`);
  console.log(`   IP: ${ip}`);
  console.log(`   User-Agent: ${userAgent}`);
  
  if (req.body && Object.keys(req.body).length > 0) {
    console.log(`   Body:`, JSON.stringify(req.body, null, 2));
  }
  
  if (req.query && Object.keys(req.query).length > 0) {
    console.log(`   Query:`, JSON.stringify(req.query, null, 2));
  }
  
  if (req.headers.authorization) {
    console.log(`   Authorization: ${req.headers.authorization.substring(0, 20)}...`);
  }
  
  // Log response
  const originalSend = res.send;
  res.send = function(data) {
    console.log(`   Response Status: ${res.statusCode}`);
    if (data && typeof data === 'string') {
      try {
        const parsed = JSON.parse(data);
        console.log(`   Response Data:`, JSON.stringify(parsed, null, 2));
      } catch (e) {
        console.log(`   Response Data: ${data.substring(0, 200)}${data.length > 200 ? '...' : ''}`);
      }
    }
    console.log(`   ⏱️  Request completed at: ${new Date().toISOString()}\n`);
    originalSend.call(this, data);
  };
  
  next();
});

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
