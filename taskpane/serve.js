/**
 * Simple HTTP server for Excel Add-in development
 * Usage: node serve.js
 */
import http from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;

const MIME_TYPES = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".xml": "text/xml"
};

const server = http.createServer((req, res) => {
  let filePath = path.join(__dirname, req.url === "/" ? "landing.html" : req.url);
  
  const ext = path.extname(filePath);
  const contentType = MIME_TYPES[ext] || "application/octet-stream";
  
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not Found");
      return;
    }
    res.writeHead(200, { "Content-Type": contentType });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`🖥️  Server running at http://localhost:${PORT}`);
  console.log(`📋 Use manifest.xml to sideload in Excel`);
});
