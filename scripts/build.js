/**
 * Build script for Excel Add-in
 * Creates a zip file that can be sideloaded in Excel
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { execSync } from "child_process";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.join(__dirname, "..");

async function build() {
  console.log("📦 Building Excel AI Add-in...");
  
  // Ensure dist exists
  const distDir = path.join(rootDir, "dist");
  if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir, { recursive: true });
  }
  
  // Copy taskpane files
  const taskpaneDir = path.join(rootDir, "taskpane");
  const distTaskpane = path.join(distDir, "taskpane");
  
  if (fs.existsSync(distTaskpane)) {
    fs.rmSync(distTaskpane, { recursive: true });
  }
  
  copyDir(taskpaneDir, distTaskpane);
  
  // Copy manifest
  fs.copyFileSync(
    path.join(taskpaneDir, "manifest.xml"),
    path.join(distDir, "manifest.xml")
  );
  
  // Create zip for distribution
  const zipName = "excel-ai-addin.zip";
  const zipPath = path.join(rootDir, zipName);
  
  try {
    execSync(`powershell -Command "Compress-Archive -Path '${distDir}\*' -DestinationPath '${zipPath}' -Force"`, {
      cwd: rootDir
    });
    console.log("✅ Build complete:", zipPath);
  } catch (e) {
    console.log("⚠️ ZIP creation failed, dist folder ready:", distTaskpane);
  }
}

function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  
  const entries = fs.readdirSync(src, { withFileTypes: true });
  
  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    
    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

build().catch(console.error);
