const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const { spawn, execSync } = require("child_process");
const readline = require("readline");

let mainWindow;
let backendProcess = null;
let shellPATH = null;

// On macOS, apps launched from Finder get a minimal PATH.
// Resolve the user's full shell PATH so the backend can find claude, pandoc, etc.
function getShellPATH() {
  if (shellPATH !== null) return shellPATH;
  if (process.platform === "darwin") {
    try {
      const shell = process.env.SHELL || "/bin/zsh";
      shellPATH = execSync(`${shell} -ilc 'echo $PATH'`, {
        encoding: "utf8",
        timeout: 5000,
      }).trim();
      return shellPATH;
    } catch {
      // Fall through to process.env.PATH
    }
  }
  shellPATH = process.env.PATH || "";
  return shellPATH;
}

function getBackendPath() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "backend", "zotellm_backend");
  }
  // Dev mode: look for PyInstaller output
  return path.join(__dirname, "..", "dist", "zotellm_backend", "zotellm_backend");
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 720,
    height: 800,
    minWidth: 600,
    minHeight: 600,
    title: "zotellm",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(() => {
  createWindow();
  // On macOS, ensure the app appears in the dock and comes to foreground
  if (process.platform === "darwin") {
    app.dock.show();
  }
  app.focus({ steal: true });
});

app.on("activate", () => {
  // macOS: re-create window when dock icon is clicked and no windows exist
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on("window-all-closed", () => {
  if (backendProcess) {
    backendProcess.kill();
    backendProcess = null;
  }
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// ---------- IPC Handlers ----------

ipcMain.handle("open-file-dialog", async (event, options) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: options.filters || [
      { name: "Documents", extensions: ["docx", "md", "markdown", "txt"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle("save-file-dialog", async () => {
  const result = await dialog.showSaveDialog(mainWindow, {
    filters: [
      { name: "Word Documents", extensions: ["docx"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  return result.canceled ? null : result.filePath;
});

ipcMain.handle("start-processing", async (event, args) => {
  const backendPath = getBackendPath();

  try {
    // Use full shell PATH so the backend can find pandoc, claude, ollama, etc.
    const env = { ...process.env, PATH: getShellPATH() };

    backendProcess = spawn(backendPath, [], {
      stdio: ["pipe", "pipe", "pipe"],
      env,
    });

    // Parse NDJSON from backend stdout
    const rl = readline.createInterface({ input: backendProcess.stdout });

    rl.on("line", (line) => {
      try {
        const msg = JSON.parse(line);
        mainWindow.webContents.send("backend-message", msg);
      } catch {
        // Non-JSON output, treat as log
        mainWindow.webContents.send("backend-message", {
          type: "log",
          text: line,
        });
      }
    });

    // Capture stderr as log messages
    const stderrRl = readline.createInterface({ input: backendProcess.stderr });
    stderrRl.on("line", (line) => {
      mainWindow.webContents.send("backend-message", {
        type: "log",
        text: `[stderr] ${line}`,
      });
    });

    backendProcess.on("error", (err) => {
      mainWindow.webContents.send("backend-message", {
        type: "done",
        success: false,
        message: `Failed to start backend: ${err.message}. Make sure you've built the backend with build_backend.sh.`,
      });
      backendProcess = null;
    });

    backendProcess.on("close", (code) => {
      if (code !== 0 && code !== null) {
        mainWindow.webContents.send("backend-message", {
          type: "done",
          success: false,
          message: `Backend exited with code ${code}`,
        });
      }
      backendProcess = null;
    });

    // Send the start command
    const startMsg = JSON.stringify({ type: "start", args }) + "\n";
    backendProcess.stdin.write(startMsg);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

ipcMain.handle("resolve-response", async (event, response) => {
  if (backendProcess && backendProcess.stdin.writable) {
    const msg = JSON.stringify(response) + "\n";
    backendProcess.stdin.write(msg);
  }
});

ipcMain.handle("get-default-zotero-db", async () => {
  const fs = require("fs");
  const homedir = require("os").homedir();
  const dbPath = path.join(homedir, "Zotero", "zotero.sqlite");
  if (fs.existsSync(dbPath)) {
    return dbPath;
  }
  return null;
});
