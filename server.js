const http = require("http");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

const PORT = 5500;

const MIME_TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".jpg": "image/jpg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".txt": "text/plain",
  ".ico": "image/x-icon",
};

let lastHeartbeat = Date.now();

const server = http.createServer((req, res) => {
  if (req.url === "/heartbeat") {
    lastHeartbeat = Date.now();
    res.writeHead(200);
    res.end("ok");
    return;
  }

  let filePath = "." + req.url;
  if (filePath === "./") {
    filePath = "./index.html";
  }

  const extname = path.extname(filePath);
  const contentType = MIME_TYPES[extname] || "application/octet-stream";

  const serveFile = (targetPath) => {
    fs.readFile(targetPath, (error, content) => {
      if (error) {
        if (error.code === "ENOENT") {
          // If not found in root, try checking the public folder
          if (!targetPath.startsWith("./public/")) {
            const publicPath = "./public" + req.url;
            serveFile(publicPath);
          } else {
            res.writeHead(404);
            res.end("File not found");
          }
        } else {
          res.writeHead(500);
          res.end("Server error: " + error.code);
        }
      } else {
        res.writeHead(200, { "Content-Type": contentType });
        res.end(content, "utf-8");
      }
    });
  };

  if (req.url.startsWith("/api/run-report")) {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const files = url.searchParams.get("files"); // comma separated
    const listOnly = url.searchParams.get("list");
    const order = url.searchParams.get("order");
    const out = url.searchParams.get("out");

    if (!files) {
      res.writeHead(400);
      res.end("No files provided");
      return;
    }

    let cmd = `python report.py ${files.split(",").map(f => `"${f.trim()}"`).join(" ")}`;
    if (listOnly) cmd += " --list";
    if (order) cmd += ` --order "${order}"`;
    if (out) cmd += ` --out "${out}"`;

    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: error.message, stderr }));
        return;
      }
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ output: stdout.trim() }));
    });
    return;
  }

  if (req.url.startsWith("/api/list-input")) {
    const inputDir = path.join(__dirname, "input");
    fs.readdir(inputDir, { withFileTypes: true }, (err, files) => {
      if (err) {
        res.writeHead(500);
        res.end(JSON.stringify({ error: err.message }));
        return;
      }
      const results = files
        .filter((f) => f.isFile())
        .map((f) => {
          const stats = fs.statSync(path.join(inputDir, f.name));
          return { name: f.name, size: stats.size };
        });
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify(results));
    });
    return;
  }

  serveFile(filePath);
});

// Auto-shutdown if no heartbeat for 10 seconds
setInterval(() => {
  if (Date.now() - lastHeartbeat > 10000) {
    process.exit(0);
  }
}, 5000);

const startServer = (port) => {
  server.listen(port, () => {
    const url = `http://localhost:${port}`;
    console.log(`[SUCCESS] Server running at ${url}`);

    // Open the browser - specifically target chrome on Windows if possible
    let command;
    if (process.platform === "win32") {
      command = `start chrome ${url}`;
    } else if (process.platform === "darwin") {
      command = `open -a "Google Chrome" ${url}`;
    } else {
      command = `xdg-open ${url}`;
    }

    exec(command, (err) => {
      if (err) {
        // Fallback to default opener if chrome fails
        const fallback =
          process.platform === "darwin"
            ? "open"
            : process.platform === "win32"
              ? "start"
              : "xdg-open";
        exec(`${fallback} ${url}`);
      }
    });
  }).on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.log(`[INFO] Port ${port} is busy, trying ${port + 1}...`);
      startServer(port + 1);
    } else {
      console.error('[ERROR] Server error:', err);
    }
  });
};

startServer(PORT);
