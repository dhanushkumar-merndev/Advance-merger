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

  serveFile(filePath);
});

// Auto-shutdown if no heartbeat for 10 seconds
setInterval(() => {
  if (Date.now() - lastHeartbeat > 10000) {
    process.exit(0);
  }
}, 5000);

server.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;

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
});
