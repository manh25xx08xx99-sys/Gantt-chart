// Add-in 用の簡易静的サーバ（依存なし・Nodeのみ）
// 使い方: node server.js  → http://localhost:8080 でこのフォルダを配信する
const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 8080;
const ROOT = __dirname;
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xml": "text/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".md": "text/markdown; charset=utf-8",
};

const server = http.createServer((req, res) => {
  // ディレクトリトラバーサル対策：URLを安全に解決する
  let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  if(urlPath === "/") urlPath = "/taskpane.html";
  const filePath = path.normalize(path.join(ROOT, urlPath));
  if(!filePath.startsWith(ROOT)){
    res.writeHead(403); res.end("Forbidden"); return;
  }
  fs.readFile(filePath, (err, data) => {
    if(err){ res.writeHead(404, {"Content-Type": "text/plain; charset=utf-8"}); res.end("404 Not Found: " + urlPath); return; }
    const ext = path.extname(filePath).toLowerCase();
    // WebView(Excelのタスクペイン)に古いjs/cssを使わせないため、毎回検証させる
    res.writeHead(200, {"Content-Type": MIME[ext] || "application/octet-stream", "Cache-Control": "no-cache"});
    res.end(data);
  });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log("工程表ツールのサーバを起動しました: http://localhost:" + PORT);
  console.log("この窓は開いたままにしてください（閉じるとExcelのタスクペインが読み込めなくなります）。");
  console.log("終了するときは Ctrl+C を押してください。");
});
