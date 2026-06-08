/*
 * Minimal HTTPS static file server for local add-in development.
 *
 * Office add-ins must be served over HTTPS, even on localhost. This serves the
 * project folder on https://localhost:3000 using the trusted dev certificate
 * created by `npm run dev-certs` (office-addin-dev-certs).
 *
 * No third-party deps — just Node's https + fs.
 */

const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
  ".xml": "text/xml; charset=utf-8",
  ".svg": "image/svg+xml",
};

function send(res, status, body, headers) {
  res.writeHead(status, Object.assign({ "Access-Control-Allow-Origin": "*" }, headers || {}));
  res.end(body);
}

async function main() {
  // Returns { cert, key } for the locally-trusted dev certificate.
  const httpsOptions = await devCerts.getHttpsServerOptions();

  https
    .createServer(httpsOptions, function (req, res) {
      // Resolve the request path inside ROOT, blocking traversal.
      const urlPath = decodeURIComponent(req.url.split("?")[0]);
      let filePath = path.join(ROOT, urlPath === "/" ? "/index.html" : urlPath);
      if (!filePath.startsWith(ROOT)) {
        return send(res, 403, "Forbidden");
      }

      fs.readFile(filePath, function (err, data) {
        if (err) {
          return send(res, 404, "Not found: " + urlPath);
        }
        const ext = path.extname(filePath).toLowerCase();
        send(res, 200, data, { "Content-Type": MIME[ext] || "application/octet-stream" });
      });
    })
    .listen(PORT, function () {
      console.log("Add-in dev server running at https://localhost:" + PORT + "/");
      console.log("Manifest references this host. Leave it running while you test.");
    });
}

main().catch(function (err) {
  console.error("Failed to start dev server:", err.message);
  console.error('Did you run "npm run dev-certs" first?');
  process.exit(1);
});
