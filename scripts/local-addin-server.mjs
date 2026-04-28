import { createReadStream } from "node:fs";
import { access, readFile, stat } from "node:fs/promises";
import { request as httpRequest } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import path from "node:path";

const DEFAULT_HOST = "127.0.0.1";
const DEFAULT_PORT = 18443;
const DEFAULT_HELPER_HOST = "127.0.0.1";
const DEFAULT_HELPER_PORT = 17926;

function parseArgs(argv) {
  const values = new Map();
  for (let index = 2; index < argv.length; index += 1) {
    const key = argv[index];
    const value = argv[index + 1];
    if (!key?.startsWith("--") || value == null) {
      continue;
    }
    values.set(key, value);
    index += 1;
  }

  return {
    host: values.get("--host") ?? DEFAULT_HOST,
    port: Number(values.get("--port") ?? DEFAULT_PORT),
    root: values.get("--root"),
    cert: values.get("--cert"),
    key: values.get("--key"),
    helperHost: values.get("--helper-host") ?? DEFAULT_HELPER_HOST,
    helperPort: Number(values.get("--helper-port") ?? DEFAULT_HELPER_PORT),
  };
}

function contentTypeFor(filePath) {
  switch (path.extname(filePath).toLowerCase()) {
    case ".html":
      return "text/html; charset=utf-8";
    case ".js":
      return "application/javascript; charset=utf-8";
    case ".css":
      return "text/css; charset=utf-8";
    case ".json":
      return "application/json; charset=utf-8";
    case ".png":
      return "image/png";
    case ".svg":
      return "image/svg+xml";
    case ".ico":
      return "image/x-icon";
    case ".map":
      return "application/json; charset=utf-8";
    default:
      return "application/octet-stream";
  }
}

function sendJson(response, statusCode, payload) {
  const body = JSON.stringify(payload);
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store",
  });
  response.end(body);
}

function safeResolvedPath(root, pathname) {
  const normalized = decodeURIComponent(pathname).replace(/^\/+/, "");
  const resolved = path.resolve(root, normalized || "index.html");
  const relative = path.relative(root, resolved);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    return null;
  }
  return resolved;
}

async function sendStaticFile(root, pathname, response) {
  const resolved = safeResolvedPath(root, pathname === "/" ? "/index.html" : pathname);
  if (!resolved) {
    sendJson(response, 403, { ok: false, message: "Forbidden" });
    return;
  }

  let target = resolved;
  try {
    const info = await stat(target);
    if (info.isDirectory()) {
      target = path.join(target, "index.html");
    }
  } catch {
    if (!path.extname(target)) {
      target = path.join(root, "index.html");
    }
  }

  try {
    await access(target);
  } catch {
    sendJson(response, 404, { ok: false, message: `Not found: ${pathname}` });
    return;
  }

  response.writeHead(200, {
    "Content-Type": contentTypeFor(target),
    "Cache-Control": target.endsWith(".html") ? "no-cache" : "public, max-age=31536000, immutable",
  });
  createReadStream(target).pipe(response);
}

function proxyToHelper(config, request, response, pathname) {
  const proxyPath = pathname.replace(/^\/native-helper/, "") || "/";
  const proxyRequest = httpRequest(
    {
      host: config.helperHost,
      port: config.helperPort,
      path: `${proxyPath}${new URL(request.url, `https://${config.host}`).search}`,
      method: request.method,
      headers: {
        ...request.headers,
        host: `${config.helperHost}:${config.helperPort}`,
      },
    },
    (proxyResponse) => {
      response.writeHead(proxyResponse.statusCode ?? 502, proxyResponse.headers);
      proxyResponse.pipe(response);
    },
  );

  proxyRequest.on("error", (error) => {
    sendJson(response, 502, {
      ok: false,
      message: `本地公式 helper 不可用：${error.message}`,
    });
  });

  request.pipe(proxyRequest);
}

async function main() {
  const config = parseArgs(process.argv);
  if (!config.root || !config.cert || !config.key) {
    console.error("Usage: node scripts/local-addin-server.mjs --root <dir> --cert <pem> --key <pem> [--host 127.0.0.1] [--port 18443] [--helper-port 17926]");
    process.exit(1);
  }

  const [cert, key] = await Promise.all([readFile(config.cert), readFile(config.key)]);
  const server = createHttpsServer({ cert, key }, async (request, response) => {
    const url = new URL(request.url ?? "/", `https://${config.host}:${config.port}`);
    if (url.pathname === "/health") {
      sendJson(response, 200, {
        ok: true,
        host: config.host,
        port: config.port,
        root: config.root,
        helperPort: config.helperPort,
      });
      return;
    }

    if (url.pathname === "/native-helper" || url.pathname.startsWith("/native-helper/")) {
      proxyToHelper(config, request, response, url.pathname);
      return;
    }

    await sendStaticFile(config.root, url.pathname, response);
  });

  server.listen(config.port, config.host, () => {
    console.log(`[SlideSCI local server] https://${config.host}:${config.port}`);
  });

  const shutdown = () => {
    server.close(() => process.exit(0));
  };
  process.on("SIGTERM", shutdown);
  process.on("SIGINT", shutdown);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack ?? error.message : String(error));
  process.exit(1);
});
