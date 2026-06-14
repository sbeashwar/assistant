import express from "express";
import { execSync } from "child_process";
import { readFileSync, readdirSync, existsSync, statSync } from "fs";
import { join, resolve as resolvePath } from "path";
import { PublicClientApplication } from "@azure/msal-node";

const PORT = parseInt(process.env.PORT || "3100", 10);
const API_SECRET = process.env.TOOL_API_SECRET || "";

const app = express();
app.use(express.json({ limit: "1mb" }));

// Health check — public (no auth required)
app.get("/health", (_req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// Auth middleware — if TOOL_API_SECRET is set, require it via Bearer token
app.use((req, res, next) => {
  if (API_SECRET) {
    const auth = req.headers.authorization;
    if (!auth || auth !== `Bearer ${API_SECRET}`) {
      return res.status(401).json({ error: "Unauthorized" });
    }
  }
  next();
});

// Tool execution endpoint
app.post("/tools/:toolName", async (req, res) => {
  const { toolName } = req.params;
  const args = req.body;

  console.log(`[tool] ${toolName}(${JSON.stringify(args)})`);

  try {
    const result = await execToolAsync(toolName, args);
    res.json({ result });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(`[tool] error: ${msg}`);
    res.status(500).json({ error: msg });
  }
});

// ---------------------------------------------------------------------------
// Tool implementations (same as github-models.ts)
// ---------------------------------------------------------------------------

const BLOCKED_PATTERNS =
  /\b(rm\s+-rf|del\s+\/[sqf]|format\s+[a-z]:|rmdir|remove-item.*-recurse.*-force|stop-process|shutdown|restart-computer|clear-recyclebin)\b/i;

function execTool(name, args) {
  switch (name) {
    case "search_files":
      return toolSearchFiles(
        args.directory,
        args.pattern,
        args.maxResults || 20
      );
    case "list_directory":
      return toolListDirectory(args.path);
    case "read_file":
      return toolReadFile(args.path);
    case "run_command":
      return toolRunCommand(args.command);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

function toolSearchFiles(directory, pattern, maxResults) {
  const dir = resolvePath(directory);
  if (!existsSync(dir)) return `Directory not found: ${dir}`;

  const results = [];
  const lowerPattern = pattern.toLowerCase();

  function walk(d, depth) {
    if (depth > 8 || results.length >= maxResults) return;
    let entries;
    try {
      entries = readdirSync(d, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      if (results.length >= maxResults) return;
      const fullPath = join(d, entry.name);
      if (entry.name.toLowerCase().includes(lowerPattern)) {
        results.push(fullPath);
      }
      if (
        entry.isDirectory() &&
        !entry.name.startsWith(".") &&
        entry.name !== "node_modules"
      ) {
        walk(fullPath, depth + 1);
      }
    }
  }

  walk(dir, 0);
  if (results.length === 0)
    return `No files matching '${pattern}' found in ${dir}`;
  return results.join("\n");
}

function toolListDirectory(path) {
  const dir = resolvePath(path);
  if (!existsSync(dir)) return `Directory not found: ${dir}`;
  const entries = readdirSync(dir, { withFileTypes: true });
  return entries.map((e) => (e.isDirectory() ? e.name + "/" : e.name)).join("\n");
}

function toolReadFile(path) {
  const file = resolvePath(path);
  if (!existsSync(file)) return `File not found: ${file}`;
  const stat = statSync(file);
  if (stat.size > 5 * 1024 * 1024)
    return `File too large (${(stat.size / 1024 / 1024).toFixed(1)} MB). Use run_command to inspect.`;
  const content = readFileSync(file, "utf-8");
  const lines = content.split("\n");
  if (lines.length > 500) {
    return (
      lines.slice(0, 500).join("\n") +
      `\n\n... (truncated, ${lines.length} total lines)`
    );
  }
  return content;
}

function toolRunCommand(command) {
  if (BLOCKED_PATTERNS.test(command)) {
    return "Blocked: this command is potentially destructive and not allowed.";
  }
  try {
    const output = execSync(
      `powershell.exe -NoProfile -NonInteractive -Command ${JSON.stringify(command)}`,
      {
        timeout: 30_000,
        encoding: "utf-8",
        maxBuffer: 512 * 1024,
        windowsHide: true,
      }
    );
    const trimmed = output.trim();
    if (trimmed.length > 10_000) {
      return trimmed.slice(0, 10_000) + "\n... (truncated)";
    }
    return trimmed || "(no output)";
  } catch (e) {
    const msg = e instanceof Error ? e.stderr || e.message : String(e);
    return `Command failed: ${msg}`;
  }
}

// ---------------------------------------------------------------------------
// Email tools (Microsoft Graph via MSAL token cache)
// ---------------------------------------------------------------------------

const MSAL_CONFIG_PATH = "C:\\mcp-servers\\outlook-email\\config.json";
const TOKEN_CACHE_PATH = "C:\\mcp-servers\\outlook-email\\.token_cache\\token_cache.json";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MAIL_SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"];

let msalApp = null;

function getMsalApp() {
  if (msalApp) return msalApp;
  const config = JSON.parse(readFileSync(MSAL_CONFIG_PATH, "utf-8"));
  msalApp = new PublicClientApplication({
    auth: {
      clientId: config.client_id,
      authority: "https://login.microsoftonline.com/consumers",
    },
  });
  // Load cached tokens
  if (existsSync(TOKEN_CACHE_PATH)) {
    const cache = readFileSync(TOKEN_CACHE_PATH, "utf-8");
    msalApp.getTokenCache().deserialize(cache);
  }
  return msalApp;
}

async function getMailToken() {
  const app = getMsalApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) throw new Error("No cached accounts. Run MCP server auth first.");
  const result = await app.acquireTokenSilent({
    account: accounts[0],
    scopes: MAIL_SCOPES,
  });
  return result.accessToken;
}

async function graphGet(endpoint, params = {}) {
  const token = await getMailToken();
  const url = new URL(`${GRAPH_BASE}${endpoint}`);
  for (const [k, v] of Object.entries(params)) {
    url.searchParams.set(k, v);
  }
  const res = await fetch(url.toString(), {
    headers: { Authorization: `Bearer ${token}` },
    signal: AbortSignal.timeout(55_000),
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPost(endpoint, body) {
  const token = await getMailToken();
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(55_000),
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.status === 202 ? {} : res.json().catch(() => ({}));
}

// Strip HTML tags for readable email content
function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/?p[^>]*>/gi, "\n")
    .replace(/<\/?div[^>]*>/gi, "\n")
    .replace(/<li[^>]*>/gi, "\n- ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#\d+;/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function toolListInbox(count) {
  const n = Math.min(count || 10, 25);
  const data = await graphGet("/me/messages", {
    $top: String(n),
    $select: "id,subject,from,receivedDateTime,isRead,bodyPreview",
    $orderby: "receivedDateTime desc",
  });
  const msgs = data.value || [];
  if (msgs.length === 0) return "Inbox is empty.";
  return msgs.map((m, i) => {
    const from = m.from?.emailAddress?.address || "unknown";
    const date = new Date(m.receivedDateTime).toLocaleString();
    const read = m.isRead ? "" : " [UNREAD]";
    return `${i + 1}. ${m.subject}${read}\n   From: ${from} | ${date}\n   ID: ${m.id}\n   Preview: ${(m.bodyPreview || "").slice(0, 120)}`;
  }).join("\n\n");
}

async function toolReadEmail(messageId) {
  const data = await graphGet(`/me/messages/${messageId}`, {
    $select: "subject,from,toRecipients,receivedDateTime,body,bodyPreview",
  });
  const from = data.from?.emailAddress?.address || "unknown";
  const to = (data.toRecipients || []).map(r => r.emailAddress?.address).join(", ");
  const body = data.body?.contentType === "html" ? htmlToText(data.body.content) : (data.body?.content || data.bodyPreview);
  return `Subject: ${data.subject}\nFrom: ${from}\nTo: ${to}\nDate: ${new Date(data.receivedDateTime).toLocaleString()}\n\n${body}`;
}

async function toolSearchEmail(query, count) {
  const n = Math.min(count || 10, 25);
  const data = await graphGet("/me/messages", {
    $search: `"${query}"`,
    $top: String(n),
    $select: "id,subject,from,receivedDateTime,isRead,bodyPreview",
  });
  const msgs = data.value || [];
  if (msgs.length === 0) return `No emails found for: ${query}`;
  return msgs.map((m, i) => {
    const from = m.from?.emailAddress?.address || "unknown";
    const date = new Date(m.receivedDateTime).toLocaleString();
    return `${i + 1}. ${m.subject}\n   From: ${from} | ${date}\n   ID: ${m.id}\n   Preview: ${(m.bodyPreview || "").slice(0, 120)}`;
  }).join("\n\n");
}

async function toolSendEmail(to, subject, body) {
  await graphPost("/me/sendMail", {
    message: {
      subject,
      body: { contentType: "Text", content: body },
      toRecipients: [{ emailAddress: { address: to } }],
    },
  });
  return `Email sent to ${to}: "${subject}"`;
}

async function toolReplyEmail(messageId, body) {
  await graphPost(`/me/messages/${messageId}/reply`, {
    comment: body,
  });
  return "Reply sent.";
}

// Make tool execution async for email tools
async function execToolAsync(name, args) {
  switch (name) {
    case "search_files":
      return toolSearchFiles(args.directory, args.pattern, args.maxResults || 20);
    case "list_directory":
      return toolListDirectory(args.path);
    case "read_file":
      return toolReadFile(args.path);
    case "run_command":
      return toolRunCommand(args.command);
    case "list_inbox":
      return await toolListInbox(args.count);
    case "read_email":
      return await toolReadEmail(args.messageId);
    case "search_email":
      return await toolSearchEmail(args.query, args.count);
    case "send_email":
      return await toolSendEmail(args.to, args.subject, args.body);
    case "reply_email":
      return await toolReplyEmail(args.messageId, args.body);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// ---------------------------------------------------------------------------

app.listen(PORT, () => {
  console.log(`[local-tools] listening on http://localhost:${PORT}`);
  console.log(`[local-tools] auth: ${API_SECRET ? "enabled" : "DISABLED (set TOOL_API_SECRET)"}`);
});
