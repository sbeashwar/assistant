"""
Reminders scanner — runs every N minutes, asks claude -p to judge each active
reminder's trigger condition, and fires a Web Push when fired=true.

Triggers can describe any condition: "an email from Delta arrives",
"the price of MSFT drops below $380", "the user has a calendar event titled
'1:1 with X' added", "the file C:\\reports\\daily.csv is updated".

Claude has the full Assistant workspace context (CLAUDE.md, skills, slash
commands) plus the outlook-email MCP server with all 11 tools (inbox, search,
folders, calendar). Add more MCPs in .claude/settings.local.json and they
become available here automatically.

Notification flow:
  fired=true  →  POST to push_config.json's fire_url
              →  Web Push fanout to subscribed devices
              →  user taps notification, opens /assistant?c=<conv> deep link
"""

import json
import logging
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import requests

# ---------------------------------------------------------------------------
# Paths & Config
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).parent
WORKSPACE_ROOT = Path(r"C:\Users\sbeas\OneDrive\Assistant")
REMINDERS_DIR = WORKSPACE_ROOT / "notes" / "reminders"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Push config lives with the outlook-email server so all scanners share it
_PUSH_CONFIG_PATH = Path(r"C:\git\mcp-servers\outlook-email\push_config.json")

CLAUDE_PATH = r"C:\Users\sbeas\AppData\Roaming\npm\claude.cmd"
MCP_CONFIG_PATH = WORKSPACE_ROOT / ".claude" / "settings.local.json"

# claude -p is allowed to use these tools when judging triggers.
# Wildcard for the outlook-email MCP gives it all 11 email + calendar tools.
ALLOWED_TOOLS = (
    "mcp__outlook-email__*,"
    "Bash,Read,Glob,Grep,WebSearch,WebFetch"
)

# All reminder judgments share ONE on-disk claude session, deterministic UUID.
# Why one shared session instead of one-per-reminder:
#   - Session list stays at exactly 1 entry ("Reminders") forever, no matter
#     how many reminders exist or how many ticks fire. Clean.
#   - The judge sees prior-tick context across all reminders. Mostly helpful
#     ("I already checked email this tick for reminder A, here's what was
#     there") — the prompt always names which reminder it's judging now.
#   - Single JSONL grows over time but only one file to manage.
# The session is NEVER deleted — it's a long-lived workhorse like the
# AssistantKeepAlive task. If it grows too large, delete it manually and
# the next tick creates a fresh one with the same ID.
SHARED_SESSION_ID = "a8c79b51-8a3f-4bcb-8a7e-2f5e5f4a3b21"
SHARED_SESSION_NAME = "Reminders"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

log = logging.getLogger("reminders")
log.setLevel(logging.INFO)
log_file = LOG_DIR / "reminders.log"
fh = logging.FileHandler(log_file, encoding="utf-8")
fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
log.addHandler(fh)
log.addHandler(logging.StreamHandler(sys.stderr))


# ---------------------------------------------------------------------------
# Frontmatter parsing (minimal — avoid the PyYAML dep)
# ---------------------------------------------------------------------------

_FRONTMATTER_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n(.*)$", re.DOTALL)


def parse_frontmatter(text: str) -> tuple[dict, str]:
    """Parse 'key: value' frontmatter. Returns (metadata, body)."""
    m = _FRONTMATTER_RE.match(text)
    if not m:
        return {}, text
    meta_text, body = m.groups()
    meta = {}
    for line in meta_text.splitlines():
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        meta[key.strip()] = val.strip()
    return meta, body


def serialize_frontmatter(meta: dict, body: str) -> str:
    lines = ["---"]
    for k, v in meta.items():
        lines.append(f"{k}: {v}")
    lines.append("---")
    lines.append("")
    return "\n".join(lines) + body


# ---------------------------------------------------------------------------
# Push fan-out
# ---------------------------------------------------------------------------

def _load_push_config() -> dict:
    if not _PUSH_CONFIG_PATH.exists():
        log.warning(f"No push config at {_PUSH_CONFIG_PATH}; notifications disabled")
        return {}
    return json.loads(_PUSH_CONFIG_PATH.read_text(encoding="utf-8"))


def fire_push(title: str, body: str, url: str = "/") -> bool:
    cfg = _load_push_config()
    fire_url = cfg.get("fire_url")
    token = cfg.get("token")
    if not fire_url or not token:
        log.warning("Push not configured (missing fire_url or token); skipping")
        return False
    try:
        resp = requests.post(
            fire_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"title": title, "body": body, "url": url, "tag": "reminder"},
            timeout=10,
        )
        resp.raise_for_status()
        log.info(f"Push fired: {resp.json()}")
        return True
    except Exception as e:
        log.error(f"Push fire failed: {e}")
        return False


# ---------------------------------------------------------------------------
# Claude judge
# ---------------------------------------------------------------------------

# The first line is what shows up in the VS Code session list. Keep it short
# and reminder-specific so users can recognize the entry at a glance.
JUDGE_PROMPT = """Reminder: {title}

You are evaluating whether this reminder's trigger condition is currently met.

REMINDER:
  id: {rid}
  title: {title}
  last_checked: {last_checked}

TRIGGER (in user's words):
{trigger_text}

Your job:
1. Use any tools you need (email, calendar, web search, file reads) to gather evidence.
2. If you check email: only consider items received AFTER {last_checked} — earlier matches
   already triggered or were intentionally ignored.
3. Be conservative. Default to fired=false unless you have positive evidence the
   condition is met.
4. Respond with a single JSON object on the LAST line of your output, NOTHING after it:

   {{"fired": <true|false>, "summary": "<one-line summary, max 100 chars, used as push body>", "evidence": "<2-3 sentences describing what you found, for the log>"}}

Do not include any text after the JSON object.
"""


def judge_reminder(reminder: dict) -> Optional[dict]:
    """Spawn claude -p, return parsed JSON verdict or None on failure."""
    prompt = JUDGE_PROMPT.format(
        rid=reminder["id"],
        title=reminder["title"],
        last_checked=reminder["last_checked_iso"],
        trigger_text=reminder["trigger_text"],
    )

    cmd = [
        CLAUDE_PATH,
        "-p",
        "--session-id", SHARED_SESSION_ID,
        "--name", SHARED_SESSION_NAME,
        "--mcp-config", str(MCP_CONFIG_PATH),
        "--allowedTools", ALLOWED_TOOLS,
        "--dangerously-skip-permissions",
        prompt,
    ]

    log.info(f"Judging reminder: {reminder['id']}")
    try:
        result = subprocess.run(
            cmd,
            cwd=str(WORKSPACE_ROOT),
            capture_output=True,
            text=True,
            timeout=600,
            shell=False,
            creationflags=0x08000000,  # CREATE_NO_WINDOW on Windows
        )
    except subprocess.TimeoutExpired:
        log.error(f"claude -p timed out for {reminder['id']}")
        return None
    except Exception as e:
        log.error(f"claude -p spawn failed for {reminder['id']}: {e}")
        return None

    if result.returncode != 0:
        log.error(f"claude -p exit {result.returncode} for {reminder['id']}: {result.stderr[:500]}")
        return None

    # Last line of stdout should be the JSON verdict
    output = result.stdout.strip()
    last_line = output.rsplit("\n", 1)[-1].strip()
    try:
        verdict = json.loads(last_line)
        if "fired" not in verdict:
            raise ValueError("missing 'fired' key")
        return verdict
    except Exception as e:
        log.error(f"Could not parse verdict for {reminder['id']}: {e}\nLast line was: {last_line[:200]}")
        return None


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def load_active_reminders() -> list[dict]:
    if not REMINDERS_DIR.exists():
        return []
    out = []
    for f in sorted(REMINDERS_DIR.glob("*.md")):
        text = f.read_text(encoding="utf-8")
        meta, body = parse_frontmatter(text)
        if meta.get("status", "active") != "active":
            continue
        # Extract Trigger section from body
        trig_match = re.search(r"##\s*Trigger\s*\n+(.+?)(?:\n##\s|\Z)", body, re.DOTALL)
        trigger_text = trig_match.group(1).strip() if trig_match else body.strip()
        out.append({
            "path": f,
            "id": meta.get("id", f.stem),
            "title": meta.get("title", f.stem),
            "mode": meta.get("mode", "oneshot"),
            "status": meta.get("status", "active"),
            "last_checked_iso": meta.get("last_checked_iso", "1970-01-01T00:00:00Z"),
            "notify_title": meta.get("notify_title", meta.get("title", "Reminder")),
            "notify_url": meta.get("notify_url", "/"),
            "trigger_text": trigger_text,
            "meta": meta,
            "body": body,
        })
    return out


def update_reminder(reminder: dict, *, fired: bool):
    """Bump last_checked_iso; set status=done if fired and mode=oneshot."""
    meta = reminder["meta"]
    meta["last_checked_iso"] = now_iso()
    if fired and reminder["mode"] == "oneshot":
        meta["status"] = "done"
    reminder["path"].write_text(
        serialize_frontmatter(meta, reminder["body"]),
        encoding="utf-8",
    )


def main():
    log.info("=== reminders tick ===")
    reminders = load_active_reminders()
    log.info(f"Active reminders: {len(reminders)}")

    for r in reminders:
        verdict = judge_reminder(r)
        if verdict is None:
            # On judge failure, still bump last_checked to avoid hammering
            update_reminder(r, fired=False)
            continue

        fired = bool(verdict.get("fired"))
        summary = verdict.get("summary", r["title"])[:200]
        evidence = verdict.get("evidence", "")
        log.info(f"  {r['id']} → fired={fired}  summary={summary!r}  evidence={evidence!r}")

        if fired:
            fire_push(
                title=r["notify_title"],
                body=summary,
                url=r["notify_url"],
            )

        update_reminder(r, fired=fired)

    log.info("=== tick complete ===")


if __name__ == "__main__":
    main()
