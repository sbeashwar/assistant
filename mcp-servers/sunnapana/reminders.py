"""
SuNaPaNa Reminders — proactive follow-up scanner

Walks `notes/reminders/*.md`, asks Claude (via `claude -p`) to judge
whether each active reminder's trigger condition has been met since
`last_checked_iso`, and fires a Web Push notification through the
`/api/push/fire` endpoint when it has. Idempotent: bumps
`last_checked_iso` on every tick, sets `status: done` for `oneshot`
reminders that fire.

This is invoked from `sunnapana.py:main()` after the email-reply loop.
"""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import httpx

log = logging.getLogger("sunnapana.reminders")

# ---------------------------------------------------------------------------
# Paths & config
# ---------------------------------------------------------------------------

WORKSPACE_ROOT = Path(r"C:\Users\sbeas\OneDrive\Assistant")
REMINDERS_DIR = WORKSPACE_ROOT / "notes" / "reminders"

# Web Push fan-out endpoint (fire everywhere). Token + URL live in
# C:\git\mcp-servers\outlook-email\push_config.json so the scheduled task
# (which runs python.exe directly with no env-var setup) can pick them up.
# Env vars override if set, for ad-hoc local testing.
_PUSH_CONFIG_PATH = Path(r"C:\git\mcp-servers\outlook-email\push_config.json")


def _load_push_config() -> dict:
    if _PUSH_CONFIG_PATH.exists():
        try:
            return json.loads(_PUSH_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning(f"Cannot read {_PUSH_CONFIG_PATH}: {e}")
    return {}


_PUSH_CFG = _load_push_config()
PUSH_FIRE_URL = os.environ.get("PUSH_FIRE_URL") or _PUSH_CFG.get("fire_url") or "http://localhost:3000/api/push/fire"
ASSISTANT_FIRE_URL = (
    os.environ.get("ASSISTANT_FIRE_URL")
    or _PUSH_CFG.get("assistant_fire_url")
    or "http://localhost:3000/api/assistant/fire"
)
PUSH_TOKEN = os.environ.get("PUSH_INTERNAL_TOKEN") or _PUSH_CFG.get("token") or ""

CLAUDE_PATH = r"C:\Users\sbeas\AppData\Roaming\npm\claude.cmd"

# ---------------------------------------------------------------------------
# Frontmatter parser (no PyYAML dependency — keep it simple and explicit)
# ---------------------------------------------------------------------------

_FRONT_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n(.*)$", re.DOTALL)


def parse_reminder(path: Path) -> Optional[dict]:
    """Parse a reminder markdown file. Returns dict with `meta` + `body` keys,
    or None if the file is malformed or not active.
    """
    try:
        raw = path.read_text(encoding="utf-8")
    except Exception as e:
        log.warning(f"Cannot read {path.name}: {e}")
        return None

    m = _FRONT_RE.match(raw)
    if not m:
        log.warning(f"{path.name}: no YAML frontmatter")
        return None

    meta = {}
    for line in m.group(1).splitlines():
        line = line.split("#", 1)[0].rstrip()  # strip inline comments
        if not line.strip():
            continue
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        meta[key.strip()] = val.strip().strip('"').strip("'")

    return {"meta": meta, "body": m.group(2), "path": path}


def write_reminder(reminder: dict) -> None:
    """Round-trip a reminder back to disk with updated metadata."""
    meta = reminder["meta"]
    body = reminder["body"]
    lines = ["---"]
    for k, v in meta.items():
        lines.append(f"{k}: {v}")
    lines.append("---")
    lines.append("")
    reminder["path"].write_text("\n".join(lines) + body, encoding="utf-8")


# ---------------------------------------------------------------------------
# LLM judge
# ---------------------------------------------------------------------------

JUDGE_PROMPT_TEMPLATE = """You are a reminder judge. Decide whether the trigger condition
described below has been met at any point since {last_checked_iso}. Use whatever
tools you have available — inbox search, web search, URL fetch, etc. — to check.

Today is {today_iso}.

=== REMINDER TRIGGER ===
{body}
=== END REMINDER TRIGGER ===

Respond with a single line of strict JSON, no prose, no markdown:

{{"fired": <true|false>, "summary": "<≤140 chars, what to put in the push notification body, or empty string if not fired>"}}

Examples of valid responses:
{{"fired": false, "summary": ""}}
{{"fired": true, "summary": "Ride1Up Portola is $300 off until May 15 — promo email from Ride1Up just arrived."}}

Be conservative: if you cannot confirm the condition with high confidence, return fired=false.
"""


def judge_reminder(reminder: dict) -> Optional[dict]:
    """Ask Claude to evaluate the reminder. Returns parsed {fired, summary} or None on failure."""
    meta = reminder["meta"]
    last_checked = meta.get("last_checked_iso", "1970-01-01T00:00:00Z")
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    prompt = JUDGE_PROMPT_TEMPLATE.format(
        last_checked_iso=last_checked,
        today_iso=today,
        body=reminder["body"].strip(),
    )

    cmd = [
        CLAUDE_PATH, "-p",
        "--output-format", "text",
        "--dangerously-skip-permissions",
        "--mcp-config", str(WORKSPACE_ROOT / ".claude" / "settings.local.json"),
        "--allowedTools", "mcp__outlook-email__*", "WebSearch", "WebFetch",
    ]

    log.info(f"Judging reminder: {meta.get('id', reminder['path'].stem)}")
    try:
        result = subprocess.run(
            cmd,
            input=prompt,
            capture_output=True,
            text=True,
            timeout=600,
            encoding="utf-8",
            cwd=str(WORKSPACE_ROOT),
        )
    except subprocess.TimeoutExpired:
        log.error(f"Claude timed out judging {meta.get('id')}")
        return None
    except Exception as e:
        log.error(f"Judge call failed for {meta.get('id')}: {e}")
        return None

    raw = (result.stdout or "").strip()
    if not raw:
        log.error(f"Empty judge response for {meta.get('id')}: stderr={result.stderr[:300]}")
        return None

    # Find the first JSON object on its own line
    json_str = None
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            json_str = line
            break
    if not json_str:
        # Try a non-strict extraction
        m = re.search(r"\{.*?\}", raw, re.DOTALL)
        json_str = m.group(0) if m else None

    if not json_str:
        log.error(f"No JSON in judge response for {meta.get('id')}: {raw[:300]}")
        return None

    try:
        verdict = json.loads(json_str)
    except json.JSONDecodeError as e:
        log.error(f"Invalid JSON from judge for {meta.get('id')}: {e}; raw={json_str[:300]}")
        return None

    if not isinstance(verdict, dict) or "fired" not in verdict:
        log.error(f"Malformed verdict for {meta.get('id')}: {verdict}")
        return None

    return verdict


# ---------------------------------------------------------------------------
# Push fan-out
# ---------------------------------------------------------------------------


def fire_push(title: str, body: str, url: str = "/", tag: str = "") -> bool:
    """POST to the local /api/push/fire endpoint. Returns True if accepted."""
    if not PUSH_TOKEN:
        log.error("PUSH_INTERNAL_TOKEN not set; cannot fire push")
        return False

    payload = {"title": title, "body": body, "url": url}
    if tag:
        payload["tag"] = tag

    try:
        with httpx.Client(timeout=15.0) as client:
            resp = client.post(
                PUSH_FIRE_URL,
                headers={
                    "Authorization": f"Bearer {PUSH_TOKEN}",
                    "Content-Type": "application/json",
                },
                json=payload,
            )
        if resp.status_code != 200:
            log.error(f"Push fire failed: {resp.status_code} {resp.text[:200]}")
            return False
        data = resp.json()
        log.info(f"Push fired: sent={data.get('sent')} failed={data.get('failed')} removed={data.get('removed')}")
        return True
    except Exception as e:
        log.error(f"Push fire error: {e}")
        return False


def post_chat_message(content: str, conversation_id: Optional[str] = None) -> bool:
    """POST a server-initiated assistant message to the PWA chat. Failure is
    logged but never fatal — the push is the user-facing surface; the chat
    entry is supplementary context.
    """
    if not PUSH_TOKEN:
        log.warning("PUSH_INTERNAL_TOKEN not set; skipping chat post")
        return False
    payload: dict = {"content": content, "push": False}
    if conversation_id:
        payload["conversationId"] = conversation_id
    try:
        with httpx.Client(timeout=15.0) as client:
            resp = client.post(
                ASSISTANT_FIRE_URL,
                headers={
                    "Authorization": f"Bearer {PUSH_TOKEN}",
                    "Content-Type": "application/json",
                },
                json=payload,
            )
        if resp.status_code != 200:
            log.warning(f"Chat post failed: {resp.status_code} {resp.text[:200]}")
            return False
        return True
    except Exception as e:
        log.warning(f"Chat post error: {e}")
        return False


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------


def scan_reminders() -> None:
    """Iterate every active reminder, judge, fire, and persist."""
    if not REMINDERS_DIR.exists():
        log.info(f"No reminders directory at {REMINDERS_DIR}; skipping")
        return

    files = sorted(REMINDERS_DIR.glob("*.md"))
    if not files:
        log.info("No reminders to scan")
        return

    log.info(f"Scanning {len(files)} reminder file(s)")
    for path in files:
        reminder = parse_reminder(path)
        if not reminder:
            continue
        meta = reminder["meta"]
        if meta.get("status", "active").lower() != "active":
            continue

        verdict = judge_reminder(reminder)
        # Always bump last_checked, even on judge failure, to avoid hammering
        meta["last_checked_iso"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        if verdict and verdict.get("fired"):
            summary = (verdict.get("summary") or "").strip() or "Reminder triggered"
            title = meta.get("notify_title") or meta.get("title") or "Reminder"
            url = meta.get("notify_url") or "/"
            tag = f"reminder-{meta.get('id', path.stem)}"
            ok = fire_push(title=title, body=summary, url=url, tag=tag)
            # Also drop the summary into the PWA chat as an assistant message
            # so it shows up in conversation history, not just the lock screen.
            chat_text = f"🔔 **{title}**\n\n{summary}"
            post_chat_message(chat_text)
            if ok and meta.get("mode", "recurring").lower() == "oneshot":
                meta["status"] = "done"
                log.info(f"Reminder {meta.get('id')} fired (oneshot) — marked done")
            elif ok:
                log.info(f"Reminder {meta.get('id')} fired (recurring) — staying active")

        write_reminder(reminder)


if __name__ == "__main__":
    # Allow standalone execution for testing: `python reminders.py`
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    scan_reminders()
