"""
SuNaPaNa — Async Email Assistant

Scans for unread emails with subject "SuNaPaNa", processes them with
Claude Code CLI (claude -p), and replies with the AI response.

Runs as a scheduled task via Windows Task Scheduler.
"""

import json
import subprocess
import sys
import os
import re
import logging
from pathlib import Path
from datetime import datetime

import httpx
import msal
import uuid
import hashlib

# ---------------------------------------------------------------------------
# Paths & Config
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).parent
MCP_DIR = Path(r"C:\git\mcp-servers\outlook-email")
CONFIG_PATH = MCP_DIR / "config.json"
TOKEN_CACHE_PATH = MCP_DIR / ".token_cache" / "token_cache.json"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

WORKSPACE_ROOT = Path(r"C:\Users\sbeas\OneDrive\Assistant")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"]
TRIGGER_SUBJECT = "SuNaPaNa"
SIGNATURE_MARKER = "<!-- sunnapana-reply -->"
ALLOWED_SENDERS = {"sbeashwar@outlook.com"}
ALLOWED_RECIPIENTS = {"sbeashwar@outlook.com"}
CLAUDE_PATH = r"C:\Users\sbeas\AppData\Roaming\npm\claude.cmd"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

log_file = LOG_DIR / f"sunnapana_{datetime.now().strftime('%Y%m%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("sunnapana")

# ---------------------------------------------------------------------------
# Auth — dual tokens: sbeashwar (read inbox), SuNaPaNa_ (send replies)
# ---------------------------------------------------------------------------

SUNNAPANA_TOKEN_CACHE_PATH = MCP_DIR / ".token_cache" / "sunnapana_token_cache.json"


def _get_token_for_cache(cache_path: Path) -> str:
    """Acquire a token silently from a given cache file."""
    config = json.loads(CONFIG_PATH.read_text())
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text())

    app = msal.PublicClientApplication(
        config["client_id"],
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        log.error(f"Token expired for cache {cache_path.name}. Re-auth needed.")
        sys.exit(1)

    if cache.has_state_changed:
        cache_path.write_text(cache.serialize())

    return result["access_token"]


def get_token() -> str:
    """Get token for reading inbox (sbeashwar account)."""
    return _get_token_for_cache(TOKEN_CACHE_PATH)


def get_send_token() -> str:
    """Get token for sending replies (SuNaPaNa_ account)."""
    return _get_token_for_cache(SUNNAPANA_TOKEN_CACHE_PATH)

    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------


def graph_get(token: str, endpoint: str, params: dict = None) -> dict:
    with httpx.Client() as client:
        resp = client.get(
            f"{GRAPH_BASE}{endpoint}",
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=30.0,
        )
        resp.raise_for_status()
        return resp.json()


def graph_post(token: str, endpoint: str, body: dict):
    with httpx.Client() as client:
        resp = client.post(
            f"{GRAPH_BASE}{endpoint}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30.0,
        )
        resp.raise_for_status()


def graph_patch(token: str, endpoint: str, body: dict):
    with httpx.Client() as client:
        resp = client.patch(
            f"{GRAPH_BASE}{endpoint}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30.0,
        )
        resp.raise_for_status()


# ---------------------------------------------------------------------------
# Find unread SuNaPaNa emails
# ---------------------------------------------------------------------------


def find_trigger_emails(token: str) -> list[dict]:
    """Find unread emails with the trigger subject."""
    # Use $search for subject text, then filter isRead in code
    # (Graph API $filter on subject is limited for personal accounts)
    params = {
        "$search": f'"subject:{TRIGGER_SUBJECT}"',
        "$top": 10,
        "$select": "id,subject,from,receivedDateTime,body,bodyPreview,isRead",
    }
    data = graph_get(token, "/me/messages", params)
    # Filter to only unread messages with trigger subject
    # Skip our own replies (identified by hidden signature marker in body)
    results = []
    for m in data.get("value", []):
        if m.get("isRead", True):
            continue
        if TRIGGER_SUBJECT.lower() not in m.get("subject", "").lower():
            continue

        # SECURITY: Only process emails from allowed senders
        sender = m.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        if sender not in ALLOWED_SENDERS:
            log.warning(f"[SECURITY] Rejected email from unauthorized sender: {sender} (subject: {m.get('subject', '')})")
            graph_patch(token, f"/me/messages/{m['id']}", {"isRead": True})
            continue

        body_content = m.get("body", {}).get("content", "")
        if SIGNATURE_MARKER in body_content:
            # This is our own reply — mark read and skip
            graph_patch(token, f"/me/messages/{m['id']}", {"isRead": True})
            log.info(f"Skipped own reply {m['id']}")
            continue
        results.append(m)
    return results


# ---------------------------------------------------------------------------
# Extract plain text from HTML body
# ---------------------------------------------------------------------------


def html_to_text(html: str) -> str:
    """Basic HTML to text conversion."""
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"</?p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?div[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<li[^>]*>", "\n- ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&#\d+;", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ---------------------------------------------------------------------------
# Call Claude
# ---------------------------------------------------------------------------


def call_ai(user_message: str, sender: str, subject: str = "") -> str:
    """Send a prompt to Claude Code CLI and return the response.

    Uses claude -p with the Assistant workspace so Claude has access to
    CLAUDE.md, slash commands, MCP servers, and full workspace context.
    Uses --resume to continue thread conversations, falls back to new session.
    """
    # Derive a deterministic session name from the thread subject
    base_subject = re.sub(r"^(re|fw|fwd)\s*:\s*", "", subject, flags=re.IGNORECASE).strip()
    session_name = f"SuNaPaNa: {base_subject[:50]}" if base_subject else "SuNaPaNa"

    prompt = f"""You are replying to an email from {sender} sent to SuNaPaNa.
Format your response as clean HTML suitable for email.
Use <p> for paragraphs, <ul>/<ol> for lists, <strong> for emphasis, <code> for inline code.
Do NOT include <html>, <head>, or <body> tags — just the inner content.
Do NOT use markdown syntax. Output valid HTML only.
Be concise and action-oriented.
Use your web search and email tools to research and provide thorough answers.

The current date is {datetime.now().strftime('%Y-%m-%d')}.

The user's email says:
{user_message}"""

    cmd = [
        CLAUDE_PATH, "-p",
        "--output-format", "text",
        "--dangerously-skip-permissions",
        "--name", session_name,
        "--mcp-config", str(WORKSPACE_ROOT / ".claude" / "settings.local.json"),
        "--allowedTools", "mcp__outlook-email__*", "WebSearch", "WebFetch",
    ]

    log.info(f"Calling Claude session: {session_name}")

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

        response = result.stdout.strip()
        stderr = result.stderr.strip() if result.stderr else ""

        if not response:
            log.error(f"Claude returned empty response. stderr: {stderr[:500]}")
            return None  # Signal failure without sending error email

        return response

    except subprocess.TimeoutExpired:
        log.error("Claude timed out after 600s")
        return None
    except Exception as e:
        log.error(f"AI call failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Reply to email
# ---------------------------------------------------------------------------


def reply_to_message(read_token: str, send_token: str, message_id: str, reply_body: str, sender: str, subject: str):
    """Send a reply from SuNaPaNa_@outlook.com and mark the original as read.

    Uses send_token (SuNaPaNa_ account) to send, read_token (sbeashwar) to mark read.
    SECURITY: Only sends to ALLOWED_RECIPIENTS.
    """
    if sender.lower() not in ALLOWED_RECIPIENTS:
        log.warning(f"[SECURITY] Blocked reply to unauthorized recipient: {sender}")
        graph_patch(read_token, f"/me/messages/{message_id}", {"isRead": True})
        return

    # Send as SuNaPaNa_ (not /reply, since we're a different account)
    tagged_body = f"{SIGNATURE_MARKER}\n{reply_body}"
    graph_post(
        send_token,
        "/me/sendMail",
        {
            "message": {
                "subject": f"Re: {subject}",
                "body": {
                    "contentType": "HTML",
                    "content": tagged_body,
                },
                "toRecipients": [
                    {"emailAddress": {"address": sender}}
                ],
            }
        },
    )
    # Mark original as read in sbeashwar's inbox
    graph_patch(read_token, f"/me/messages/{message_id}", {"isRead": True})



# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------


def main():
    log.info("SuNaPaNa email scan starting...")

    try:
        read_token = get_token()       # sbeashwar — read inbox
        send_token = get_send_token()   # SuNaPaNa_ — send replies
    except Exception as e:
        log.error(f"Auth failed: {e}")
        sys.exit(1)

    emails = find_trigger_emails(read_token)
    log.info(f"Found {len(emails)} unread SuNaPaNa email(s)")

    if not emails:
        return

    for email in emails:
        msg_id = email["id"]
        body_html = email.get("body", {}).get("content", "")
        body_text = html_to_text(body_html) if body_html else email.get("bodyPreview", "")
        sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")

        # Skip empty messages
        if not body_text.strip():
            log.info(f"Skipping empty message {msg_id}")
            graph_patch(read_token, f"/me/messages/{msg_id}", {"isRead": True})
            continue

        log.info(f"Processing email from {sender}: {body_text[:100]}...")

        # Call AI with thread subject for session grouping
        subject = email.get("subject", "")
        ai_response = call_ai(body_text, sender, subject)

        if ai_response is None:
            # AI failed — mark as read to stop the loop, but don't send error reply
            log.error(f"AI failed for message {msg_id} — marking read, no reply sent")
            graph_patch(read_token, f"/me/messages/{msg_id}", {"isRead": True})
            continue

        log.info(f"AI response length: {len(ai_response)} chars")

        # Reply from SuNaPaNa_ account
        try:
            reply_to_message(read_token, send_token, msg_id, ai_response, sender, subject)
            log.info(f"Replied to message {msg_id} from SuNaPaNa_")
            # Mirror the reply into the PWA chat as a server-initiated message
            try:
                from reminders import post_chat_message  # reuse helper
                snippet = ai_response[:200].replace("<", "&lt;")
                post_chat_message(
                    f"📧 Replied to **{sender}** (Subject: {subject}):\n\n{snippet}{'…' if len(ai_response) > 200 else ''}"
                )
            except Exception as e:
                log.warning(f"Failed to mirror reply into chat: {e}")
        except Exception as e:
            log.error(f"Failed to reply to {msg_id}: {e}")

    log.info("SuNaPaNa email scan complete.")

    # Reminder scan — proactive follow-ups (push notifications)
    try:
        from reminders import scan_reminders
        scan_reminders()
    except Exception as e:
        log.error(f"Reminder scan failed: {e}")


if __name__ == "__main__":
    main()
