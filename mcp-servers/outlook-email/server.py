"""
Outlook.com Email MCP Server

Connects to Microsoft Graph API via MSAL device-code flow.
Provides tools to read, search, send, and manage Outlook.com email.

Setup:
  1. Register an app at https://portal.azure.com → App registrations
  2. Set redirect URI to http://localhost
  3. Add delegated permissions: Mail.Read, Mail.ReadWrite, Mail.Send
  4. Copy the Application (client) ID into config.json
  5. Run this server — first launch will prompt device-code auth in terminal
"""

import json
import os
import sys
import asyncio
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import Optional

import httpx
import msal
from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Config & Auth
# ---------------------------------------------------------------------------

SERVER_DIR = Path(__file__).parent
CONFIG_PATH = SERVER_DIR / "config.json"
TOKEN_CACHE_PATH = SERVER_DIR / ".token_cache" / "token_cache.json"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.Read"]


def _load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing {CONFIG_PATH}. Create it with: "
            '{"client_id": "<your-azure-app-client-id>"}'
        )
    with open(CONFIG_PATH) as f:
        return json.load(f)


def _build_msal_app(client_id: str) -> msal.PublicClientApplication:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )
    return app, cache


def _save_cache(cache: msal.SerializableTokenCache):
    TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_PATH.write_text(cache.serialize())


def _acquire_token(app: msal.PublicClientApplication, cache) -> str:
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow failed: {flow}")
        # Print to stderr so it shows in the MCP server logs
        print(
            f"\n{'='*60}\n"
            f"SIGN IN REQUIRED\n"
            f"Go to: {flow['verification_uri']}\n"
            f"Enter code: {flow['user_code']}\n"
            f"{'='*60}\n",
            file=sys.stderr,
            flush=True,
        )
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    _save_cache(cache)
    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------

async def _graph_get(token: str, endpoint: str, params: Optional[dict] = None) -> dict:
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}{endpoint}",
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=30.0,
        )
        resp.raise_for_status()
        return resp.json()


async def _graph_post(token: str, endpoint: str, body: dict) -> dict:
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}{endpoint}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30.0,
        )
        resp.raise_for_status()
        if resp.status_code == 202 or not resp.content:
            return {"status": "success"}
        return resp.json()


async def _graph_patch(token: str, endpoint: str, body: dict) -> dict:
    async with httpx.AsyncClient() as client:
        resp = await client.patch(
            f"{GRAPH_BASE}{endpoint}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30.0,
        )
        resp.raise_for_status()
        if not resp.content:
            return {"status": "success"}
        return resp.json()


def _format_message(msg: dict, include_body: bool = False) -> dict:
    """Extract key fields from a Graph message object."""
    result = {
        "id": msg.get("id", ""),
        "subject": msg.get("subject", "(no subject)"),
        "from": msg.get("from", {}).get("emailAddress", {}).get("address", "unknown"),
        "date": msg.get("receivedDateTime", ""),
        "isRead": msg.get("isRead", False),
        "hasAttachments": msg.get("hasAttachments", False),
    }
    to_list = msg.get("toRecipients", [])
    if to_list:
        result["to"] = [r.get("emailAddress", {}).get("address", "") for r in to_list]
    if include_body:
        body = msg.get("body", {})
        result["bodyPreview"] = msg.get("bodyPreview", "")
        # Return plain text content when possible, fall back to HTML
        if body.get("contentType") == "text":
            result["body"] = body.get("content", "")
        else:
            result["body"] = body.get("content", "")
            result["bodyType"] = "html"
    return result


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp_server = FastMCP(
    "outlook-email",
    instructions=(
        "Access Outlook.com email via Microsoft Graph. "
        "Tools: list inbox, read messages, search email, send email, "
        "manage read status, list folders."
    ),
)

# Auth state (initialized on first tool call)
_config = None
_msal_app = None
_msal_cache = None


def _get_token() -> str:
    global _config, _msal_app, _msal_cache
    if _config is None:
        _config = _load_config()
        _msal_app, _msal_cache = _build_msal_app(_config["client_id"])
    return _acquire_token(_msal_app, _msal_cache)


@mcp_server.tool()
async def list_inbox(
    count: int = 10,
    skip: int = 0,
    unread_only: bool = False,
) -> str:
    """List recent inbox messages. Returns subject, sender, date, read status.

    Args:
        count: Number of messages to return (max 50, default 10)
        skip: Number of messages to skip for pagination (default 0)
        unread_only: If true, only return unread messages
    """
    token = _get_token()
    count = min(count, 50)

    params = {
        "$top": count,
        "$skip": skip,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,toRecipients",
    }
    if unread_only:
        params["$filter"] = "isRead eq false"

    data = await _graph_get(token, "/me/mailFolders/inbox/messages", params)
    messages = [_format_message(m) for m in data.get("value", [])]
    return json.dumps(messages, indent=2)


@mcp_server.tool()
async def read_email(message_id: str, mark_as_read: bool = True) -> str:
    """Read a specific email by its ID. Returns full body content.

    Args:
        message_id: The message ID (from list_inbox or search_email results)
        mark_as_read: Whether to mark the message as read (default true)
    """
    token = _get_token()
    params = {
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,body,bodyPreview,toRecipients,ccRecipients",
    }
    data = await _graph_get(token, f"/me/messages/{message_id}", params)

    if mark_as_read and not data.get("isRead", True):
        await _graph_patch(token, f"/me/messages/{message_id}", {"isRead": True})

    return json.dumps(_format_message(data, include_body=True), indent=2)


@mcp_server.tool()
async def search_email(
    query: str,
    count: int = 10,
) -> str:
    """Search emails using natural language or Outlook search syntax.
    Searches subject, body, sender, and recipients.

    Args:
        query: Search query (e.g. "from:john tax documents", "subject:invoice", "meeting tomorrow")
        count: Max results to return (max 25, default 10)
    """
    token = _get_token()
    count = min(count, 25)
    params = {
        "$search": f'"{query}"',
        "$top": count,
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,toRecipients",
    }
    data = await _graph_get(token, "/me/messages", params)
    messages = []
    for m in data.get("value", []):
        formatted = _format_message(m)
        formatted["bodyPreview"] = m.get("bodyPreview", "")[:200]
        messages.append(formatted)
    return json.dumps(messages, indent=2)


@mcp_server.tool()
async def send_email(
    to: list[str],
    subject: str,
    body: str,
    cc: Optional[list[str]] = None,
    importance: str = "normal",
) -> str:
    """Send an email from the user's Outlook.com account.

    Args:
        to: List of recipient email addresses
        subject: Email subject line
        body: Email body (plain text)
        cc: Optional list of CC email addresses
        importance: Priority: "low", "normal", or "high" (default "normal")
    """
    token = _get_token()

    message = {
        "subject": subject,
        "body": {"contentType": "text", "content": body},
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in to
        ],
        "importance": importance,
    }
    if cc:
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc
        ]

    await _graph_post(token, "/me/sendMail", {"message": message})
    return json.dumps({"status": "sent", "to": to, "subject": subject})


@mcp_server.tool()
async def reply_to_email(
    message_id: str,
    body: str,
    reply_all: bool = False,
) -> str:
    """Reply to an email.

    Args:
        message_id: The message ID to reply to
        body: Reply body (plain text)
        reply_all: If true, reply to all recipients (default false)
    """
    token = _get_token()
    endpoint_action = "replyAll" if reply_all else "reply"
    await _graph_post(
        token,
        f"/me/messages/{message_id}/{endpoint_action}",
        {"comment": body},
    )
    return json.dumps({"status": "replied", "reply_all": reply_all})


@mcp_server.tool()
async def list_mail_folders(top_level_only: bool = True) -> str:
    """List mail folders (Inbox, Sent Items, Drafts, etc.).

    Args:
        top_level_only: If true, only list top-level folders (default true)
    """
    token = _get_token()
    params = {
        "$select": "id,displayName,totalItemCount,unreadItemCount",
        "$top": 50,
    }
    data = await _graph_get(token, "/me/mailFolders", params)
    folders = [
        {
            "id": f["id"],
            "name": f["displayName"],
            "total": f.get("totalItemCount", 0),
            "unread": f.get("unreadItemCount", 0),
        }
        for f in data.get("value", [])
    ]
    return json.dumps(folders, indent=2)


@mcp_server.tool()
async def list_folder_messages(
    folder_id: str,
    count: int = 10,
    skip: int = 0,
) -> str:
    """List messages in a specific mail folder.

    Args:
        folder_id: Folder ID (from list_mail_folders) or well-known name (inbox, sentitems, drafts, junkemail, deleteditems)
        count: Number of messages to return (max 50, default 10)
        skip: Number of messages to skip for pagination
    """
    token = _get_token()
    count = min(count, 50)
    params = {
        "$top": count,
        "$skip": skip,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,toRecipients",
    }
    data = await _graph_get(token, f"/me/mailFolders/{folder_id}/messages", params)
    messages = [_format_message(m) for m in data.get("value", [])]
    return json.dumps(messages, indent=2)


@mcp_server.tool()
async def mark_as_read(message_ids: list[str], is_read: bool = True) -> str:
    """Mark one or more messages as read or unread.

    Args:
        message_ids: List of message IDs to update
        is_read: True to mark as read, False to mark as unread (default true)
    """
    token = _get_token()
    results = []
    for mid in message_ids:
        try:
            await _graph_patch(token, f"/me/messages/{mid}", {"isRead": is_read})
            results.append({"id": mid, "status": "updated"})
        except Exception as e:
            results.append({"id": mid, "status": "error", "error": str(e)})
    return json.dumps(results, indent=2)


# ---------------------------------------------------------------------------
# Calendar tools (Calendars.Read scope)
# ---------------------------------------------------------------------------


def _format_event(evt: dict) -> dict:
    """Extract key fields from a Graph calendar event."""
    attendees = []
    for a in evt.get("attendees", []) or []:
        addr = a.get("emailAddress", {})
        attendees.append({
            "name": addr.get("name", ""),
            "email": addr.get("address", ""),
            "type": a.get("type", "required"),
            "response": a.get("status", {}).get("response", "none"),
        })
    location = evt.get("location", {}) or {}
    return {
        "id": evt.get("id", ""),
        "subject": evt.get("subject", "(no subject)"),
        "start": evt.get("start", {}).get("dateTime", ""),
        "end": evt.get("end", {}).get("dateTime", ""),
        "timezone": evt.get("start", {}).get("timeZone", ""),
        "isAllDay": evt.get("isAllDay", False),
        "location": location.get("displayName", ""),
        "organizer": evt.get("organizer", {}).get("emailAddress", {}).get("address", ""),
        "isOnlineMeeting": evt.get("isOnlineMeeting", False),
        "onlineMeetingUrl": evt.get("onlineMeeting", {}).get("joinUrl") if evt.get("onlineMeeting") else None,
        "responseStatus": evt.get("responseStatus", {}).get("response", "none"),
        "attendees": attendees,
        "bodyPreview": evt.get("bodyPreview", ""),
    }


@mcp_server.tool()
async def list_calendar_events(
    days_ahead: int = 7,
    days_behind: int = 0,
    count: int = 25,
) -> str:
    """List calendar events in a time window around now.

    Args:
        days_ahead: How many days into the future to include (default 7)
        days_behind: How many days into the past to include (default 0, i.e. only future)
        count: Max events to return (max 50, default 25)
    """
    token = _get_token()
    count = min(count, 50)
    now = datetime.now(timezone.utc)
    start = (now.replace(hour=0, minute=0, second=0, microsecond=0)
             - timedelta(days=days_behind)).isoformat().replace("+00:00", "Z")
    end = (now + timedelta(days=days_ahead)).isoformat().replace("+00:00", "Z")
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$top": count,
        "$orderby": "start/dateTime asc",
        "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,isOnlineMeeting,onlineMeeting,responseStatus,bodyPreview",
    }
    # calendarView expands recurring instances; better than /events for "what's on my schedule"
    data = await _graph_get(token, "/me/calendarView", params)
    events = [_format_event(e) for e in data.get("value", [])]
    return json.dumps(events, indent=2)


@mcp_server.tool()
async def search_calendar_events(query: str, count: int = 25) -> str:
    """Search calendar events by subject / body / attendee. Returns recent matches.

    Use this for 'when is my meeting with X' or 'find the offsite event'.

    Args:
        query: Search text (matches subject, body, attendees)
        count: Max results (max 50, default 25)
    """
    token = _get_token()
    count = min(count, 50)
    params = {
        "$search": f'"{query}"',
        "$top": count,
        "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,isOnlineMeeting,onlineMeeting,responseStatus,bodyPreview",
    }
    # /me/events with $search requires a ConsistencyLevel header? Actually Outlook
    # personal endpoints accept $search without it. If we ever hit issues, add:
    #   headers={"ConsistencyLevel": "eventual"}
    data = await _graph_get(token, "/me/events", params)
    events = [_format_event(e) for e in data.get("value", [])]
    return json.dumps(events, indent=2)


@mcp_server.tool()
async def get_calendar_event(event_id: str) -> str:
    """Get full details of a single calendar event by ID.

    Args:
        event_id: Event ID (from list_calendar_events or search_calendar_events)
    """
    token = _get_token()
    data = await _graph_get(token, f"/me/events/{event_id}")
    evt = _format_event(data)
    # Add full body for single-event fetch
    body = data.get("body", {})
    if body:
        evt["bodyType"] = body.get("contentType", "html")
        evt["body"] = body.get("content", "")
    return json.dumps(evt, indent=2)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp_server.run(transport="stdio")
