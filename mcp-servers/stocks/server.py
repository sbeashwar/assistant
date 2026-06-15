"""
Stocks MCP Server

Thin HTTP wrapper around the Next.js stocks miniapp routes at
http://localhost:3000/api/stocks/* so that any surface running `claude -p`
against the OneDrive workspace (Claude Code in VS Code AND the PWA chat at
https://assistant.tail4621a0.ts.net) gets live quotes, fundamentals, news,
options, watchlist, and sector data.

Why wrap, not duplicate: the miniapp already uses yahoo-finance2 with real-time,
free, no-API-key data. Wrapping its routes (vs reimplementing in Python) means
the PWA UI and the chat always see the same numbers — no drift risk.

Requires the Next.js app to be running on localhost:3000 (it always is, per
AssistantKeepAlive). Override via STOCKS_API_BASE env var if needed.
"""

import json
import os
from typing import Optional

import httpx
from mcp.server.fastmcp import FastMCP

BASE_URL = os.environ.get("STOCKS_API_BASE", "http://localhost:3000")
REQUEST_TIMEOUT = 30.0


async def _get(path: str, params: Optional[dict] = None) -> dict:
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{BASE_URL}{path}",
            params=params,
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp_server = FastMCP(
    "stocks",
    instructions=(
        "Live stock market data backed by the MySuperApp stocks miniapp. "
        "Same data the user sees in the PWA stocks page. "
        "Tools: quotes, fundamentals, news, options chains, symbol search, "
        "watchlist, sector assignments. Read-only — chat cannot mutate the "
        "watchlist or sectors today; user does that in the UI."
    ),
)


@mcp_server.tool()
async def get_quotes(symbols: str) -> str:
    """Get current price + day's change/high/low for one or more tickers.

    Args:
        symbols: Comma-separated tickers (e.g. "AAPL,MSFT,NVDA"). Case-insensitive.

    Returns JSON: { quotes: { SYMBOL: { price, change, changePercent, high,
    low, open, previousClose } } }
    """
    data = await _get("/api/stocks/quotes", params={"symbols": symbols})
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def get_fundamentals(symbols: str) -> str:
    """Get fundamentals (52-week range, P/E, EPS, market cap, dividend yield,
    beta, average volume, debt/equity) for one or more tickers.

    Args:
        symbols: Comma-separated tickers (e.g. "AAPL,MSFT").

    Returns JSON keyed by symbol. Fields may be null if Yahoo doesn't report
    them for that security (common for ETFs).
    """
    data = await _get("/api/stocks/fundamentals", params={"symbols": symbols})
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def get_news(symbol: str) -> str:
    """Get the latest 8 news articles for a ticker (title, publisher, link,
    publishedAt ISO timestamp, thumbnail URL).

    Args:
        symbol: A single ticker (e.g. "AAPL").
    """
    data = await _get("/api/stocks/news", params={"symbol": symbol})
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def search_symbols(query: str) -> str:
    """Search for ticker symbols by company name or partial symbol.

    Args:
        query: Free text (e.g. "Microsoft", "rivian", "MSFT").

    Returns JSON: { results: [{ symbol, name }], mock?: bool }. `mock: true`
    means no Finnhub API key is configured and a small static fallback was
    used — fine for common names, won't find obscure tickers.
    """
    data = await _get("/api/stocks/search", params={"q": query})
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def get_watchlist() -> str:
    """Get the user's saved watchlist from the stocks miniapp. Read-only —
    the user adds/removes tickers in the PWA UI."""
    data = await _get("/api/stocks/watchlist")
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def get_sectors() -> str:
    """Get the user's sector assignments (which sector each watchlist symbol
    is filed under) plus the full list of available sector names."""
    data = await _get("/api/stocks/sectors")
    return json.dumps(data, indent=2)


@mcp_server.tool()
async def get_options_chain(symbol: str, expiration: Optional[str] = None) -> str:
    """Get the options chain (calls + puts) for a ticker.

    Args:
        symbol: Underlying ticker (e.g. "AAPL").
        expiration: Optional ISO date "YYYY-MM-DD" to filter to one
            expiration. Omit to get the nearest expiration (default Yahoo
            behavior — typically the next weekly).

    Returns JSON: {
        symbol, underlyingPrice,
        expirationDates: [ISO dates],          # all available expirations
        strikes: [numbers],                    # all strikes across all exps
        expirations: [{ expirationDate, calls, puts }],
    }
    Each contract is trimmed to: contractSymbol, strike, lastPrice, bid, ask,
    volume, openInterest, impliedVolatility, inTheMoney. The full chain for
    one expiration is typically ~200 strikes; budget tokens accordingly.

    To explore the at-the-money region, first call without expiration to see
    the underlyingPrice and the list of expirationDates, then re-call with
    a specific expiration and filter strikes near the underlying in the
    response yourself.
    """
    params = {"symbol": symbol}
    if expiration:
        params["expiration"] = expiration
    data = await _get("/api/stocks/options", params=params)
    return json.dumps(data, indent=2)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp_server.run(transport="stdio")
