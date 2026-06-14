"""One-time auth script for SuNaPaNa_ account.
Run this interactively to get a token for SuNaPaNa_@outlook.com.
"""
import json
from pathlib import Path
import msal

CONFIG_PATH = Path(r"C:\mcp-servers\outlook-email\config.json")
CACHE_PATH = Path(r"C:\mcp-servers\outlook-email\.token_cache\sunnapana_token_cache.json")
SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"]

config = json.loads(CONFIG_PATH.read_text())
cache = msal.SerializableTokenCache()

app = msal.PublicClientApplication(
    config["client_id"],
    authority="https://login.microsoftonline.com/consumers",
    token_cache=cache,
)

import time
# Device code flow — will print a URL and code to enter
flow = app.initiate_device_flow(scopes=SCOPES)
# Force a generous polling window (Microsoft device codes are valid ~15 min)
flow["expires_at"] = int(time.time()) + 900
print(flow["message"])

# Also write code to a file so it can be read externally
Path(r"C:\mcp-servers\sunnapana\device_code.txt").write_text(flow["message"])

print("\n>>> Sign in with SuNaPaNa_@outlook.com <<<\n")

result = app.acquire_token_by_device_flow(flow)
if "access_token" in result:
    CACHE_PATH.write_text(cache.serialize())
    print(f"\nSuccess! Token cached at {CACHE_PATH}")
    print(f"Account: {result.get('id_token_claims', {}).get('preferred_username', 'unknown')}")
else:
    print(f"\nFailed: {result.get('error_description', result)}")
