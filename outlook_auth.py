"""Microsoft Graph auth — slaat token op in Supabase i.p.v. een lokaal bestand."""
import msal, requests
import storage

CLIENT_ID = "ea9e61fd-7233-4740-8073-a51868445bc1"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES    = ["Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite"]
GRAPH     = "https://graph.microsoft.com/v1.0"


def get_token() -> str:
    cache = msal.SerializableTokenCache()
    cached = storage.get_token_cache()
    if cached:
        cache.deserialize(cached)

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result or "access_token" not in result:
        raise RuntimeError(
            "Microsoft token niet beschikbaar of verlopen. "
            "Run setup_supabase_auth.py op je PC."
        )

    if cache.has_state_changed:
        storage.save_token_cache(cache.serialize())

    return result["access_token"]


def hdrs(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_all_folders(token: str) -> dict:
    """Return {displayName: folder_object}"""
    folders = {}
    url = (f"{GRAPH}/me/mailFolders"
           f"?$top=100&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount")
    while url:
        r = requests.get(url, headers=hdrs(token)); r.raise_for_status()
        data = r.json()
        for f in data.get("value", []):
            folders[f["displayName"]] = f
            if f["childFolderCount"] > 0:
                cr = requests.get(
                    f"{GRAPH}/me/mailFolders/{f['id']}/childFolders"
                    f"?$top=100&$select=id,displayName,totalItemCount,unreadItemCount",
                    headers=hdrs(token)
                )
                for child in cr.json().get("value", []):
                    folders[child["displayName"]] = child
        url = data.get("@odata.nextLink")
    return folders
