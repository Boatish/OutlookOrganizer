"""
Verwijder alle inbox-regels die naar 'inhoudelijke' mappen sturen
(school, facturen, beveiliging, werk, tickets, developer...).
Behoud alleen: nieuwsbrieven, aanbiedingen, archief.
De coach sorteert de rest voortaan vanuit inbox.
"""
import sys, msal, requests
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CLIENT_ID = "ea9e61fd-7233-4740-8073-a51868445bc1"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES    = ["Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite"]
GRAPH     = "https://graph.microsoft.com/v1.0"
CACHE     = r"C:\Users\Roldi\OutlookOrganizer\.token_cache"

def get_token():
    cache = msal.SerializableTokenCache()
    with open(CACHE) as f: cache.deserialize(f.read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    r = app.acquire_token_silent(SCOPES, account=app.get_accounts()[0])
    with open(CACHE, "w") as f: f.write(cache.serialize())
    return r["access_token"]

def hdrs(t): return {"Authorization": f"Bearer {t}", "Content-Type": "application/json"}

def get_folder_id_map(token):
    """Return {folder_id: folder_name}"""
    id_to_name = {}
    url = f"{GRAPH}/me/mailFolders?$top=100&$select=id,displayName,childFolderCount"
    while url:
        r = requests.get(url, headers=hdrs(token)); r.raise_for_status()
        data = r.json()
        for f in data.get("value", []):
            id_to_name[f["id"]] = f["displayName"]
            if f["childFolderCount"] > 0:
                cr = requests.get(f"{GRAPH}/me/mailFolders/{f['id']}/childFolders?$top=50&$select=id,displayName", headers=hdrs(token))
                for c in cr.json().get("value", []):
                    id_to_name[c["id"]] = c["displayName"]
        url = data.get("@odata.nextLink")
    return id_to_name

# Mappen die auto-regels MOGEN hebben (gebruiker hoeft ze niet te zien)
AUTO_OK = {
    "📰 Nieuwsbrieven",
    "📡 Tech & Security News",
    "🛒 Aanbiedingen & Shops",
    "📦 Archief",
    "Deleted Items",
}

def main():
    token      = get_token()
    id_to_name = get_folder_id_map(token)

    r = requests.get(f"{GRAPH}/me/mailFolders/inbox/messageRules", headers=hdrs(token))
    rules = r.json().get("value", [])
    print(f"{len(rules)} regels gevonden.\n")

    behouden = deleted = 0
    for rule in rules:
        rid      = rule["id"]
        naam     = rule["displayName"]
        folder_id = rule.get("actions", {}).get("moveToFolder", "")
        dest_name = id_to_name.get(folder_id, "")

        # Verwijder ook 'For all messages from...' regels zonder afzenders (Microsoft UI-regels)
        has_conditions = bool(
            rule.get("conditions", {}).get("senderContains") or
            rule.get("conditions", {}).get("subjectContains") or
            rule.get("conditions", {}).get("from")
        )
        no_dest = not folder_id

        if dest_name in AUTO_OK:
            print(f"  BEHOUDEN  {naam[:60]}")
            behouden += 1
        elif no_dest or not has_conditions:
            dr = requests.delete(f"{GRAPH}/me/mailFolders/inbox/messageRules/{rid}", headers=hdrs(token))
            status = "✓" if dr.status_code == 204 else f"✗({dr.status_code})"
            print(f"  {status} VERWIJDERD (geen filter)  {naam[:55]}")
            deleted += 1
        else:
            dr = requests.delete(f"{GRAPH}/me/mailFolders/inbox/messageRules/{rid}", headers=hdrs(token))
            status = "✓" if dr.status_code == 204 else f"✗({dr.status_code})"
            print(f"  {status} VERWIJDERD → {dest_name:<30}  {naam[:40]}")
            deleted += 1

    print(f"\n{behouden} regels behouden, {deleted} verwijderd.")
    print("Voortaan sorteert de coach inbox-mails zelf.")

if __name__ == "__main__":
    main()
