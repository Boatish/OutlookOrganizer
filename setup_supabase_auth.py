"""
Eenmalig uitvoeren op de PC om het Microsoft token in Supabase op te slaan.
Zet SUPABASE_URL en SUPABASE_KEY als env vars of vul ze hieronder in.

  python setup_supabase_auth.py
"""
import msal, os, sys

# ── Vul hier je Supabase gegevens in als je ze niet als env var hebt ──────────
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
# ─────────────────────────────────────────────────────────────────────────────

if not SUPABASE_URL or not SUPABASE_KEY:
    print("Stel SUPABASE_URL en SUPABASE_KEY in als env vars of vul ze in dit script in.")
    sys.exit(1)

# Zet env vars zodat storage.py ze kan gebruiken
os.environ["SUPABASE_URL"] = SUPABASE_URL
os.environ["SUPABASE_KEY"] = SUPABASE_KEY

import storage

CLIENT_ID = "ea9e61fd-7233-4740-8073-a51868445bc1"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES    = ["Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite"]
LOCAL_CACHE = r"C:\Users\Roldi\OutlookOrganizer\.token_cache"


def main():
    cache = msal.SerializableTokenCache()

    # Probeer bestaande lokale cache te laden
    if os.path.exists(LOCAL_CACHE):
        with open(LOCAL_CACHE) as f:
            cache.deserialize(f.read())
        print(f"Bestaande lokale token cache gevonden.")

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            storage.save_token_cache(cache.serialize())
            print(f"✅ Token opgeslagen in Supabase (account: {accounts[0]['username']})")
            return

    # Device flow als er geen cache is
    print("Geen bestaand token gevonden, nieuwe login nodig...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    print(f"\n{flow['message']}\n")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        storage.save_token_cache(cache.serialize())
        print("✅ Token opgeslagen in Supabase. De cloud bot kan nu inloggen.")
    else:
        print(f"❌ Auth mislukt: {result.get('error')}: {result.get('error_description')}")


if __name__ == "__main__":
    main()
