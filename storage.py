"""
Supabase persistence layer — vervangt lokale JSON bestanden.

Aanmaken in Supabase SQL editor (eenmalig):
  CREATE TABLE config (key TEXT PRIMARY KEY, value TEXT NOT NULL);
"""
import json, os
from supabase import create_client

_client = None

def _db():
    global _client
    if _client is None:
        _client = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])
    return _client

def _get(key, default=None):
    r = _db().table("config").select("value").eq("key", key).execute()
    return json.loads(r.data[0]["value"]) if r.data else default

def _set(key, value):
    _db().table("config").upsert({"key": key, "value": json.dumps(value, ensure_ascii=False)}).execute()

def get_actie_lijst():            return _get("actie_lijst", [])
def save_actie_lijst(lst):        _set("actie_lijst", lst)

def get_token_cache() -> str:     return _get("msal_token_cache", "")
def save_token_cache(s: str):     _set("msal_token_cache", s)

def get_last_check() -> str:      return _get("last_check", "")
def save_last_check(iso: str):    _set("last_check", iso)

def get_sort_log() -> list:       return _get("sort_log", [])
def append_sort_log(entries: list):
    _set("sort_log", (entries + get_sort_log())[:200])

def get_conversation(chat_id) -> list:
    return _get(f"conv_{chat_id}", [])

def save_conversation(chat_id, messages: list):
    _set(f"conv_{chat_id}", messages[-24:])  # laatste 24 berichten bewaren
