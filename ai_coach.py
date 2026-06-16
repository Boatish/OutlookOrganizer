#!/usr/bin/env python3
"""
AI Mail Coach
Monitort mailbox elke 3 uur, beheert een persistente actielijst,
stuurt Telegram notificaties en verwijdert junk.

Gebruik:
  python ai_coach.py                  -- normale run
  python ai_coach.py --get-chat-id    -- Telegram chat-ID ophalen
  python ai_coach.py --briefing       -- stuur dagelijkse briefing nu
"""

import json
import os
import re
import sys
import uuid
import msal
import requests
import time
from datetime import datetime, timezone, timedelta
from anthropic import Anthropic

# ── Outlook / Graph ──────────────────────────────────────────────────────────
CLIENT_ID   = "ea9e61fd-7233-4740-8073-a51868445bc1"
AUTHORITY   = "https://login.microsoftonline.com/consumers"
SCOPES      = ["Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite"]
GRAPH       = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE = r"C:\Users\Roldi\OutlookOrganizer\.token_cache"
LAST_CHECK  = r"C:\Users\Roldi\OutlookOrganizer\.last_check"
ACTIE_FILE  = r"C:\Users\Roldi\OutlookOrganizer\actie_lijst.json"

# Actie-mappen: ALLE mails (gelezen + ongelezen), want de map zelf = de actielijst
ACTION_FOLDERS = [
    "📌 Actie Vereist",
    "⚠️ Openstaand",
]
# Inbox: alle ongelezen (ook oude)
PRIORITY_FOLDERS = [
    "Inbox",
]
# Mappen die alleen op nieuwe mails gescand worden
NEW_ONLY_FOLDERS = [
    "💸 Financieel",
    "🎒 School — HL Leiden",
    "💻 Developer & Tech",
    "🎫 Tickets & Events",
    "🔐 Beveiliging",
]

JUNK_FOLDER     = "Junk Email"
SORT_LOG        = r"C:\Users\Roldi\OutlookOrganizer\sort_log.json"
MAX_PRIORITY    = 30   # max ongelezen per prioriteitsmap (alle ongelezen)
MAX_NEW         = 10   # max nieuwe mails per overige map
MAX_JUNK        = 40
MAX_BODY        = 600

# Mappen waarnaar inbox mails automatisch gesorteerd mogen worden
SORTEER_MAPPEN = [
    "Inbox",                    # persoonlijke mails blijven hier
    "🧾 Facturen & Bonnetjes",
    "📱 Abonnementen",
    "🎒 School — HL Leiden",
    "🔐 Beveiliging",
    "🎫 Tickets & Events",
    "💻 Developer & Tech",
    "☁️ Cloud & Hosting",
    "🤖 AI & Tools",
    "💼 Werk — Linden-IT",
    "⚠️ Openstaand",
    "📌 Actie Vereist",
    "📦 Archief",
]

# ── Credentials ──────────────────────────────────────────────────────────────
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT  = os.environ.get("TELEGRAM_CHAT_ID", "")

# ── Google Calendar ───────────────────────────────────────────────────────────
GCAL_CREDENTIALS = r"C:\Users\Roldi\OutlookOrganizer\google_credentials.json"
GCAL_TOKEN       = r"C:\Users\Roldi\OutlookOrganizer\.google_token.json"
GCAL_SCOPES      = ["https://www.googleapis.com/auth/calendar.events"]


# ════════════════════════════════════════════════════════════════════════════
# Actielijst — persistent JSON bestand
# ════════════════════════════════════════════════════════════════════════════

def load_actie_lijst():
    try:
        with open(ACTIE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def save_actie_lijst(lijst):
    with open(ACTIE_FILE, "w", encoding="utf-8") as f:
        json.dump(lijst, f, ensure_ascii=False, indent=2)


def format_actie_lijst_telegram(lijst):
    open_items = [a for a in lijst if a.get("status") != "gedaan"]
    if not open_items:
        return "✅ Geen open actiepunten."

    today = datetime.now().date()
    urgent, middel, laag = [], [], []

    for a in open_items:
        dl = a.get("deadline")
        days_left = None
        if dl:
            try:
                days_left = (datetime.strptime(dl, "%Y-%m-%d").date() - today).days
            except ValueError:
                pass

        if a.get("urgentie") == "hoog" or (days_left is not None and days_left <= 7):
            urgent.append((a, days_left))
        elif a.get("urgentie") == "middel":
            middel.append((a, days_left))
        else:
            laag.append((a, days_left))

    lines = []
    for label, emoji, items in [("URGENT", "🔴", urgent), ("AANDACHT", "🟡", middel), ("LATER", "⚪", laag)]:
        if not items:
            continue
        lines.append(f"*{emoji} {label}:*")
        for a, days_left in items[:5]:
            dl_str = ""
            if days_left is not None:
                if days_left < 0:
                    dl_str = f" _(VERLOPEN!_ {abs(days_left)}d geleden)"
                elif days_left == 0:
                    dl_str = " _(vandaag!)_"
                else:
                    dl_str = f" _(nog {days_left}d)_"
            lines.append(f"• {a['beschrijving']}{dl_str}")

    if len(open_items) > 10:
        lines.append(f"\n_...en {len(open_items) - 10} andere open punten_")

    return "\n".join(lines)


# ════════════════════════════════════════════════════════════════════════════
# Outlook authenticatie
# ════════════════════════════════════════════════════════════════════════════

def get_outlook_token():
    cache = msal.SerializableTokenCache()
    try:
        with open(TOKEN_CACHE) as f:
            cache.deserialize(f.read())
    except FileNotFoundError:
        pass

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "error" in flow:
            raise Exception(f"Device flow fout: {flow['error']}")
        print(f"\n{flow['message']}\n")
        result = app.acquire_token_by_device_flow(flow)

    with open(TOKEN_CACHE, "w") as f:
        f.write(cache.serialize())

    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Inloggen mislukt: {result.get('error')}")


def hdrs(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


# ════════════════════════════════════════════════════════════════════════════
# Outlook helpers
# ════════════════════════════════════════════════════════════════════════════

def get_all_folders(token):
    folders = {}
    url = f"{GRAPH}/me/mailFolders?$top=100&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount"
    while url:
        r = requests.get(url, headers=hdrs(token))
        r.raise_for_status()
        data = r.json()
        for f in data.get("value", []):
            folders[f["displayName"]] = f
            if f["childFolderCount"] > 0:
                cu = (f"{GRAPH}/me/mailFolders/{f['id']}/childFolders"
                      f"?$top=100&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount")
                cr = requests.get(cu, headers=hdrs(token))
                cr.raise_for_status()
                for child in cr.json().get("value", []):
                    folders[child["displayName"]] = child
        url = data.get("@odata.nextLink")
    return folders


def strip_html(html):
    text = re.sub(r'<style[^>]*>.*?</style>', ' ', html, flags=re.DOTALL)
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', text, flags=re.DOTALL)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;|&#160;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    return re.sub(r'\s+', ' ', text).strip()


def fetch_emails(token, folder_id, since_dt=None, max_count=20, all_mails=False):
    select = "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,body"
    if all_mails:
        flt = ""
    elif since_dt:
        iso = since_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        flt = f"&$filter=isRead eq false and receivedDateTime ge {iso}"
    else:
        flt = "&$filter=isRead eq false"

    url = (
        f"{GRAPH}/me/mailFolders/{folder_id}/messages"
        f"?$top=50&$select={select}{flt}"
        f"&$orderby=receivedDateTime asc"
    )
    emails = []
    while url and len(emails) < max_count:
        r = requests.get(url, headers=hdrs(token))
        if r.status_code != 200:
            break
        data = r.json()
        for msg in data.get("value", []):
            raw   = msg.get("body", {}).get("content", "") or ""
            btype = msg.get("body", {}).get("contentType", "text")
            body  = strip_html(raw) if btype == "html" else re.sub(r'\s+', ' ', raw).strip()
            emails.append({
                "id":         msg["id"],
                "subject":    msg.get("subject", "") or "",
                "from_name":  msg.get("from", {}).get("emailAddress", {}).get("name", ""),
                "from_addr":  msg.get("from", {}).get("emailAddress", {}).get("address", ""),
                "date":       msg.get("receivedDateTime", "")[:10],
                "important":  msg.get("importance", "") == "high",
                "attachment": msg.get("hasAttachments", False),
                "body":       body[:MAX_BODY],
            })
            if len(emails) >= max_count:
                break
        url = data.get("@odata.nextLink")
        time.sleep(0.05)
    return emails


def fetch_junk(token, folder_id, max_count=MAX_JUNK):
    select = "id,subject,from,receivedDateTime"
    url = (
        f"{GRAPH}/me/mailFolders/{folder_id}/messages"
        f"?$top=50&$select={select}&$orderby=receivedDateTime desc"
    )
    emails = []
    while url and len(emails) < max_count:
        r = requests.get(url, headers=hdrs(token))
        if r.status_code != 200:
            break
        data = r.json()
        for msg in data.get("value", []):
            emails.append({
                "id":        msg["id"],
                "subject":   msg.get("subject", "") or "",
                "from_name": msg.get("from", {}).get("emailAddress", {}).get("name", ""),
                "from_addr": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
                "date":      msg.get("receivedDateTime", "")[:10],
            })
            if len(emails) >= max_count:
                break
        url = data.get("@odata.nextLink")
    return emails


def move_msg(token, msg_id, dest_id):
    r = requests.post(
        f"{GRAPH}/me/messages/{msg_id}/move",
        headers=hdrs(token),
        json={"destinationId": dest_id},
    )
    return r.status_code in (200, 201)


def delete_msg(token, msg_id):
    r = requests.delete(f"{GRAPH}/me/messages/{msg_id}", headers=hdrs(token))
    return r.status_code == 204


# ════════════════════════════════════════════════════════════════════════════
# Google Calendar
# ════════════════════════════════════════════════════════════════════════════

def get_calendar_service():
    if not os.path.exists(GCAL_CREDENTIALS):
        return None
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request as GRequest
        from googleapiclient.discovery import build
    except ImportError:
        return None

    creds = None
    if os.path.exists(GCAL_TOKEN):
        creds = Credentials.from_authorized_user_file(GCAL_TOKEN, GCAL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(GRequest())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GCAL_CREDENTIALS, GCAL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(GCAL_TOKEN, "w") as f:
            f.write(creds.to_json())
    from googleapiclient.discovery import build
    return build("calendar", "v3", credentials=creds)


def get_upcoming_events(service, days=7):
    if not service:
        return []
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days)
    result = service.events().list(
        calendarId="primary",
        timeMin=now.isoformat(),
        timeMax=end.isoformat(),
        maxResults=10,
        singleEvents=True,
        orderBy="startTime",
    ).execute()
    events = []
    for e in result.get("items", []):
        start = e["start"].get("dateTime", e["start"].get("date", ""))
        events.append({"title": e.get("summary", ""), "start": start[:10]})
    return events


def create_calendar_event(service, ev):
    if not service or not ev.get("datum"):
        return False
    try:
        datum = ev["datum"]
        tijd  = ev.get("tijd")
        if tijd:
            tz     = "Europe/Amsterdam"
            start  = {"dateTime": f"{datum}T{tijd}:00", "timeZone": tz}
            dt     = datetime.fromisoformat(f"{datum}T{tijd}:00")
            end_dt = dt + timedelta(minutes=int(ev.get("duur_minuten", 60)))
            end    = {"dateTime": end_dt.isoformat(), "timeZone": tz}
        else:
            start = {"date": datum}
            end   = {"date": datum}
        service.events().insert(calendarId="primary", body={
            "summary":     ev.get("titel", ""),
            "description": ev.get("beschrijving", ""),
            "start": start,
            "end":   end,
        }).execute()
        return True
    except Exception:
        return False


# ════════════════════════════════════════════════════════════════════════════
# Claude prompt
# ════════════════════════════════════════════════════════════════════════════

def build_prompt(priority_emails, new_emails, junk_emails, actie_lijst, calendar_events):
    today = datetime.now().strftime("%A %d %B %Y")
    open_acties = [a for a in actie_lijst if a.get("status") != "gedaan"]

    lines = [f"Vandaag is het {today}.\n"]

    # Huidige open actielijst meegeven
    if open_acties:
        lines.append("## Huidige open actiepunten (door jou bijgehouden):")
        for a in open_acties:
            dl = f", deadline: {a['deadline']}" if a.get("deadline") else ""
            lines.append(f"  [ID:{a['id']}] [{a['urgentie'].upper()}] {a['beschrijving']} (van: {a.get('afzender','?')}){dl}")
        lines.append("")

    # Prioriteitsmappen — alle ongelezen (ook oude!)
    lines.append("## Prioriteitsmappen (alle ongelezen, ook oude):")
    heeft_prioriteit = False
    for folder_name, emails in priority_emails.items():
        if not emails:
            continue
        heeft_prioriteit = True
        lines.append(f"### {folder_name} ({len(emails)} ongelezen):")
        for e in emails:
            flags = (["BELANGRIJK"] if e["important"] else []) + (["BIJLAGE"] if e["attachment"] else [])
            flag_str = f" [{', '.join(flags)}]" if flags else ""
            lines.append(
                f"  [{e['date']}]{flag_str} Van: {e['from_name']} <{e['from_addr']}>\n"
                f"  Onderwerp: {e['subject']}\n"
                f"  Inhoud: {e['body']}\n"
            )
    if not heeft_prioriteit:
        lines.append("  (geen ongelezen in prioriteitsmappen)\n")

    # Nieuwe mails in overige mappen
    if any(new_emails.values()):
        lines.append("## Nieuwe mails overige mappen:")
        for folder_name, emails in new_emails.items():
            if not emails:
                continue
            lines.append(f"### {folder_name} ({len(emails)} nieuw):")
            for e in emails:
                lines.append(f"  [{e['date']}] {e['from_name']} — {e['subject']}")
        lines.append("")

    if junk_emails:
        lines.append(f"## Junk ({len(junk_emails)} stuks — beoordeel valse positieven):")
        for i, e in enumerate(junk_emails):
            lines.append(f"  #{i} [{e['date']}] {e['from_name']} <{e['from_addr']}> — {e['subject']}")
        lines.append("")

    if calendar_events:
        lines.append("## Agenda komende 7 dagen:")
        for ev in calendar_events:
            lines.append(f"  • {ev['start']} — {ev['title']}")
        lines.append("")

    prompt = f"""Je bent de persoonlijke AI-assistent van Roldi (Nederlandse student/developer, 2026).
Jouw taak: beheer zijn actielijst op basis van zijn mailbox. Wees direct en concreet.

{chr(10).join(lines)}

Geef je antwoord ALLEEN als geldig JSON:

```json
{{
  "actielijst_updates": [
    {{
      "actie": "toevoegen|afsluiten|updaten",
      "id": "bestaand-id-of-null-bij-toevoegen",
      "beschrijving": "Wat Roldi concreet moet doen (niet vaag)",
      "urgentie": "hoog|middel|laag",
      "afzender": "naam/organisatie",
      "deadline": "YYYY-MM-DD of null",
      "status": "open|gedaan"
    }}
  ],
  "kalender_events": [
    {{
      "titel": "Korte titel",
      "datum": "YYYY-MM-DD",
      "tijd": "HH:MM of null",
      "duur_minuten": 60,
      "beschrijving": ""
    }}
  ],
  "junk_valse_positieven": [],
  "heeft_nieuws": true,
  "telegram_bericht": "Max 15 regels. Gebruik *vet* voor headers. Wees concreet over wat er moet gebeuren."
}}
```

Regels voor actielijst_updates:
- "toevoegen": nieuw item (id = null)
- "afsluiten": bestaand item markeren als gedaan (geef het ID mee, status = "gedaan")
- "updaten": urgentie of deadline aanpassen (geef het ID mee)
- Voeg ALLEEN toe als je een echte actie ziet: betaling, aanmaning, schooltaak, technische actie, deadline
- Nieuwsbrieven, promoties, beveiliging-notificaties zonder actie = NOOIT toevoegen
- Als een mail duidelijk maakt dat iets al geregeld is → sluit het bijbehorende item af

Telegram bericht formaat:
- Begin met een korte statusregel: "X open taken, Y urgent"
- Lijst de urgente items met deadline
- Sluit af met wat er nieuw is of als alles rustig is"""

    return prompt


def call_claude(prompt):
    client = Anthropic(api_key=ANTHROPIC_KEY)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


# ════════════════════════════════════════════════════════════════════════════
# Inbox sorteren
# ════════════════════════════════════════════════════════════════════════════

def _append_sort_log(entries):
    try:
        try:
            with open(SORT_LOG, encoding="utf-8") as f:
                log = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            log = []
        log = (entries + log)[:200]
        with open(SORT_LOG, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def sort_inbox(token, folders):
    """
    Haal alle ongelezen inbox mails op, laat Claude ze indelen,
    verplaats ze naar de juiste mappen.
    Geeft (verplaatst_dict, sort_entries) terug.
    """
    inbox_info = folders.get("Inbox")
    if not inbox_info:
        return {}, []

    emails = fetch_emails(token, inbox_info["id"], since_dt=None, max_count=50)
    if not emails:
        print("  Inbox: geen ongelezen mails om te sorteren.")
        return {}, []

    print(f"  Inbox sorteren: {len(emails)} ongelezen mail(s)...")

    beschikbaar = [m for m in SORTEER_MAPPEN if m == "Inbox" or m in folders]
    beschikbaar_str = "\n".join(f"  - {m}" for m in beschikbaar)

    mail_lines = []
    for i, m in enumerate(emails):
        mail_lines.append(
            f"#{i} [{m['date']}] Van: {m['from_name']} <{m['from_addr']}>\n"
            f"   Onderwerp: {m['subject']}\n"
            f"   Inhoud: {m['body'][:300]}"
        )

    prompt = f"""Je bent de e-mail sorterassistent van Roldi (Nederlandse student/developer, Den Haag, 2026).
Deel de volgende inbox mails in naar de juiste mappen.

BESCHIKBARE MAPPEN:
{beschikbaar_str}

INBOX MAILS:
{chr(10).join(mail_lines)}

SORTEERREGELS:
- Persoonlijke mails van echte mensen (vrienden, familie) → "Inbox" (niet verplaatsen)
- Facturen, rekeningen, betalingen → "🧾 Facturen & Bonnetjes"
- Abonnements-mails (Youfone, Spotify, Nintendo, etc.) → "📱 Abonnementen"
- School (HL Leiden, DUO, Studielink, Brightspace) → "🎒 School — HL Leiden"
- Login-alerts, 2FA, security-mails → "🔐 Beveiliging"
- Tickets, reserveringen, evenementen → "🎫 Tickets & Events"
- GitHub, Vercel, Supabase, developer tools → "💻 Developer & Tech"
- Azure, Google Cloud, hosting → "☁️ Cloud & Hosting"
- AI tools (OpenAI, Anthropic, Mistral) → "🤖 AI & Tools"
- Werk (Linden-IT) → "💼 Werk — Linden-IT"
- Aanmaningen, dringende openstaande rekeningen → "⚠️ Openstaand"
- Onbekend maar lijkt dringend → "📌 Actie Vereist"
- Alles zonder duidelijke categorie → "📦 Archief"

Geef antwoord ALLEEN als geldige JSON array:
```json
[
  {{"index": 0, "naar_map": "exacte mapnaam uit de lijst", "reden": "één zin"}},
  {{"index": 1, "naar_map": "Inbox", "reden": "persoonlijke mail"}}
]
```
Alle indices 0 t/m {len(emails)-1} moeten aanwezig zijn."""

    client = Anthropic(api_key=ANTHROPIC_KEY)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text

    m = re.search(r'```json\s*(.*?)\s*```', raw, re.DOTALL)
    try:
        beslissingen = json.loads(m.group(1) if m else raw.strip())
    except (json.JSONDecodeError, AttributeError):
        print("  Inbox sort: parse mislukt.")
        return {}, []

    verplaatst = {}
    sort_entries = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    for item in beslissingen:
        idx  = item.get("index", -1)
        naar = item.get("naar_map", "Inbox")
        reden = item.get("reden", "")

        if idx < 0 or idx >= len(emails):
            continue

        mail = emails[idx]
        sort_entries.append({
            "tijdstip":        now_str,
            "mail_id":         mail["id"],
            "afzender":        mail["from_addr"],
            "onderwerp":       mail["subject"],
            "datum_ontvangen": mail["date"],
            "naar_map":        naar,
            "reden":           reden,
        })

        if naar == "Inbox":
            continue  # persoonlijke mail, niet verplaatsen

        dst_info = folders.get(naar, {})
        dst_id = dst_info.get("id") if isinstance(dst_info, dict) else None
        if not dst_id:
            continue

        ok = move_msg(token, mail["id"], dst_id)
        if ok:
            verplaatst[naar] = verplaatst.get(naar, 0) + 1
            print(f"  ✓ #{idx:02d} → {naar[:40]}  ({mail['from_addr'][:30]})")
        else:
            print(f"  ✗ #{idx:02d} move mislukt → {naar}")
        time.sleep(0.05)

    _append_sort_log(sort_entries)
    return verplaatst, sort_entries


def parse_response(text):
    m = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        return None


# ════════════════════════════════════════════════════════════════════════════
# Telegram
# ════════════════════════════════════════════════════════════════════════════

def send_telegram(text):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT:
        print(f"  [Telegram niet geconfigureerd]\n{text}")
        return False
    r = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
        json={"chat_id": TELEGRAM_CHAT, "text": text, "parse_mode": "Markdown"},
        timeout=10,
    )
    return r.status_code == 200


def get_telegram_chat_id():
    if not TELEGRAM_TOKEN:
        print("Stel TELEGRAM_BOT_TOKEN in als omgevingsvariabele.")
        return
    r = requests.get(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates", timeout=10)
    updates = r.json().get("result", [])
    if updates:
        print(f"\nJouw TELEGRAM_CHAT_ID: {updates[-1]['message']['chat']['id']}")
    else:
        print("Geen berichten gevonden. Stuur eerst een bericht naar je bot.")


# ════════════════════════════════════════════════════════════════════════════
# Tijdstempel
# ════════════════════════════════════════════════════════════════════════════

def load_last_check():
    try:
        with open(LAST_CHECK) as f:
            return datetime.fromisoformat(f.read().strip()).replace(tzinfo=timezone.utc)
    except (FileNotFoundError, ValueError):
        return datetime.now(timezone.utc) - timedelta(hours=24)


def save_last_check():
    with open(LAST_CHECK, "w") as f:
        f.write(datetime.now(timezone.utc).isoformat())


# ════════════════════════════════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════════════════════════════════

def main():
    if "--get-chat-id" in sys.argv:
        get_telegram_chat_id()
        return

    briefing_only = "--briefing" in sys.argv

    now_str = datetime.now().strftime("%d %b %Y %H:%M")
    print(f"[{now_str}] AI Coach gestart")

    if not ANTHROPIC_KEY:
        print("Fout: ANTHROPIC_API_KEY niet ingesteld.")
        return

    if briefing_only:
        # Stuur de huidige actielijst als dagelijkse briefing
        actie_lijst = load_actie_lijst()
        msg = f"*📋 Dagelijkse briefing — {datetime.now().strftime('%d %b')}*\n\n"
        msg += format_actie_lijst_telegram(actie_lijst)
        send_telegram(msg)
        print("  Briefing verstuurd.")
        return

    token   = get_outlook_token()
    folders = get_all_folders(token)
    since   = load_last_check()
    print(f"  Mappen geladen. Checkt nieuw sinds {since.strftime('%d %b %H:%M')} UTC")

    # ── Stap 1: Inbox sorteren ────────────────────────────────────────────────
    sort_result, _sort_log = sort_inbox(token, folders)

    # Herlaad mappen zodat tellingen kloppen na verplaatsen
    if sort_result:
        folders = get_all_folders(token)

    # ── Stap 2: Actie-mappen scannen ─────────────────────────────────────────
    priority_emails = {}
    for name in ACTION_FOLDERS:
        f = folders.get(name)
        if not f:
            continue
        emails = fetch_emails(token, f["id"], all_mails=True, max_count=MAX_PRIORITY)
        if emails:
            priority_emails[name] = emails
            print(f"  {name}: {len(emails)} mails")

    # Inbox: alle ongelezen (ook oude)
    for name in PRIORITY_FOLDERS:
        f = folders.get(name)
        if not f:
            continue
        emails = fetch_emails(token, f["id"], since_dt=None, max_count=MAX_PRIORITY)
        if emails:
            priority_emails[name] = emails
            print(f"  {name}: {len(emails)} ongelezen")

    # Overige mappen: alleen nieuw
    new_emails = {}
    for name in NEW_ONLY_FOLDERS:
        f = folders.get(name)
        if not f:
            continue
        emails = fetch_emails(token, f["id"], since_dt=since, max_count=MAX_NEW)
        if emails:
            new_emails[name] = emails
            print(f"  {name}: {len(emails)} nieuw")

    # Junk
    junk_folder = folders.get(JUNK_FOLDER)
    junk_emails = fetch_junk(token, junk_folder["id"]) if junk_folder else []
    if junk_emails:
        print(f"  Junk: {len(junk_emails)} te beoordelen")

    # Google Calendar
    cal      = get_calendar_service()
    upcoming = get_upcoming_events(cal) if cal else []
    print(f"  Google Calendar: {'niet gekoppeld' if not cal else f'{len(upcoming)} events'}")

    # Huidige actielijst laden
    actie_lijst = load_actie_lijst()

    # Niets te analyseren?
    has_priority = any(priority_emails.values())
    has_new      = any(new_emails.values())
    result       = None

    if has_priority or has_new or junk_emails:
        # Claude
        print("  Claude analyseert...")
        prompt = build_prompt(priority_emails, new_emails, junk_emails, actie_lijst, upcoming)
        raw    = call_claude(prompt)
        result = parse_response(raw)

        if not result:
            print("  Parse mislukt — opgeslagen als coach_debug.txt")
            with open(r"C:\Users\Roldi\OutlookOrganizer\coach_debug.txt", "w", encoding="utf-8") as f:
                f.write(raw)
    elif not sort_result:
        print("  Geen ongelezen of nieuwe mails. Coach slaat over.")
        save_last_check()
        return

    # Actielijst bijwerken
    updates = result.get("actielijst_updates", []) if result else []
    id_map  = {a["id"]: a for a in actie_lijst}
    changed = 0

    for upd in updates:
        actie = upd.get("actie", "toevoegen")
        uid   = upd.get("id")

        if actie == "toevoegen":
            new_id = str(uuid.uuid4())[:8]
            actie_lijst.append({
                "id":           new_id,
                "beschrijving": upd.get("beschrijving", ""),
                "urgentie":     upd.get("urgentie", "middel"),
                "afzender":     upd.get("afzender", ""),
                "deadline":     upd.get("deadline"),
                "status":       "open",
                "aangemaakt":   datetime.now().strftime("%Y-%m-%d"),
            })
            changed += 1

        elif actie in ("afsluiten", "updaten") and uid and uid in id_map:
            item = id_map[uid]
            if actie == "afsluiten":
                item["status"] = "gedaan"
            else:
                if upd.get("urgentie"):
                    item["urgentie"] = upd["urgentie"]
                if upd.get("deadline") is not None:
                    item["deadline"] = upd["deadline"]
            changed += 1

    if changed:
        save_actie_lijst(actie_lijst)
        print(f"  Actielijst: {changed} wijziging(en) opgeslagen")

    # Kalender
    cal_created = 0
    for ev in (result.get("kalender_events", []) if result else []):
        if create_calendar_event(cal, ev):
            cal_created += 1
            print(f"  Kalender: '{ev.get('titel')}' ({ev.get('datum')})")

    # Junk
    inbox_id      = folders.get("Inbox", {}).get("id")
    false_pos_idx = set(int(i) for i in (result.get("junk_valse_positieven", []) if result else []) if isinstance(i, (int, str)) and str(i).isdigit())
    rescued = 0
    deleted = 0

    for idx, email in enumerate(junk_emails):
        if idx in false_pos_idx and inbox_id:
            if move_msg(token, email["id"], inbox_id):
                rescued += 1
        else:
            if delete_msg(token, email["id"]):
                deleted += 1
        time.sleep(0.05)

    if junk_emails:
        print(f"  Junk: {rescued} gered, {deleted} verwijderd")

    # Bouw sort-samenvatting voor Telegram
    if sort_result:
        sort_lines = [f"📬 *{sum(sort_result.values())} mail(s) gesorteerd:*"]
        for map_naam, n in sorted(sort_result.items(), key=lambda x: -x[1]):
            sort_lines.append(f"  • {n} → {map_naam}")
        sort_summary = "\n".join(sort_lines)
    else:
        sort_summary = ""

    # Telegram
    heeft_nieuws = result.get("heeft_nieuws", True) if result else False
    if sort_result or heeft_nieuws:
        tg_msg = ""
        if sort_summary:
            tg_msg += sort_summary + "\n\n"
        if result and heeft_nieuws:
            tg_msg += result.get("telegram_bericht", "")
        footer = []
        if cal_created:
            footer.append(f"📅 {cal_created} agenda-item(s) aangemaakt")
        if rescued:
            footer.append(f"♻️ {rescued} mail uit junk gered")
        if deleted:
            footer.append(f"🗑️ {deleted} junk verwijderd")
        if footer:
            tg_msg += "\n\n" + " · ".join(footer)
        tg_msg = tg_msg.strip()
        if tg_msg:
            ok = send_telegram(tg_msg)
            print(f"  Telegram: {'✓' if ok else '✗'}")
    else:
        print("  Alles rustig — geen Telegram.")

    save_last_check()
    print("  Klaar.")


if __name__ == "__main__":
    main()
