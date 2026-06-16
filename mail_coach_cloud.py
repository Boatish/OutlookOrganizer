"""
Mail Coach — cloud versie.
Gebruikt storage.py (Supabase) i.p.v. lokale bestanden.
Roep run_check() of run_briefing() aan vanuit server.py.
"""
import json, os, re, uuid, requests, time
from datetime import datetime, timezone, timedelta
from anthropic import Anthropic
import storage, outlook_auth

GRAPH       = "https://graph.microsoft.com/v1.0"
ANTHR_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
TG_TOKEN    = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TG_CHAT     = os.environ.get("TELEGRAM_CHAT_ID", "")
MAX_BODY    = 600

ACTION_FOLDERS   = ["📌 Actie Vereist", "⚠️ Openstaand"]
NEW_ONLY_FOLDERS = ["🧾 Facturen & Bonnetjes", "🎒 School — HL Leiden",
                    "💻 Developer & Tech", "🎫 Tickets & Events", "🔐 Beveiliging"]
SORTEER_MAPPEN   = [
    "Inbox", "🧾 Facturen & Bonnetjes", "📱 Abonnementen",
    "🎒 School — HL Leiden", "🔐 Beveiliging", "🎫 Tickets & Events",
    "💻 Developer & Tech", "☁️ Cloud & Hosting", "🤖 AI & Tools",
    "💼 Werk — Linden-IT", "⚠️ Openstaand", "📌 Actie Vereist", "📦 Archief",
]

# ── Telegram ──────────────────────────────────────────────────────────────────

def _tg(text: str):
    if not TG_TOKEN or not TG_CHAT:
        print(f"[TG] {text}")
        return
    requests.post(
        f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage",
        json={"chat_id": TG_CHAT, "text": text, "parse_mode": "Markdown"},
        timeout=10,
    )

# ── Mail ophalen ──────────────────────────────────────────────────────────────

def _strip_html(html: str) -> str:
    text = re.sub(r'<style[^>]*>.*?</style>', ' ', html, flags=re.DOTALL)
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', text, flags=re.DOTALL)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;|&#160;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    return re.sub(r'\s+', ' ', text).strip()


def _fetch(token, folder_id, since_dt=None, max_n=20, all_mails=False):
    select = "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,body"
    if all_mails:
        flt = ""
    elif since_dt:
        flt = f"&$filter=isRead eq false and receivedDateTime ge {since_dt.strftime('%Y-%m-%dT%H:%M:%SZ')}"
    else:
        flt = "&$filter=isRead eq false"

    url = (f"{GRAPH}/me/mailFolders/{folder_id}/messages"
           f"?$top=50&$select={select}{flt}&$orderby=receivedDateTime asc")
    emails = []
    while url and len(emails) < max_n:
        r = requests.get(url, headers=outlook_auth.hdrs(token))
        if r.status_code != 200:
            break
        data = r.json()
        for m in data.get("value", []):
            raw   = m.get("body", {}).get("content", "") or ""
            btype = m.get("body", {}).get("contentType", "text")
            body  = _strip_html(raw) if btype == "html" else re.sub(r'\s+', ' ', raw).strip()
            emails.append({
                "id":         m["id"],
                "subject":    m.get("subject", "") or "",
                "from_name":  m.get("from", {}).get("emailAddress", {}).get("name", ""),
                "from_addr":  m.get("from", {}).get("emailAddress", {}).get("address", ""),
                "date":       m.get("receivedDateTime", "")[:10],
                "important":  m.get("importance", "") == "high",
                "attachment": m.get("hasAttachments", False),
                "body":       body[:MAX_BODY],
            })
            if len(emails) >= max_n:
                break
        url = data.get("@odata.nextLink")
        time.sleep(0.05)
    return emails


def _move(token, msg_id, dst_id) -> bool:
    r = requests.post(
        f"{GRAPH}/me/messages/{msg_id}/move",
        headers=outlook_auth.hdrs(token),
        json={"destinationId": dst_id},
    )
    return r.status_code in (200, 201)

# ── Inbox sorteren ────────────────────────────────────────────────────────────

def _sort_inbox(token, folders: dict):
    inbox = folders.get("Inbox")
    if not inbox:
        return {}

    emails = _fetch(token, inbox["id"], max_n=50)
    if not emails:
        return {}

    print(f"  Inbox sorteren: {len(emails)} ongelezen mail(s)...")
    beschikbaar = [m for m in SORTEER_MAPPEN if m == "Inbox" or m in folders]

    mail_lines = []
    for i, m in enumerate(emails):
        mail_lines.append(
            f"#{i} [{m['date']}] Van: {m['from_name']} <{m['from_addr']}>\n"
            f"   Onderwerp: {m['subject']}\n"
            f"   Inhoud: {m['body'][:250]}"
        )

    prompt = f"""Je bent de e-mail sorterassistent van Roldi (Nederlandse student/developer, Den Haag, 2026).
Deel de volgende inbox mails in naar de juiste mappen.

BESCHIKBARE MAPPEN:
{chr(10).join('  - ' + m for m in beschikbaar)}

MAILS:
{chr(10).join(mail_lines)}

REGELS:
- Persoonlijke mails van echte mensen → "Inbox"
- Facturen/rekeningen/betalingen → "🧾 Facturen & Bonnetjes"
- Abonnementen (Youfone, Spotify, etc.) → "📱 Abonnementen"
- School (HL Leiden, DUO, Studielink) → "🎒 School — HL Leiden"
- Login-alerts, 2FA → "🔐 Beveiliging"
- Tickets/reserveringen → "🎫 Tickets & Events"
- GitHub/Vercel/Supabase → "💻 Developer & Tech"
- Azure/Google Cloud → "☁️ Cloud & Hosting"
- AI tools (OpenAI, Anthropic) → "🤖 AI & Tools"
- Werk (Linden-IT) → "💼 Werk — Linden-IT"
- Aanmaningen/dringende rekeningen → "⚠️ Openstaand"
- Overig → "📦 Archief"

Antwoord ALLEEN als JSON array:
```json
[{{"index": 0, "naar_map": "exacte mapnaam", "reden": "één zin"}}]
```
Alle indices 0 t/m {len(emails)-1} moeten aanwezig zijn."""

    anth   = Anthropic(api_key=ANTHR_KEY)
    raw    = anth.messages.create(
        model="claude-sonnet-4-6", max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    ).content[0].text

    m = re.search(r'```json\s*(.*?)\s*```', raw, re.DOTALL)
    try:
        beslissingen = json.loads(m.group(1) if m else raw.strip())
    except Exception:
        print("  Sort parse mislukt.")
        return {}

    verplaatst = {}
    log_entries = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    for item in beslissingen:
        idx  = item.get("index", -1)
        naar = item.get("naar_map", "Inbox")
        if idx < 0 or idx >= len(emails):
            continue
        mail = emails[idx]
        log_entries.append({
            "tijdstip": now_str, "mail_id": mail["id"],
            "afzender": mail["from_addr"], "onderwerp": mail["subject"],
            "datum_ontvangen": mail["date"], "naar_map": naar,
            "reden": item.get("reden", ""),
        })
        if naar == "Inbox":
            continue
        dst = folders.get(naar, {}).get("id") if isinstance(folders.get(naar), dict) else None
        if dst and _move(token, mail["id"], dst):
            verplaatst[naar] = verplaatst.get(naar, 0) + 1
            print(f"  ✓ #{idx:02d} → {naar[:40]}")
        time.sleep(0.05)

    if log_entries:
        storage.append_sort_log(log_entries)

    return verplaatst

# ── Claude analyse ────────────────────────────────────────────────────────────

def _build_prompt(priority_emails, new_emails, actie_lijst):
    today      = datetime.now().strftime("%A %d %B %Y")
    open_acties = [a for a in actie_lijst if a.get("status") != "gedaan"]
    lines = [f"Vandaag is het {today}.\n"]

    if open_acties:
        lines.append("## Huidige open actiepunten:")
        for a in open_acties:
            dl = f", deadline: {a['deadline']}" if a.get("deadline") else ""
            lines.append(f"  [ID:{a['id']}] [{a['urgentie'].upper()}] {a['beschrijving']}{dl}")
        lines.append("")

    lines.append("## Prioriteitsmappen (alle ongelezen):")
    for folder_name, emails in priority_emails.items():
        if not emails: continue
        lines.append(f"### {folder_name} ({len(emails)}):")
        for e in emails:
            flags = (["BELANGRIJK"] if e["important"] else []) + (["BIJLAGE"] if e["attachment"] else [])
            flag_str = f" [{', '.join(flags)}]" if flags else ""
            lines.append(
                f"  [{e['date']}]{flag_str} Van: {e['from_name']} <{e['from_addr']}>\n"
                f"  Onderwerp: {e['subject']}\n  Inhoud: {e['body']}\n"
            )

    if any(new_emails.values()):
        lines.append("## Nieuwe mails overige mappen:")
        for folder_name, emails in new_emails.items():
            if not emails: continue
            lines.append(f"### {folder_name} ({len(emails)} nieuw):")
            for e in emails:
                lines.append(f"  [{e['date']}] {e['from_name']} — {e['subject']}")
        lines.append("")

    return f"""Je bent de persoonlijke AI-assistent van Roldi (Nederlandse student/developer, 2026).
Jouw taak: beheer zijn actielijst op basis van zijn mailbox.

{chr(10).join(lines)}

Geef je antwoord ALLEEN als geldig JSON:

```json
{{
  "actielijst_updates": [
    {{
      "actie": "toevoegen|afsluiten|updaten",
      "id": "bestaand-id-of-null",
      "beschrijving": "Wat Roldi concreet moet doen",
      "urgentie": "hoog|middel|laag",
      "afzender": "naam/organisatie",
      "deadline": "YYYY-MM-DD of null",
      "status": "open|gedaan"
    }}
  ],
  "heeft_nieuws": true,
  "telegram_bericht": "Max 15 regels. Gebruik *vet* voor headers. Direct en concreet."
}}
```

Regels:
- "toevoegen": nieuw item (id = null)
- "afsluiten": bestaand item is gedaan (geef ID mee)
- "updaten": urgentie/deadline wijzigen (geef ID mee)
- Voeg alleen toe bij echte actie: betaling, aanmaning, schooltaak, deadline
- Nieuwsbrieven/promoties = nooit toevoegen"""


def _parse(text: str):
    m = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    try:
        return json.loads(m.group(1) if m else text.strip())
    except Exception:
        return None


def _update_actie_lijst(actie_lijst, updates):
    id_map  = {a["id"]: a for a in actie_lijst}
    changed = 0
    for upd in updates:
        actie = upd.get("actie", "toevoegen")
        uid   = upd.get("id")
        if actie == "toevoegen":
            actie_lijst.append({
                "id":           str(uuid.uuid4())[:8],
                "beschrijving": upd.get("beschrijving", ""),
                "urgentie":     upd.get("urgentie", "middel"),
                "afzender":     upd.get("afzender", ""),
                "deadline":     upd.get("deadline"),
                "status":       "open",
                "aangemaakt":   datetime.now().strftime("%Y-%m-%d"),
            })
            changed += 1
        elif actie in ("afsluiten", "updaten") and uid in id_map:
            item = id_map[uid]
            if actie == "afsluiten":
                item["status"] = "gedaan"
            else:
                if upd.get("urgentie"): item["urgentie"] = upd["urgentie"]
                if upd.get("deadline") is not None: item["deadline"] = upd["deadline"]
            changed += 1
    return changed

# ── Publieke entry points ─────────────────────────────────────────────────────

def run_check():
    """Volledige mail check: inbox sorteren + analyse + Telegram rapport."""
    if not ANTHR_KEY:
        print("ANTHROPIC_API_KEY niet ingesteld.")
        return

    print(f"[{datetime.now().strftime('%d %b %H:%M')}] Mail coach gestart")

    try:
        token = outlook_auth.get_token()
    except RuntimeError as e:
        _tg(f"⚠️ Mail coach: {e}")
        return

    folders = outlook_auth.get_all_folders(token)
    since_iso = storage.get_last_check()
    since = (datetime.fromisoformat(since_iso).replace(tzinfo=timezone.utc)
             if since_iso else datetime.now(timezone.utc) - timedelta(hours=24))
    print(f"  Checkt nieuw sinds {since.strftime('%d %b %H:%M')} UTC")

    # Inbox sorteren
    sort_result = _sort_inbox(token, folders)
    if sort_result:
        folders = outlook_auth.get_all_folders(token)

    # Actie-mappen
    priority_emails = {}
    for name in ACTION_FOLDERS:
        f = folders.get(name)
        if f:
            emails = _fetch(token, f["id"], all_mails=True, max_n=30)
            if emails:
                priority_emails[name] = emails

    # Overige mappen — alleen nieuw
    new_emails = {}
    for name in NEW_ONLY_FOLDERS:
        f = folders.get(name)
        if f:
            emails = _fetch(token, f["id"], since_dt=since, max_n=10)
            if emails:
                new_emails[name] = emails

    actie_lijst = storage.get_actie_lijst()
    result      = None

    if any(priority_emails.values()) or any(new_emails.values()):
        print("  Claude analyseert...")
        prompt = _build_prompt(priority_emails, new_emails, actie_lijst)
        raw    = Anthropic(api_key=ANTHR_KEY).messages.create(
            model="claude-sonnet-4-6", max_tokens=3000,
            messages=[{"role": "user", "content": prompt}],
        ).content[0].text
        result = _parse(raw)

        if result:
            changed = _update_actie_lijst(actie_lijst, result.get("actielijst_updates", []))
            if changed:
                storage.save_actie_lijst(actie_lijst)
                print(f"  Actielijst: {changed} wijziging(en)")
        else:
            print("  Parse mislukt.")

    # Telegram rapport
    tg_parts = []
    if sort_result:
        lines = [f"📬 *{sum(sort_result.values())} mail(s) gesorteerd:*"]
        for naam, n in sorted(sort_result.items(), key=lambda x: -x[1]):
            lines.append(f"  • {n} → {naam}")
        tg_parts.append("\n".join(lines))

    if result and result.get("heeft_nieuws"):
        tg_parts.append(result.get("telegram_bericht", ""))

    if tg_parts:
        _tg("\n\n".join(tg_parts))
    else:
        print("  Alles rustig — geen Telegram.")

    storage.save_last_check(datetime.now(timezone.utc).isoformat())
    print("  Klaar.")


def run_briefing():
    """Dagelijkse briefing: stuur de huidige actielijst."""
    actie_lijst = storage.get_actie_lijst()
    open_items  = [a for a in actie_lijst if a.get("status") != "gedaan"]

    if not open_items:
        _tg(f"*📋 Dagelijkse briefing — {datetime.now().strftime('%d %b')}*\n\n✅ Geen open actiepunten.")
        return

    today  = datetime.now().date()
    urgent, middel, laag = [], [], []
    for a in open_items:
        dl  = a.get("deadline")
        dld = None
        if dl:
            try: dld = (datetime.strptime(dl, "%Y-%m-%d").date() - today).days
            except ValueError: pass
        if a.get("urgentie") == "hoog" or (dld is not None and dld <= 7):
            urgent.append((a, dld))
        elif a.get("urgentie") == "middel":
            middel.append((a, dld))
        else:
            laag.append((a, dld))

    lines = [f"*📋 Briefing {datetime.now().strftime('%d %b')} — {len(open_items)} open taken:*\n"]
    for lbl, emo, items in [("URGENT", "🔴", urgent), ("AANDACHT", "🟡", middel), ("LATER", "⚪", laag)]:
        if not items: continue
        lines.append(f"*{emo} {lbl}:*")
        for a, dld in items[:5]:
            dl_str = ""
            if dld is not None:
                dl_str = (" _(verlopen!)_" if dld < 0 else
                          " _(vandaag!)_"  if dld == 0 else
                          f" _(nog {dld}d)_")
            lines.append(f"• {a['beschrijving']}{dl_str}")

    _tg("\n".join(lines))
