#!/usr/bin/env python3
"""
Telegram Command Bot voor AI Mail Coach.
Draait elke minuut via Task Scheduler.
Verwerkt Telegram opdrachten en reageert direct.

Commando's:
  /help     - Toon beschikbare commando's
  /check    - Voer mail check direct uit
  /status   - Toon samenvatting van je mailbox
  /log      - Laatste regels van coach.log
  /stop     - Bevestig dat de bot actief is
"""

import json
import os
import sys
import msal
import requests
import subprocess
import time
from datetime import datetime, timezone

ACTIE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "actie_lijst.json")
SORT_LOG   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sort_log.json")

# ── Config ───────────────────────────────────────────────────────────────────
COACH_DIR    = r"C:\Users\Roldi\OutlookOrganizer"
OFFSET_FILE  = os.path.join(COACH_DIR, ".telegram_offset")
LOG_FILE     = os.path.join(COACH_DIR, "coach.log")
BOT_LOG      = os.path.join(COACH_DIR, "telegram_bot.log")
TOKEN_CACHE  = os.path.join(COACH_DIR, ".token_cache")
COACH_SCRIPT = os.path.join(COACH_DIR, "ai_coach.py")

# Schrijf altijd naar logbestand (werkt ook zonder console via pythonw.exe)
def log(msg):
    line = f"[{datetime.now().strftime('%d %b %H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(BOT_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

CLIENT_ID  = "ea9e61fd-7233-4740-8073-a51868445bc1"
AUTHORITY  = "https://login.microsoftonline.com/consumers"
SCOPES     = ["Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite"]
GRAPH      = "https://graph.microsoft.com/v1.0"

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT  = os.environ.get("TELEGRAM_CHAT_ID", "")
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")


# ── Telegram helpers ──────────────────────────────────────────────────────────

def tg(method, **kwargs):
    r = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/{method}",
        json=kwargs, timeout=10,
    )
    return r.json()


def send(text, parse_mode="Markdown"):
    tg("sendMessage", chat_id=TELEGRAM_CHAT, text=text, parse_mode=parse_mode)


def get_updates(offset=None):
    params = {"timeout": 0, "allowed_updates": ["message"]}
    if offset:
        params["offset"] = offset
    r = requests.get(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates",
        params=params, timeout=10,
    )
    return r.json().get("result", [])


def load_offset():
    try:
        with open(OFFSET_FILE) as f:
            return int(f.read().strip())
    except (FileNotFoundError, ValueError):
        return None


def save_offset(offset):
    with open(OFFSET_FILE, "w") as f:
        f.write(str(offset))


# ── Outlook helpers ───────────────────────────────────────────────────────────

def get_outlook_token():
    cache = msal.SerializableTokenCache()
    try:
        with open(TOKEN_CACHE) as f:
            cache.deserialize(f.read())
    except FileNotFoundError:
        return None
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        with open(TOKEN_CACHE, "w") as f:
            f.write(cache.serialize())
        return result["access_token"]
    return None


def hdrs(token):
    return {"Authorization": f"Bearer {token}"}


def get_folder_summary(token):
    """Laad alle mappen en geef een status-overzicht."""
    url = f"{GRAPH}/me/mailFolders?$top=100&$select=displayName,totalItemCount,unreadItemCount,childFolderCount"
    r = requests.get(url, headers=hdrs(token), timeout=10)
    if r.status_code != 200:
        return None

    lines = []
    show = {"Inbox", "📌 Actie Vereist", "⚠️ Openstaand", "💸 Financieel",
            "🎒 School — HL Leiden", "🎫 Tickets & Events", "💻 Developer & Tech",
            "🔐 Beveiliging", "📦 Archief", "📰 Nieuwsbrieven", "💼 Werk — Linden-IT"}

    for f in sorted(r.json().get("value", []), key=lambda x: x["displayName"]):
        name = f["displayName"]
        if name not in show:
            continue
        total  = f["totalItemCount"]
        unread = f["unreadItemCount"]
        if unread > 0:
            lines.append(f"  {name}: {total} ({unread} ongelezen)")
        else:
            lines.append(f"  {name}: {total}")

    return "\n".join(lines)


# ── Command handlers ──────────────────────────────────────────────────────────

def cmd_taken():
    try:
        with open(ACTIE_FILE, encoding="utf-8") as f:
            lijst = json.load(f)
    except FileNotFoundError:
        send("📋 Nog geen actielijst — doe eerst een /check.")
        return

    open_items = [a for a in lijst if a.get("status") != "gedaan"]
    gedaan     = [a for a in lijst if a.get("status") == "gedaan"]

    if not open_items:
        send(f"✅ Geen open actiepunten.\n_({len(gedaan)} afgehandeld)_")
        return

    today = datetime.now().date()
    urgent, middel, laag = [], [], []

    for a in open_items:
        dl        = a.get("deadline")
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

    lines = [f"*📋 Actielijst ({len(open_items)} open, {len(gedaan)} gedaan):*\n"]

    for label, emoji, items in [("URGENT", "🔴", urgent), ("AANDACHT", "🟡", middel), ("LATER", "⚪", laag)]:
        if not items:
            continue
        lines.append(f"*{emoji} {label}:*")
        for a, days_left in items:
            dl_str = ""
            if days_left is not None:
                if days_left < 0:
                    dl_str = f" _(verlopen! {abs(days_left)}d geleden)_"
                elif days_left == 0:
                    dl_str = " _(vandaag!)_"
                else:
                    dl_str = f" _(nog {days_left}d)_"
            lines.append(f"• {a['beschrijving']}{dl_str}")

    send("\n".join(lines))


def cmd_herstel(args=""):
    """Toon recente sorteringen. /herstel of /herstel [n] om mail terug naar inbox te zetten."""
    try:
        with open(SORT_LOG, encoding="utf-8") as f:
            log = json.load(f)
    except FileNotFoundError:
        send("📋 Geen sorteer-log gevonden. Doe eerst een /check.")
        return

    moved = [e for e in log if e.get("naar_map") != "Inbox"][:15]
    if not moved:
        send("📋 Nog geen mails automatisch gesorteerd.")
        return

    # Als er een getal meegestuurd wordt → undo die mail
    idx_str = args.strip()
    if idx_str.isdigit():
        idx = int(idx_str)
        if idx < 0 or idx >= len(moved):
            send(f"❌ Ongeldig nummer. Kies 0–{len(moved)-1}.")
            return
        entry = moved[idx]
        token = get_outlook_token()
        if not token:
            send("❌ Geen Outlook token — token verlopen?")
            return
        r = requests.post(
            f"https://graph.microsoft.com/v1.0/me/messages/{entry['mail_id']}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": "inbox"},
            timeout=10,
        )
        if r.status_code in (200, 201):
            send(f"✅ Mail teruggeplaatst in inbox:\n_{entry['onderwerp'][:80]}_")
        else:
            send(f"❌ Mislukt (status {r.status_code}) — mail is mogelijk al verplaatst.")
        return

    # Toon lijst
    lines = ["*📂 Recente sorteringen (stuur /herstel [n] om terug te plaatsen):*\n"]
    for i, e in enumerate(moved):
        lines.append(
            f"*{i}* — {e.get('afzender','?')[:35]}\n"
            f"   _{e.get('onderwerp','')[:60]}_\n"
            f"   → {e.get('naar_map','')} _(op {e.get('tijdstip','?')[:10]})_"
        )
    send("\n".join(lines))


def cmd_help():
    send(
        "*AI Mail Coach — Commando's:*\n\n"
        "/taken — Open actiepunten\n"
        "/check — Mail check nu uitvoeren\n"
        "/status — Mailbox overzicht\n"
        "/herstel — Sorteer-log bekijken / mail terugzetten\n"
        "/log — Laatste log regels\n"
        "/help — Dit bericht\n\n"
        "_Coach checkt automatisch elke 3 uur. Dagelijkse briefing om 08:00._"
    )


def cmd_stop():
    send("✅ Bot is actief en reageert op commando's.")


def cmd_status():
    send("📊 Mailbox ophalen...")
    token = get_outlook_token()
    if not token:
        send("❌ Kan niet inloggen bij Outlook. Token verlopen?")
        return
    summary = get_folder_summary(token)
    if summary:
        now = datetime.now().strftime("%d %b %H:%M")
        send(f"*📬 Mailbox status ({now}):*\n\n{summary}")
    else:
        send("❌ Kon mailbox niet ophalen.")


def cmd_log():
    try:
        with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        last = "".join(lines[-30:]).strip()
        if len(last) > 3500:
            last = "..." + last[-3500:]
        send(f"*📋 Coach log (laatste regels):*\n```\n{last}\n```")
    except FileNotFoundError:
        send("⚠️ Log bestand niet gevonden.")


def cmd_check():
    send("🔍 Mail check gestart...")
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as logf:
            proc = subprocess.Popen(
                [sys.executable, COACH_SCRIPT],
                env=env,
                cwd=COACH_DIR,
                stdout=logf,
                stderr=subprocess.STDOUT,
            )
            proc.wait(timeout=300)

        if proc.returncode != 0:
            send(f"⚠️ Coach afgerond met code {proc.returncode}. Gebruik /log voor details.")
            return

        # Lees laatste logregels om te kijken of er nieuws was
        try:
            with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
                last = "".join(f.readlines()[-10:])
        except Exception:
            last = ""

        if "Geen nieuwe mails" in last or "slaat over" in last:
            send("✅ Check klaar — geen nieuwe mails gevonden.")
        elif "Klaar" in last or "Telegram" in last:
            pass  # coach stuurde al een bericht
        else:
            send("✅ Check klaar. Gebruik /log voor details.")

    except subprocess.TimeoutExpired:
        proc.kill()
        send("⏱️ Coach timeout (5 min) — check /log.")
    except Exception as e:
        send(f"❌ Fout bij starten coach: {e}")


COMMANDS = {
    "/help":       cmd_help,
    "/taken":      cmd_taken,
    "/check":      cmd_check,
    "/controleer": cmd_check,
    "/status":     cmd_status,
    "/herstel":    None,   # afgehandeld apart vanwege argumenten
    "/log":        cmd_log,
    "/stop":       cmd_stop,
    "/start":      cmd_help,
}


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT:
        log("TELEGRAM_BOT_TOKEN of TELEGRAM_CHAT_ID niet ingesteld.")
        return

    offset = load_offset()
    updates = get_updates(offset)

    if not updates:
        log(f"Polling... geen commando's (offset={offset})")
        return

    for update in updates:
        update_id = update["update_id"]
        msg = update.get("message", {})
        chat_id = str(msg.get("chat", {}).get("id", ""))
        text = (msg.get("text") or "").strip().lower()

        # Alleen verwerken als het van de juiste chat komt
        if chat_id != str(TELEGRAM_CHAT):
            save_offset(update_id + 1)
            continue

        # Extraheer het commando en eventuele argumenten
        parts   = text.split(None, 1)
        command = parts[0].split("@")[0] if parts else text
        cmd_arg = parts[1] if len(parts) > 1 else ""

        log(f"Commando ontvangen: {command!r} arg={cmd_arg!r}")

        if command == "/herstel":
            try:
                cmd_herstel(cmd_arg)
            except Exception as e:
                send(f"❌ Fout bij /herstel: {e}")
        else:
            handler = COMMANDS.get(command)
            if handler:
                try:
                    handler()
                except Exception as e:
                    send(f"❌ Fout bij verwerken van {command}: {e}")
            elif command in COMMANDS:
                pass  # None handler = afgehandeld elders
            else:
                if text.startswith("/"):
                    send(f"❓ Onbekend commando: `{command}`\nGebruik /help voor een overzicht.")

        save_offset(update_id + 1)


if __name__ == "__main__":
    main()
