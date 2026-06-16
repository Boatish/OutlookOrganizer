"""
FastAPI server — draait op Railway (of lokaal).
- Verwerkt Telegram webhook berichten
- APScheduler: mail check elke 3 uur + briefing 08:00
- Chatbot via chat_handler.py
"""
import os, logging
from datetime import datetime
from contextlib import asynccontextmanager

import httpx
from fastapi import FastAPI, Request
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

import chat_handler, mail_coach_cloud

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

TG_TOKEN  = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TG_CHAT   = os.environ.get("TELEGRAM_CHAT_ID", "")
WEBHOOK   = os.environ.get("WEBHOOK_URL", "")          # bijv. https://xxx.railway.app/telegram
AMS       = pytz.timezone("Europe/Amsterdam")
scheduler = AsyncIOScheduler(timezone=AMS)


# ── Scheduler jobs ────────────────────────────────────────────────────────────

def job_mail_check():
    log.info("Scheduled mail check gestart")
    try:
        mail_coach_cloud.run_check()
    except Exception as e:
        log.error(f"Mail check fout: {e}")


def job_briefing():
    log.info("Dagelijkse briefing verstuurd")
    try:
        mail_coach_cloud.run_briefing()
    except Exception as e:
        log.error(f"Briefing fout: {e}")


# ── App lifecycle ─────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Webhook registreren
    if TG_TOKEN and WEBHOOK:
        async with httpx.AsyncClient() as client:
            r = await client.post(
                f"https://api.telegram.org/bot{TG_TOKEN}/setWebhook",
                json={"url": WEBHOOK, "allowed_updates": ["message"]},
            )
            log.info(f"Webhook ingesteld: {r.json()}")

    # Scheduler starten
    scheduler.add_job(job_mail_check, CronTrigger(hour="*/3", minute=5, timezone=AMS))
    scheduler.add_job(job_briefing,   CronTrigger(hour=8,     minute=0, timezone=AMS))
    scheduler.start()
    log.info("Scheduler gestart (mail check */3u, briefing 08:00 AMS)")

    yield

    scheduler.shutdown()


app = FastAPI(lifespan=lifespan)


# ── Telegram helpers ──────────────────────────────────────────────────────────

async def send(text: str, parse_mode="Markdown"):
    async with httpx.AsyncClient() as client:
        await client.post(
            f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage",
            json={"chat_id": TG_CHAT, "text": text, "parse_mode": parse_mode},
            timeout=10,
        )


# ── Directe commando handlers (geen Claude) ───────────────────────────────────

async def cmd_log():
    log_path = os.path.join(os.path.dirname(__file__), "coach.log")
    try:
        with open(log_path, encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        last = "".join(lines[-30:]).strip()
        if len(last) > 3500:
            last = "..." + last[-3500:]
        await send(f"*📋 Log (laatste regels):*\n```\n{last}\n```")
    except FileNotFoundError:
        await send("⚠️ Geen log bestand.")


async def cmd_status():
    import storage
    try:
        import outlook_auth, requests as req
        token = outlook_auth.get_token()
        r = req.get(
            f"https://graph.microsoft.com/v1.0/me/mailFolders"
            f"?$top=100&$select=displayName,totalItemCount,unreadItemCount",
            headers=outlook_auth.hdrs(token), timeout=10,
        )
        show = {"Inbox", "📌 Actie Vereist", "⚠️ Openstaand", "🧾 Facturen & Bonnetjes",
                "🎒 School — HL Leiden", "🎫 Tickets & Events", "💻 Developer & Tech",
                "🔐 Beveiliging", "📰 Nieuwsbrieven", "💼 Werk — Linden-IT"}
        lines = []
        for f in sorted(r.json().get("value", []), key=lambda x: x["displayName"]):
            if f["displayName"] not in show: continue
            u = f["unreadItemCount"]
            t = f["totalItemCount"]
            lines.append(f"  {f['displayName']}: {t}" + (f" ({u} ongelezen)" if u else ""))
        now = datetime.now(AMS).strftime("%d %b %H:%M")
        await send(f"*📬 Mailbox ({now}):*\n\n" + "\n".join(lines))
    except Exception as e:
        await send(f"❌ Status ophalen mislukt: {e}")


# ── Webhook endpoint ──────────────────────────────────────────────────────────

@app.post("/telegram")
async def telegram_webhook(request: Request):
    try:
        data    = await request.json()
        update  = data.get("message", {})
        chat_id = str(update.get("chat", {}).get("id", ""))
        text    = (update.get("text") or "").strip()

        if not text or chat_id != str(TG_CHAT):
            return {"ok": True}

        log.info(f"Bericht ontvangen: {text[:60]!r}")

        # Directe commando's zonder Claude
        cmd = text.lower().split()[0].split("@")[0]
        if cmd == "/log":
            await cmd_log()
        elif cmd == "/status":
            await cmd_status()
        elif cmd in ("/start", "/help"):
            await send(
                "*Hallo Roldi! 👋*\n\n"
                "Ik ben je persoonlijke mail-assistent. Stel me gewoon vragen, zoals:\n\n"
                "• _Wat moet ik deze week doen?_\n"
                "• _Heb ik nog urgente taken?_\n"
                "• _Check mijn mail_\n"
                "• _Maak taak: bel Infomedics voor vrijdag_\n"
                "• _De tandarts is betaald_\n\n"
                "Handige commando's: /status /log"
            )
        else:
            # Alles andere → Claude chatbot
            resp = chat_handler.handle_message(chat_id, text)
            await send(resp)

    except Exception as e:
        log.error(f"Webhook fout: {e}", exc_info=True)

    return {"ok": True}


@app.get("/")
async def health():
    return {
        "status": "ok",
        "time":   datetime.now(AMS).strftime("%Y-%m-%d %H:%M %Z"),
        "jobs":   [str(j) for j in scheduler.get_jobs()],
    }
