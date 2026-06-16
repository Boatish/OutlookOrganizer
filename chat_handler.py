"""
Conversational Claude handler voor de Telegram chatbot.
Ondersteunt vrije tekst, tool use voor taken/mail/agenda.
"""
import json, os, uuid, re, requests, threading
from datetime import datetime

def _strip_html(html: str) -> str:
    text = re.sub(r'<style[^>]*>.*?</style>', ' ', html, flags=re.DOTALL)
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', text, flags=re.DOTALL)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;|&#160;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    return re.sub(r'\s+', ' ', text).strip()

from anthropic import Anthropic
import storage, outlook_auth

GRAPH = "https://graph.microsoft.com/v1.0"
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

SYSTEM = """Je bent de persoonlijke AI-assistent van Roldi via Telegram.
Roldi is een Nederlandse student/developer (2026) in Den Haag.

Je helpt hem met zijn e-mail, taken, agenda en dagelijkse organisatie.
Je bent direct, persoonlijk en beknopt — dit is een chat, geen rapport.
Gebruik Nederlandse taal. Gebruik *vet* voor nadruk. Geen onnodige emoji.

Als je in mails deadlines of actie-items ziet → stel proactief voor om die op te nemen.
Als een taak een deadline heeft die bijna verstrijkt → benoem dat duidelijk.

Vandaag: {today}

Huidige open taken ({n_open}):
{taken_samenvatting}
"""

TOOLS = [
    {
        "name": "get_taken",
        "description": "Haal de volledige actielijst op (open + gedaan)",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "maak_taak",
        "description": "Maak een nieuw actiepunt aan in de actielijst",
        "input_schema": {
            "type": "object",
            "required": ["beschrijving", "urgentie"],
            "properties": {
                "beschrijving": {"type": "string"},
                "urgentie":     {"type": "string", "enum": ["hoog", "middel", "laag"]},
                "deadline":     {"type": "string", "description": "YYYY-MM-DD of null"},
                "afzender":     {"type": "string"},
            },
        },
    },
    {
        "name": "sluit_taak",
        "description": "Markeer een taak als gedaan/afgehandeld",
        "input_schema": {
            "type": "object",
            "required": ["taak_id"],
            "properties": {
                "taak_id": {"type": "string", "description": "8-char ID van de taak"},
            },
        },
    },
    {
        "name": "update_taak",
        "description": "Pas urgentie of deadline van een bestaande taak aan",
        "input_schema": {
            "type": "object",
            "required": ["taak_id"],
            "properties": {
                "taak_id":  {"type": "string"},
                "urgentie": {"type": "string", "enum": ["hoog", "middel", "laag"]},
                "deadline": {"type": "string", "description": "YYYY-MM-DD"},
            },
        },
    },
    {
        "name": "get_mailbox_status",
        "description": "Haal ongelezen aantallen per mailmap op",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "trigger_mail_check",
        "description": "Start een volledige mail check op de achtergrond (sorteert inbox, analyseert, stuurt Telegram rapport)",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_gesorteerde_mails",
        "description": "Toon recent automatisch gesorteerde mails (voor herstel of overzicht)",
        "input_schema": {
            "type": "object",
            "properties": {
                "max": {"type": "integer", "description": "Max aantal (standaard 10)"},
            },
        },
    },
    {
        "name": "herstel_mail",
        "description": "Zet een automatisch gesorteerde mail terug naar de inbox",
        "input_schema": {
            "type": "object",
            "required": ["mail_id"],
            "properties": {
                "mail_id": {"type": "string", "description": "Graph API mail ID uit de sort log"},
            },
        },
    },
    {
        "name": "lees_map",
        "description": (
            "Lees de inhoud van één of meerdere specifieke mailmappen. "
            "Gebruik dit als je de daadwerkelijke mails wil zien om een vraag te beantwoorden, "
            "bijvoorbeeld 'kijk in mijn schoolmap' of 'zijn er nieuwe facturen'."
        ),
        "input_schema": {
            "type": "object",
            "required": ["mappen"],
            "properties": {
                "mappen": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Lijst van exacte mapnamen, bijv. "
                        "[\"🎒 School — HL Leiden\", \"🧾 Facturen & Bonnetjes\"]. "
                        "Gebruik 'Inbox' voor de inbox."
                    ),
                },
                "alleen_ongelezen": {
                    "type": "boolean",
                    "description": "true = alleen ongelezen mails (standaard), false = ook al gelezen mails",
                },
                "max_per_map": {
                    "type": "integer",
                    "description": "Maximaal aantal mails per map terug te geven (standaard 10, max 25)",
                },
            },
        },
    },
]


def _build_system():
    actie_lijst = storage.get_actie_lijst()
    open_items  = [a for a in actie_lijst if a.get("status") != "gedaan"]
    today       = datetime.now().strftime("%A %d %B %Y")

    if open_items:
        lines = []
        for a in open_items[:8]:
            dl = f" | deadline {a['deadline']}" if a.get("deadline") else ""
            lines.append(f"  [{a['id']}] [{a['urgentie'].upper()}] {a['beschrijving']}{dl}")
        if len(open_items) > 8:
            lines.append(f"  ... en {len(open_items) - 8} meer")
        samenvatting = "\n".join(lines)
    else:
        samenvatting = "  (geen open taken)"

    return SYSTEM.format(today=today, n_open=len(open_items), taken_samenvatting=samenvatting)


def _execute_tool(name: str, inputs: dict):
    if name == "get_taken":
        return storage.get_actie_lijst()

    elif name == "maak_taak":
        lst    = storage.get_actie_lijst()
        new_id = str(uuid.uuid4())[:8]
        lst.append({
            "id":           new_id,
            "beschrijving": inputs["beschrijving"],
            "urgentie":     inputs.get("urgentie", "middel"),
            "afzender":     inputs.get("afzender", ""),
            "deadline":     inputs.get("deadline"),
            "status":       "open",
            "aangemaakt":   datetime.now().strftime("%Y-%m-%d"),
        })
        storage.save_actie_lijst(lst)
        return {"ok": True, "id": new_id}

    elif name == "sluit_taak":
        lst = storage.get_actie_lijst()
        for a in lst:
            if a["id"] == inputs["taak_id"]:
                a["status"] = "gedaan"
                storage.save_actie_lijst(lst)
                return {"ok": True, "beschrijving": a["beschrijving"]}
        return {"ok": False, "error": "Niet gevonden"}

    elif name == "update_taak":
        lst = storage.get_actie_lijst()
        for a in lst:
            if a["id"] == inputs["taak_id"]:
                if "urgentie" in inputs:
                    a["urgentie"] = inputs["urgentie"]
                if "deadline" in inputs:
                    a["deadline"] = inputs["deadline"]
                storage.save_actie_lijst(lst)
                return {"ok": True}
        return {"ok": False, "error": "Niet gevonden"}

    elif name == "get_mailbox_status":
        try:
            token = outlook_auth.get_token()
            r = requests.get(
                f"{GRAPH}/me/mailFolders?$top=100"
                f"&$select=displayName,totalItemCount,unreadItemCount",
                headers=outlook_auth.hdrs(token), timeout=10,
            )
            show = {
                "Inbox", "📌 Actie Vereist", "⚠️ Openstaand", "💸 Financieel",
                "🎒 School — HL Leiden", "🎫 Tickets & Events", "💻 Developer & Tech",
                "🔐 Beveiliging", "📰 Nieuwsbrieven", "💼 Werk — Linden-IT",
                "🧾 Facturen & Bonnetjes",
            }
            return [
                {
                    "map":      f["displayName"],
                    "totaal":   f["totalItemCount"],
                    "ongelezen": f["unreadItemCount"],
                }
                for f in r.json().get("value", [])
                if f["displayName"] in show
            ]
        except Exception as e:
            return {"error": str(e)}

    elif name == "trigger_mail_check":
        def _run():
            try:
                import mail_coach_cloud
                mail_coach_cloud.run_check()
            except Exception as e:
                _send_telegram(f"⚠️ Mail check mislukt: {e}")
        threading.Thread(target=_run, daemon=True).start()
        return {"bericht": "Mail check gestart. Je krijgt een bericht zodra er iets gevonden is."}

    elif name == "get_gesorteerde_mails":
        max_n = inputs.get("max", 10)
        log   = storage.get_sort_log()
        return [e for e in log if e.get("naar_map") != "Inbox"][:max_n]

    elif name == "herstel_mail":
        try:
            token = outlook_auth.get_token()
            r = requests.post(
                f"{GRAPH}/me/messages/{inputs['mail_id']}/move",
                headers=outlook_auth.hdrs(token),
                json={"destinationId": "inbox"}, timeout=10,
            )
            return {"ok": r.status_code in (200, 201)}
        except Exception as e:
            return {"error": str(e)}

    elif name == "lees_map":
        try:
            token        = outlook_auth.get_token()
            folders_data = outlook_auth.get_all_folders(token)
            mappen       = inputs.get("mappen", [])
            ongelezen    = inputs.get("alleen_ongelezen", True)
            max_n        = min(int(inputs.get("max_per_map", 10)), 25)
            select       = "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,body"
            resultaat    = {}

            for map_naam in mappen:
                folder = folders_data.get(map_naam)
                if not folder:
                    # Probeer hoofdletterongevoelig te matchen
                    match = next(
                        (v for k, v in folders_data.items() if k.lower() == map_naam.lower()),
                        None,
                    )
                    if not match:
                        resultaat[map_naam] = {"fout": f"Map '{map_naam}' niet gevonden"}
                        continue
                    folder = match

                flt = "&$filter=isRead eq false" if ongelezen else ""
                url = (
                    f"{GRAPH}/me/mailFolders/{folder['id']}/messages"
                    f"?$top={max_n}&$select={select}{flt}"
                    f"&$orderby=receivedDateTime desc"
                )
                r = requests.get(url, headers=outlook_auth.hdrs(token), timeout=15)
                if r.status_code != 200:
                    resultaat[map_naam] = {"fout": f"API fout {r.status_code}"}
                    continue

                mails = []
                for m in r.json().get("value", []):
                    raw   = m.get("body", {}).get("content", "") or ""
                    btype = m.get("body", {}).get("contentType", "text")
                    body  = _strip_html(raw) if btype == "html" else re.sub(r'\s+', ' ', raw).strip()
                    mails.append({
                        "onderwerp":   m.get("subject", "") or "",
                        "van":         m.get("from", {}).get("emailAddress", {}).get("name", ""),
                        "van_adres":   m.get("from", {}).get("emailAddress", {}).get("address", ""),
                        "datum":       m.get("receivedDateTime", "")[:10],
                        "gelezen":     m.get("isRead", False),
                        "belangrijk":  m.get("importance", "") == "high",
                        "bijlage":     m.get("hasAttachments", False),
                        "inhoud":      body[:600],
                    })
                resultaat[map_naam] = mails

            return resultaat
        except Exception as e:
            return {"error": str(e)}

    return {"error": f"Onbekend tool: {name}"}


def handle_message(chat_id: str, text: str) -> str:
    """
    Verwerk een gebruikersbericht. Retourneert de assistant-reactie als string.
    Converseert in een tool-use loop totdat Claude een definitief antwoord geeft.
    """
    history  = storage.get_conversation(chat_id)
    history.append({"role": "user", "content": text})

    messages = list(history)
    system   = _build_system()

    for _ in range(8):  # max 8 tool-rondes per bericht
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            system=system,
            tools=TOOLS,
            messages=messages,
        )

        if resp.stop_reason == "end_turn":
            antwoord = next(
                (b.text for b in resp.content if hasattr(b, "text")), ""
            )
            history.append({"role": "assistant", "content": antwoord})
            storage.save_conversation(chat_id, history)
            return antwoord

        if resp.stop_reason == "tool_use":
            # Serialiseer content voor de messages-lijst
            raw_content = [
                {"type": "text", "text": b.text} if hasattr(b, "text")
                else {"type": "tool_use", "id": b.id, "name": b.name, "input": b.input}
                for b in resp.content
            ]
            messages.append({"role": "assistant", "content": raw_content})

            tool_results = []
            for b in resp.content:
                if b.type == "tool_use":
                    result = _execute_tool(b.name, b.input)
                    tool_results.append({
                        "type":        "tool_result",
                        "tool_use_id": b.id,
                        "content":     json.dumps(result, ensure_ascii=False),
                    })
            messages.append({"role": "user", "content": tool_results})

        else:
            break

    return "Sorry, er ging iets mis bij het verwerken van je bericht."


def _send_telegram(text: str):
    token   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID", "")
    if not token or not chat_id:
        return
    requests.post(
        f"https://api.telegram.org/bot{token}/sendMessage",
        json={"chat_id": chat_id, "text": text, "parse_mode": "Markdown"},
        timeout=10,
    )
