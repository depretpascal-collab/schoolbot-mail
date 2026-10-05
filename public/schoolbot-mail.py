#!/usr/bin/env python3
"""SchoolBot Mail (MailPilot) - tri d'emails par IA, en local. Offert par SchoolBot.be. Aucune dépendance (Python 3.8+).

Au premier lancement, un assistant demande les paramètres IMAP/SMTP et la clé API
(stockés localement dans ~/.mailpilot/config.json). Détection automatique des serveurs à partir de l'adresse.
IA : locale via Ollama (recommandé, les mails ne quittent pas le PC) ou Claude (clé API, ANTHROPIC_API_KEY).
Lancer : python mailpilot.py  puis ouvrir http://127.0.0.1:8765
"""
import os, re, json, imaplib, smtplib, ssl, urllib.request, urllib.error, datetime, html
from email import message_from_bytes
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, make_msgid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

E = os.environ.get
CONF_PATH = os.path.join(os.path.expanduser("~"), ".mailpilot", "config.json")
CFG = {}


def load_cfg():
    global CFG
    try:
        CFG = json.load(open(CONF_PATH, encoding="utf-8"))
    except Exception:
        CFG = {}


def save_cfg(c):
    os.makedirs(os.path.dirname(CONF_PATH), exist_ok=True)
    CFG.update(c)
    with open(CONF_PATH, "w", encoding="utf-8") as f:
        json.dump(CFG, f)
    try:
        os.chmod(CONF_PATH, 0o600)
    except Exception:
        pass


load_cfg()


def imap_connect(c=None):
    c = c or CFG
    ctx, port, sec = ssl.create_default_context(), int(c.get("imap_port") or 993), c.get("imap_sec", "ssl")
    if sec == "ssl":
        M = imaplib.IMAP4_SSL(c["imap_host"], port, ssl_context=ctx)
    else:
        M = imaplib.IMAP4(c["imap_host"], port)
        if sec == "starttls":
            M.starttls(ctx)
    M.login(c.get("user") or c["email"], c["pass"])
    return M


def smtp_connect(c=None):
    c = c or CFG
    ctx, port, sec = ssl.create_default_context(), int(c.get("smtp_port") or 587), c.get("smtp_sec", "starttls")
    if sec == "ssl":
        s = smtplib.SMTP_SSL(c["smtp_host"], port, context=ctx, timeout=30)
    else:
        s = smtplib.SMTP(c["smtp_host"], port, timeout=30)
        if sec == "starttls":
            s.starttls(context=ctx)
    s.login(c.get("user") or c["email"], c["pass"])
    return s
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MAILS = {}  # id -> mail dict (cache mémoire)


def dh(v):
    try:
        return str(make_header(decode_header(v or "")))
    except Exception:
        return v or ""


def body_of(msg):
    plain = htm = ""
    for p in msg.walk():
        if p.get_content_maintype() != "text" or p.get_filename():
            continue
        raw = p.get_payload(decode=True) or b""
        txt = raw.decode(p.get_content_charset() or "utf-8", "replace")
        if p.get_content_type() == "text/plain":
            plain += txt
        elif p.get_content_type() == "text/html":
            htm += txt
    if plain.strip():
        return plain.strip()
    t = re.sub(r"(?is)<(script|style).*?</\1>", "", htm)
    t = re.sub(r"(?i)<br\s*/?>|</p>", "\n", t)
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def fetch_today(limit=60):
    d = datetime.date.today() - datetime.timedelta(days=7)  # les 7 derniers jours
    since = f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"
    M = imap_connect()
    M.select("INBOX", readonly=True)  # lecture seule : rien n'est marqué comme lu
    _, data = M.search(None, "SINCE", since)
    out = []
    for num in data[0].split()[-limit:][::-1]:
        _, d_ = M.fetch(num, "(BODY.PEEK[])")
        msg = message_from_bytes(d_[0][1])
        mid = num.decode()
        out.append({
            "id": mid, "from": dh(msg["From"]), "subject": dh(msg["Subject"]) or "(sans objet)",
            "date": msg["Date"] or "", "body": body_of(msg)[:6000],
            "reply_to": parseaddr(msg["Reply-To"] or msg["From"])[1],
            "message_id": msg["Message-ID"] or "", "category": "info", "summary": "",
        })
    M.logout()
    return out


KNOWN = {
    "office365": {"imap_host": "outlook.office365.com", "imap_port": "993", "imap_sec": "ssl",
                  "smtp_host": "smtp.office365.com", "smtp_port": "587", "smtp_sec": "starttls"},
    "google": {"imap_host": "imap.gmail.com", "imap_port": "993", "imap_sec": "ssl",
               "smtp_host": "smtp.gmail.com", "smtp_port": "587", "smtp_sec": "starttls"},
}


def _get(url, timeout=6):
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "SchoolBotMail"}), timeout=timeout) as r:
        return r.read().decode("utf-8", "replace")


def _parse_autoconfig(xml, email):
    import xml.etree.ElementTree as ET
    root, out = ET.fromstring(xml), {}
    local, dom = email.split("@")
    for tag, pre in (("incomingServer", "imap"), ("outgoingServer", "smtp")):
        for srv in root.iter(tag):
            if srv.get("type") not in ("imap", "smtp"):
                continue
            g = lambda k: (srv.findtext(k) or "").strip()
            sock = g("socketType").upper()
            out.update({pre + "_host": g("hostname").replace("%EMAILDOMAIN%", dom), pre + "_port": g("port"),
                        pre + "_sec": "ssl" if sock == "SSL" else "starttls" if sock == "STARTTLS" else "none"})
            u = g("username")
            if u and "user" not in out:
                out["user"] = u.replace("%EMAILADDRESS%", email).replace("%EMAILLOCALPART%", local)
            break
    return out if out.get("imap_host") and out.get("smtp_host") else None


def autodetect(email):
    """Trouve les serveurs IMAP/SMTP. Seul le nom de domaine est interrogé, jamais un mail."""
    if "@" not in email:
        return {"error": "adresse invalide"}
    dom = email.split("@")[1].lower().strip()
    for url in (f"https://autoconfig.{dom}/mail/config-v1.1.xml?emailaddress={email}",
                f"https://{dom}/.well-known/autoconfig/mail/config-v1.1.xml",
                f"https://autoconfig.thunderbird.net/v1.1/{dom}"):
        try:
            r = _parse_autoconfig(_get(url), email)
            if r:
                return dict(r, source="autoconfig")
        except Exception:
            pass
    try:  # enregistrements MX (via DNS public) pour repérer Microsoft 365 / Google
        mx = " ".join(a.get("data", "") for a in json.loads(_get(f"https://dns.google/resolve?name={dom}&type=MX")).get("Answer", [])).lower()
        if "outlook.com" in mx:
            return dict(KNOWN["office365"], user=email, source="Microsoft 365")
        if "google.com" in mx or "googlemail.com" in mx:
            return dict(KNOWN["google"], user=email, source="Google")
    except Exception:
        pass
    return {"imap_host": "imap." + dom, "imap_port": "993", "imap_sec": "ssl", "smtp_host": "smtp." + dom,
            "smtp_port": "587", "smtp_sec": "starttls", "user": email, "source": "supposition"}


PREF = ["mistral-small", "mistral-nemo", "gemma2:9b", "gemma2", "llama3.2", "mistral"]
_MODEL = {}


def ollama_models():
    """Modèles installés sur ce PC (aucun mail n'est transmis)."""
    try:
        with urllib.request.urlopen((CFG.get("ollama_url") or "http://127.0.0.1:11434") + "/api/tags", timeout=5) as r:
            return [m["name"] for m in json.load(r).get("models", [])]
    except Exception:
        return []


def pick_model():
    """Le meilleur modèle déjà installé sur ce PC, choisi tout seul."""
    if _MODEL.get("m"):
        return _MODEL["m"]
    got = ollama_models()
    if not got:
        return CFG.get("ollama_model") or "mistral"
    for p in PREF:
        for g in got:
            if g == p or g.split(":")[0] == p:
                _MODEL["m"] = g
                return g
    _MODEL["m"] = got[0]
    return got[0]


def ai(system, user, max_tokens=2000):
    if (CFG.get("ai") or "ollama") == "ollama":
        model = CFG.get("ollama_model") or "auto"
        if model in ("", "auto"):
            model = pick_model()
        req = urllib.request.Request(
            (CFG.get("ollama_url") or "http://127.0.0.1:11434") + "/api/chat",
            data=json.dumps({"model": model, "stream": False,
                             "options": {"num_predict": max_tokens},
                             "messages": [{"role": "system", "content": system},
                                          {"role": "user", "content": user}]}).encode(),
            headers={"content-type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                return json.load(r)["message"]["content"]
        except urllib.error.HTTPError as e:
            if e.code == 404:
                names = ", ".join(ollama_models()) or "aucun"
                raise RuntimeError("Le modèle '" + model + "' n'est pas (encore) téléchargé dans Ollama. "
                                   "Modèles disponibles sur ce PC : " + names +
                                   ". Corrigez le nom dans ⚙ « Modèle local », ou attendez la fin du téléchargement.")
            raise RuntimeError("Erreur de l'IA locale (" + str(e.code) + ") : " + e.read().decode(errors="ignore")[:200])
        except (TimeoutError, OSError) as e:
            if "timed out" in str(e):
                raise RuntimeError("L'IA locale met trop de temps à répondre : le modèle '" + model +
                                   "' est sans doute trop lourd pour ce PC. Essayez un modèle plus léger.")
            raise RuntimeError("Ollama ne répond pas : lancez l'application Ollama (icône lama près de l'horloge), puis réessayez.")
    return claude(system, user, max_tokens)


def claude(system, user, max_tokens=2000):
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": CFG.get("model") or "claude-sonnet-5-5", "max_tokens": max_tokens, "system": system,
                         "messages": [{"role": "user", "content": user}]}).encode(),
        headers={"content-type": "application/json", "x-api-key": CFG.get("api_key") or E("ANTHROPIC_API_KEY", ""),
                 "anthropic-version": "2023-06-01"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return "".join(b.get("text", "") for b in json.load(r)["content"])


CATS = ("urgent", "repondre", "transmettre", "administratif", "info", "pub")


def classify(mails):
    if not mails:
        return
    sys_ = ("Tu tries la boîte mail d'une école (direction / secrétariat). Pour chaque mail, choisis UNE catégorie : "
            "urgent (ne peut pas attendre : absence imprévue, sécurité, délai aujourd'hui), "
            "repondre (une personne attend une réponse : parent, enseignant, partenaire), "
            "transmettre (concerne surtout un collègue ou un autre service), "
            "administratif (factures, circulaires, documents officiels à ranger), "
            "info (à lire, aucune action), pub (publicités, newsletters, notifications automatiques). "
            "Ajoute un résumé d'une phrase en français : qui écrit, pourquoi, ce que ça demande. Réponds UNIQUEMENT par un "
            'tableau JSON : [{"id":"..","category":"..","summary":".."}]')
    for i in range(0, len(mails), 15):  # par lots : plus fiable avec un modèle local
        items = [{"id": m["id"], "from": m["from"], "subject": m["subject"], "extrait": m["body"][:400]}
                 for m in mails[i:i + 15]]
        try:
            txt = ai(sys_, json.dumps(items, ensure_ascii=False), 3000)
            for r in json.loads(re.search(r"\[.*\]", txt, re.S).group(0)):
                rid = str(r.get("id"))
                if rid in MAILS and r.get("category") in CATS:
                    MAILS[rid]["category"], MAILS[rid]["summary"] = r["category"], r.get("summary", "")
        except RuntimeError:
            raise
        except Exception:
            pass


def split_thread(body):
    """Sépare le nouveau message de l'historique cité (lignes '>' ou 'Le ... a écrit :')."""
    lines = body.splitlines()
    for i, l in enumerate(lines):
        s = l.strip()
        if s.startswith(">") or re.match(r"(?i)^(le .{5,120}a écrit|on .{5,120}wrote|-{3,}\s*(original|message))", s) \
                or re.match(r"(?i)^\*?(from|de)\s*:\*?", s):
            new = "\n".join(lines[:i]).strip()
            old = "\n".join(re.sub(r"^\s*>+ ?", "", x) for x in lines[i:]).strip()
            return new or body.strip(), old
    return body.strip(), ""


def draft(m):
    me = CFG.get("email", "")
    sig = CFG.get("signature", "").strip()
    sender_name, sender_addr = parseaddr(m["from"])
    new, history = split_thread(m["body"])
    today = datetime.date.today().strftime("%A %d/%m/%Y")
    sys_ = (
        "Tu es l'assistant de l'UTILISATEUR et tu rédiges SA réponse à un email. "
        f"L'UTILISATEUR est le propriétaire de la boîte {me}"
        + (f", qui signe : « {sig.splitlines()[0]} »" if sig else "") + ". "
        f"L'INTERLOCUTEUR est l'expéditeur du dernier message : {sender_name or sender_addr} <{sender_addr}>. "
        "Règles strictes : tu écris AU NOM DE L'UTILISATEUR, À L'INTERLOCUTEUR ; la formule d'appel s'adresse "
        "à l'interlocuteur (jamais à l'utilisateur) ; ne signe jamais du nom de l'interlocuteur et ne reprends "
        "jamais sa signature. Réponds au DERNIER message ; l'historique sert uniquement de contexte "
        "(ce qui a déjà été proposé, accepté ou demandé, et par qui). Ne réaffirme pas ce que l'interlocuteur a "
        "dit comme si c'était l'utilisateur. Réponds aux questions posées ; si une information manque "
        "(ex. un choix que seul l'utilisateur peut faire), laisse un [crochet] à compléter. "
        "Langue du mail reçu, ton professionnel, concis. Réponds UNIQUEMENT avec le corps de la réponse "
        "(formule d'appel et formule de politesse), SANS signature et sans objet."
    )
    user = (f"Date du jour : {today}\nObjet : {m['subject']}\n\n"
            f"=== DERNIER MESSAGE, reçu le {m['date']}, écrit par l'INTERLOCUTEUR ({sender_name or sender_addr}) ===\n"
            f"{new}\n\n"
            + (f"=== HISTORIQUE PRÉCÉDENT (messages plus anciens, du plus récent au plus ancien ; "
               f"les messages de {me} sont ceux de l'UTILISATEUR) ===\n{history[:4000]}\n" if history else ""))
    out = ai(sys_, user, 1200).strip()
    if sender_name and sender_name.split()[-1].lower() in out.lower().splitlines()[-1].lower():
        out = "\n".join(out.splitlines()[:-1]).strip()  # retire une signature erronée de l'interlocuteur
    return out + (f"\n\n{sig}" if sig else "")


def send(p):
    msg = EmailMessage()
    msg["From"], msg["To"] = CFG.get("email") or CFG["user"], p["to"]
    subj = p["subject"]
    msg["Subject"] = subj if subj.lower().startswith("re:") else "Re: " + subj
    if p.get("message_id"):
        msg["In-Reply-To"] = msg["References"] = p["message_id"]
    msg["Message-ID"] = make_msgid()
    msg.set_content(p["body"])
    s = smtp_connect()
    s.send_message(msg)
    s.quit()


PAGE = """<!doctype html><html lang="fr"><meta charset="utf-8"><title>SchoolBot Mail</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{--bg:#f6f4ef;--card:#ffffff;--bd:#e6e1d6;--fg:#1e2430;--mut:#6b7280;--ac:#2563eb;--ac2:#1d4ed8;--amber:#f59e0b;
--red:#dc2626;--org:#ea580c;--yel:#ca8a04;--blu:#2563eb;--grn:#16a34a;--gry:#6b7280}
*{box-sizing:border-box}html,body{height:100%}
body{margin:0;font:15px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif;color:var(--fg);
background:radial-gradient(900px 420px at 85% -10%,#dbe7ff 0%,transparent 60%),radial-gradient(700px 400px at -5% 110%,#fdeecd 0%,transparent 55%),var(--bg);background-attachment:fixed;overflow-x:hidden}
header{position:sticky;top:0;z-index:5;display:flex;gap:12px;align-items:center;padding:12px 24px;background:rgba(246,244,239,.85);backdrop-filter:blur(12px);border-bottom:1px solid var(--bd)}
.brand{display:flex;align-items:center;gap:10px;flex:1;font-weight:800;font-size:17px;letter-spacing:-.01em}
.brand small{font-weight:400;color:var(--mut);font-size:12px}.brand a{color:var(--ac);text-decoration:none;font-weight:600}
.orb{width:32px;height:32px;border-radius:10px;flex:none;background:linear-gradient(135deg,var(--ac),var(--ac2));display:flex;align-items:center;justify-content:center;box-shadow:0 4px 14px -4px rgba(37,99,235,.5)}
.orb:after{content:"";width:12px;height:12px;border-radius:50%;background:#fff;box-shadow:inset -3px -3px 0 var(--amber)}
.orb.big{width:60px;height:60px;border-radius:18px}.orb.big:after{width:22px;height:22px;box-shadow:inset -5px -5px 0 var(--amber)}
.orb.pulse{animation:pulse 1.6s ease-in-out infinite}
@keyframes pulse{50%{transform:scale(1.07);box-shadow:0 6px 24px -4px rgba(37,99,235,.65)}}
#who{color:var(--mut);font-size:13px}
button{font:inherit;font-weight:600;cursor:pointer;border:0;border-radius:12px;padding:9px 16px;color:#fff;background:var(--ac);box-shadow:0 4px 14px -4px rgba(37,99,235,.45);transition:transform .15s,background .15s}
button:hover{background:var(--ac2);transform:translateY(-1px)}
button.g{background:var(--card);color:var(--fg);border:1px solid var(--bd);box-shadow:none}
button.g:hover{background:#f0ede6}
main{max-width:1100px;margin:auto;padding:28px 24px}
.panel{background:var(--card);border:1px solid var(--bd);border-radius:18px;padding:26px;box-shadow:0 10px 30px -18px rgba(30,36,48,.25);animation:in .4s ease}
@keyframes in{from{opacity:0;transform:translateY(8px)}}
.hero{display:flex;gap:18px;align-items:center}.hero h2{margin:0;font-size:26px;letter-spacing:-.02em}.hero p{margin:2px 0 0;color:var(--mut)}
.bar{height:8px;border-radius:9px;background:#ece8de;margin:22px 0 6px;overflow:hidden}
.bar i{display:block;height:100%;width:0;background:linear-gradient(90deg,var(--amber),var(--ac));border-radius:9px;transition:width .4s}
.count{text-align:right;font-weight:700;color:var(--fg)}
h3.big{font-size:30px;line-height:1.15;margin:26px 0 18px;letter-spacing:-.02em}
.chips{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}
.chip{display:flex;align-items:center;gap:10px;padding:14px 16px;border-radius:14px;background:var(--card);border:1px solid var(--bd);cursor:pointer;transition:.15s}
.chip:hover{border-color:var(--ac);box-shadow:0 6px 18px -10px rgba(37,99,235,.4)}.chip b{font-size:20px}
.dot{width:9px;height:9px;border-radius:50%;flex:none}
.layout{display:grid;grid-template-columns:230px 1fr;gap:20px}
.side{display:flex;flex-direction:column;gap:6px}
.side div{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:12px;cursor:pointer;color:var(--mut);font-weight:500}
.side div:hover{background:#ece8de;color:var(--fg)}
.side div.on{background:#e3ecfd;color:var(--ac);font-weight:700}.side span{margin-left:auto;font-weight:700}
.mail{display:flex;gap:12px;align-items:flex-start;padding:14px 16px;border-radius:14px;background:var(--card);border:1px solid var(--bd);margin-bottom:10px;cursor:pointer;transition:.15s;animation:in .3s ease both}
.mail:hover{border-color:var(--ac);box-shadow:0 6px 18px -10px rgba(37,99,235,.35)}.mail .dot{margin-top:8px}
.mail b{display:block}.mail small{color:var(--mut)}.mail .tag{margin-left:auto;font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;border:1px solid currentColor;white-space:nowrap}
.ready{font-size:11px;font-weight:600;color:var(--grn);margin-top:4px}
.split{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
.label{font-size:12px;letter-spacing:.12em;font-weight:700;color:var(--mut);display:flex;align-items:center;gap:8px}
pre{white-space:pre-wrap;font:inherit;margin:12px 0 0;color:#3a4150;max-height:65vh;overflow:auto}
.reply{background:linear-gradient(160deg,#eef4ff,#fdf6e7);border:1px solid #c7d8f8;box-shadow:0 16px 40px -20px rgba(37,99,235,.35)}
textarea{width:100%;min-height:330px;margin-top:14px;font:inherit;line-height:1.6;color:var(--fg);background:#fff;border:1px solid var(--bd);border-radius:14px;padding:16px;resize:vertical}
textarea[readonly]{background:transparent;border-color:transparent}
.note{font-size:12px;color:var(--mut);margin:8px 0 12px}
.row{display:flex;gap:10px;align-items:center}.row .grow{flex:1}
label{display:block;font-size:13px;color:var(--mut);margin:12px 0 4px}
input,select{width:100%;padding:10px 12px;font:inherit;color:var(--fg);background:#fff;border:1px solid var(--bd);border-radius:10px}
input:focus,select:focus,textarea:focus{outline:2px solid rgba(37,99,235,.35);border-color:var(--ac)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:0 14px}
#st,#wst{color:var(--mut);font-size:13px}
@media(max-width:850px){.layout,.split,.grid2{grid-template-columns:1fr}}
</style>
<header><div class="brand"><div class="orb"></div>SchoolBot Mail <small>offert par <a href="https://schoolbot.be" target="_blank">SchoolBot.be</a></small></div>
<span id="who"></span><button class="g" onclick="gear()">⚙ Réglages</button><button onclick="load()">Traiter mes mails</button></header>
<main id="app"></main>
<script>
const $=s=>document.querySelector(s);const app=$('#app');let mails=[],cur=null,filter=null,cfg={};
const CATS={urgent:['Urgent','var(--red)','urgents'],repondre:['À répondre','var(--org)','réponses à rédiger'],transmettre:['À transmettre','var(--blu)','à transmettre'],administratif:['Administratif','var(--yel)','documents à ranger'],info:['À lire','var(--grn)','à lire'],pub:['Pubs & notifications','var(--gry)','pubs et notifications']};
const esc=s=>String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const name=f=>(f||'').replace(/<.*>/,'').replace(/"/g,'').trim()||f;
const F=[["email","Adresse e-mail"],["user","Identifiant de connexion"],["pass","Mot de passe","password"],["imap_host","Serveur entrant (IMAP)"],["imap_port","Port IMAP"],["imap_sec","Sécurité IMAP","sec"],["smtp_host","Serveur sortant (SMTP)"],["smtp_port","Port SMTP"],["smtp_sec","Sécurité SMTP","sec"],["ai","Intelligence artificielle","ai"],["ollama_model","Modèle local (Ollama)"],["api_key","Clé API Anthropic (seulement si IA Claude)","password"],["signature","Signature des réponses"]];
function wizard(c){c=c||{};app.innerHTML='<div class="panel" style="max-width:720px;margin:auto"><div class="hero"><div class="orb big"></div><div><h2>Connectons votre boîte mail</h2><p>Paramètres fournis par votre service informatique. Ils restent sur cet ordinateur.</p></div></div><div class="grid2" id="fields"></div><div class="row" style="margin-top:20px"><button class="g" onclick="testIt()">Tester la connexion</button><button onclick="saveIt()">Enregistrer et commencer</button></div><p id="wst"></p></div>';
 const f=$('#fields');
 for(const[k,t,ty]of F){const w=document.createElement('div');const l=document.createElement('label');l.textContent=t;let i;
  if(ty==='sec'){i=document.createElement('select');for(const[v,n]of[['ssl','SSL/TLS'],['starttls','STARTTLS'],['none','Aucune']])i.add(new Option(n,v))}
  else if(ty==='ai'){i=document.createElement('select');for(const[v,n]of[['ollama','Locale — les mails restent sur ce PC'],['claude','Claude (en ligne, clé API)']])i.add(new Option(n,v))}
  else{i=document.createElement('input');i.type=ty||'text'}
  i.id='f_'+k;if(c[k])i.value=c[k];w.append(l,i);f.append(w)}
 if(!c.smtp_sec)$('#f_smtp_sec').value='starttls';if(!c.ollama_model)$('#f_ollama_model').value='mistral';$('#f_email').onblur=guess}
async function guess(){const e=$('#f_email').value;if(!e.includes('@'))return;$('#wst').textContent='Recherche des paramètres…';
 const r=await(await fetch('/api/detect',{method:'POST',body:JSON.stringify({email:e})})).json();if(r.error){$('#wst').textContent=r.error;return}
 for(const k of['user','imap_host','imap_port','imap_sec','smtp_host','smtp_port','smtp_sec'])if(r[k])$('#f_'+k).value=r[k];
 $('#wst').textContent=r.source==='supposition'?'Paramètres supposés : vérifiez-les avec « Tester la connexion ».':'Paramètres trouvés ('+r.source+'). Entrez le mot de passe puis testez.'}
const vals=()=>Object.fromEntries(F.map(([k])=>[k,$('#f_'+k).value]));
async function testIt(){$('#wst').textContent='Test en cours…';const r=await(await fetch('/api/test',{method:'POST',body:JSON.stringify(vals())})).json();$('#wst').textContent='Réception : '+r.imap+' · Envoi : '+r.smtp}
async function saveIt(){await fetch('/api/config',{method:'POST',body:JSON.stringify(vals())});load()}
async function gear(){wizard((await(await fetch('/api/config')).json()).cfg)}
async function load(){const c=await(await fetch('/api/config')).json();cfg=c.cfg||{};$('#who').textContent=cfg.email||'';if(!c.configured||!c.has_key){wizard(c.cfg);return}
 app.innerHTML='<div class="panel"><div class="hero"><div class="orb big pulse"></div><div><h2 id="t">Je lis vos mails</h2><p id="s">Un par un. Qui écrit, pourquoi, et ce que ça demande de vous.</p></div></div><div class="bar"><i id="bar"></i></div><div class="count" id="cnt"></div></div>';
 let p=0;const tick=setInterval(()=>{p+=(92-p)*0.03;$('#bar').style.width=p+'%'},300);
 const phr=['Je lis les sujets et les expéditeurs…','Je repère ce qui est urgent…','Je classe par catégorie…','Je résume chaque mail…'];let k=0;const pt=setInterval(()=>{$('#s').textContent=phr[k++%phr.length]},2600);
 const r=await fetch('/api/mails');const j=await r.json();clearInterval(tick);clearInterval(pt);
 if(!Array.isArray(j)){app.innerHTML='<div class="panel"><h2>Oups.</h2><p>'+esc(j.error)+'</p></div>';return}
 mails=j;$('#bar').style.width='100%';$('#cnt').textContent=mails.length+' / '+mails.length;setTimeout(summary,500);prefetch()}
function counts(){const c={};for(const m of mails)c[m.category]=(c[m.category]||0)+1;return c}
function summary(){const c=counts();const urg=c.urgent||0;
 let h='<div class="panel"><div class="hero"><div class="orb big"></div><div><h2>Compris.</h2><p>'+(urg?urg+' mail(s) urgent(s) vous attendent.':"Rien d'urgent qui vous attend. Tout a une place.")+'</p></div></div>';
 h+='<h3 class="big">'+mails.length+' mails compris.<br>Voilà ce que ça donne :</h3><div class="chips">';
 for(const[k,[,col,lab]]of Object.entries(CATS))if(c[k])h+='<div class="chip" onclick="list(&quot;'+k+'&quot;)"><span class="dot" style="color:'+col+';background:'+col+'"></span><b>'+c[k]+'</b> '+lab+'</div>';
 h+='</div><div class="row" style="margin-top:22px"><span class="grow"></span><button onclick="list(c0())">Voir les mails →</button></div></div>';app.innerHTML=h}
const c0=()=>Object.keys(CATS).find(k=>counts()[k])||null;
function list(f){filter=f;const c=counts();
 let h='<div class="layout"><div class="side">';
 for(const[k,[lab,col]]of Object.entries(CATS))if(c[k])h+='<div class="'+(k===filter?'on':'')+'" onclick="list(&quot;'+k+'&quot;)"><span class="dot" style="color:'+col+';background:'+col+'"></span>'+lab+'<span>'+c[k]+'</span></div>';
 h+='<div onclick="summary()" style="margin-top:10px">← Vue d’ensemble</div></div><div>';
 for(const m of mails.filter(m=>m.category===filter)){const[lab,col]=CATS[m.category];
  h+='<div class="mail" onclick="openMail(&quot;'+m.id+'&quot;)"><span class="dot" style="color:'+col+';background:'+col+'"></span><div><b>'+esc(name(m.from))+' · '+esc(m.subject)+'</b><small>'+esc(m.summary)+'</small>'+(m.draft?'<div class="ready">● Réponse prête</div>':'')+'</div><span class="tag" style="color:'+col+'">'+lab+'</span></div>'}
 app.innerHTML=h+'</div></div>'}
function getDraft(m){if(!m.p)m.p=fetch('/api/draft',{method:'POST',body:JSON.stringify({id:m.id})}).then(r=>r.json()).then(j=>{if(j.draft)m.draft=j.draft;else m.p=null;return j});return m.p}
async function prefetch(){for(const m of mails.filter(m=>m.category==='urgent'||m.category==='repondre')){try{await getDraft(m)}catch(e){}}}
async function openMail(id){cur=mails.find(m=>m.id===id);
 app.innerHTML='<div class="row" style="margin-bottom:14px"><button class="g" onclick="list(filter)">← Retour</button></div><div class="split"><div class="panel"><div class="label">MAIL REÇU</div><h2 style="margin:8px 0 2px;font-size:20px">'+esc(cur.subject)+'</h2><small style="color:var(--mut)">'+esc(cur.from)+' · '+esc(cur.date)+'</small><pre>'+esc(cur.body)+'</pre></div>'+
 '<div class="panel reply"><div class="label"><span class="dot" style="color:var(--yel);background:var(--yel)"></span>RÉPONSE PRÉPARÉE PAR L’IA</div><h2 style="margin:8px 0 2px;font-size:18px">Re: '+esc(cur.subject.replace(/^re: */i,''))+'</h2><small style="color:var(--mut)">À : '+esc(cur.from)+'</small>'+
 '<textarea id="reply" readonly>Rédaction en cours…</textarea><div class="note">✦ Rédigée à partir de l’historique — vous validez, il envoie.</div><div class="row"><button class="g" onclick="edit()">✎ Modifier</button><button class="g" onclick="redo()">↻ Regénérer</button><button class="grow" onclick="sendIt()">✉ Envoyer</button></div><p id="st"></p></div></div>';
 const j=await getDraft(cur);$('#reply').value=cur.draft||('Erreur : '+j.error)}
function edit(){const t=$('#reply');t.readOnly=false;t.focus()}
async function redo(){cur.p=null;cur.draft=null;$('#reply').value='Rédaction en cours…';const j=await getDraft(cur);$('#reply').value=cur.draft||('Erreur : '+j.error)}
async function sendIt(){if(!confirm('Envoyer cette réponse à '+cur.reply_to+' ?'))return;$('#st').textContent='Envoi…';
 const r=await fetch('/api/send',{method:'POST',body:JSON.stringify({to:cur.reply_to,subject:cur.subject,body:$('#reply').value,message_id:cur.message_id})});
 const j=await r.json();$('#st').textContent=j.ok?'✅ Envoyé':'Erreur : '+j.error}
load();
</script></html>"""


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def reply(self, obj, code=200, ctype="application/json"):
        data = obj.encode() if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path == "/":
            return self.reply(PAGE, ctype="text/html")
        if self.path == "/api/config":
            safe = {k: v for k, v in CFG.items() if k not in ("pass", "api_key")}
            return self.reply({"configured": bool(CFG.get("imap_host") and CFG.get("pass")), "cfg": safe,
                               "has_key": (CFG.get("ai") or "ollama") == "ollama" or bool(CFG.get("api_key") or E("ANTHROPIC_API_KEY"))})
        if self.path == "/api/mails":
            try:
                ms = fetch_today()
                MAILS.clear()
                MAILS.update({m["id"]: m for m in ms})
                classify(ms)
                return self.reply(ms)
            except Exception as e:
                return self.reply({"error": str(e)}, 500)
        self.reply({"error": "not found"}, 404)

    def do_POST(self):
        p = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or "{}")
        try:
            if self.path == "/api/detect":
                return self.reply(autodetect(p.get("email", "")))
            if self.path == "/api/test":
                c = dict(CFG)
                c.update({k: v for k, v in p.items() if v != ""})
                res = {}
                for name, fn in (("imap", imap_connect), ("smtp", smtp_connect)):
                    try:
                        x = fn(c)
                        x.logout() if name == "imap" else x.quit()
                        res[name] = "OK"
                    except Exception as e:
                        res[name] = "échec (" + str(e)[:120] + ")"
                return self.reply(res)
            if self.path == "/api/config":
                save_cfg({k: v for k, v in p.items() if v != ""})
                return self.reply({"ok": True})
            if self.path == "/api/draft":
                return self.reply({"draft": draft(MAILS[p["id"]])})
            if self.path == "/api/send":
                send(p)
                return self.reply({"ok": True})
        except Exception as e:
            return self.reply({"error": str(e)}, 500)
        self.reply({"error": "not found"}, 404)


if __name__ == "__main__":
    print("SchoolBot Mail → http://127.0.0.1:8765")
    try:
        import webbrowser; webbrowser.open("http://127.0.0.1:8765")
    except Exception:
        pass
    ThreadingHTTPServer(("127.0.0.1", 8765), H).serve_forever()
