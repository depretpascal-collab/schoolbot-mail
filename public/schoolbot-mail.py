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
from http.server import BaseHTTPRequestHandler, HTTPServer

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
            "message_id": msg["Message-ID"] or "", "category": "a_lire", "summary": "",
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


def ai(system, user, max_tokens=2000):
    if (CFG.get("ai") or "ollama") == "ollama":
        req = urllib.request.Request(
            (CFG.get("ollama_url") or "http://127.0.0.1:11434") + "/api/chat",
            data=json.dumps({"model": CFG.get("ollama_model") or "mistral", "stream": False,
                             "options": {"num_predict": max_tokens},
                             "messages": [{"role": "system", "content": system},
                                          {"role": "user", "content": user}]}).encode(),
            headers={"content-type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                return json.load(r)["message"]["content"]
        except urllib.error.URLError:
            raise RuntimeError("IA locale introuvable : installez Ollama (ollama.com) puis lancez "
                               "'ollama pull " + (CFG.get("ollama_model") or "mistral") + "'")
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


def classify(mails):
    if not mails:
        return
    items = [{"id": m["id"], "from": m["from"], "subject": m["subject"], "extrait": m["body"][:500]} for m in mails]
    sys_ = ("Tu tries une boîte mail. Pour chaque mail, choisis une catégorie parmi : urgent, a_traiter, "
            "a_lire, postposable. Ajoute un résumé d'une phrase en français. Réponds UNIQUEMENT par un "
            'tableau JSON : [{"id":"..","category":"..","summary":".."}]')
    txt = ai(sys_, json.dumps(items, ensure_ascii=False), 4000)
    try:
        res = json.loads(re.search(r"\[.*\]", txt, re.S).group(0))
        for r in res:
            if r["id"] in MAILS and r.get("category") in ("urgent", "a_traiter", "a_lire", "postposable"):
                MAILS[r["id"]]["category"], MAILS[r["id"]]["summary"] = r["category"], r.get("summary", "")
    except Exception:
        pass


def draft(m):
    sys_ = ("Tu rédiges des réponses d'email professionnelles, claires et concises, dans la langue du mail "
            "reçu. Réponds UNIQUEMENT avec le corps de la réponse (formule d'appel et de politesse incluses), "
            "sans objet. Si une information manque, laisse un [crochet] à compléter.")
    out = ai(sys_, f"De : {m['from']}\nObjet : {m['subject']}\n\n{m['body']}", 1200).strip()
    return out + (f"\n\n{CFG['signature']}" if CFG.get("signature") else "")


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
:root{--bg:#fff;--fg:#1d1d1f;--mut:#6b7280;--bd:#e5e7eb;--card:#f7f7f8;--ac:#2563eb}
@media(prefers-color-scheme:dark){:root{--bg:#16161a;--fg:#ececf0;--mut:#9ca3af;--bd:#2c2c34;--card:#1e1e24;--ac:#60a5fa}}
*{box-sizing:border-box}body{margin:0;font:15px system-ui,sans-serif;background:var(--bg);color:var(--fg)}
header{display:flex;gap:12px;align-items:center;padding:12px 20px;border-bottom:1px solid var(--bd)}
header h1{font-size:17px;margin:0;flex:1}button{background:var(--ac);color:#fff;border:0;border-radius:8px;padding:8px 14px;font:inherit;cursor:pointer}
button.g{background:var(--card);color:var(--fg);border:1px solid var(--bd)}
#list{padding:16px 20px;max-width:900px;margin:auto}h2{font-size:14px;color:var(--mut);margin:22px 0 8px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:10px 14px;margin-bottom:8px;cursor:pointer}
.card:hover{border-color:var(--ac)}.card b{display:block}.card small{color:var(--mut)}
#split{display:none;grid-template-columns:1fr 1fr;gap:0;height:calc(100vh - 54px)}
.pane{padding:16px 20px;overflow:auto;display:flex;flex-direction:column;gap:10px}.pane+.pane{border-left:1px solid var(--bd)}
pre{white-space:pre-wrap;font:inherit;margin:0}textarea{flex:1;min-height:300px;font:inherit;color:var(--fg);background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:12px;resize:none}
#wiz{display:none;max-width:620px;margin:auto;padding:20px}#wiz label{display:block;font-size:13px;color:var(--mut);margin:12px 0 3px}
input,select{width:100%;padding:8px 10px;font:inherit;color:var(--fg);background:var(--card);border:1px solid var(--bd);border-radius:8px}
.row{display:flex;gap:8px;align-items:center}@media(max-width:800px){#split{grid-template-columns:1fr;height:auto}}
</style>
<header><h1>📬 SchoolBot Mail <small style="font-weight:400;color:var(--mut);font-size:12px">offert par <a href="https://schoolbot.be" target="_blank" style="color:var(--ac)">SchoolBot.be</a></small></h1><button class="g" id="back" style="display:none" onclick="home()">← Liste</button><button class="g" onclick="gear()">⚙</button><button onclick="load()">Actualiser</button></header>
<div id="list"></div>
<div id="wiz"><h3>Connexion de votre boîte mail</h3><small style="color:var(--mut)">Paramètres fournis par votre service informatique ou votre PO. Ils restent sur cet ordinateur.</small>
<div id="fields"></div><div class="row" style="margin-top:18px"><button class="g" onclick="testIt()">Tester la connexion</button><button onclick="saveIt()">Enregistrer et commencer</button></div><p id="wst" style="color:var(--mut)"></p></div>
<div id="split"><div class="pane"><b id="subj"></b><small id="meta"></small><pre id="orig"></pre></div>
<div class="pane"><div class="row"><b style="flex:1">Réponse proposée</b><button class="g" onclick="redo()">↻ Regénérer</button></div>
<textarea id="reply"></textarea><div class="row"><button onclick="sendIt()">Envoyer</button><small id="st"></small></div></div></div>
<script>
const $=s=>document.querySelector(s);let mails=[],cur=null;
const CATS=[["urgent","🔴 Urgent"],["a_traiter","🟠 À traiter"],["a_lire","🔵 À lire"],["postposable","⚪ Postposable"]];
function home(){$('#wiz').style.display='none';$('#split').style.display='none';$('#list').style.display='block';$('#back').style.display='none'}
const F=[["email","Adresse e-mail"],["user","Identifiant de connexion (souvent l'adresse e-mail)"],["pass","Mot de passe","password"],["imap_host","Serveur entrant (IMAP)"],["imap_port","Port IMAP"],["imap_sec","Sécurité IMAP","sec"],["smtp_host","Serveur sortant (SMTP)"],["smtp_port","Port SMTP"],["smtp_sec","Sécurité SMTP","sec"],["ai","Intelligence artificielle","ai"],["ollama_model","Modèle local (Ollama)"],["api_key","Clé API Anthropic (seulement si IA Claude)","password"],["signature","Signature des réponses"]];
function wizard(c){c=c||{};$('#list').style.display='none';$('#split').style.display='none';$('#wiz').style.display='block';$('#back').style.display='none';
 const f=$('#fields');f.textContent='';
 for(const[k,t,ty]of F){const l=document.createElement('label');l.textContent=t;let i;
  if(ty==='sec'){i=document.createElement('select');for(const[v,n]of[['ssl','SSL/TLS'],['starttls','STARTTLS'],['none','Aucune']])i.add(new Option(n,v))}
  else if(ty==='ai'){i=document.createElement('select');for(const[v,n]of[['ollama','Locale (Ollama) — les mails restent sur ce PC'],['claude','Claude (en ligne, clé API)']])i.add(new Option(n,v))}
  else{i=document.createElement('input');i.type=ty||'text'}
  i.id='f_'+k;if(c[k])i.value=c[k];f.append(l,i)}
 if(!c.smtp_sec)$('#f_smtp_sec').value='starttls';if(!c.ollama_model)$('#f_ollama_model').value='mistral';$('#f_email').onblur=guess}
async function guess(){const e=$('#f_email').value;if(!e.includes('@'))return;$('#wst').textContent='Recherche des paramètres…';
 const r=await(await fetch('/api/detect',{method:'POST',body:JSON.stringify({email:e})})).json();if(r.error){$('#wst').textContent=r.error;return}
 for(const k of['user','imap_host','imap_port','imap_sec','smtp_host','smtp_port','smtp_sec'])if(r[k])$('#f_'+k).value=r[k];
 $('#wst').textContent=r.source==='supposition'?'Paramètres supposés : vérifiez-les avec « Tester la connexion ».':'Paramètres trouvés ('+r.source+'). Entrez le mot de passe puis testez.'}
const vals=()=>Object.fromEntries(F.map(([k])=>[k,$('#f_'+k).value]));
async function testIt(){$('#wst').textContent='Test en cours…';const r=await(await fetch('/api/test',{method:'POST',body:JSON.stringify(vals())})).json();$('#wst').textContent='Réception : '+r.imap+' · Envoi : '+r.smtp}
async function saveIt(){await fetch('/api/config',{method:'POST',body:JSON.stringify(vals())});load()}
async function gear(){wizard((await(await fetch('/api/config')).json()).cfg)}
async function load(){home();const c=await(await fetch('/api/config')).json();if(!c.configured||!c.has_key){wizard(c.cfg);return}
 $('#list').textContent="Lecture et analyse des mails du jour…";
 const r=await fetch('/api/mails');const j=await r.json();if(!Array.isArray(j)){$('#list').textContent="Erreur : "+j.error;return}
 mails=j;render()}
function render(){const l=$('#list');l.textContent='';if(!mails.length){l.textContent="Aucun mail aujourd'hui.";return}
 for(const[c,t]of CATS){const ms=mails.filter(m=>m.category===c);if(!ms.length)continue;
  const h=document.createElement('h2');h.textContent=t+' ('+ms.length+')';l.append(h);
  for(const m of ms){const d=document.createElement('div');d.className='card';
   d.innerHTML='<b></b><small></small>';d.children[0].textContent=m.subject;
   d.children[1].textContent=m.from+' — '+m.summary;d.onclick=()=>openMail(m.id);l.append(d)}}}
async function openMail(id){cur=mails.find(m=>m.id===id);$('#list').style.display='none';$('#split').style.display='grid';$('#back').style.display='';
 $('#subj').textContent=cur.subject;$('#meta').textContent=cur.from+' · '+cur.date;$('#orig').textContent=cur.body;$('#st').textContent='';redo()}
async function redo(){$('#reply').value='Rédaction en cours…';
 const r=await fetch('/api/draft',{method:'POST',body:JSON.stringify({id:cur.id})});const j=await r.json();$('#reply').value=j.draft||('Erreur : '+j.error)}
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
    HTTPServer(("127.0.0.1", 8765), H).serve_forever()
