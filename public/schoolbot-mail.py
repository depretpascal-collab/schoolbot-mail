#!/usr/bin/env python3
"""SchoolBot Mail (MailPilot) - tri d'emails par IA, en local. Développé par Educlan Asbl. Aucune dépendance (Python 3.8+).

Au premier lancement, un assistant connecte la boîte mail :
- compte Microsoft 365 (écoles) : on se connecte avec son compte, SchoolBot Mail ne voit jamais le mot de passe ;
- ou serveurs IMAP/SMTP saisis à la main (repli), détectés automatiquement à partir de l'adresse.
IA : locale via Ollama (recommandé, les mails ne quittent pas le PC) ou Claude (clé API, ANTHROPIC_API_KEY).
Lancer : double-clic sur SchoolBot Mail (ou python schoolbot-mail.py)
"""
import os, re, json, base64, imaplib, smtplib, ssl, urllib.request, urllib.error, urllib.parse, datetime, html
from email import message_from_bytes
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, make_msgid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

E = os.environ.get
CONF_PATH = os.path.join(os.path.expanduser("~"), ".mailpilot", "config.json")
CFG = {}

# Numéro de version : à augmenter à chaque nouvelle version, en même temps que version.json
VERSION = "1.6"
# Adresse du dépôt GitHub (ex. "pascal/schoolbot-mail") ; vide = pas de vérification
GITHUB_REPO = "depretpascal-collab/schoolbot-mail"

# ---------- Microsoft 365 : on se connecte avec son compte, l'app ne voit jamais le mot de passe ----------
# Identifiant de l'application à créer une seule fois dans Microsoft Entra (gratuit).
# Une fois renseigné ici, tous les utilisateurs en bénéficient sans rien configurer.
MS_CLIENT_ID = "7659db28-79ec-47b1-b055-a36d3394e682"
MS_SCOPE = "https://graph.microsoft.com/.default offline_access"
MS_AUTH = "https://login.microsoftonline.com/common/oauth2/v2.0"
MS_API = "https://graph.microsoft.com/v1.0"
MS = {"state": "idle", "user_code": "", "url": "", "msg": ""}


def _vtuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))) or (0,)


def check_update():
    """Lit version.json sur GitHub. Seul un numéro de version est lu : aucune donnée n'est envoyée."""
    res = {"current": VERSION, "available": False}
    if not GITHUB_REPO:
        return res
    try:
        url = "https://raw.githubusercontent.com/%s/main/version.json" % GITHUB_REPO
        with urllib.request.urlopen(url, timeout=5) as r:
            info = json.loads(r.read().decode("utf-8"))
        latest = str(info.get("version", ""))
        res.update(latest=latest, notes=info.get("notes", ""),
                   url=info.get("url") or "https://github.com/%s/releases/latest" % GITHUB_REPO,
                   available=_vtuple(latest) > _vtuple(VERSION))
    except Exception:
        pass
    return res


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
    return strip_html(htm)


def strip_html(t):
    """HTML -> texte lisible (les mails d'école arrivent souvent en HTML)."""
    t = re.sub(r"(?is)<(script|style).*?</\1>", "", t or "")
    t = re.sub(r"(?i)<br\s*/?>|</p>|</div>", "\n", t)
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def fetch_today(limit=60):
    if is_ms():
        return graph_messages(limit)
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
    if dom in ("gmail.com", "googlemail.com"):
        return dict(KNOWN["google"], user=email, source="Google")
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


# ---------- Microsoft 365 : connexion au compte de l'école ----------
def ms_client_id():
    return (CFG.get("ms_client_id") or MS_CLIENT_ID).strip()


def is_ms():
    return bool(CFG.get("ms_refresh") or CFG.get("ms_access"))


def _post(url, data=None, ctype="application/x-www-form-urlencoded", tok=None):
    body = None
    h = {"User-Agent": "SchoolBotMail"}
    if data is not None:
        body = (urllib.parse.urlencode(data).encode() if ctype == "application/x-www-form-urlencoded"
                else json.dumps(data).encode())
        h["Content-Type"] = ctype
    if tok:
        h["Authorization"] = "Bearer " + tok
    req = urllib.request.Request(url, data=body, headers=h)
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read().decode("utf-8")
    return json.loads(raw) if raw else {}


def ms_account(tok):
    """Lit l'adresse du compte dans le jeton Microsoft (aucune donnée n'est envoyée nulle part)."""
    try:
        p = tok.split(".")[1]
        p += "=" * (-len(p) % 4)
        c = json.loads(base64.urlsafe_b64decode(p))
        return c.get("mail") or c.get("email") or c.get("preferred_username") or ""
    except Exception:
        return ""


def ms_start():
    """Demande à Microsoft un code temporaire : l'utilisateur le tape, SchoolBot Mail ne voit pas le mot de passe."""
    if MS.get("state") == "ok":
        return ms_public()
    cid = ms_client_id()
    if not cid:
        raise RuntimeError("La connexion Microsoft n'est pas encore activée dans cette version de SchoolBot Mail. "
                           "Écrivez à contact@educlan.org : c'est une seule ligne à ajouter par Educlan.")
    d = _post(MS_AUTH + "/devicecode", {"client_id": cid, "scope": MS_SCOPE})
    MS.update(state="waiting", user_code=d["user_code"],
              url=d.get("verification_uri") or "https://microsoft.com/devicelogin",
              device_code=d["device_code"], interval=int(d.get("interval") or 5),
              expires=time.time() + int(d.get("expires_in") or 600),
              msg="Tapez ce code sur la page de Microsoft.")
    threading.Thread(target=ms_poll, daemon=True).start()
    return ms_public()


def ms_public():
    return {k: MS.get(k, "") for k in ("state", "user_code", "url", "msg")}


def ms_explain(code, desc):
    if code in ("access_denied", "interaction_required", "consent_required"):
        return ("Microsoft a refusé : « " + (desc or "")[:160] + " ». Si l'école interdit d'autoriser elle-même "
                "des applications, le service informatique de l'école devra donner l'accès.")
    if code == "expired_token":
        return "Le code a expiré : relancez « Se connecter avec Microsoft »."
    return "Microsoft a refusé la connexion : " + (desc or code or "erreur inconnue")[:200]


def ms_poll():
    cid = ms_client_id()
    time.sleep(MS.get("interval", 5))
    while MS.get("state") == "waiting":
        if time.time() > MS.get("expires", 0):
            MS.update(state="error", msg="Le code a expiré. Relancez la connexion Microsoft.")
            return
        try:
            r = _post(MS_AUTH + "/token", {"client_id": cid,
                                           "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                                           "device_code": MS["device_code"]})
        except urllib.error.HTTPError as e:
            try:
                err = json.loads(e.read().decode("utf-8", "replace"))
            except Exception:
                err = {}
            c = err.get("error", "")
            if c == "authorization_pending":
                time.sleep(MS.get("interval", 5)); continue
            if c == "slow_down":
                MS["interval"] = int(MS.get("interval", 5)) + 5; time.sleep(MS["interval"]); continue
            MS.update(state="error", msg=ms_explain(c, err.get("error_description", "")))
            return
        except Exception:
            time.sleep(MS.get("interval", 5)); continue
        acct = ms_account(r.get("id_token") or "")
        save_cfg({"ms_access": r["access_token"], "ms_refresh": r.get("refresh_token", ""),
                  "ms_expiry": time.time() + int(r.get("expires_in", 3600)),
                  "ms_email": acct or CFG.get("email", ""),
                  "email": acct or CFG.get("email", ""), "provider": "microsoft"})
        MS.update(state="ok", msg="Connecté !")
        _MODEL.clear()
        return


def ms_token():
    """Jeton d'accès en cours de validité, renouvelé en silence si nécessaire."""
    if CFG.get("ms_access") and time.time() < float(CFG.get("ms_expiry", 0) or 0) - 60:
        return CFG["ms_access"]
    rt = CFG.get("ms_refresh")
    if not rt:
        raise RuntimeError("La connexion Microsoft est périmée : reconnectez-vous avec votre compte.")
    try:
        r = _post(MS_AUTH + "/token", {"client_id": ms_client_id(), "grant_type": "refresh_token",
                                       "refresh_token": rt, "scope": MS_SCOPE})
    except urllib.error.HTTPError:
        save_cfg({"ms_access": "", "ms_refresh": "", "ms_expiry": 0})
        raise RuntimeError("Microsoft ne reconnaît plus cette session : reconnectez-vous avec votre compte.")
    save_cfg({"ms_access": r["access_token"], "ms_refresh": r.get("refresh_token", rt),
              "ms_expiry": time.time() + int(r.get("expires_in", 3600))})
    return r["access_token"]


def graph(path, method="GET", body=None):
    req = urllib.request.Request(MS_API + path, method=method,
                                 data=json.dumps(body).encode() if body is not None else None,
                                 headers={"Authorization": "Bearer " + ms_token(),
                                          "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            raw = r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        det = e.read().decode("utf-8", "replace")[:250]
        if e.code in (401, 403):
            raise RuntimeError("Microsoft a refusé l'accès à la boîte (" + str(e.code) + "). "
                               "Reconnectez-vous avec votre compte (bouton ⚙). Détail : " + det)
        raise RuntimeError("Erreur Microsoft 365 (" + str(e.code) + ") : " + det)
    return json.loads(raw) if raw else {}


def ms_date(iso):
    try:
        dt = datetime.datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone()
    except Exception:
        return iso or ""
    j = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."][dt.weekday()]
    m = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."][dt.month - 1]
    return "%s %d %s %02d:%02d" % (j, dt.day, m, dt.hour, dt.minute)


def graph_messages(limit=60):
    """Lit les 7 derniers jours de la boîte via Microsoft (sans jamais demander de mot de passe)."""
    since = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
    q = urllib.parse.urlencode({"$top": limit, "$orderby": "receivedDateTime desc",
                                "$select": "id,from,receivedDateTime,subject,bodyPreview,conversationId,internetMessageId",
                                "$filter": "receivedDateTime ge " + since})
    out = []
    for i, m in enumerate(graph("/me/mailFolders/inbox/messages?" + q).get("value", [])):
        f = m.get("from", {}).get("emailAddress", {})
        out.append({"id": str(i + 1), "gid": m.get("id", ""), "conv": m.get("conversationId", ""),
                    "from": "%s <%s>" % (f.get("name") or f.get("address") or "", f.get("address") or ""),
                    "subject": m.get("subject") or "(sans objet)",
                    "date": ms_date(m.get("receivedDateTime", "")),
                    "body": (m.get("bodyPreview") or "").strip(),
                    "reply_to": f.get("address") or "", "message_id": m.get("internetMessageId") or "",
                    "category": "info", "summary": "", "full": False})
    return out


def graph_open(m):
    """Récupère le texte complet du mail et l'historique de la conversation (auteurs et dates séparés)."""
    if m.get("full"):
        return
    try:
        full = graph("/me/messages/" + m["gid"] + "?$select=body,from,receivedDateTime")
    except RuntimeError:
        m["full"] = True
        return
    m["body"] = strip_html(full.get("body", {}).get("content") or m.get("body", ""))[:6000]
    m["full"] = True
    conv = m.get("conv")
    if not conv:
        return
    q = urllib.parse.urlencode({"$top": 8, "$orderby": "receivedDateTime desc",
                                "$select": "id,body,from,receivedDateTime",
                                "$filter": "conversationId eq '" + conv.replace("'", "") + "'"})
    try:
        th = graph("/me/messages?" + q).get("value", [])
    except RuntimeError:
        return
    hist = []
    for p in th:
        if p.get("id") == m["gid"]:
            continue
        f = p.get("from", {}).get("emailAddress", {})
        who = f.get("name") or f.get("address") or "?"
        txt = strip_html(p.get("body", {}).get("content") or "")[:1200]
        hist.append("De : %s, le %s\n%s" % (who, ms_date(p.get("receivedDateTime", "")), txt))
    if hist:
        m["history"] = "\n\n".join(hist)


def graph_reply(m, text):
    """Prépare la réponse dans Microsoft (destinataire et objet déjà bons), puis l'envoie telle que modifiée."""
    d = graph("/me/messages/" + m["gid"] + "/createReply", "POST", {})
    did = d.get("id")
    if not did:
        raise RuntimeError("Microsoft n'a pas préparé la réponse.")
    graph("/me/messages/" + did, "PATCH", {"body": {"contentType": "Text", "content": text}})
    graph("/me/messages/" + did + "/send", "POST", {})


def graph_mail(p, subj):
    graph("/me/sendMail", "POST", {"message": {"subject": subj,
                                              "body": {"contentType": "Text", "content": p["body"]},
                                              "toRecipients": [{"emailAddress": {"address": p["to"]}}]},
                                   "saveToSentItems": True})


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
    if is_ms():
        try:
            graph_open(m)
        except RuntimeError:
            pass
    me = CFG.get("ms_email") or CFG.get("email", "")
    sig = CFG.get("signature", "").strip()
    sender_name, sender_addr = parseaddr(m["from"])
    if m.get("history"):
        new, history = m["body"], m["history"]  # historique propre fourni par Microsoft (messages séparés)
    else:
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
    subj = p["subject"]
    subj = subj if subj.lower().startswith("re:") else "Re: " + subj
    if is_ms():
        m = MAILS.get(str(p.get("id")))
        if m and m.get("gid"):
            graph_reply(m, p["body"])
            return
        graph_mail(p, subj)
        return
    msg = EmailMessage()
    msg["From"], msg["To"] = CFG.get("email") or CFG["user"], p["to"]
    msg["Subject"] = subj
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
.msbox{margin:18px 0;padding:18px;border-radius:16px;background:linear-gradient(160deg,#eef4ff,#fdf6e7);border:1px solid #c7d8f8}
.code{font-size:26px;font-weight:800;letter-spacing:.16em;color:var(--ac2)}
@media(max-width:850px){.layout,.split,.grid2{grid-template-columns:1fr}}
</style>
<header><div class="brand"><div class="orb"></div>SchoolBot Mail <small>développé par Educlan Asbl</small></div>
<span id="who"></span><button class="g" onclick="gear()">⚙ Réglages</button><button class="g" id="btnout" style="display:none" onclick="logout()">Déconnexion</button><button onclick="load()">Traiter mes mails</button></header>
<div id="upd" style="display:none;margin:10px auto 0;max-width:1100px;padding:12px 16px;border-radius:14px;background:#fff7e6;border:1px solid #f3c56b;font-size:14px"></div>
<main id="app"></main>
<script>
const $=s=>document.querySelector(s);const app=$('#app');let mails=[],cur=null,filter=null,cfg={},msFound=0;
fetch('/api/update').then(r=>r.json()).then(u=>{if(!u.available)return;const b=$('#upd');b.style.display='block';
 b.innerHTML='Une nouvelle version de SchoolBot Mail est disponible ('+u.latest+', vous avez la '+u.current+'). '+(u.notes?'<br><small>'+u.notes.replace(/[&<>]/g,'')+'</small><br>':'')+' <a href="'+u.url+'" target="_blank"><b>Télécharger la mise à jour</b></a> · <a href="#" onclick="this.parentNode.remove();return false">Plus tard</a>'}).catch(()=>{});
const CATS={urgent:['Urgent','var(--red)','urgents'],repondre:['À répondre','var(--org)','réponses à rédiger'],transmettre:['À transmettre','var(--blu)','à transmettre'],administratif:['Administratif','var(--yel)','documents à ranger'],info:['À lire','var(--grn)','à lire'],pub:['Pubs & notifications','var(--gry)','pubs et notifications']};
const esc=s=>String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const name=f=>(f||'').replace(/<.*>/,'').replace(/"/g,'').trim()||f;
const F=[["email","Adresse e-mail"],["user","Identifiant de connexion"],["pass","Mot de passe","password"],["imap_host","Serveur entrant (IMAP)"],["imap_port","Port IMAP"],["imap_sec","Sécurité IMAP","sec"],["smtp_host","Serveur sortant (SMTP)"],["smtp_port","Port SMTP"],["smtp_sec","Sécurité SMTP","sec"],["ai","Intelligence artificielle","ai"],["ollama_model","Modèle local (Ollama)","model"],["api_key","Clé API Anthropic (seulement si IA Claude)","password"],["signature","Signature des réponses"]];
const CATALOG=[['mistral-small','meilleur français, 16 Go de mémoire conseillés (~14 Go)'],['mistral-nemo','bon français, 8-16 Go (~7 Go)'],['gemma2:9b','correct, 8 Go (~5 Go)'],['llama3.2','léger, PC modeste (~2 Go)'],['mistral','très léger, français moyen (~4 Go)']];
async function fillModels(sel,cur){sel.add(new Option('Automatique — le meilleur modèle installé','auto'));
 try{const r=await(await fetch('/api/models')).json();
  for(const m of r.models)sel.add(new Option(m+(m===r.recommended?' — conseillé':''),m));
  const g=document.createElement('optgroup');g.label='Télécharger un autre modèle…';
  for(const[m,d]of CATALOG)if(!r.models.some(x=>x===m||x.startsWith(m+':')))g.append(new Option('⬇ '+m+' — '+d,'dl:'+m));
  if(g.children.length)sel.add(g)}
 catch(e){for(const m of['mistral-small','mistral-nemo','mistral','llama3.2'])sel.add(new Option(m,m))}
 sel.value=(cur&&[...sel.options].some(o=>o.value===cur))?cur:'auto'}
function wizard(c){c=c||{};msFound=0;app.innerHTML='<div class="panel" style="max-width:720px;margin:auto"><div class="hero"><div class="orb big"></div><div><h2>Connectons votre boîte mail</h2><p>Paramètres fournis par votre service informatique. Ils restent sur cet ordinateur.</p></div></div><div id="mszone" style="display:none"></div><div class="grid2" id="fields"></div><div class="row" style="margin-top:20px"><button class="g" onclick="testIt()">Tester la connexion</button><button onclick="saveIt()">Enregistrer et commencer</button></div><p id="wst"></p></div>';
 const f=$('#fields');
 for(const[k,t,ty]of F){const w=document.createElement('div');const l=document.createElement('label');l.textContent=t;let i;
   if(ty==='sec'){i=document.createElement('select');for(const[v,n]of[['ssl','SSL/TLS'],['starttls','STARTTLS'],['none','Aucune']])i.add(new Option(n,v))}
  else if(ty==='ai'){i=document.createElement('select');for(const[v,n]of[['ollama','Locale — les mails restent sur ce PC'],['claude','Claude (en ligne, clé API)']])i.add(new Option(n,v))}
  else if(ty==='model'){i=document.createElement('select');fillModels(i,c[k])}
  else{i=document.createElement('input');i.type=ty||'text'}
  i.id='f_'+k;if(ty!=='model'&&c[k])i.value=c[k];w.append(l,i);f.append(w)}
 if(!c.smtp_sec)$('#f_smtp_sec').value='starttls';$('#f_email').onblur=guess;
 if(c.ms)msLogged(c)}
function msLogged(c){msFound=1;const z=$('#mszone');z.style.display='block';
 z.innerHTML='<div class="msbox">Compte Microsoft connecté : <b>'+esc(c.ms_email||c.email||'')+'</b>'+
 '<div class="row" style="margin-top:12px"><button class="g" onclick="msOut()">Se déconnecter</button></div></div>'}
async function guess(){const e=$('#f_email').value;if(!e.includes('@'))return;$('#wst').textContent='Recherche des paramètres…';
 const r=await(await fetch('/api/detect',{method:'POST',body:JSON.stringify({email:e})})).json();if(r.error){$('#wst').textContent=r.error;return}
 for(const k of['user','imap_host','imap_port','imap_sec','smtp_host','smtp_port','smtp_sec'])if(r[k])$('#f_'+k).value=r[k];
  if((r.source||'').indexOf('Microsoft')>=0){msAsk();return}
  if((r.source||'').indexOf('Google')>=0){gmailHelp();return}
  $('#wst').textContent=r.source==='supposition'?'Paramètres supposés : vérifiez-les avec « Tester la connexion ».':'Paramètres trouvés ('+r.source+'). Entrez le mot de passe puis testez.'}
function gmailHelp(){const z=$('#mszone');z.style.display='block';
 z.innerHTML='<div class="msbox"><b>Votre boîte est chez Google (Gmail).</b><p style="margin:6px 0 10px;color:var(--mut)">Google refuse votre mot de passe habituel pour un programme comme SchoolBot Mail. Il faut un <b>mot de passe d\u2019application</b>, une seule fois :</p>'+
 '<ol style="margin:0 0 12px 18px;color:var(--mut);font-size:14px;line-height:1.6"><li>Activez la validation en deux étapes sur <a href="https://myaccount.google.com/security" target="_blank">myaccount.google.com/security</a></li>'+
 '<li>Ouvrez <a href="https://myaccount.google.com/apppasswords" target="_blank">myaccount.google.com/apppasswords</a> et créez un mot de passe nommé « SchoolBot Mail »</li>'+
 '<li>Collez le code de 16 lettres dans le champ <b>Mot de passe</b> ci-dessous, puis testez la connexion</li></ol>'+
 '<p style="margin:0;color:var(--mut);font-size:13px">Votre mot de passe habituel ne change pas : il sert toujours à consulter vos mails. Si l\u2019option est absente, l\u2019administrateur Google de votre école doit l\u2019autoriser.</p></div>';
 $('#wst').textContent='Paramètres Gmail remplis. Créez le mot de passe d\u2019application (voir ci-dessus) puis testez.'}
function msAsk(){msFound=1;const z=$('#mszone');z.style.display='block';
 z.innerHTML='<div class="msbox"><b>Votre boîte est chez Microsoft 365.</b><p style="margin:6px 0 14px;color:var(--mut)">Le plus simple : connectez-vous avec votre compte habituel. SchoolBot Mail ne verra jamais votre mot de passe.</p>'+
 '<div class="row"><button onclick="msGo()">Se connecter avec Microsoft</button><button class="g" onclick="msHide()">Ou saisir un mot de passe</button></div><p id="msmsg"></p></div>'}
function msHide(){const z=$('#mszone');z.style.display='none';msFound=0;$('#wst').textContent='Entrez le mot de passe puis testez la connexion.'}
async function msGo(){const m=$('#msmsg');m.textContent='Préparation de la connexion…';
 const r=await(await fetch('/api/ms/start',{method:'POST',body:'{}'})).json();
 if(r.error){m.textContent=r.error;return}msPoll()}
async function msPoll(){const r=await(await fetch('/api/ms/status')).json();const m=$('#msmsg');if(!m)return;
 if(r.state==='waiting'){m.innerHTML='Ouvrez <a href="'+r.url+'" target="_blank">la page de Microsoft</a> et tapez ce code :<div class="code">'+esc(r.user_code)+'</div>';
  setTimeout(msPoll,1500)}
 else if(r.state==='ok'){m.innerHTML='✅ Connecté. Appuyez sur « Enregistrer et commencer ».'}
 else if(r.state==='error'){m.textContent=r.msg}}
async function msOut(){await fetch('/api/ms/out',{method:'POST',body:'{}'});msFound=0;$('#mszone').style.display='none';$('#wst').textContent='Connexion Microsoft coupée.'}
const vals=()=>Object.fromEntries(F.map(([k])=>[k,$('#f_'+k).value]));
async function testIt(){$('#wst').textContent='Test en cours…';const r=await(await fetch('/api/test',{method:'POST',body:JSON.stringify(vals())})).json();$('#wst').textContent='Réception : '+r.imap+' · Envoi : '+r.smtp}
async function saveIt(){const v=vals();let dl=null;if((v.ollama_model||'').startsWith('dl:')){dl=v.ollama_model.slice(3);v.ollama_model=dl}
 await fetch('/api/config',{method:'POST',body:JSON.stringify(v)});
 if(dl)await fetch('/api/setup',{method:'POST',body:JSON.stringify({model:dl})});load()}
async function gear(){wizard((await(await fetch('/api/config')).json()).cfg)}
async function logout(){if(!confirm('Se déconnecter de ce compte mail ? Vos réglages IA et votre signature seront conservés.'))return;
 await fetch('/api/logout',{method:'POST',body:'{}'});cfg={};$('#who').textContent='';$('#btnout').style.display='none';wizard({})}
async function setupScreen(){let s=await(await fetch('/api/setup')).json();if(s.stage==='ready')return true;
 app.innerHTML='<div class="panel"><div class="hero"><div class="orb big pulse"></div><div><h2>Préparation de votre assistant</h2><p id="sm">'+(s.want?'SchoolBot Mail télécharge le modèle <b>'+esc(s.want)+'</b>. Vos mails ne quittent pas cet ordinateur. Selon la connexion, comptez quelques minutes.':'Une seule fois : SchoolBot Mail installe son intelligence artificielle sur cet ordinateur. Vos mails ne le quitteront jamais. Comptez 10 à 20 minutes.')+'</p></div></div><div class="bar"><i id="bar"></i></div><div class="count" id="cnt"></div></div>';
 s=await(await fetch('/api/setup',{method:'POST',body:'{}'})).json();
 while(s.stage!=='ready'){$('#bar').style.width=(s.pct||1)+'%';$('#cnt').textContent=s.msg||'';
  if(s.stage==='error'){$('#cnt').innerHTML=esc(s.msg)+' <button onclick="load()">Réessayer</button>';return false}
  await new Promise(r=>setTimeout(r,1000));s=await(await fetch('/api/setup')).json()}
 return true}
async function load(){if(!await setupScreen())return;const c=await(await fetch('/api/config')).json();cfg=c.cfg||{};$('#who').textContent=cfg.email||'';$('#btnout').style.display=c.configured?'':'none';if(!c.configured||!c.has_key){wizard(c.cfg);return}
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
 app.innerHTML='<div class="row" style="margin-bottom:14px"><button class="g" onclick="list(filter)">← Retour</button></div><div class="split"><div class="panel"><div class="label">MAIL REÇU</div><h2 style="margin:8px 0 2px;font-size:20px">'+esc(cur.subject)+'</h2><small style="color:var(--mut)">'+esc(cur.from)+' · '+esc(cur.date)+'</small><pre id="mb">'+esc(cur.body)+'</pre></div>'+
 '<div class="panel reply"><div class="label"><span class="dot" style="color:var(--yel);background:var(--yel)"></span>RÉPONSE PRÉPARÉE PAR L’IA</div><h2 style="margin:8px 0 2px;font-size:18px">Re: '+esc(cur.subject.replace(/^re: */i,''))+'</h2><small style="color:var(--mut)">À : '+esc(cur.from)+'</small>'+
 '<textarea id="reply" readonly>Rédaction en cours…</textarea><div class="note">✦ Rédigée à partir de l’historique — vous validez, il envoie.</div><div class="row"><button class="g" onclick="edit()">✎ Modifier</button><button class="g" onclick="redo()">↻ Regénérer</button><button class="grow" onclick="sendIt()">✉ Envoyer</button></div><p id="st"></p></div></div>';
 try{const j=await(await fetch('/api/open',{method:'POST',body:JSON.stringify({id})})).json();
  if(j.body&&$('#mb'))$('#mb').textContent=j.body}catch(e){}
 const j=await getDraft(cur);$('#reply').value=cur.draft||('Erreur : '+j.error)}
function edit(){const t=$('#reply');t.readOnly=false;t.focus()}
async function redo(){cur.p=null;cur.draft=null;$('#reply').value='Rédaction en cours…';const j=await getDraft(cur);$('#reply').value=cur.draft||('Erreur : '+j.error)}
async function sendIt(){if(!confirm('Envoyer cette réponse à '+cur.reply_to+' ?'))return;$('#st').textContent='Envoi…';
 const r=await fetch('/api/send',{method:'POST',body:JSON.stringify({id:cur.id,to:cur.reply_to,subject:cur.subject,body:$('#reply').value,message_id:cur.message_id})});
 const j=await r.json();$('#st').textContent=j.ok?'✅ Envoyé':'Erreur : '+j.error}
load();
</script></html>"""


# ---------- Installation automatique de l'IA locale (aucune commande à taper) ----------
import sys, threading, subprocess, platform, tempfile, time, shutil, zipfile
SETUP = {"stage": "idle", "pct": 0, "msg": ""}
IS_WIN, IS_MAC = os.name == "nt", sys.platform == "darwin"


def ram_gb():
    try:
        if IS_WIN:
            import ctypes
            class M(ctypes.Structure):
                _fields_ = [("l", ctypes.c_ulong), ("m", ctypes.c_ulong), ("t", ctypes.c_ulonglong)] + [(f"x{i}", ctypes.c_ulonglong) for i in range(6)]
            m = M(); m.l = ctypes.sizeof(M); ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
            return m.t / 2**30
        if IS_MAC:
            return int(subprocess.check_output(["sysctl", "-n", "hw.memsize"])) / 2**30
        return os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES") / 2**30
    except Exception:
        return 8


def model_for_pc():
    g = ram_gb()
    # mistral-small = qualité de français exigée ; dès 16 Go (Windows en rapporte ~15,8)
    return "mistral-small" if g >= 15 else "mistral-nemo"


def ollama_up():
    try:
        urllib.request.urlopen("http://127.0.0.1:11434/api/tags", timeout=3); return True
    except Exception:
        return False


def ollama_exe():
    cands = [shutil.which("ollama") or ""]
    if IS_WIN:
        cands.append(os.path.join(E("LOCALAPPDATA", ""), "Programs", "Ollama", "ollama.exe"))
    if IS_MAC:
        cands += [os.path.expanduser("~/Applications/Ollama.app/Contents/Resources/ollama"), "/Applications/Ollama.app/Contents/Resources/ollama"]
    return next((c for c in cands if c and os.path.exists(c)), None)


def download(url, dest, lo, hi, msg):
    with urllib.request.urlopen(url, timeout=30) as r, open(dest, "wb") as f:
        tot, n = int(r.headers.get("Content-Length") or 0), 0
        while True:
            b = r.read(1 << 20)
            if not b: break
            f.write(b); n += len(b)
            if tot: SETUP.update(pct=lo + (hi - lo) * n / tot, msg=f"{msg} ({n >> 20} / {tot >> 20} Mo)")


def start_ollama(exe):
    # Lancé détaché : pas de fenêtre noire, et Ollama ne garde aucun fichier du programme ouvert.
    flags = (0x08000000 | 0x00000008 | 0x00000200) if IS_WIN else 0
    env = {k: v for k, v in os.environ.items() if not k.startswith("_MEI") and k not in ("_PYI_APPLICATION_HOME_DIR", "_PYI_ARCHIVE_FILE", "_PYI_PARENT_PROCESS_LEVEL")}
    subprocess.Popen([exe, "serve"], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                     creationflags=flags, cwd=os.path.expanduser("~"), env=env, close_fds=True,
                     start_new_session=not IS_WIN)
    for _ in range(40):
        if ollama_up(): return True
        time.sleep(1)
    return False


def pull_model(m):
    SETUP.update(stage="model", pct=30, msg=f"Téléchargement de l'assistant ({m})…")
    req = urllib.request.Request("http://127.0.0.1:11434/api/pull", data=json.dumps({"name": m}).encode())
    with urllib.request.urlopen(req, timeout=600) as r:
        for line in r:
            d = json.loads(line or "{}")
            if d.get("error"): raise RuntimeError(d["error"])
            if d.get("total"):
                SETUP.update(pct=30 + 69 * d.get("completed", 0) / d["total"],
                             msg=f"Téléchargement de {m} : {d.get('completed',0) >> 20} / {d['total'] >> 20} Mo")
    _MODEL.clear()


def run_setup(want=None):
    try:
        SETUP.update(stage="engine", pct=2, msg="Vérification du moteur d'IA…")
        if not ollama_up():
            exe = ollama_exe()
            if not exe:
                tmp = tempfile.mkdtemp()
                if IS_WIN:
                    f = os.path.join(tmp, "OllamaSetup.exe")
                    download("https://ollama.com/download/OllamaSetup.exe", f, 2, 25, "Téléchargement du moteur d'IA")
                    SETUP.update(pct=26, msg="Installation du moteur d'IA…")
                    subprocess.run([f, "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES"], creationflags=0x08000000)
                elif IS_MAC:
                    f = os.path.join(tmp, "Ollama.zip")
                    download("https://ollama.com/download/Ollama-darwin.zip", f, 2, 25, "Téléchargement du moteur d'IA")
                    apps = os.path.expanduser("~/Applications"); os.makedirs(apps, exist_ok=True)
                    subprocess.run(["ditto", "-x", "-k", f, apps])
                else:
                    raise RuntimeError("Installez Ollama depuis ollama.com")
                exe = ollama_exe()
            if not exe or not (ollama_up() or start_ollama(exe)):
                raise RuntimeError("Le moteur d'IA n'a pas pu démarrer. Redémarrez l'ordinateur puis relancez SchoolBot Mail.")
        if want and want not in ollama_models():
            pull_model(want)
        elif not ollama_models():
            pull_model(model_for_pc())
        SETUP.update(stage="ready", pct=100, msg="Prêt !")
    except Exception as e:
        SETUP.update(stage="error", msg=str(e))


def setup_status():
    if SETUP["stage"] == "idle" and (CFG.get("ai") or "ollama") == "ollama" and ollama_up() and ollama_models():
        SETUP.update(stage="ready", pct=100)
    if (CFG.get("ai") or "ollama") != "ollama":
        SETUP.update(stage="ready", pct=100)
    return SETUP


def desktop_shortcut():
    """Crée l'icône sur le bureau au premier lancement de l'application (.exe / .app)."""
    if not getattr(sys, "frozen", False): return
    try:
        home = os.path.expanduser("~")
        if IS_WIN:
            desk = subprocess.check_output(["powershell", "-NoProfile", "-Command", "[Environment]::GetFolderPath('Desktop')"],
                                           creationflags=0x08000000, text=True).strip() or os.path.join(home, "Desktop")
            lnk = os.path.join(desk, "SchoolBot Mail.lnk")
            if os.path.exists(lnk): return
            exe = sys.executable
            ps = (f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}');$s.TargetPath='{exe}';"
                  f"$s.WorkingDirectory='{os.path.dirname(exe)}';$s.IconLocation='{exe},0';$s.Save()")
            subprocess.run(["powershell", "-NoProfile", "-Command", ps], creationflags=0x08000000)
        elif IS_MAC:
            app = os.path.abspath(os.path.join(sys.executable, "..", "..", ".."))
            link = os.path.join(home, "Desktop", "SchoolBot Mail")
            if app.endswith(".app") and not os.path.lexists(link): os.symlink(app, link)
    except Exception:
        pass


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
            safe = {k: v for k, v in CFG.items()
                    if k not in ("pass", "api_key") and not k.startswith("ms_access") and not k.startswith("ms_refresh")}
            safe["ms"] = is_ms()
            return self.reply({"configured": is_ms() or bool(CFG.get("imap_host") and CFG.get("pass")), "cfg": safe,
                               "has_key": (CFG.get("ai") or "ollama") == "ollama" or bool(CFG.get("api_key") or E("ANTHROPIC_API_KEY"))})
        if self.path == "/api/setup":
            return self.reply(setup_status())
        if self.path == "/api/models":
            return self.reply({"models": ollama_models(), "recommended": pick_model()})
        if self.path == "/api/update":
            return self.reply(check_update())
        if self.path == "/api/ms/status":
            return self.reply(ms_public())
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
            if self.path == "/api/setup":
                want = (p.get("model") or "").strip() or None
                if SETUP["stage"] in ("idle", "error") or (want and SETUP["stage"] == "ready"):
                    SETUP.update(stage="engine", pct=1, msg="Démarrage…", want=want or "")
                    threading.Thread(target=run_setup, args=(want,), daemon=True).start()
                return self.reply(SETUP)
            if self.path == "/api/detect":
                return self.reply(autodetect(p.get("email", "")))
            if self.path == "/api/ms/start":
                return self.reply(ms_start())
            if self.path == "/api/ms/out":
                save_cfg({"ms_access": "", "ms_refresh": "", "ms_expiry": 0, "provider": ""})
                MS.update(state="idle", user_code="", msg="")
                return self.reply({"ok": True})
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
            if self.path == "/api/logout":
                for k in ("email", "user", "pass", "imap_host", "imap_port", "imap_sec",
                          "smtp_host", "smtp_port", "smtp_sec", "provider",
                          "ms_access", "ms_refresh", "ms_expiry", "ms_email"):
                    CFG.pop(k, None)
                save_cfg({})
                MS.update(state="idle", user_code="", msg="")
                MAILS.clear()
                return self.reply({"ok": True})
            if self.path == "/api/open":
                m = MAILS.get(str(p.get("id")))
                if not m:
                    return self.reply({"error": "mail introuvable"}, 404)
                if is_ms():
                    graph_open(m)
                return self.reply({"body": m.get("body", "")})
            if self.path == "/api/draft":
                return self.reply({"draft": draft(MAILS[p["id"]])})
            if self.path == "/api/send":
                send(p)
                return self.reply({"ok": True})
        except Exception as e:
            return self.reply({"error": str(e)}, 500)
        self.reply({"error": "not found"}, 404)


if __name__ == "__main__":
    URL = "http://127.0.0.1:8765"
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", 8765), H)
    except OSError:  # déjà lancé : on rouvre simplement la fenêtre
        import webbrowser; webbrowser.open(URL); sys.exit(0)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    desktop_shortcut()
    print("SchoolBot Mail → " + URL)
    try:
        import webview  # fenêtre propre à l'application (version .exe / .app)
        webview.create_window("SchoolBot Mail", URL, width=1280, height=860, min_size=(900, 600))
        webview.start()
    except Exception:
        import webbrowser; webbrowser.open(URL)
        try:
            while True: time.sleep(3600)
        except KeyboardInterrupt:
            pass
