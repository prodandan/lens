"""
Contact Scraper — extrage email + telefon din pagina /contact, header, footer
"""

# ─────────────────────────────────────────────────────────────────────────────
SITES = [
    "spy-shop.ro", "aden.ro", "teleshopromania.ro", "telemarket.ro",
    "mywebshop.ro", "elefunstore.ro", "mizar24.ro", "mdorice.ro",
    "autojeep.ro", "carsound.ro", "mudster.ro", "optimusdigital.ro",
    "leoauto.ro", "frilla.ro", "vigoshop.ro", "vigoexpress.ro",
    "pansiprod.ro", "raijucarii.ro", "houseconcept.ro", "bricomall.ro",
    "ledbazar.ro", "utb-shop.ro", "elareducere.ro", "pieseauto.ro",
    "flexzon.ro", "e-jumbo.ro", "ledurimasina.ro", "crazy4led.ro",
    "ssvauto.ro", "utilajeagricolecraiova.ro", "bazar-shop.ro",
    "techstar.ro", "deluxetuning.ro", "baduglobal.ro", "becleduri.ro",
    "km100.ro", "xenon-drive.ro", "miromoto.ro", "esell.ro",
    "tehnoelectric.ro", "ledbar.ro", "sculebgs.ro", "generalmotor.ro",
    "vectro.ro", "bricolaj.ro", "becuri-led-auto.ro", "kardaf.ro",
    "autodoc24.ro", "e-glasspandoor.ro", "gave.ro", "piese-auto.ro",
    "magazinulcuscule.ro", "home.store.ro", "autosema.ro", "sersimo.ro",
    "preturirezonabile.ro", "accesoriiautotuning.ro", "sealauto.ro",
    "polimerizarefaruri.ro", "artool.ro", "epiesa.ro", "iaureduceri.ro",
    "fivo.ro", "egospodarul.ro", "diversmag.ro", "tor-online.ro",
    "ecomplex.ro", "sculesiechipamente.ro", "daciashop.ro", "moka-gsm.ro",
    "iprofesional.ro", "pni.ro", "lndauto.ro", "vagauto.ro", "moemax.ro",
    "autobliz.ro", "autopieseonline24.ro", "shopu.ro", "targuldepiese.ro",
    "vagstore.ro", "romanianmag.ro", "promotiinonstop.ro",
    "anunturi-auto.4tuning.ro", "tuningexpert.ro", "autoparadoxsystem.ro",
    "crazyshop.ro", "eldepo.ro", "pretzmic.ro", "azoro.ro",
    "angelsauto.ro", "a2t.ro", "mezoni.ro", "nichiduta.ro", "nextly.ro",
    "edshop.ro", "trolii-auto.ro", "industrial.utilajul.ro", "gmoto.ro",
    "officedirect.ro", "electroniclight.ro", "comenzi.ro", "shop.roben.ro",
    "ghirlandegradina.ro", "bec-expert.ro", "accesorii-autoutilitare.ro",
    "akheronserv.ro", "linhai-atv.ro", "esemromania.ro", "tuningcox.ro",
    "euautopiese.ro", "mobexpert.ro", "ssaauto.ro", "mam-bricolaj.ro",
    "ozone.ro", "xband.ro", "blusmart.ro", "monirom.ro",
    "lampisilumini.ro", "asmarket.ro", "axauto.ro", "unshop.ro",
    "bilioner.ro", "strenger.ro", "auto-led-shop.ro", "dacianmag.ro",
    "preturireduse.ro", "autohelix.ro", "homepc.ro", "rsbtuning.ro",
    "produsredus.ro", "onlinediscount.ro", "poloshop.ro", "xxxlutz.ro",
    "pret-minim.ro", "ebrazi.ro", "decocraciun.ro", "elsales.ro",
    "barete-leduri.ro", "frize.ro", "vand-xenon.ro", "auxito.ro",
    "qmax.ro", "dmsauto4x4.ro", "rollfast.ro", "multimasimex.ro",
    "rapidauto.ro", "lidermarket.ro", "simpletools.ro", "electrospot.ro",
    "smallgear.ro", "homeplus.ro", "lidl.ro", "narpo.ro", "uluitor.ro",
    "magazelo.ro", "flymusic.ro", "produstop.ro", "shopexo.ro",
    "mrg-shop.ro", "sculeagro.ro", "agrafa.ro", "dinaelectronics.ro",
    "debicicleta.ro", "europart.ro", "offroadgold.ro", "mtools.ro",
    "folina.ro", "giftexpress.ro", "happymax.ro", "polytron.ro",
    "piese.ro", "hafele.ro", "kraftdele.ro", "bizoo.ro",
    "piese-camioane.ro", "evo-moto.ro", "totalfishing.ro", "vonmag.ro",
    "autocorect.ro", "emudding.ro", "eshopmania.ro", "ham-bebe.ro",
    "karlatoys.ro", "magazin-agro.ro", "today-mag.ro", "rabalux.ro",
    "led4you.ro", "kosmo.ro", "cmall.ro", "nwradu.ro", "vwforum.ro",
    "esderman.ro", "magflix.ro", "xenon4u.ro", "leditup.ro", "c-led.ro",
    "art4x4.ro", "atvrom.ro", "maxmania.ro", "marosbike.ro", "avex.ro",
    "vixo.ro", "depozituldeutilaje.ro",
]

OUTPUT_EXCEL = r"C:\Users\Dan\Desktop\contacte_site-uri.xlsx"
CHECKPOINT   = r"C:\Users\Dan\Desktop\contacte_progress.json"

SCRAPE_DELAY    = 1.0
REQUEST_TIMEOUT = 15

# sluguri de contact încercate în ordine — primele sunt cele mai probabile
CONTACT_SLUGS = [
    "contact", "contacts", "contacte", "contact-us", "contactati-ne",
    "despre-noi", "despre", "informatii-contact", "info",
]

# emailuri de la aceste domenii sunt platforme/servicii, nu contacte reale
PLATFORM_DOMAINS = re.compile(
    r'@.*(sentry|wix|wordpress|woocommerce|mailchimp|sendgrid|hubspot|'
    r'google|facebook|microsoft|apple|amazon|cloudflare|gravatar|'
    r'prestashop|opencart|magento|shopify|squarespace|example|'
    r'yourdomain|domain\.com|test\.com|localhost)',
    re.I
)

# prefixe de email care indică adrese sistem, nu contact
NOREPLY_RE = re.compile(
    r'^(noreply|no-reply|donotreply|do-not-reply|bounce|'
    r'mailer-daemon|postmaster|webmaster|admin@(?!.*\.ro))',
    re.I
)
# ─────────────────────────────────────────────────────────────────────────────

import json
import re
import signal
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

PHONE_REGEX = re.compile(
    r'(?<!\d)(?:'
    r'\+40[\s\.\-]?[23789]\d{2}[\s\.\-]?\d{3}[\s\.\-]?\d{3}'
    r'|0[23789]\d{2}[\s\.\-]?\d{3}[\s\.\-]?\d{3}'
    r')(?!\d)'
)
EMAIL_REGEX = re.compile(r'\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b')
VALID_PHONE = re.compile(r'^\+40[237]\d{8}$')

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "ro-RO,ro;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

_shutdown = False


def _sigint(sig, frame):
    global _shutdown
    log("Ctrl+C — oprire curată după site-ul curent …")
    _shutdown = True


signal.signal(signal.SIGINT, _sigint)


# ── utilități ────────────────────────────────────────────────────────────────

def log(msg, idx=None, total=None):
    ts = datetime.now().strftime("%H:%M:%S")
    prefix = f"[{idx}/{total}] " if idx is not None else ""
    print(f"[{ts}] {prefix}{msg}", flush=True)


def load_cp():
    p = Path(CHECKPOINT)
    if p.exists():
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {"done": [], "results": {}}


def save_cp(state):
    with open(CHECKPOINT, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ── normalizare ───────────────────────────────────────────────────────────────

def norm_phone(raw: str):
    digits = re.sub(r"[^\d+]", "", raw)
    if digits.startswith("+40"):
        num = "+40" + digits[3:]
    elif digits.startswith("40") and len(digits) == 11:
        num = "+" + digits
    elif digits.startswith("0") and len(digits) == 10:
        num = "+40" + digits[1:]
    elif len(digits) == 9 and digits[0] in "237":
        num = "+40" + digits
    else:
        return None
    num = num.replace(" ", "")
    return num if VALID_PHONE.match(num) else None


def ok_email(addr: str, domain: str = "") -> bool:
    """Filtrează emailurile false sau de platformă."""
    addr = addr.lower().strip()
    if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]{2,}$', addr):
        return False
    if PLATFORM_DOMAINS.search(addr):
        return False
    if NOREPLY_RE.match(addr.split("@")[0]):
        return False
    if re.search(r'\.(png|jpg|gif|svg|webp|ico|css|js)$', addr):
        return False
    return True


def rank_email(addr: str, domain: str) -> int:
    """Scor de prioritate: emailuri de pe domeniul propriu > restul."""
    root = domain.replace("www.", "").split(".")[0]
    if f"@{domain}" in addr or f"@{root}" in addr:
        return 0   # cel mai bun
    return 1


# ── fetch ─────────────────────────────────────────────────────────────────────

def fetch(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HTTP_HEADERS, timeout=REQUEST_TIMEOUT,
                         allow_redirects=True)
        if r.status_code < 400:
            return r.text
    except Exception:
        pass
    return None


# ── extragere din HTML ────────────────────────────────────────────────────────

def extract(html: str, domain: str) -> tuple[list[str], list[str]]:
    """
    Extrage emailuri și telefoane dintr-un HTML.
    Ordinea de prioritate:
      1. mailto: și tel: link-uri (cel mai sigur)
      2. JSON-LD schema.org
      3. Zonele vizibile: header, footer, bandă top, secțiuni contact
      (NU face regex pe tot HTML-ul brut — evită false positive din cod/template)
    """
    emails: set[str] = set()
    phones: set[str] = set()

    soup = BeautifulSoup(html, "lxml")

    # elimină taguri invizibile care conțin text de template
    for tag in soup.find_all(["script", "style", "noscript", "template", "meta"]):
        tag.decompose()

    # 1. mailto: și tel: — sursa cea mai de încredere
    for a in soup.find_all("a", href=True):
        h = a["href"].strip()
        if h.lower().startswith("mailto:"):
            addr = h[7:].split("?")[0].strip().lower()
            if ok_email(addr, domain):
                emails.add(addr)
        elif h.lower().startswith("tel:"):
            n = norm_phone(h[4:].strip())
            if n:
                phones.add(n)

    # 2. JSON-LD schema.org
    for sc in soup.find_all("script", type="application/ld+json"):
        try:
            _jsonld(json.loads(sc.string or "{}"), emails, phones, domain)
        except Exception:
            pass

    # 3. Zone vizibile prioritare — text curat, fără cod
    priority_zones = []

    # <header>, <footer>, <nav>
    for tag in ["header", "footer", "nav"]:
        priority_zones.extend(soup.find_all(tag))

    # elemente cu clase/id care sugerează contact sau bandă de info
    contact_sel = re.compile(
        r'contact|footer|header|topbar|top[-_]bar|top[-_]info|'
        r'info[-_]bar|address|phone|email|tel[-_]|fax|'
        r'social|widget[-_]contact|block[-_]contact',
        re.I
    )
    for el in soup.find_all(class_=contact_sel):
        priority_zones.append(el)
    for el in soup.find_all(id=contact_sel):
        priority_zones.append(el)

    # extrage text curat din zone prioritare și aplică regex
    seen_texts: set[str] = set()
    for zone in priority_zones:
        text = zone.get_text(" ", strip=True)
        if text in seen_texts or len(text) < 5:
            continue
        seen_texts.add(text)
        _regex_text(text, emails, phones, domain)

    # 4. Fallback: întreg body-ul ca text curat (nu HTML brut)
    # Folosim text vizibil, nu codul sursă — evită emailuri din JS/atribute
    body = soup.find("body")
    if body:
        body_text = body.get_text(" ", strip=True)
        _regex_text(body_text, emails, phones, domain)

    # sortare: emailuri proprii primele
    sorted_emails = sorted(emails, key=lambda e: rank_email(e, domain))
    return sorted_emails, sorted(phones)


def _regex_text(text: str, emails: set, phones: set, domain: str):
    for m in EMAIL_REGEX.findall(text):
        if ok_email(m, domain):
            emails.add(m.lower())
    for m in PHONE_REGEX.findall(text):
        n = norm_phone(m)
        if n:
            phones.add(n)


def _jsonld(obj, emails, phones, domain):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in ("email", "contactEmail") and isinstance(v, str):
                if ok_email(v, domain):
                    emails.add(v.lower())
            elif k in ("telephone", "faxNumber") and isinstance(v, str):
                n = norm_phone(v)
                if n:
                    phones.add(n)
            else:
                _jsonld(v, emails, phones, domain)
    elif isinstance(obj, list):
        for i in obj:
            _jsonld(i, emails, phones, domain)


# ── scraping per domeniu ──────────────────────────────────────────────────────

def scrape_site(domain: str) -> tuple[list, list, str]:
    """
    Strategie:
    1. Încearcă /contact, /contacts, /contacte (și variantele lor) — PRIMUL
    2. Dacă nu găsește nimic acolo, cade pe homepage (header/footer)
    Returnează (emails, phones, sursa_găsită)
    """
    # determină baza (https cu fallback http)
    base = f"https://{domain}"
    test = fetch(base)
    if test is None:
        base = f"http://{domain}"

    emails: set[str] = set()
    phones: set[str] = set()
    source = ""

    # ── PASUL 1: pagini dedicate de contact ──────────────────────────────────
    for slug in CONTACT_SLUGS:
        if _shutdown:
            break
        url = f"{base}/{slug}"
        html = fetch(url)
        if not html or len(html) < 300:
            continue

        e, p = extract(html, domain)
        if e or p:
            emails.update(e)
            phones.update(p)
            source = url
            log(f"    ✓ contact găsit la /{slug}", None, None)
            break   # am găsit — nu mai căutăm alte sluguri

    # ── PASUL 2: homepage — header + footer ca fallback ───────────────────────
    if not emails and not phones:
        html = fetch(base) if test is None else test
        if html:
            e, p = extract(html, domain)
            if e or p:
                emails.update(e)
                phones.update(p)
                source = base
                log(f"    ✓ contact găsit în homepage", None, None)

    if not emails and not phones:
        source = "—"

    sorted_emails = sorted(emails, key=lambda x: rank_email(x, domain))
    return sorted_emails, sorted(phones), source


# ── Excel ─────────────────────────────────────────────────────────────────────

HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
OK_FILL    = PatternFill("solid", fgColor="E8F5E9")   # verde — contact găsit
NOK_FILL   = PatternFill("solid", fgColor="FFEBEE")   # roșu  — negăsit
ALT_OK     = PatternFill("solid", fgColor="C8E6C9")
ALT_NOK    = PatternFill("solid", fgColor="FFCDD2")


def _hdr(cell, text):
    cell.value = text
    cell.font  = HDR_FONT
    cell.fill  = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_excel(state: dict, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Contacte Site-uri"

    for ci, h in enumerate(["Nr", "Domeniu", "Email-uri", "Telefoane", "Sursa", "Status"], 1):
        _hdr(ws.cell(row=1, column=ci), h)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 25

    row = 2
    for domain in SITES:
        data = state["results"].get(domain)
        if data is None:
            continue

        has = bool(data["emails"] or data["phones"])
        fill = (OK_FILL if row % 2 == 0 else ALT_OK) if has else (NOK_FILL if row % 2 == 0 else ALT_NOK)
        status = "✓ găsit" if has else "— negăsit"

        for ci, v in enumerate([
            row - 1,
            domain,
            ", ".join(data["emails"]),
            ", ".join(data["phones"]),
            data.get("source", ""),
            status,
        ], 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 12

    wb.save(path)


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    total = len(SITES)
    state = load_cp()
    done  = set(state.get("done", []))

    if done:
        log(f"Checkpoint găsit — {len(done)}/{total} deja procesate, reluăm de unde am rămas.")

    print("=" * 55)
    print(f"  Site-uri total:       {total}")
    print(f"  Deja procesate:       {len(done)}")
    print(f"  De procesat acum:     {total - len(done)}")
    print("=" * 55)

    for idx, domain in enumerate(SITES, 1):
        if _shutdown:
            break
        if domain in done:
            log(f"Skip: {domain}", idx, total)
            continue

        log(f"→ {domain}", idx, total)
        time.sleep(SCRAPE_DELAY)

        emails, phones, source = scrape_site(domain)
        log(f"  email={emails}  tel={phones}", idx, total)

        state["results"][domain] = {
            "emails": emails,
            "phones": phones,
            "source": source,
        }
        state["done"].append(domain)
        save_cp(state)

        try:
            build_excel(state, OUTPUT_EXCEL)
        except Exception as e:
            log(f"  Avertisment Excel: {e}", idx, total)

    results = state["results"]
    found    = sum(1 for d in results.values() if d["emails"] or d["phones"])
    no_email = sum(1 for d in results.values() if not d["emails"])
    no_phone = sum(1 for d in results.values() if not d["phones"])

    print("\n" + "=" * 55)
    print("  SUMAR FINAL")
    print("=" * 55)
    print(f"  Site-uri procesate:        {len(results)}")
    print(f"  Cu cel puțin un contact:   {found}")
    print(f"  Fără email găsit:          {no_email}")
    print(f"  Fără telefon găsit:        {no_phone}")
    print(f"\n  Excel: {OUTPUT_EXCEL}")
    if _shutdown:
        print("  (Oprit cu Ctrl+C — progresul salvat)")


if __name__ == "__main__":
    main()
