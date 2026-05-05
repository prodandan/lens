"""
Contact Scraper — extrage email + telefon din header/footer/pagina de contact
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
# ─────────────────────────────────────────────────────────────────────────────

import json
import re
import signal
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── regex ────────────────────────────────────────────────────────────────────
EMAIL_REGEX = re.compile(r'\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b')
PHONE_REGEX = re.compile(
    r'(?<!\d)(?:'
    r'\+40[\s\-\.]?[23789]\d{8}'
    r'|0[23789]\d{8}'
    r'|[23789]\d{2}[\s\-\.]?\d{3}[\s\-\.]?\d{3}'
    r')(?!\d)'
)
FAKE_EMAIL = re.compile(
    r'sentry\.io|wixpress\.com|googleapis\.com|gstatic\.com'
    r'|schema\.org|example\.com|@\d+x\.'
    r'|\.(png|jpg|gif|svg|webp|ico|css|js)$',
    re.I
)
VALID_PHONE_RE = re.compile(r'^\+40[237]\d{8}$')

# pagini de contact comune de încercat
CONTACT_SLUGS = [
    "contact", "contacte", "contact-us", "despre-noi", "despre",
    "informatii-contact", "info", "ajutor", "help", "support",
]

HEADERS = {
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
    return num if VALID_PHONE_RE.match(num) else None


def ok_email(addr: str) -> bool:
    if FAKE_EMAIL.search(addr):
        return False
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]{2,}$', addr))


# ── fetch ─────────────────────────────────────────────────────────────────────

def fetch(url: str):
    try:
        r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT,
                         allow_redirects=True)
        if r.status_code < 400:
            return r.text
    except Exception:
        pass
    return None


# ── extragere contact dintr-un HTML ──────────────────────────────────────────

def extract_from_html(html: str, base_url: str):
    emails: set[str] = set()
    phones: set[str] = set()

    soup = BeautifulSoup(html, "lxml")

    # 1. mailto: și tel: — cea mai fiabilă sursă
    for a in soup.find_all("a", href=True):
        h = a["href"]
        if h.startswith("mailto:"):
            addr = h[7:].split("?")[0].strip().lower()
            if ok_email(addr):
                emails.add(addr)
        elif h.startswith("tel:"):
            n = norm_phone(h[4:].strip())
            if n:
                phones.add(n)

    # 2. JSON-LD
    for sc in soup.find_all("script", type="application/ld+json"):
        try:
            _jsonld(json.loads(sc.string or ""), emails, phones)
        except Exception:
            pass

    # 3. header + footer — zone prioritare pentru regex
    for zone_tag in ["header", "footer", "nav"]:
        for zone in soup.find_all(zone_tag):
            _regex_zone(zone.get_text(" ", strip=True), emails, phones)

    # zone cu clase/id sugestive
    for sel in ["contact", "footer", "header", "topbar", "top-bar",
                "info", "address", "sidebar"]:
        for el in soup.find_all(class_=re.compile(sel, re.I)):
            _regex_zone(el.get_text(" ", strip=True), emails, phones)
        for el in soup.find_all(id=re.compile(sel, re.I)):
            _regex_zone(el.get_text(" ", strip=True), emails, phones)

    # 4. fallback regex pe tot HTML-ul
    _regex_zone(html, emails, phones)

    return sorted(emails), sorted(phones)


def _regex_zone(text: str, emails: set, phones: set):
    for m in EMAIL_REGEX.findall(text):
        if ok_email(m):
            emails.add(m.lower())
    for m in PHONE_REGEX.findall(text):
        n = norm_phone(m)
        if n:
            phones.add(n)


def _jsonld(obj, emails, phones):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in ("email", "contactEmail") and isinstance(v, str):
                if ok_email(v):
                    emails.add(v.lower())
            elif k in ("telephone", "faxNumber") and isinstance(v, str):
                n = norm_phone(v)
                if n:
                    phones.add(n)
            else:
                _jsonld(v, emails, phones)
    elif isinstance(obj, list):
        for i in obj:
            _jsonld(i, emails, phones)


# ── scraping complet per domeniu ──────────────────────────────────────────────

def scrape_site(domain: str):
    base = f"https://{domain}"
    emails: set[str] = set()
    phones: set[str] = set()
    pages_tried: list[str] = []

    # --- homepage ---
    html = fetch(base)
    if html is None:
        base = f"http://{domain}"
        html = fetch(base)

    if html:
        pages_tried.append(base)
        e, p = extract_from_html(html, base)
        emails.update(e)
        phones.update(p)

    # --- pagini de contact ---
    # încearcă slugurile standard; oprire la primul răspuns valid
    for slug in CONTACT_SLUGS:
        if _shutdown:
            break
        contact_url = f"{base}/{slug}"
        time.sleep(0.4)
        html2 = fetch(contact_url)
        if html2 and len(html2) > 500:
            pages_tried.append(contact_url)
            e, p = extract_from_html(html2, base)
            emails.update(e)
            phones.update(p)
            # dacă am găsit ceva, nu mai mergem mai departe
            if emails or phones:
                break

    return sorted(emails), sorted(phones), pages_tried


# ── Excel ─────────────────────────────────────────────────────────────────────

HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
ALT_FILL   = PatternFill("solid", fgColor="D6E4F0")
PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")
OK_FILL    = PatternFill("solid", fgColor="E8F5E9")
NOK_FILL   = PatternFill("solid", fgColor="FFEBEE")


def _hdr(cell, text):
    cell.value = text
    cell.font  = HDR_FONT
    cell.fill  = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_excel(state: dict, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Contacte Site-uri"

    headers = ["Nr", "Domeniu", "Email-uri", "Telefoane", "Status"]
    for ci, h in enumerate(headers, 1):
        _hdr(ws.cell(row=1, column=ci), h)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 25

    row = 2
    for domain in SITES:
        data = state["results"].get(domain)
        if data is None:
            continue

        has_data = bool(data["emails"] or data["phones"])
        if row % 2 == 0:
            fill = OK_FILL if has_data else NOK_FILL
        else:
            fill = OK_FILL if has_data else PatternFill("solid", fgColor="FFF9C4")

        status = "✓ găsit" if has_data else "— negăsit"
        vals = [
            row - 1,
            domain,
            ", ".join(data["emails"]),
            ", ".join(data["phones"]),
            status,
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 12

    wb.save(path)


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    total = len(SITES)
    state = load_cp()
    done  = set(state.get("done", []))

    if done:
        log(f"Checkpoint găsit — {len(done)}/{total} site-uri deja procesate.")

    print("=" * 55)
    print(f"  Site-uri de procesat: {total}")
    print(f"  Deja procesate:       {len(done)}")
    print(f"  Rămase:               {total - len(done)}")
    print("=" * 55)

    for idx, domain in enumerate(SITES, 1):
        if _shutdown:
            break
        if domain in done:
            log(f"Skip (deja procesat): {domain}", idx, total)
            continue

        log(f"Scraping: {domain}", idx, total)
        time.sleep(SCRAPE_DELAY)

        emails, phones, pages = scrape_site(domain)

        log(f"  email={emails}  tel={phones}", idx, total)

        state["results"][domain] = {
            "emails": emails,
            "phones": phones,
            "pages":  pages,
        }
        state["done"].append(domain)
        save_cp(state)

        try:
            build_excel(state, OUTPUT_EXCEL)
        except Exception as e:
            log(f"  Avertisment Excel: {e}", idx, total)

    # sumar
    results = state["results"]
    found    = sum(1 for d in results.values() if d["emails"] or d["phones"])
    no_email = sum(1 for d in results.values() if not d["emails"])
    no_phone = sum(1 for d in results.values() if not d["phones"])

    print("\n" + "=" * 55)
    print("  SUMAR FINAL")
    print("=" * 55)
    print(f"  Site-uri procesate:       {len(results)}")
    print(f"  Cu cel puțin un contact:  {found}")
    print(f"  Fără email:               {no_email}")
    print(f"  Fără telefon:             {no_phone}")
    print(f"\n  Excel salvat la: {OUTPUT_EXCEL}")
    if _shutdown:
        print("  (Oprit cu Ctrl+C — progresul e salvat)")


if __name__ == "__main__":
    main()
