"""
Reverse Image Search + Contact Extractor
Google Lens (via SerpApi) + ImgBB upload + Excel output
"""

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURARE
# ─────────────────────────────────────────────────────────────────────────────
SERPAPI_KEY   = "91da37a76b3ac1c32479cbbfd1a821bd772820065099b691b236c859bf7971bb"
IMGBB_KEY     = "77efb1da3ad80ccdcdad2f1ea6d350a1"

IMAGES_FOLDER = r"C:\Users\Dan\Desktop\Poze produse"
OUTPUT_EXCEL  = r"C:\Users\Dan\Desktop\rezultate_lens.xlsx"
CHECKPOINT    = r"C:\Users\Dan\Desktop\lens_progress.json"

LIMIT_IMAGES  = None    # None = toate imaginile din folder
START_FROM    = 249     # procesează doar imaginile cu număr >= această valoare

BLACKLIST_DOMAINS = {
    "ieftinmag.ro", "cel.ro", "autoconfort.ro", "evomag.ro",
    "vexio.ro", "flanco.ro", "altex.ro", "domo.ro",
    "okazii.ro", "roveli.ro", "teamdeals.ro", "emagiacumparaturilor.ro",
}
# ─────────────────────────────────────────────────────────────────────────────

import base64
import json
import re
import signal
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── constante ────────────────────────────────────────────────────────────────
SERPAPI_DELAY   = 1.5
SCRAPE_DELAY    = 1.0
REQUEST_TIMEOUT = 12

RO_PATTERN  = re.compile(r'\.ro(/|$)', re.I)
EMAIL_REGEX = re.compile(r'\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b')
PHONE_REGEX = re.compile(
    r'(?<!\d)(?:'
    r'\+40\s*[\s\-\.]?[23789]\d{8}'
    r'|0[23789]\d{8}'
    r'|(?:\+40)?[\s\-\.]?[23789]\d{2}[\s\-\.]?\d{3}[\s\-\.]?\d{3}'
    r')(?!\d)'
)
FAKE_POSITIVE = re.compile(
    r'sentry\.io|wixpress\.com|googleapis\.com|gstatic\.com'
    r'|schema\.org|example\.com|@\d+x\.'
    r'|\.(png|jpg|gif|svg|webp|ico|css|js)$',
    re.I
)
VALID_PHONE_RE = re.compile(r'^\+40[237]\d{8}$')

_shutdown = False


def _handle_sigint(sig, frame):
    global _shutdown
    log("Ctrl+C detectat — oprire curată după imaginea curentă …")
    _shutdown = True


signal.signal(signal.SIGINT, _handle_sigint)


# ── utilități ────────────────────────────────────────────────────────────────

def log(msg: str, idx: int = None, total: int = None):
    ts = datetime.now().strftime("%H:%M:%S")
    prefix = f"[{idx}/{total}] " if idx is not None else ""
    print(f"[{ts}] {prefix}{msg}", flush=True)


def load_checkpoint() -> dict:
    p = Path(CHECKPOINT)
    if p.exists():
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {"done": [], "domain_cache": {}, "results": {}}


def save_checkpoint(state: dict):
    with open(CHECKPOINT, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ── ImgBB upload ─────────────────────────────────────────────────────────────

def upload_imgbb(image_path: Path) -> str | None:
    """Uploadează imaginea pe ImgBB (expiră în 1h) și returnează URL-ul public."""
    try:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        resp = requests.post(
            "https://api.imgbb.com/1/upload",
            data={"key": IMGBB_KEY, "image": b64, "expiration": 3600},
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("success"):
            return data["data"]["url"]
        log(f"  ImgBB eroare răspuns: {data}")
        return None
    except Exception as e:
        log(f"  ImgBB upload eșuat ({image_path.name}): {e}")
        return None


def list_local_images() -> list[Path]:
    """Returnează imaginile numerice din IMAGES_FOLDER, sortate numeric."""
    folder = Path(IMAGES_FOLDER)
    if not folder.exists():
        print(f"EROARE: Folderul nu există: {IMAGES_FOLDER}")
        sys.exit(1)
    files = sorted(
        [p for p in folder.glob("*.jpg") if p.stem.isdigit() and int(p.stem) >= START_FROM],
        key=lambda p: int(p.stem),
    )
    if LIMIT_IMAGES:
        files = files[:LIMIT_IMAGES]
    return files


# ── SerpApi Google Lens ───────────────────────────────────────────────────────

def serpapi_lens(image_url: str) -> list[dict]:
    try:
        r = requests.get(
            "https://serpapi.com/search",
            params={
                "engine": "google_lens",
                "url": image_url,
                "country": "ro",
                "hl": "ro",
                "api_key": SERPAPI_KEY,
            },
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        return r.json().get("visual_matches", [])
    except Exception as e:
        log(f"SerpApi eroare: {e}")
        return []


def top_ro_matches(matches: list[dict]) -> list[dict]:
    """Returnează TOATE site-urile .ro distincte, excluzând blacklist-ul."""
    seen_domains: set[str] = set()
    results = []
    for m in matches:
        link = m.get("link", "")
        if not RO_PATTERN.search(link):
            continue
        domain = urlparse(link).netloc.lower().lstrip("www.")
        if domain in seen_domains:
            continue
        if domain in BLACKLIST_DOMAINS:
            continue
        seen_domains.add(domain)
        results.append({"domain": domain, "url": link})
    return results


# ── extragere contact ─────────────────────────────────────────────────────────

def _normalize_phone(raw: str) -> str | None:
    digits = re.sub(r"[^\d+]", "", raw)
    if digits.startswith("+40"):
        num = "+40" + digits[3:]
    elif digits.startswith("40") and len(digits) == 11:
        num = "+" + digits
    elif digits.startswith("0") and len(digits) == 10:
        num = "+40" + digits[1:]
    else:
        return None
    num = re.sub(r"\s", "", num)
    return num if VALID_PHONE_RE.match(num) else None


def _filter_email(addr: str) -> bool:
    if FAKE_POSITIVE.search(addr):
        return False
    if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', addr):
        return False
    return True


def extract_contact(url: str) -> tuple[list[str], list[str]]:
    emails: set[str] = set()
    phones: set[str] = set()

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "ro-RO,ro;q=0.9",
    }

    try:
        r = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        html = r.text
    except Exception as e:
        log(f"  Scrape eșuat {url}: {e}")
        return [], []

    soup = BeautifulSoup(html, "lxml")

    # 1. mailto: și tel:
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("mailto:"):
            addr = href[7:].split("?")[0].strip()
            if _filter_email(addr):
                emails.add(addr.lower())
        elif href.startswith("tel:"):
            raw = href[4:].strip()
            n = _normalize_phone(raw)
            if n:
                phones.add(n)

    # 2. JSON-LD schema.org
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            _extract_jsonld(data, emails, phones)
        except Exception:
            pass

    # 3. regex fallback
    for m in EMAIL_REGEX.findall(html):
        if _filter_email(m):
            emails.add(m.lower())
    for m in PHONE_REGEX.findall(html):
        n = _normalize_phone(m)
        if n:
            phones.add(n)

    return sorted(emails), sorted(phones)


def _extract_jsonld(obj, emails: set, phones: set):
    if isinstance(obj, dict):
        for key, val in obj.items():
            if key in ("email", "contactEmail") and isinstance(val, str):
                if _filter_email(val):
                    emails.add(val.lower())
            elif key in ("telephone", "faxNumber") and isinstance(val, str):
                n = _normalize_phone(val)
                if n:
                    phones.add(n)
            else:
                _extract_jsonld(val, emails, phones)
    elif isinstance(obj, list):
        for item in obj:
            _extract_jsonld(item, emails, phones)


def get_contact_cached(domain: str, url: str, domain_cache: dict) -> tuple[list, list]:
    if domain in domain_cache:
        return domain_cache[domain]["emails"], domain_cache[domain]["phones"]
    time.sleep(SCRAPE_DELAY)
    emails, phones = extract_contact(url)
    domain_cache[domain] = {"emails": emails, "phones": phones}
    return emails, phones


# ── Excel builder ─────────────────────────────────────────────────────────────

HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
ALT_FILL   = PatternFill("solid", fgColor="D6E4F0")
PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")

# culori ciclice pentru grupuri de site-uri
_SITE_PALETTE = [
    "E8F5E9", "FFF9C4", "FCE4EC", "EDE7F6", "E0F2F1",
    "FBE9E7", "E3F2FD", "F9FBE7", "FCE4EC", "E8EAF6",
    "EFEBE9", "F3E5F5", "E0F7FA", "FFFDE7", "F1F8E9",
]


def _site_fill(i: int) -> PatternFill:
    color = _SITE_PALETTE[i % len(_SITE_PALETTE)]
    return PatternFill("solid", fgColor=color)


def _hdr(cell, text):
    cell.value = text
    cell.font  = HDR_FONT
    cell.fill  = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_excel(state: dict, output_path: str):
    wb = Workbook()

    # ── calculează numărul maxim de site-uri din toate rezultatele ────────────
    max_sites = max(
        (len(d.get("sites", [])) for d in state["results"].values()),
        default=1,
    )
    max_sites = max(max_sites, 1)   # cel puțin 1 coloană site
    last_data_col = 2 + max_sites * 4   # A=1, B=2, apoi 4 col/site

    # ── Sheet 1: Rezultate per Imagine ────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Rezultate per Imagine"

    _hdr(ws1["A1"], "Nr")
    _hdr(ws1["B1"], "Imagine")
    col = 3
    for i in range(1, max_sites + 1):
        sf = _site_fill(i - 1)
        for sub in ("Domeniu", "URL", "Email", "Telefon"):
            c = ws1.cell(row=1, column=col, value=f"Site {i} — {sub}")
            c.font = Font(bold=True, size=9)
            c.fill = sf
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            col += 1

    ws1.freeze_panes = "A2"
    ws1.row_dimensions[1].height = 30

    row = 2
    for img_key, data in sorted(state["results"].items(), key=lambda x: int(x[0])):
        fill = ALT_FILL if (row % 2 == 0) else PLAIN_FILL
        ws1.cell(row=row, column=1, value=int(img_key)).fill = fill
        ws1.cell(row=row, column=2, value=data.get("filename", f"{img_key}.jpg")).fill = fill
        col = 3
        for site in data.get("sites", []):
            for v in [
                site.get("domain", ""),
                site.get("url", ""),
                ", ".join(site.get("emails", [])),
                ", ".join(site.get("phones", [])),
            ]:
                c = ws1.cell(row=row, column=col, value=v)
                c.fill = fill
                c.alignment = Alignment(wrap_text=True, vertical="top")
                col += 1
        # completează coloanele rămase goale până la ultima coloană header
        while col <= last_data_col:
            ws1.cell(row=row, column=col).fill = fill
            col += 1
        row += 1

    # lățimi coloane Sheet 1
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 18
    col_widths = [18, 40, 30, 20]
    for i in range(3, last_data_col + 1):
        ws1.column_dimensions[get_column_letter(i)].width = col_widths[(i - 3) % 4]

    # ── Sheet 2: Cumulat per Site ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Cumulat per Site")
    for ci, h in enumerate(
        ["Domeniu", "Nr produse", "Email-uri", "Telefoane", "Lista numere imagini"], 1
    ):
        _hdr(ws2.cell(row=1, column=ci), h)

    ws2.freeze_panes = "A2"
    ws2.row_dimensions[1].height = 25

    domain_agg: dict[str, dict] = {}
    for img_key, data in state["results"].items():
        for site in data.get("sites", []):
            d = site.get("domain", "")
            if not d:
                continue
            if d not in domain_agg:
                domain_agg[d] = {"count": 0, "emails": set(), "phones": set(), "images": []}
            domain_agg[d]["count"] += 1
            domain_agg[d]["emails"].update(site.get("emails", []))
            domain_agg[d]["phones"].update(site.get("phones", []))
            domain_agg[d]["images"].append(img_key)

    row = 2
    for domain, agg in sorted(domain_agg.items(), key=lambda x: -x[1]["count"]):
        fill = ALT_FILL if (row % 2 == 0) else PLAIN_FILL
        for ci, v in enumerate([
            domain,
            agg["count"],
            ", ".join(sorted(agg["emails"])),
            ", ".join(sorted(agg["phones"])),
            ", ".join(sorted(agg["images"], key=lambda x: int(x))),
        ], 1):
            c = ws2.cell(row=row, column=ci, value=v)
            c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 40
    ws2.column_dimensions["D"].width = 30
    ws2.column_dimensions["E"].width = 35

    wb.save(output_path)


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    global _shutdown

    # 1. Listează imaginile locale
    images = list_local_images()
    total  = len(images)

    if total == 0:
        print(f"EROARE: Nicio imagine .jpg cu nume numeric găsită în:\n  {IMAGES_FOLDER}")
        sys.exit(1)

    cost_est = total * 0.02
    print("=" * 60)
    print(f"  Imagini găsite:        {total}")
    print(f"  Cost estimat SerpApi:  ~$0.02 × {total} = ~${cost_est:.2f}")
    if LIMIT_IMAGES:
        print(f"  MOD TEST:              LIMIT_IMAGES={LIMIT_IMAGES}")
    print("=" * 60)

    # 2. Checkpoint
    state    = load_checkpoint()
    done_set = set(str(p.stem) for p in images if str(p.stem) in state.get("done", []))

    skipped = len(done_set)
    if skipped:
        log(f"Checkpoint găsit — sar peste {skipped} imagini deja procesate.")

    # 3. Procesare
    for idx, img_path in enumerate(images, 1):
        if _shutdown:
            break

        img_key = img_path.stem   # "1", "2", ...

        if img_key in state.get("done", []):
            log(f"Deja procesat ({img_path.name}), skip.", idx, total)
            continue

        # Upload ImgBB
        log(f"Upload ImgBB: {img_path.name}", idx, total)
        pub_url = upload_imgbb(img_path)
        if not pub_url:
            log(f"  Upload eșuat — înregistrez fără site-uri.", idx, total)
            state["results"][img_key] = {"filename": img_path.name, "sites": []}
            state["done"].append(img_key)
            save_checkpoint(state)
            build_excel(state, OUTPUT_EXCEL)
            continue

        log(f"  → {pub_url[:70]}", idx, total)

        # SerpApi
        log(f"Google Lens …", idx, total)
        time.sleep(SERPAPI_DELAY)

        if _shutdown:
            break

        try:
            matches = serpapi_lens(pub_url)
        except Exception as e:
            log(f"  SerpApi excepție neașteptată: {e} — salvez progresul și continui.", idx, total)
            matches = []

        if not matches and pub_url:
            # serpapi_lens a returnat [] fie din eroare, fie că nu există rezultate —
            # înregistrăm oricum ca să nu pierdem progresul
            log(f"  Niciun rezultat SerpApi (eroare sau produs negăsit).", idx, total)

        ro_sites = top_ro_matches(matches)
        log(f"  → {len(ro_sites)} site-uri .ro (după filtrare blacklist)", idx, total)

        sites_data = []
        for site in ro_sites:
            if _shutdown:
                break
            domain = site["domain"]
            url    = site["url"]
            log(f"  Scrape: {domain}", idx, total)
            emails, phones = get_contact_cached(domain, url, state["domain_cache"])
            log(f"    email={emails}  tel={phones}", idx, total)
            sites_data.append({
                "domain": domain,
                "url":    url,
                "emails": emails,
                "phones": phones,
            })

        state["results"][img_key] = {"filename": img_path.name, "sites": sites_data}
        state["done"].append(img_key)
        save_checkpoint(state)
        log(f"Checkpoint + Excel salvat.", idx, total)
        try:
            build_excel(state, OUTPUT_EXCEL)
        except Exception as e:
            log(f"  Avertisment: Excel rebuild eșuat: {e}", idx, total)

    # 4. Sumar
    processed  = len(state["results"])
    with_sites = sum(1 for v in state["results"].values() if v["sites"])
    no_sites   = processed - with_sites

    domain_count: dict[str, int] = {}
    for data in state["results"].values():
        for s in data["sites"]:
            d = s.get("domain", "")
            if d:
                domain_count[d] = domain_count.get(d, 0) + 1
    top10 = sorted(domain_count.items(), key=lambda x: -x[1])[:10]

    print("\n" + "=" * 60)
    print("  SUMAR FINAL")
    print("=" * 60)
    print(f"  Imagini procesate:       {processed}")
    print(f"  Cu ≥1 site .ro:          {with_sites}")
    print(f"  Fără site-uri .ro:       {no_sites}")
    if top10:
        print(f"\n  TOP 10 DOMENII:")
        for i, (d, c) in enumerate(top10, 1):
            print(f"    {i:2}. {d:<35} {c} produse")
    print("=" * 60)
    print(f"\n  Excel salvat la: {OUTPUT_EXCEL}")
    if _shutdown:
        print("  (Oprit manual cu Ctrl+C — progresul e salvat)")


if __name__ == "__main__":
    main()
