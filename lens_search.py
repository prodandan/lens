"""
Reverse Image Search + Contact Extractor
Google Lens (via SerpApi) + ImgBB hosting + Excel output
"""

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURARE — pune cheile tale aici
# ─────────────────────────────────────────────────────────────────────────────
SERPAPI_KEY  = "CHEIA_TA_SERPAPI"
IMGBB_KEY    = "CHEIA_TA_IMGBB"

IMAGES_FOLDER = r"C:\Users\Dan\Desktop\Poze produse"
OUTPUT_EXCEL  = r"C:\Users\Dan\Desktop\rezultate_lens.xlsx"
CHECKPOINT    = r"C:\Users\Dan\Desktop\lens_progress.json"

LIMIT_IMAGES  = 5          # None = toate imaginile; 5 = test pilot
# ─────────────────────────────────────────────────────────────────────────────

import base64
import json
import os
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
SERPAPI_DELAY  = 1.5   # secunde între apeluri SerpApi
SCRAPE_DELAY   = 1.0   # secunde între scrape-uri
REQUEST_TIMEOUT = 12   # timeout per request HTTP

RO_PATTERN      = re.compile(r'\.ro(/|$)', re.I)
EMAIL_REGEX     = re.compile(
    r'\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b'
)
PHONE_REGEX     = re.compile(
    r'(?<!\d)(?:'
    r'\+40\s*[\s\-\.]?[23789]\d{8}'
    r'|0[23789]\d{8}'
    r'|(?:\+40)?[\s\-\.]?[23789]\d{2}[\s\-\.]?\d{3}[\s\-\.]?\d{3}'
    r')(?!\d)'
)
FAKE_POSITIVE   = re.compile(
    r'sentry\.io|wixpress\.com|googleapis\.com|gstatic\.com'
    r'|schema\.org|example\.com|@\d+x\.'
    r'|\.(png|jpg|gif|svg|webp|ico|css|js)$',
    re.I
)
VALID_PHONE_RE  = re.compile(r'^\+40[237]\d{8}$')

# stare globală (pentru SIGINT curat)
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


def image_files(folder: str) -> list[Path]:
    base = Path(folder)
    files = sorted(
        [p for p in base.glob("*.jpg") if p.stem.isdigit()],
        key=lambda p: int(p.stem),
    )
    if LIMIT_IMAGES:
        files = files[:LIMIT_IMAGES]
    return files


# ── ImgBB upload ─────────────────────────────────────────────────────────────

def upload_imgbb(image_path: Path) -> str | None:
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    try:
        r = requests.post(
            "https://api.imgbb.com/1/upload",
            data={"key": IMGBB_KEY, "image": b64, "expiration": 3600},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        return r.json()["data"]["url"]
    except Exception as e:
        log(f"ImgBB upload eșuat pentru {image_path.name}: {e}")
        return None


# ── SerpApi Google Lens ───────────────────────────────────────────────────────

def serpapi_lens(image_url: str) -> list[dict]:
    """Returnează lista visual_matches din SerpApi Google Lens."""
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


def top_ro_matches(matches: list[dict], max_per_image: int = 5) -> list[dict]:
    """Primele max_per_image rezultate distincte cu domeniu .ro."""
    seen_domains = set()
    results = []
    for m in matches:
        link = m.get("link", "")
        if not RO_PATTERN.search(link):
            continue
        domain = urlparse(link).netloc.lower().lstrip("www.")
        if domain in seen_domains:
            continue
        seen_domains.add(domain)
        results.append({"domain": domain, "url": link})
        if len(results) >= max_per_image:
            break
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
    """Returnează (emails, phones) pentru o pagină URL."""
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

    # 1. mailto: și tel: link-uri
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

    # 3. regex fallback pe HTML brut
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


# ── scraper cu cache per domeniu ──────────────────────────────────────────────

def get_contact_cached(domain: str, url: str, domain_cache: dict) -> tuple[list, list]:
    if domain in domain_cache:
        return domain_cache[domain]["emails"], domain_cache[domain]["phones"]
    time.sleep(SCRAPE_DELAY)
    emails, phones = extract_contact(url)
    domain_cache[domain] = {"emails": emails, "phones": phones}
    return emails, phones


# ── Excel builder ─────────────────────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
ALT_FILL  = PatternFill("solid", fgColor="D6E4F0")
PLAIN_FILL= PatternFill("solid", fgColor="FFFFFF")

SITE_COLORS = ["E8F5E9", "FFF9C4", "FCE4EC", "EDE7F6", "E0F2F1"]


def _hdr(cell, text):
    cell.value = text
    cell.font  = HDR_FONT
    cell.fill  = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_excel(state: dict, output_path: str):
    wb = Workbook()

    # ── Sheet 1: Rezultate per Imagine ────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Rezultate per Imagine"

    # headere rând 1
    _hdr(ws1["A1"], "Nr")
    _hdr(ws1["B1"], "Imagine")
    col = 3
    for i in range(1, 6):
        fill = PatternFill("solid", fgColor=SITE_COLORS[i - 1])
        for sub in ("Domeniu", "URL", "Email", "Telefon"):
            c = ws1.cell(row=1, column=col, value=f"Site {i} — {sub}")
            c.font = Font(bold=True, size=9)
            c.fill = fill
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
        for site in data.get("sites", [])[:5]:
            vals = [
                site.get("domain", ""),
                site.get("url", ""),
                ", ".join(site.get("emails", [])),
                ", ".join(site.get("phones", [])),
            ]
            for v in vals:
                c = ws1.cell(row=row, column=col, value=v)
                c.fill = fill
                c.alignment = Alignment(wrap_text=True, vertical="top")
                col += 1
        # completare coloane lipsă
        while col <= 22:
            ws1.cell(row=row, column=col).fill = fill
            col += 1
        row += 1

    # lățimi coloane Sheet 1
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 18
    for i in range(3, 23):
        letter = get_column_letter(i)
        sub_idx = (i - 3) % 4
        ws1.column_dimensions[letter].width = [18, 40, 30, 20][sub_idx]

    # ── Sheet 2: Cumulat per Site ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Cumulat per Site")
    headers2 = ["Domeniu", "Nr produse", "Email-uri", "Telefoane", "Lista numere imagini"]
    for ci, h in enumerate(headers2, 1):
        _hdr(ws2.cell(row=1, column=ci), h)

    ws2.freeze_panes = "A2"
    ws2.row_dimensions[1].height = 25

    # agregare
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

    sorted_domains = sorted(domain_agg.items(), key=lambda x: -x[1]["count"])

    row = 2
    for domain, agg in sorted_domains:
        fill = ALT_FILL if (row % 2 == 0) else PLAIN_FILL
        cells = [
            domain,
            agg["count"],
            ", ".join(sorted(agg["emails"])),
            ", ".join(sorted(agg["phones"])),
            ", ".join(sorted(agg["images"], key=lambda x: int(x))),
        ]
        for ci, v in enumerate(cells, 1):
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


# ── main loop ─────────────────────────────────────────────────────────────────

def main():
    global _shutdown

    images = image_files(IMAGES_FOLDER)
    total  = len(images)

    if total == 0:
        print("Nicio imagine .jpg cu nume numeric găsită în folderul specificat.")
        sys.exit(1)

    cost_est = total * 0.02
    print("=" * 60)
    print(f"  Imagini găsite:  {total}")
    print(f"  Cost estimat:    ~$0.02/imagine × {total} = ~${cost_est:.2f} SerpApi")
    if LIMIT_IMAGES:
        print(f"  MOD TEST:        LIMIT_IMAGES={LIMIT_IMAGES}")
    print("=" * 60)

    state = load_checkpoint()
    done_set = set(state.get("done", []))

    skipped = len(done_set)
    if skipped:
        log(f"Checkpoint găsit — sar peste {skipped} imagini deja procesate.")

    for idx, img_path in enumerate(images, 1):
        if _shutdown:
            break

        img_key = img_path.stem  # "1", "2", etc.

        if img_key in done_set:
            log(f"Deja procesat, skip.", idx, total)
            continue

        log(f"Upload ImgBB: {img_path.name}", idx, total)
        pub_url = upload_imgbb(img_path)
        if not pub_url:
            log(f"  !! Upload eșuat, înregistrez fără site-uri.", idx, total)
            state["results"][img_key] = {"filename": img_path.name, "sites": []}
            state["done"].append(img_key)
            save_checkpoint(state)
            build_excel(state, OUTPUT_EXCEL)
            continue

        log(f"  URL: {pub_url}", idx, total)
        log(f"SerpApi Google Lens …", idx, total)
        time.sleep(SERPAPI_DELAY)

        if _shutdown:
            break

        matches = serpapi_lens(pub_url)
        ro_sites = top_ro_matches(matches)
        log(f"  → {len(ro_sites)} site-uri .ro găsite", idx, total)

        sites_data = []
        for site in ro_sites:
            if _shutdown:
                break
            domain = site["domain"]
            url    = site["url"]
            log(f"  Scrape: {domain}", idx, total)
            emails, phones = get_contact_cached(domain, url, state["domain_cache"])
            log(f"    email={emails} tel={phones}", idx, total)
            sites_data.append({
                "domain": domain,
                "url":    url,
                "emails": emails,
                "phones": phones,
            })

        state["results"][img_key] = {"filename": img_path.name, "sites": sites_data}
        state["done"].append(img_key)
        save_checkpoint(state)
        log(f"Checkpoint salvat + Excel rebuilt.", idx, total)
        build_excel(state, OUTPUT_EXCEL)

    # ── sumar final ──────────────────────────────────────────────────────────
    processed = len(state["results"])
    with_sites  = sum(1 for v in state["results"].values() if v["sites"])
    no_sites    = processed - with_sites

    # top 10 domenii
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
    print(f"  Imagini procesate:          {processed}")
    print(f"  Cu ≥1 site .ro:             {with_sites}")
    print(f"  Fără site-uri .ro:          {no_sites}")
    print(f"\n  TOP 10 DOMENII:")
    for i, (d, c) in enumerate(top10, 1):
        print(f"    {i:2}. {d:<35} {c} produse")
    print("=" * 60)
    print(f"\n  Excel salvat la: {OUTPUT_EXCEL}")
    if _shutdown:
        print("  (Oprit manual cu Ctrl+C — progresul e salvat)")


if __name__ == "__main__":
    main()
