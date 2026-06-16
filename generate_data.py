#!/usr/bin/env python3
"""
BeyDeck Data Generator
Fetches ALL Beyblade X components from Beyblade Fandom Wiki API,
merges with WBO competitive data from Excel for initial tier setup.
Output: data.js (loaded by index.html)

Usage:  python3 generate_data.py
Needs:  WBO Excel file in ./temp/
Output: ./data.js
"""

import urllib.request
import urllib.parse
import json
import time
import zipfile
import xml.etree.ElementTree as ET
import re
import os
import sys

# ─── Config ───────────────────────────────────────────────────────────────────

WIKI_API   = "https://beyblade.fandom.com/api.php"
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "temp",
             "WBO Winning Combinations Thread _ Anti Wizard Wizard Club Data Analysis Tool Results.xlsx")
OUT_PATH   = os.path.join(os.path.dirname(__file__), "data.js")
DELAY      = 0.25   # seconds between API calls

# ─── Tier maps (based on WBO data analysis) ───────────────────────────────────

# Normalised name (no spaces, lowercase) → tier label
BLADE_TIER = {
    "wizardrod":      "GOD",
    "sharkscale":     "S",
    "phoenixwing":    "S",
    "cobaltdragoon":  "A",
    "aeropegasus":    "A",
    "silverwolf":     "A",
    "tyrannobeat":    "A",
    "meteordragoon":  "A",
    "hoverwyvern":    "B",
    "golemrock":      "B",
    "bulletgriffon":  "B",
    "clockmirage":    "B",
    "tricerapress":   "B",
    "dranbuster":     "B",
    "knightmail":     "B",
    "impactdrake":    "B",
    "mummycurse":     "B",
    "scorpiospear":   "B",
    "helisscythe":    "B",
    "hellsscythe":    "B",
    "blitz":          "B",
    "dranstrike":     "B",
}

BIT_TIER = {
    "rush":      "S",
    "hexa":      "S",
    "lowrush":   "S",
    "freeball":  "S",
    "elevate":   "S",
    "kick":      "S",
    "ball":      "A",
    "level":     "A",
    "point":     "A",
    "taper":     "A",
    "flat":      "A",
    "loworb":    "A",
    "jolt":      "A",
    "unite":     "A",
    "underneedle": "A",
    "wedge":     "A",
    "orb":       "A",
    "glide":     "A",
    "rubberaccel": "A",
}

RATCHET_TIER = {
    "1-60": "TOP",
    "9-60": "TOP",
    "3-60": "TOP",
    "1-50": "TOP",
}

ASSIST_TIER = {
    "heavy": "S",
    "wheel": "S",
    "flow":  "A",
    "jaggy": "A",
    "slash": "A",
}

OVER_TIER = {
    "peak":  "A",
    "break": "A",
    "flow":  "A",
    "guard": "A",
}

# Lock chips: all named chips = Metal (Tier S = 2 pts)
# The "plastic" chip is implicit (Tier C = 0 pts) — not on Fandom as its own page

# ─── Fandom API helpers ───────────────────────────────────────────────────────

def api(params: dict) -> dict:
    params["format"] = "json"
    url = WIKI_API + "?" + urllib.parse.urlencode(params)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "BeyDeckBot/1.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"    [warn] API error: {e}")
        return {}


def search_pages(query: str, limit: int = 500) -> list[str]:
    """Return all page titles matching srsearch query."""
    titles, offset = [], 0
    while True:
        data = api({
            "action":    "query",
            "list":      "search",
            "srsearch":  query,
            "srnamespace": 0,
            "srlimit":   min(50, limit - len(titles)),
            "sroffset":  offset,
        })
        hits = data.get("query", {}).get("search", [])
        titles.extend(h["title"] for h in hits)
        time.sleep(DELAY)

        cont = data.get("continue", {})
        if not cont or len(titles) >= limit:
            break
        offset = cont.get("sroffset", offset + len(hits))

    return titles


def page_images(page_title: str) -> list[str]:
    """Return image file titles listed on a wiki page."""
    data = api({"action": "query", "titles": page_title, "prop": "images"})
    time.sleep(DELAY)
    pages = data.get("query", {}).get("pages", {})
    imgs = []
    for p in pages.values():
        imgs.extend(i["title"] for i in p.get("images", [])
                    if i["title"].lower().endswith((".png", ".jpg", ".jpeg")))
    return imgs


def image_cdn_url(file_title: str) -> str | None:
    """Resolve a File:XYZ.png title to its CDN URL."""
    data = api({
        "action": "query",
        "titles": file_title,
        "prop":   "imageinfo",
        "iiprop": "url",
    })
    time.sleep(DELAY)
    for p in data.get("query", {}).get("pages", {}).values():
        info = p.get("imageinfo", [])
        if info:
            return info[0].get("url")
    return None


def best_image(page_title: str, keyword: str) -> str | None:
    """
    Pick the best image for a component page.
    Prefer images whose filename contains the keyword (e.g. 'Blade', 'Ratchet').
    """
    imgs = page_images(page_title)
    if not imgs:
        return None

    # Skip obviously wrong images (tournament banners, product box shots, etc.)
    skip = {"stadium", "arena", "booster", "decklayer", "photo", "manga",
            "anime", "infobox", "logo", "icon", "banner", "map", "flag"}
    imgs = [i for i in imgs if not any(s in i.lower() for s in skip)]

    if not imgs:
        return None

    # Prefer images whose filename starts with or contains the keyword
    kw = keyword.lower()
    preferred = [i for i in imgs if kw in i.lower().replace(" ", "").replace("-", "")]
    candidates = preferred if preferred else imgs

    return image_cdn_url(candidates[0])


# ─── Normalise helpers ────────────────────────────────────────────────────────

def norm(s: str) -> str:
    """Lowercase, strip spaces & hyphens for tier lookup."""
    return re.sub(r"[\s\-]", "", s).lower()


def clean(title: str, prefix: str) -> str:
    """'Blade - PhoenixWing' → 'PhoenixWing'"""
    return title.removeprefix(f"{prefix} - ").strip()


# ─── Read Excel for competitive data ─────────────────────────────────────────

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

def _shared_strings(zf) -> list[str]:
    try:
        with zf.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
        result = []
        for si in tree.findall(f".//{{{NS}}}si"):
            texts = [t.text or "" for t in si.iter(f"{{{NS}}}t")]
            result.append("".join(texts))
        return result
    except Exception:
        return []


def _read_sheet(zf, sheet_idx: int, ss: list[str], max_rows=2000) -> list[list[str]]:
    rows = []
    with zf.open(f"xl/worksheets/sheet{sheet_idx}.xml") as f:
        tree = ET.parse(f)
    for i, row in enumerate(tree.findall(f".//{{{NS}}}row")):
        if i >= max_rows:
            break
        cells = []
        for c in row.findall(f"{{{NS}}}c"):
            t = c.get("t", "")
            v = c.find(f"{{{NS}}}v")
            if v is not None and v.text is not None:
                cells.append(ss[int(v.text)] if t == "s" else v.text)
            else:
                cells.append("")
        rows.append(cells)
    return rows


def load_excel_competitive_data() -> dict:
    """
    Returns dict: normalised_name → {"usage": int, "comp_score": float, "type": str, "anatomy": str}
    Reads the "Full Thread Data" sheet (sheet 2) for part-level stats.
    Blades that appear in CX combos (rows where LockChip col is non-empty) get anatomy="cx".
    """
    result = {}
    if not os.path.exists(EXCEL_PATH):
        print(f"[warn] Excel not found at {EXCEL_PATH} — skipping competitive data")
        return result

    try:
        with zipfile.ZipFile(EXCEL_PATH) as zf:
            ss = _shared_strings(zf)
            rows = _read_sheet(zf, 2, ss, 3000)

        # Part statistics sidebar (from col 27 onwards):
        # [27]=LockChip Name, [28]=LC pts, [29]=LC usage, [30]=LC comp
        # [32]=OBlade Name, [33]=OB pts, [34]=OB usage, [35]=OB comp
        # [37]=Blade Name, [38]=B pts, [39]=B usage, [40]=B comp
        # [43]=Assist Name, [44]=A pts, [45]=A usage, [46]=A comp
        # [48]=Ratchet Name, [49]=R pts, [50]=R usage, [51]=R comp
        # [53]=Bit Name, [54]=Bt pts, [55]=Bt usage, [56]=Bt comp

        type_cols = {
            "lockchip":   (27, 29, 30),
            "overblade":  (32, 34, 35),
            "blade":      (37, 39, 40),
            "assistblade":(43, 45, 46),
            "ratchet":    (48, 50, 51),
            "bit":        (53, 55, 56),
        }

        for row in rows[1:]:
            for t, (name_col, usage_col, comp_col) in type_cols.items():
                if name_col >= len(row):
                    continue
                name = row[name_col].strip()
                if not name:
                    continue
                key = norm(name)
                if key not in result:
                    try:
                        usage = int(float(row[usage_col])) if usage_col < len(row) else 0
                        comp  = float(row[comp_col]) if comp_col < len(row) else 0.0
                    except (ValueError, IndexError):
                        usage, comp = 0, 0.0
                    result[key] = {"usage": usage, "comp_score": comp, "type": t, "name": name}


    except Exception as e:
        print(f"[warn] Excel read error: {e}")

    print(f"  Loaded {len(result)} competitive data entries from Excel")
    return result


# ─── Tier assignment ──────────────────────────────────────────────────────────

def assign_tier(name: str, comp_type: str, excel: dict) -> str:
    key = norm(name)

    if comp_type == "blade":
        return BLADE_TIER.get(key, "C")

    if comp_type == "ratchet":
        # Ratchets are named like "1-60", "9-60"
        raw = name.strip()
        return RATCHET_TIER.get(raw, "STANDARD")

    if comp_type == "bit":
        return BIT_TIER.get(key, "C")

    if comp_type == "assistblade":
        return ASSIST_TIER.get(key, "C")

    if comp_type == "overblade":
        return OVER_TIER.get(key, "C")

    if comp_type == "lockchip":
        return "S"  # All named lock chips are Metal (weighted) = Tier S

    return "C"


# ─── Per-component-type fetch ─────────────────────────────────────────────────

TIER_POINTS = {
    # blades
    "GOD": 8, "S": 6, "A": 4, "B": 2, "C": 0,
    # ratchets
    "TOP": 1, "STANDARD": 0,
}

COMPONENT_TYPES = [
    # (search_query, prefix_to_strip, internal_type, image_keyword, anatomy)
    # anatomy is only set for blades (standard vs cx); None for other types
    ("Blade - ",         "Blade",        "blade",        "Blade",  "standard"),
    ("Main Blade - ",    "Main Blade",   "blade",        "Blade",  "cx"),
    ("Ratchet - ",       "Ratchet",      "ratchet",      "Ratchet", None),
    ("Bit - ",           "Bit",          "bit",          "Bit",     None),
    ("Assist Blade - ",  "Assist Blade", "assistblade",  "Assist",  None),
    ("Over Blade - ",    "Over Blade",   "overblade",    "Over",    None),
    ("Lock Chip - ",     "Lock Chip",    "lockchip",     "Lock",    None),
]

# Pages to skip (not actual Beyblade X components)
SKIP_TITLES = {
    "bit chip", "bit beast", "bit protector", "disc - ratchet",
    "ratchet integrated bit", "ratchet-integrated blade",
    "ratchet integrated blade", "basic line",
}

def should_skip(title: str) -> bool:
    t = title.lower()
    for s in SKIP_TITLES:
        if t.startswith(s) or s in t:
            return True
    # Skip old-gen items (Hasbro variants, Star Wars, etc.)
    skip_words = ["hasbro", "stormtrooper", "darth", "vader", "luke", "grogu",
                  "optimus", "bumblebee", "venom", "miles", "stag beast",
                  "hack viking", "hackviking"]
    return any(w in t for w in skip_words)


def fetch_components(search_query: str, prefix: str, comp_type: str,
                     img_keyword: str, excel: dict, anatomy: str | None = None) -> list[dict]:
    print(f"\n  Searching: '{search_query}' ...")
    titles = search_pages(search_query, limit=200)
    print(f"  Found {len(titles)} pages")

    components = []
    for i, title in enumerate(titles):
        if should_skip(title):
            continue
        name = clean(title, prefix)
        if not name or " - " in name:
            # e.g. "Blade - Ultimate (Hasbro)" — secondary check
            continue

        tier  = assign_tier(name, comp_type, excel)
        pts   = TIER_POINTS.get(tier, 0)
        edata = excel.get(norm(name), {})

        print(f"  [{i+1}/{len(titles)}] {name:30s} tier={tier}  usage={edata.get('usage','-')}")

        img = best_image(title, img_keyword)

        entry = {
            "id":         norm(name),
            "name":       name,
            "type":       comp_type,
            "tier":       tier,
            "points":     pts,
            "usage":      edata.get("usage", 0),
            "comp_score": round(edata.get("comp_score", 0.0), 2),
            "image":      img,
        }
        if anatomy is not None:
            entry["anatomy"] = anatomy
        components.append(entry)

    return components


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print(" BeyDeck Data Generator")
    print("=" * 60)

    # 1. Load Excel competitive data
    print("\n[1] Loading Excel competitive data...")
    excel = load_excel_competitive_data()

    # 2. Fetch all component types from Fandom
    print("\n[2] Fetching components from Fandom API...")
    all_components: dict[str, list] = {}

    for search_q, prefix, comp_type, img_kw, anatomy in COMPONENT_TYPES:
        comps = fetch_components(search_q, prefix, comp_type, img_kw, excel, anatomy)
        if comp_type not in all_components:
            all_components[comp_type] = []
        all_components[comp_type].extend(comps)
        print(f"  → {len(comps)} {comp_type} ({anatomy or 'n/a'}) components fetched")

    # Add implicit Plastic Lock Chip (not on Fandom as own page)
    all_components["lockchip"].append({
        "id": "plastic",
        "name": "Plastic",
        "type": "lockchip",
        "tier": "C",
        "points": 0,
        "usage": 0,
        "comp_score": 0.0,
        "image": None,
    })

    # 3. Build config (default — admin can override in localStorage)
    config = {
        "budgetMax":  30,
        "cxDiscount": -2,
        "malusCombos": [
            {
                "id":     "rod-160-hexa",
                "label":  "WizardRod 1-60 Hexa",
                "blade":  "wizardrod",
                "ratchet": "160",
                "bit":    "hexa",
                "malus":  2,
            },
            {
                "id":     "shark-lowrush",
                "label":  "SharkScale + LowRush",
                "blade":  "sharkscale",
                "ratchet": None,   # any ratchet
                "bit":    "lowrush",
                "malus":  2,
            },
            {
                "id":     "cobalt-elevate",
                "label":  "CobaltDragoon + Elevate",
                "blade":  "cobaltdragoon",
                "ratchet": None,
                "bit":    "elevate",
                "malus":  2,
            },
        ],
    }

    # 4. Serialize to data.js
    print("\n[3] Writing data.js...")

    payload = {
        "config":     config,
        "components": all_components,
    }

    js = "// BeyDeck — auto-generated by generate_data.py. Do not edit manually.\n"
    js += "const BBX_DATA = " + json.dumps(payload, ensure_ascii=False, indent=2) + ";\n"

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(js)

    total = sum(len(v) for v in all_components.values())
    print(f"\n✓ Done! {total} components written to {OUT_PATH}")
    print(f"  Blades:       {len(all_components.get('blade', []))}")
    print(f"  Ratchets:     {len(all_components.get('ratchet', []))}")
    print(f"  Bits:         {len(all_components.get('bit', []))}")
    print(f"  Assist Blades:{len(all_components.get('assistblade', []))}")
    print(f"  Over Blades:  {len(all_components.get('overblade', []))}")
    print(f"  Lock Chips:   {len(all_components.get('lockchip', []))}")


if __name__ == "__main__":
    main()
