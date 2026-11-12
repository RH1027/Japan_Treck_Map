"""Build data/shops.json from the recommendations Excel file.

Usage:
    python3 tools/build_data.py path/to/Japan_Treck_recommendations_list.xlsx

- Each sheet (Tokyo, Kyoto, Osaka, Hiroshima, ...) becomes a city tab.
- Columns are matched by header name, so new columns (e.g. "Tabelog",
  "Google Map", "Address") can be added in any order.
- Hyperlinks in the HP / Tabelog / Google Map cells are used as link targets.
- Locations come from data/locations.json (name -> address/lat/lng). Shops
  without an entry there are geocoded from "Address" (or name + area) with
  the GSI address search API and added to locations.json for review.
"""
import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
LOC_FILE = ROOT / "data" / "locations.json"
OUT_FILE = ROOT / "data" / "shops.json"
TR_FILE = ROOT / "data" / "translations.json"


def has_japanese(t):
    return any("\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" for ch in (t or ""))

ALIASES = {
    "name": ["name"],
    "area": ["area"],
    "description": ["description", "discription"],
    "hp": ["hp", "website"],
    "reservation": ["reservation", "resercation"],
    "note": ["note"],
    "tabelog": ["tabelog", "食べログ"],
    "gmap": ["google map", "googlemap", "google maps"],
    "address": ["address"],
    "google_rating": ["google rating", "google map rating", "google maps rating"],
    "tabelog_rating": ["tabelog rating", "tabelog score"],
    "rating_date": ["rating date", "ratings as of"],
}


def key_for(header):
    h = (header or "").strip().lower()
    for k, names in ALIASES.items():
        if h in names:
            return k
    return None


def gsi_geocode(q):
    url = "https://msearch.gsi.go.jp/address-search/AddressSearch?q=" + urllib.parse.quote(q)
    with urllib.request.urlopen(url, timeout=20) as r:
        res = json.load(r)
    if not res:
        return None
    lng, lat = res[0]["geometry"]["coordinates"]
    return {"lat": lat, "lng": lng, "matched": res[0]["properties"]["title"]}


def main(xlsx):
    locs = json.loads(LOC_FILE.read_text(encoding="utf-8")) if LOC_FILE.exists() else {}
    # Japanese name in the Excel file -> English name / description / note.
    trs = json.loads(TR_FILE.read_text(encoding="utf-8")) if TR_FILE.exists() else {}
    untranslated = []
    wb = openpyxl.load_workbook(xlsx)
    cities = []
    missing = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows())
        if not rows:
            continue
        cols = {}
        for c in rows[0]:
            k = key_for(c.value)
            if k and k not in cols:
                cols[k] = c.column - 1
        shops = []
        for row in rows[1:]:
            def val(k):
                i = cols.get(k)
                if i is None or i >= len(row):
                    return None
                v = row[i].value
                return str(v).strip() if v is not None else None

            def link(k):
                i = cols.get(k)
                if i is None or i >= len(row):
                    return None
                cell = row[i]
                if cell.hyperlink and cell.hyperlink.target:
                    return {"label": val(k) or "Link", "url": cell.hyperlink.target}
                v = val(k)
                if v and v.startswith("http"):
                    return {"label": "Link", "url": v}
                return None

            def num(k):
                # A number typed in the "食べログ" / "Google Map" column is a rating.
                i = cols.get(k)
                if i is None or i >= len(row):
                    return None
                v = row[i].value
                return str(v) if isinstance(v, (int, float)) else None

            name = val("name")
            if not name:
                continue
            loc = locs.get(name)
            if not loc:
                q = val("address") or f"{name} {val('area') or ''} {ws.title}"
                g = gsi_geocode(q) if val("address") else None
                if g:
                    loc = {"address": val("address"), **g, "source": "GSI geocode (auto)"}
                    locs[name] = loc
                else:
                    missing.append(f"{ws.title}: {name}")
            tr = trs.get(name, {})
            name_en = tr.get("name") or name
            name_jp = (loc or {}).get("name_jp") or (name if has_japanese(name) else None)
            desc = tr.get("description") or val("description")
            note = tr.get("note") or val("note")
            if has_japanese(name_en) or has_japanese(desc) or has_japanese(note):
                untranslated.append(f"{ws.title}: {name}")
            shops.append({
                "name": name_en,
                "name_jp": name_jp,
                "area": val("area"),
                "description": desc,
                "hp": link("hp"),
                "reservation": val("reservation"),
                "note": note,
                "tabelog": link("tabelog") or ({"label": "Tabelog", "url": loc["tabelog_url"]} if loc and loc.get("tabelog_url") else None),
                "gmap": link("gmap"),
                "google_rating": val("google_rating") or num("gmap"),
                "tabelog_rating": val("tabelog_rating") or num("tabelog"),
                "rating_date": val("rating_date"),
                "address": (loc or {}).get("address"),
                "lat": (loc or {}).get("lat"),
                "lng": (loc or {}).get("lng"),
            })
        cities.append({"city": ws.title, "shops": shops})
    LOC_FILE.write_text(json.dumps(locs, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    OUT_FILE.write_text(json.dumps({"cities": cities}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {OUT_FILE} ({sum(len(c['shops']) for c in cities)} shops)")
    if untranslated:
        print("Needs English translation (add to data/translations.json):")
        for m in untranslated:
            print("  -", m)
    if missing:
        print("No location yet (add to data/locations.json):")
        for m in missing:
            print("  -", m)


if __name__ == "__main__":
    main(sys.argv[1])
