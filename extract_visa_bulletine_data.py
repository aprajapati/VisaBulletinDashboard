import re
import json
import hashlib
from datetime import datetime
from urllib.parse import urljoin

import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE = "https://travel.state.gov"
INDEX_URL = "https://travel.state.gov/content/travel/en/legal/visa-law0/visa-bulletin.html"

MONTHS = {m.lower(): i for i, m in enumerate(
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], start=1
)}

def sha256_text(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()

def parse_dd_mmm_yy(s: str):
    s = s.strip()
    if s in {"C", "U"}:
        return {"kind": "STATUS", "status": s}
    # site uses dd-mmm-yy, sometimes with spaces
    m = re.match(r"^\s*(\d{1,2})[-\s]?([A-Za-z]{3})[-\s]?(\d{2})\s*$", s)
    if not m:
        return {"kind": "STATUS", "status": "UNKNOWN"} if s else {"kind": "STATUS", "status": "NA"}
    dd = int(m.group(1))
    mon = MONTHS[m.group(2).lower()]
    yy = int(m.group(3))
    # heuristic: DOS bulletin uses 2-digit year; assume 00-79 => 2000+, else 1900+
    year = 2000 + yy if yy <= 79 else 1900 + yy
    iso = f"{year:04d}-{mon:02d}-{dd:02d}"
    return {"kind": "DATE", "date": iso, "asOfText": s.strip()}

def slug(s: str) -> str:
    s = re.sub(r"\s+", " ", s.strip().lower())
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_") or "unknown"

def canonical_col_id(label: str) -> str:
    l = label.lower()
    if "all chargeability" in l or "except those listed" in l:
        return "WORLDWIDE"
    if "china" in l and "mainland" in l:
        return "CHINA"
    if l.strip() == "india" or " india" in l:
        return "INDIA"
    if "mexico" in l:
        return "MEXICO"
    if "philippines" in l:
        return "PHILIPPINES"
    return slug(label)

def find_bulletin_links():
    html = requests.get(INDEX_URL, timeout=30).text
    soup = BeautifulSoup(html, "html.parser")
    links = set()
    for a in soup.select("a[href]"):
        href = a["href"]
        if "visa-bulletin-for-" in href and href.endswith(".html"):
            links.add(urljoin(BASE, href))
    # index page doesn't always list all; expand by crawling year pages if needed:
    # you can optionally enumerate /visa-bulletin/2025/ etc. via site search, but
    # for stability use a secondary seed list from known year directories.
    return sorted(links)

def extract_text_blocks(soup: BeautifulSoup):
    blocks = []
    # capture headings + paragraphs; keep it simple but effective
    for i, sec in enumerate(soup.select("h2, h3, h4, p")):
        name = sec.name.lower()
        txt = sec.get_text(" ", strip=True)
        if not txt:
            continue
        btype = "OTHER"
        if name.startswith("h"):
            # heading-only blocks not useful; we'll attach heading to next paragraph
            continue
        # heuristic tags
        tags = []
        mentions = []
        for kw in ["retrogress", "oversub", "annual", "unavailable", "unauthorized", "uscis", "dhs", "supersed", "revis"]:
            if kw in txt.lower():
                mentions.append(kw)
        if mentions:
            tags.append("KEYWORDS")
        blocks.append({
            "blockId": f"b{i}",
            "type": btype,
            "heading": None,
            "text": txt,
            "tags": tags,
            "mentions": mentions
        })
    return blocks

def parse_charts_from_html(html: str):
    soup = BeautifulSoup(html, "html.parser")

    # Attempt to read all tables
    tables = pd.read_html(html)
    charts = []

    # Heuristic: walk through the page, mapping nearby headings to tables by order.
    headings = [h.get_text(" ", strip=True) for h in soup.select("h2, h3, h4")]
    # fallback titles if headings count doesn't match
    heading_iter = iter(headings)

    for t_idx, df in enumerate(tables):
        df = df.copy()
        # require at least 2 columns and 2 rows
        if df.shape[1] < 2 or df.shape[0] < 2:
            continue

        # First column tends to be category labels
        raw_cols = [str(c) for c in df.columns.tolist()]
        if raw_cols[0].lower() in {"class", "category", "preference"}:
            pass

        title = None
        # find the next heading that looks like a chart title
        for _ in range(10):
            try:
                candidate = next(heading_iter)
            except StopIteration:
                candidate = None
            if candidate and ("final action" in candidate.lower() or "dates for filing" in candidate.lower() or "visa availability" in candidate.lower()):
                title = candidate
                break
        if not title:
            title = f"Table {t_idx}"

        # infer system + chartType
        tl = title.lower()
        system = "OTHER"
        if "family" in tl:
            system = "FAMILY"
        elif "employment" in tl:
            system = "EMPLOYMENT"
        chartType = "UNKNOWN"
        if "final action" in tl:
            chartType = "FINAL_ACTION_DATES"
        elif "dates for filing" in tl:
            chartType = "DATES_FOR_FILING"

        # columns
        col_labels = raw_cols[1:]  # exclude category col
        columns = []
        for c in col_labels:
            columns.append({
                "colId": canonical_col_id(c),
                "label": c,
                "aliases": []
            })

        # rows
        rows = []
        row_ids = []
        for r in df.iloc[:, 0].astype(str).tolist():
            rid = slug(r)
            row_ids.append(rid)
            rows.append({
                "rowId": rid,
                "label": r,
                "group": None,
                "preferenceCode": None,
                "notes": None
            })

        # cells (sparse)
        cells = []
        for i, rid in enumerate(row_ids):
            for j, col in enumerate(columns, start=1):
                raw = str(df.iat[i, j]).strip()
                val = parse_dd_mmm_yy(raw)
                cells.append({
                    "rowId": rid,
                    "colId": col["colId"],
                    "value": val,
                    "rawText": raw if raw else None,
                    "sourceRef": None
                })

        charts.append({
            "system": system,
            "chartType": chartType,
            "title": title,
            "schemaHint": {"tableKey": None, "parserVersion": "v1"},
            "columns": columns,
            "rows": rows,
            "cells": cells,
            "notes": None
        })

    return soup, charts

def extract_bulletin(url: str):
    html = requests.get(url, timeout=30).text
    soup, charts = parse_charts_from_html(html)
    textBlocks = extract_text_blocks(soup)

    # month/year from title
    h1 = soup.select_one("h1")
    title = h1.get_text(" ", strip=True) if h1 else url
    m = re.search(r"for\s+([A-Za-z]+)\s+(\d{4})", title, flags=re.I)
    month, year = None, None
    if m:
        month_name = m.group(1)[:3].lower()
        month = MONTHS.get(month_name)
        year = int(m.group(2))
    bid = f"{year:04d}-{month:02d}" if (year and month) else slug(title)

    # revision detection
    full_text = soup.get_text(" ", strip=True).lower()
    is_revised = ("supersedes" in full_text) or ("revised" in full_text and "visa bulletin" in full_text)
    revision_note = None
    if "supersedes" in full_text:
        revision_note = "Contains 'supersedes' language (possible revised bulletin)."

    # pdf link (if present)
    pdf = None
    for a in soup.select("a[href$='.pdf']"):
        pdf = urljoin(BASE, a["href"])
        break

    return {
        "id": bid,
        "publication": {"month": month, "year": year, "volume": None, "number": None, "issueDate": None, "revisedDate": None},
        "sources": {"htmlUrl": url, "pdfUrl": pdf, "printerFriendlyUrl": None},
        "revision": {"isRevised": bool(is_revised), "supersedes": None, "supersededBy": None, "revisionNote": revision_note},
        "charts": charts,
        "textBlocks": textBlocks,
        "anomalies": [],
        "raw": {"htmlSha256": sha256_text(html), "extractedAt": datetime.now(datetime.UTC).isoformat() + "Z"}
    }

def build_dataset(urls):
    bulletins = []
    for u in urls:
        try:
            bulletins.append(extract_bulletin(u))
        except Exception as e:
            print("Failed:", u, e)
    return {
        "dataset": {
            "source": "travel.state.gov",
            "generatedAt": datetime.now(datetime.UTC).isoformat() + "Z",
            "schemaVersion": "1.0.0",
            "notes": "Extracted from DOS Visa Bulletin HTML pages"
        },
        "bulletins": bulletins
    }

if __name__ == "__main__":
    urls = find_bulletin_links()
    data = build_dataset(urls)
    with open("visa_bulletins.all.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("Wrote visa_bulletins.all.json with", len(data["bulletins"]), "bulletins")
