#!/usr/bin/env python3
"""
AI & LLM Ethics Living Literature Tracker

Fetches recent scholarly work from arXiv, OpenAlex and Crossref; watches selected
official AI-policy pages; deduplicates and scores candidates; and rebuilds the
reader-facing Excel workbook.

Default mode uses only free/public sources. OPENALEX_API_KEY and CONTACT_EMAIL
are optional environment variables.
"""
from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse, urlunparse

import xml.etree.ElementTree as ET
import requests
import xlsxwriter
import yaml
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config" / "sources.yaml"
DATA_DIR = ROOT / "data"

LIBRARY_PATH = DATA_DIR / "library.json"
DISCOVERED_PATH = DATA_DIR / "discovered.json"
STATE_PATH = DATA_DIR / "state.json"
LOG_PATH = DATA_DIR / "update_log.json"

SESSION = requests.Session()

CORE_LIBRARY_HEADERS = [
    ("id", "ID"),
    ("theme", "Theme"),
    ("subtopic", "Subtopic"),
    ("title", "Title"),
    ("author_body", "Author / Body"),
    ("year", "Year"),
    ("source_type", "Source Type"),
    ("status", "Status"),
    ("what_it_covers", "What it covers"),
    ("why_reader_benefits", "Why the reader benefits"),
    ("best_use", "Best use"),
    ("relevance_5", "2026 Relevance Score"),
    ("read_priority", "Read Priority"),
    ("region_scope", "Region / Scope"),
    ("key_ethical_question", "Key ethical question"),
    ("source_url", "Source URL"),
    ("notes_status", "Notes / status"),
    ("reader_mode", "Reader mode"),
]

DISCOVERY_HEADERS = [
    ("detected_date", "Detected"),
    ("publication_date", "Published"),
    ("title", "Title"),
    ("theme", "Theme"),
    ("source_name", "Source"),
    ("source_type", "Source Type"),
    ("importance_score", "Importance Score"),
    ("read_priority", "Priority"),
    ("why_flagged", "Why flagged"),
    ("source_url", "URL"),
    ("doi", "DOI"),
    ("decision", "Decision"),
    ("notes_status", "Notes"),
]

def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default

def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")

def load_config() -> dict[str, Any]:
    return yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))

def today_utc() -> date:
    return datetime.now(timezone.utc).date()

def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    s = BeautifulSoup(str(value), "html.parser").get_text(" ")
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()

def normalize_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(title).lower()).strip()

def canonical_url(url: str) -> str:
    if not url:
        return ""
    try:
        p = urlparse(url.strip())
        host = p.netloc.lower()
        path = re.sub(r"/+$", "", p.path)
        # Avoid tracking query strings/fragments for identity.
        return urlunparse((p.scheme.lower() or "https", host, path, "", "", ""))
    except Exception:
        return url.strip()

def normalize_doi(doi: str) -> str:
    doi = clean_text(doi).lower()
    doi = doi.replace("https://doi.org/", "").replace("http://doi.org/", "").replace("doi:", "").strip()
    return doi

def candidate_key(item: dict[str, Any]) -> str:
    doi = normalize_doi(item.get("doi", ""))
    if doi:
        return "doi:" + doi
    url = canonical_url(item.get("source_url", ""))
    if url:
        return "url:" + url
    return f"title:{normalize_title(item.get('title',''))}|{item.get('year','')}"

def parse_date(value: Any) -> date | None:
    if not value:
        return None
    s = str(value)[:10]
    try:
        return date.fromisoformat(s)
    except Exception:
        return None

def reconstruct_openalex_abstract(inv: dict[str, list[int]] | None) -> str:
    if not inv:
        return ""
    positioned = []
    for word, positions in inv.items():
        for pos in positions:
            positioned.append((pos, word))
    positioned.sort()
    return clean_text(" ".join(word for _, word in positioned))

def classify_theme(text: str, config: dict[str, Any]) -> tuple[str, list[str], int]:
    txt = clean_text(text).lower()
    best_theme = "Foundations"
    best_hits: list[str] = []
    best_score = 0
    for theme, keywords in config["themes"].items():
        hits = []
        score = 0
        for kw in keywords:
            kw_l = kw.lower()
            count = txt.count(kw_l)
            if count:
                hits.append(kw)
                score += min(count, 3)
        if score > best_score:
            best_theme, best_hits, best_score = theme, hits, score
    return best_theme, best_hits[:6], best_score

def recency_points(pub_date: date | None, detected: date) -> int:
    if pub_date is None:
        return 8
    days = max(0, (detected - pub_date).days)
    if days <= 14:
        return 15
    if days <= 30:
        return 13
    if days <= 90:
        return 10
    if days <= 365:
        return 6
    return 2

def score_candidate(item: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    title = clean_text(item.get("title", ""))
    abstract = clean_text(item.get("abstract", ""))
    combined = f"{title}. {abstract}"
    theme, hits, hit_score = classify_theme(combined, config)
    item["theme"] = theme
    item["matched_keywords"] = hits

    # Strongly favor title matches, then abstract/body matches.
    title_l = title.lower()
    title_hits = sum(1 for h in hits if h.lower() in title_l)
    relevance = min(45, 8 * title_hits + 5 * max(0, hit_score - title_hits))
    if relevance == 0:
        relevance = min(15, 3 * hit_score)

    authority = float(item.get("authority", 75))
    authority_pts = round(authority * 0.25)

    pub = parse_date(item.get("publication_date"))
    detected = parse_date(item.get("detected_date")) or today_utc()
    recency = recency_points(pub, detected)

    impact_hits = [s for s in config.get("impact_signals", []) if s.lower() in combined.lower()]
    impact = min(15, 4 * len(impact_hits))
    # Official policy pages are important even if the anchor text is short.
    if item.get("source_adapter") == "official":
        impact = max(impact, 10)

    score = int(min(100, relevance + authority_pts + recency + impact))
    item["importance_score"] = score

    if score >= 90:
        priority = "Must Read Candidate"
    elif score >= 80:
        priority = "High"
    elif score >= 70:
        priority = "Medium"
    else:
        priority = "Low"
    item["read_priority"] = priority

    bits = []
    if hits:
        bits.append(f"matches {theme}: {', '.join(hits[:4])}")
    if authority >= 95:
        bits.append("authoritative institutional source")
    elif authority >= 85:
        bits.append("strong scholarly metadata source")
    if impact_hits:
        bits.append("impact signal: " + ", ".join(impact_hits[:3]))
    if pub and (detected - pub).days <= 30:
        bits.append("recent")
    item["why_flagged"] = "; ".join(bits) or "matched configured AI-ethics discovery query"
    return item

def make_session_headers(config: dict[str, Any]) -> dict[str, str]:
    contact = os.getenv("CONTACT_EMAIL", "").strip()
    ua = config["runtime"].get("user_agent", "AI-Ethics-Living-Tracker/1.0")
    if contact:
        ua += f" ({contact})"
    return {"User-Agent": ua, "Accept": "*/*"}

def fetch_arxiv(query: dict[str, Any], config: dict[str, Any], since: date) -> list[dict[str, Any]]:
    endpoint = "https://export.arxiv.org/api/query"
    max_results = int(config["project"]["max_results_per_query"])
    params = {
        "search_query": query["arxiv"],
        "start": 0,
        "max_results": max_results,
        "sortBy": "submittedDate",
        "sortOrder": "descending",
    }
    r = SESSION.get(endpoint, params=params, headers=make_session_headers(config),
                    timeout=config["runtime"]["request_timeout_seconds"])
    r.raise_for_status()
    root = ET.fromstring(r.content)
    atom = {"a": "http://www.w3.org/2005/Atom", "x": "http://arxiv.org/schemas/atom"}
    out = []
    for e in root.findall("a:entry", atom):
        pub = clean_text(e.findtext("a:published", default="", namespaces=atom))[:10]
        pub_date = parse_date(pub)
        if pub_date and pub_date < since:
            continue
        authors = []
        for a in e.findall("a:author", atom):
            name = clean_text(a.findtext("a:name", default="", namespaces=atom))
            if name:
                authors.append(name)
        url = clean_text(e.findtext("a:id", default="", namespaces=atom))
        doi = clean_text(e.findtext("x:doi", default="", namespaces=atom))
        out.append({
            "title": clean_text(e.findtext("a:title", default="", namespaces=atom)),
            "author_body": ", ".join(authors),
            "year": pub_date.year if pub_date else "",
            "publication_date": pub,
            "detected_date": today_utc().isoformat(),
            "source_name": "arXiv",
            "source_adapter": "arxiv",
            "source_type": "Preprint / research",
            "status": "Current research / preprint",
            "source_url": canonical_url(url),
            "doi": normalize_doi(doi),
            "abstract": clean_text(e.findtext("a:summary", default="", namespaces=atom)),
            "authority": config["source_authority"]["arxiv"],
            "region_scope": "Research / global",
        })
    return out

def fetch_openalex(query: dict[str, Any], config: dict[str, Any], since: date) -> list[dict[str, Any]]:
    endpoint = "https://api.openalex.org/works"
    max_results = int(config["project"]["max_results_per_query"])
    params = {
        "search": query["general"],
        "filter": f"from_publication_date:{since.isoformat()},to_publication_date:{today_utc().isoformat()}",
        "sort": "publication_date:desc",
        "per_page": min(100, max_results),
    }
    key = os.getenv("OPENALEX_API_KEY", "").strip()
    if key:
        params["api_key"] = key
    r = SESSION.get(endpoint, params=params, headers=make_session_headers(config),
                    timeout=config["runtime"]["request_timeout_seconds"])
    r.raise_for_status()
    out = []
    for w in r.json().get("results", []):
        authors = []
        for a in w.get("authorships", [])[:12]:
            name = (a.get("author") or {}).get("display_name")
            if name:
                authors.append(name)
        location = w.get("primary_location") or {}
        src = (location.get("source") or {}).get("display_name") or "OpenAlex"
        pub = w.get("publication_date") or ""
        pub_date = parse_date(pub)
        doi = normalize_doi(w.get("doi") or "")
        url = w.get("doi") or w.get("id") or ""
        out.append({
            "title": clean_text(w.get("display_name") or w.get("title") or ""),
            "author_body": ", ".join(authors),
            "year": w.get("publication_year") or (pub_date.year if pub_date else ""),
            "publication_date": pub,
            "detected_date": today_utc().isoformat(),
            "source_name": src,
            "source_adapter": "openalex",
            "source_type": clean_text(w.get("type") or "Scholarly work"),
            "status": "Scholarly metadata",
            "source_url": canonical_url(url),
            "doi": doi,
            "abstract": reconstruct_openalex_abstract(w.get("abstract_inverted_index")),
            "authority": config["source_authority"]["openalex"],
            "region_scope": "Research / global",
            "citation_count": w.get("cited_by_count", 0),
        })
    return out

def crossref_date(item: dict[str, Any]) -> str:
    for key in ("published-print", "published-online", "published", "issued", "created"):
        block = item.get(key)
        if isinstance(block, dict) and block.get("date-parts"):
            parts = block["date-parts"][0]
            if not parts:
                continue
            y = int(parts[0])
            m = int(parts[1]) if len(parts) > 1 else 1
            d = int(parts[2]) if len(parts) > 2 else 1
            try:
                return date(y, m, d).isoformat()
            except Exception:
                return f"{y:04d}-{m:02d}-01"
    return ""

def fetch_crossref(query: dict[str, Any], config: dict[str, Any], since: date) -> list[dict[str, Any]]:
    endpoint = "https://api.crossref.org/works"
    max_results = int(config["project"]["max_results_per_query"])
    params = {
        "query.bibliographic": query["general"],
        "filter": f"from-pub-date:{since.isoformat()},until-pub-date:{today_utc().isoformat()}",
        "rows": max_results,
        "sort": "published",
        "order": "desc",
    }
    contact = os.getenv("CONTACT_EMAIL", "").strip()
    if contact:
        params["mailto"] = contact
    r = SESSION.get(endpoint, params=params, headers=make_session_headers(config),
                    timeout=config["runtime"]["request_timeout_seconds"])
    r.raise_for_status()
    out = []
    for w in r.json().get("message", {}).get("items", []):
        title = clean_text((w.get("title") or [""])[0])
        authors = []
        for a in w.get("author", [])[:12]:
            name = " ".join(x for x in [a.get("given", ""), a.get("family", "")] if x).strip()
            if name:
                authors.append(name)
        pub = crossref_date(w)
        pub_date = parse_date(pub)
        container = clean_text((w.get("container-title") or [""])[0])
        doi = normalize_doi(w.get("DOI", ""))
        url = f"https://doi.org/{doi}" if doi else w.get("URL", "")
        out.append({
            "title": title,
            "author_body": ", ".join(authors) or clean_text(w.get("publisher", "")),
            "year": pub_date.year if pub_date else "",
            "publication_date": pub,
            "detected_date": today_utc().isoformat(),
            "source_name": container or clean_text(w.get("publisher", "")) or "Crossref",
            "source_adapter": "crossref",
            "source_type": clean_text(w.get("type", "Published work")),
            "status": "Published / Crossref metadata",
            "source_url": canonical_url(url),
            "doi": doi,
            "abstract": clean_text(w.get("abstract", "")),
            "authority": config["source_authority"]["crossref"],
            "region_scope": "Research / global",
            "citation_count": w.get("is-referenced-by-count", 0),
        })
    return out

def relevant_anchor(text: str, href: str, config: dict[str, Any]) -> bool:
    combined = f"{text} {href}".lower()
    _, _, score = classify_theme(combined, config)
    return score > 0

def watch_policy_page(page: dict[str, Any], config: dict[str, Any]) -> list[dict[str, Any]]:
    headers = make_session_headers(config)
    r = SESSION.get(page["url"], headers=headers, timeout=config["runtime"]["request_timeout_seconds"])
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    base_domain = page["domain"].lower()
    seen_local = set()
    candidates = []
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        if len(text) < 10:
            continue
        href = canonical_url(urljoin(page["url"], a["href"]))
        parsed = urlparse(href)
        host = parsed.netloc.lower()
        if not (host == base_domain or host.endswith("." + base_domain)):
            continue
        if href in seen_local or href == canonical_url(page["url"]):
            continue
        if not relevant_anchor(text, href, config):
            continue
        seen_local.add(href)
        candidates.append({
            "title": text[:300],
            "author_body": page["name"],
            "year": today_utc().year,
            "publication_date": today_utc().isoformat(),  # detection date when page metadata is unavailable
            "detected_date": today_utc().isoformat(),
            "source_name": page["name"],
            "source_adapter": "official",
            "source_type": "Official policy / institutional update",
            "status": "Official page watch",
            "source_url": href,
            "doi": "",
            "abstract": "",
            "authority": page["authority"],
            "region_scope": "Official / policy",
        })
        if len(candidates) >= int(config["project"]["max_policy_links_per_page"]):
            break
    return candidates

def offline_fixture() -> list[dict[str, Any]]:
    d = today_utc().isoformat()
    return [
        {
            "title": "Governance of Tool-Using Language Model Agents: Accountability and Audit Requirements",
            "author_body": "Demo Authors",
            "year": today_utc().year,
            "publication_date": d,
            "detected_date": d,
            "source_name": "Offline Demo Journal",
            "source_adapter": "crossref",
            "source_type": "Journal article",
            "status": "Demo only",
            "source_url": "https://example.org/demo-agent-governance",
            "doi": "10.0000/demo.agent.2026",
            "abstract": "A governance framework for autonomous LLM agents, tool use, accountability, transparency, audit and privacy.",
            "authority": 88,
            "region_scope": "Research / global",
        },
        {
            "title": "AI Companion Design and Emotional Dependence: A Longitudinal Evaluation",
            "author_body": "Demo Authors",
            "year": today_utc().year,
            "publication_date": d,
            "detected_date": d,
            "source_name": "Offline Demo",
            "source_adapter": "arxiv",
            "source_type": "Preprint / research",
            "status": "Demo only",
            "source_url": "https://example.org/demo-companion",
            "doi": "",
            "abstract": "Longitudinal study of emotional dependence, anthropomorphism and manipulation in AI companion systems.",
            "authority": 76,
            "region_scope": "Research / global",
        },
    ]

def dedupe_candidates(candidates: list[dict[str, Any]], library: list[dict[str, Any]],
                      discovered: list[dict[str, Any]]) -> list[dict[str, Any]]:
    existing = {candidate_key(x) for x in library + discovered}
    best: dict[str, dict[str, Any]] = {}
    for item in candidates:
        k = candidate_key(item)
        if not k or k in existing:
            continue
        old = best.get(k)
        if old is None or int(item.get("importance_score", 0)) > int(old.get("importance_score", 0)):
            best[k] = item
    return list(best.values())

def enrich_for_library(item: dict[str, Any], new_id: int) -> dict[str, Any]:
    abstract = clean_text(item.get("abstract", ""))
    covers = abstract[:650] if abstract else "Newly detected source; open the linked source for full details."
    why = item.get("why_flagged", "")
    return {
        "id": new_id,
        "theme": item.get("theme", "Foundations"),
        "subtopic": item.get("theme", "AI ethics"),
        "title": item.get("title", ""),
        "author_body": item.get("author_body") or item.get("source_name", ""),
        "year": item.get("year") or today_utc().year,
        "source_type": item.get("source_type", ""),
        "status": "Auto-promoted official update — verify before formal citation",
        "what_it_covers": covers,
        "why_reader_benefits": f"Automatically flagged because it {why}.",
        "best_use": "Current-awareness reading; policy / literature update",
        "relevance_5": 5 if int(item.get("importance_score", 0)) >= 90 else 4,
        "read_priority": "Must Read" if int(item.get("importance_score", 0)) >= 95 else "Strong",
        "region_scope": item.get("region_scope", ""),
        "key_ethical_question": f"What does this new {item.get('theme','AI ethics')} development change for responsible AI practice or governance?",
        "source_url": item.get("source_url", ""),
        "notes_status": "Automatically promoted by strict official-source rule; human verification recommended.",
        "reader_mode": "Review first",
        "importance_score": item.get("importance_score", 0),
        "detected_date": item.get("detected_date", ""),
        "publication_date": item.get("publication_date", ""),
        "source_adapter": item.get("source_adapter", ""),
        "decision": "Curated",
        "doi": item.get("doi", ""),
        "abstract": item.get("abstract", ""),
    }

def is_policy_record(r: dict[str, Any]) -> bool:
    txt = f"{r.get('theme','')} {r.get('source_type','')} {r.get('status','')}".lower()
    return any(k in txt for k in ["policy", "regulation", "governance", "official", "human rights", "transparency", "accountability"])

def write_headers(ws, row: int, headers: list[str], fmt) -> None:
    for c, h in enumerate(headers):
        ws.write(row, c, h, fmt)

def write_url_or_text(ws, row: int, col: int, value: Any, fmt, url_fmt) -> None:
    s = "" if value is None else str(value)
    if s.startswith("http://") or s.startswith("https://"):
        try:
            ws.write_url(row, col, s, url_fmt, string=s)
            return
        except Exception:
            pass
    ws.write(row, col, value, fmt)

def build_formats(wb: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "title": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#15324A",
                                "font_size": 16, "align": "left", "valign": "vcenter"}),
        "subtitle": wb.add_format({"font_color": "#44546A", "bg_color": "#F7F9FB",
                                   "text_wrap": True, "valign": "vcenter"}),
        "header": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78",
                                 "border": 1, "border_color": "#D9E2F3", "align": "center",
                                 "valign": "vcenter", "text_wrap": True}),
        "body": wb.add_format({"border": 1, "border_color": "#E6EAF0", "valign": "top",
                               "text_wrap": True}),
        "url": wb.add_format({"font_color": "#0563C1", "underline": 1, "border": 1,
                              "border_color": "#E6EAF0", "valign": "top", "text_wrap": True}),
        "kpi_label": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#4472C4",
                                    "align": "center", "border": 1}),
        "kpi_value": wb.add_format({"bold": True, "font_size": 15, "bg_color": "#E2F0D9",
                                    "align": "center", "border": 1}),
        "section": wb.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1}),
        "small": wb.add_format({"font_size": 9, "font_color": "#666666", "text_wrap": True}),
    }

def sheet_title(ws, title: str, subtitle: str, cols: int, fmt: dict[str, Any]) -> None:
    ws.merge_range(0, 0, 0, cols - 1, title, fmt["title"])
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, cols - 1, subtitle, fmt["subtitle"])
    ws.set_row(1, 36)

def write_core_library(ws, rows: list[dict[str, Any]], fmt: dict[str, Any], title: str, subtitle: str) -> None:
    cols = len(CORE_LIBRARY_HEADERS)
    sheet_title(ws, title, subtitle, cols, fmt)
    write_headers(ws, 3, [h for _, h in CORE_LIBRARY_HEADERS], fmt["header"])
    for rr, item in enumerate(rows, start=4):
        for cc, (key, _) in enumerate(CORE_LIBRARY_HEADERS):
            value = item.get(key, "")
            write_url_or_text(ws, rr, cc, value, fmt["body"], fmt["url"])
    ws.freeze_panes(4, 0)
    ws.autofilter(3, 0, max(3, 3 + len(rows)), cols - 1)
    widths = [7,24,26,44,27,9,22,28,45,45,34,14,16,20,42,42,34,16]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

def write_discovery_sheet(ws, rows: list[dict[str, Any]], fmt: dict[str, Any], title: str, subtitle: str) -> None:
    cols = len(DISCOVERY_HEADERS)
    sheet_title(ws, title, subtitle, cols, fmt)
    write_headers(ws, 3, [h for _, h in DISCOVERY_HEADERS], fmt["header"])
    for rr, item in enumerate(rows, start=4):
        for cc, (key, _) in enumerate(DISCOVERY_HEADERS):
            write_url_or_text(ws, rr, cc, item.get(key, ""), fmt["body"], fmt["url"])
    if not rows:
        ws.write(4, 2, "No items currently in this view.", fmt["body"])
    ws.freeze_panes(4, 0)
    ws.autofilter(3, 0, max(4, 3 + len(rows)), cols - 1)
    ws.conditional_format(4, 6, max(4, 3 + len(rows)), 6, {
        "type": "3_color_scale", "min_color": "#FEE2E2", "mid_color": "#FEF3C7", "max_color": "#DCFCE7"
    })
    widths = [13,13,46,24,22,24,14,16,46,45,22,16,34]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

def write_start_here(ws, library: list[dict[str, Any]], fmt: dict[str, Any]) -> None:
    sheet_title(ws, "Start Here — 2026-aware AI & LLM Ethics Reading Path",
                "Use this sheet for a compact orientation. The automated discovery tabs are intentionally separate from the stable reading path.",
                8, fmt)
    total = len(library)
    y26 = sum(1 for r in library if int(r.get("year") or 0) == 2026)
    must = sum(1 for r in library if r.get("read_priority") == "Must Read")
    official = sum(1 for r in library if "official" in str(r.get("status","")).lower() or "institution" in str(r.get("status","")).lower())
    labels = ["Curated sources","2026 sources","Must Read","Official / institutional"]
    values = [total,y26,must,official]
    for c, x in enumerate(labels):
        ws.write(3,c,x,fmt["kpi_label"])
        ws.write(4,c,values[c],fmt["kpi_value"])
    ws.write(6,0,"Recommended 10-source reading path",fmt["section"])
    heads = ["Step","Read this","Year","Why now","Theme","Priority","Source URL"]
    write_headers(ws,7,heads,fmt["header"])
    ranked = sorted(library, key=lambda r: (
        0 if r.get("read_priority") == "Must Read" else 1,
        -int(r.get("year") or 0),
        -int(r.get("relevance_5") or 0)
    ))[:10]
    for i,r in enumerate(ranked, start=8):
        why = "Foundation / anchor source" if int(r.get("year") or 0) < 2025 else "Current 2025–2026 evidence or governance update"
        vals = [i-7,r.get("title",""),r.get("year",""),why,r.get("theme",""),r.get("read_priority",""),r.get("source_url","")]
        for c,v in enumerate(vals):
            write_url_or_text(ws,i,c,v,fmt["body"],fmt["url"])
    ws.set_column(0,0,8); ws.set_column(1,1,50); ws.set_column(2,2,9)
    ws.set_column(3,3,32); ws.set_column(4,4,25); ws.set_column(5,5,16); ws.set_column(6,6,45)

def write_topic_map(ws, library: list[dict[str, Any]], fmt: dict[str, Any]) -> None:
    sheet_title(ws, "Topic Map — What to Read for Each AI / LLM Ethics Question",
                "Theme-level map generated from the current curated library.", 6, fmt)
    heads = ["Theme","Sources in library","2026-heavy?","Start with","Then read","Typical reader question"]
    write_headers(ws,3,heads,fmt["header"])
    by_theme = defaultdict(list)
    for r in library:
        by_theme[r.get("theme","Other")].append(r)
    row=4
    for theme, items in sorted(by_theme.items()):
        ordered = sorted(items,key=lambda r:(0 if r.get("read_priority")=="Must Read" else 1,-int(r.get("year") or 0)))
        y26=sum(1 for x in items if int(x.get("year") or 0)==2026)
        heavy="Yes" if y26 >= max(1,len(items)/2) else "Mixed" if y26 else "No"
        start=ordered[0].get("title","") if ordered else ""
        then="; ".join(x.get("title","") for x in ordered[1:3])
        question = ordered[0].get("key_ethical_question","") if ordered else ""
        for c,v in enumerate([theme,len(items),heavy,start,then,question]):
            ws.write(row,c,v,fmt["body"])
        row+=1
    ws.set_column(0,0,28); ws.set_column(1,2,16); ws.set_column(3,4,48); ws.set_column(5,5,52)
    ws.freeze_panes(4,0)

def write_dashboard(workbook, ws, library: list[dict[str, Any]], discovered: list[dict[str, Any]], new_week: list[dict[str, Any]],
                    logs: list[dict[str, Any]], fmt: dict[str, Any]) -> None:
    sheet_title(ws, "AI & LLM Ethics — Living Literature Dashboard",
                "The stable curated library plus an automatically refreshed discovery layer from scholarly APIs and official policy pages.",
                8, fmt)
    metrics = [
        ("Curated", len(library)),
        ("2026 curated", sum(1 for r in library if int(r.get("year") or 0)==2026)),
        ("New this week", len(new_week)),
        ("Pending review", sum(1 for r in discovered if r.get("decision") in ("Pending","Watch"))),
    ]
    for c,(lab,val) in enumerate(metrics):
        ws.write(3,c,lab,fmt["kpi_label"]); ws.write(4,c,val,fmt["kpi_value"])
    ws.write(6,0,"Curated sources by theme",fmt["section"])
    theme_counts=Counter(r.get("theme","Other") for r in library)
    write_headers(ws,7,["Theme","Count"],fmt["header"])
    for rr,(theme,count) in enumerate(sorted(theme_counts.items(),key=lambda x:(-x[1],x[0])),start=8):
        ws.write(rr,0,theme,fmt["body"]); ws.write(rr,1,count,fmt["body"])
    if theme_counts:
        chart=workbook.add_chart({"type":"bar"})
        end=7+len(theme_counts)
        chart.add_series({
            "name":"Curated count",
            "categories":["Living Dashboard",8,0,end,0],
            "values":["Living Dashboard",8,1,end,1],
        })
        chart.set_title({"name":"Curated sources by theme"})
        chart.set_legend({"none":True})
        ws.insert_chart("D7",chart,{"x_scale":1.3,"y_scale":1.2})
    ws.write(6,7,"Latest run",fmt["section"])
    latest=logs[-1] if logs else {}
    latest_rows=[
        ("Run time",latest.get("run_time","")),
        ("Mode",latest.get("mode","")),
        ("Sources checked",latest.get("sources_checked","")),
        ("Candidates fetched",latest.get("candidates_fetched","")),
        ("New after dedupe",latest.get("new_after_dedupe","")),
        ("Errors / notes",latest.get("notes","")),
    ]
    for i,(k,v) in enumerate(latest_rows,start=7):
        ws.write(i,7,k,fmt["header"]); ws.write(i,8,v,fmt["body"])
    ws.set_column(0,0,28); ws.set_column(1,1,12); ws.set_column(2,6,14)
    ws.set_column(7,7,22); ws.set_column(8,8,48)

def write_source_config(ws, config: dict[str, Any], fmt: dict[str, Any]) -> None:
    sheet_title(ws, "Source Configuration", "Edit config/sources.yaml in the repository to change the automation. This sheet is a readable mirror.", 7, fmt)
    heads=["Source","Mode","Endpoint / URL","Enabled","Authority","Purpose","Notes"]
    write_headers(ws,3,heads,fmt["header"])
    rows=[
        ["arXiv","API","https://export.arxiv.org/api/query","YES",config["source_authority"]["arxiv"],"Fresh preprints","3-second delay between repeated calls"],
        ["OpenAlex","API","https://api.openalex.org/works","YES",config["source_authority"]["openalex"],"Broad scholarly discovery","OPENALEX_API_KEY optional"],
        ["Crossref","API","https://api.crossref.org/works","YES",config["source_authority"]["crossref"],"Published DOI metadata","CONTACT_EMAIL recommended"],
    ]
    for p in config.get("policy_pages",[]):
        rows.append([p["name"],"Official page watch",p["url"],"YES",p["authority"],"Policy/news link discovery","First scan bootstraps without flooding"])
    for rr,row in enumerate(rows,start=4):
        for c,v in enumerate(row):
            write_url_or_text(ws,rr,c,v,fmt["body"],fmt["url"])
    ws.set_column(0,1,24); ws.set_column(2,2,52); ws.set_column(3,4,12); ws.set_column(5,6,38)

def write_update_log(ws, logs: list[dict[str, Any]], fmt: dict[str, Any]) -> None:
    sheet_title(ws, "Automation Update Log", "Every run records discovery and deduplication counts.", 9, fmt)
    heads=["Run time","Mode","Sources checked","Candidates fetched","New after dedupe","New This Week","Review Queue","Auto-promoted","Errors / notes"]
    write_headers(ws,3,heads,fmt["header"])
    for rr,log in enumerate(logs[-100:],start=4):
        vals=[
            log.get("run_time",""),log.get("mode",""),log.get("sources_checked",0),log.get("candidates_fetched",0),
            log.get("new_after_dedupe",0),log.get("new_this_week",0),log.get("review_queue",0),
            log.get("auto_promoted",0),log.get("notes",""),
        ]
        for c,v in enumerate(vals): ws.write(rr,c,v,fmt["body"])
    ws.set_column(0,1,22); ws.set_column(2,7,16); ws.set_column(8,8,54)
    ws.freeze_panes(4,0)

def write_guide(ws, fmt: dict[str, Any]) -> None:
    sheet_title(ws, "Automation Guide", "Minimal setup: put this package in a GitHub repository and enable Actions.", 6, fmt)
    heads=["Step","Action","What happens","Required?","Where","Notes"]
    write_headers(ws,3,heads,fmt["header"])
    rows=[
        [1,"Create GitHub repo","Stores code, data and the latest XLSX","YES","GitHub","Private repo is fine"],
        [2,"Upload this package","Keeps folder structure intact","YES","Repository root","Include .github/workflows"],
        [3,"Enable Actions","Allows scheduled and manual runs","YES","Repo → Actions","Default Monday 08:00 Asia/Kolkata"],
        [4,"Add OPENALEX_API_KEY","Higher OpenAlex daily budget","NO","Actions secret","Keyless light use is supported"],
        [5,"Add CONTACT_EMAIL","Polite identification for APIs","Recommended","Actions secret","Useful for Crossref and User-Agent"],
        [6,"Review New This Week","See newly detected high-value items","Reader action","Excel","Review Queue protects curated quality"],
    ]
    for rr,row in enumerate(rows,start=4):
        for c,v in enumerate(row): ws.write(rr,c,v,fmt["body"])
    ws.set_column(0,0,8); ws.set_column(1,1,28); ws.set_column(2,2,42); ws.set_column(3,3,15); ws.set_column(4,4,34); ws.set_column(5,5,44)

def build_workbook(config: dict[str, Any], library: list[dict[str, Any]], discovered: list[dict[str, Any]],
                   logs: list[dict[str, Any]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    wb=xlsxwriter.Workbook(output)
    wb.set_properties({
        "title":"AI & LLM Ethics Living Literature Tracker",
        "subject":"Automatically updated reading guide for AI / LLM ethics",
        "author":"Living Literature Tracker",
    })
    fmt=build_formats(wb)

    today=today_utc()
    week_days=int(config["project"]["new_this_week_days"])
    week_cutoff=today-timedelta(days=week_days)
    new_week=[x for x in discovered if (parse_date(x.get("detected_date")) or date.min) >= week_cutoff and int(x.get("importance_score",0)) >= int(config["project"]["minimum_review_score"])]
    review=[x for x in discovered if x.get("decision") in ("Pending","Watch")]

    ws=wb.add_worksheet("Living Dashboard"); write_dashboard(wb,ws,library,discovered,new_week,logs,fmt)
    ws=wb.add_worksheet("Start Here"); write_start_here(ws,library,fmt)
    ws=wb.add_worksheet("2026 Reading Library"); write_core_library(ws,library,fmt,"Ethics of Artificial Intelligence & LLMs — Curated Reading Library","Stable curated sources. Automatically discovered items are triaged separately before promotion.")
    ws=wb.add_worksheet("New This Week"); write_discovery_sheet(ws,sorted(new_week,key=lambda x:(-int(x.get("importance_score",0)),x.get("publication_date",""))),fmt,"New This Week — Automatically Discovered","Recent high-scoring discoveries from scholarly APIs and official policy pages.")
    ws=wb.add_worksheet("Review Queue"); write_discovery_sheet(ws,sorted(review,key=lambda x:-int(x.get("importance_score",0))),fmt,"Review Queue — Human-in-the-loop Curation","Promote only sources that are genuinely useful. Automatic discovery is intentionally not identical to automatic curation.")
    ws=wb.add_worksheet("Topic Map"); write_topic_map(ws,library,fmt)
    ws=wb.add_worksheet("2026 Only"); write_core_library(ws,[r for r in library if int(r.get("year") or 0)==2026],fmt,"2026-Only View","Current-year curated items.")
    ws=wb.add_worksheet("Policy & Regulation"); write_core_library(ws,[r for r in library if is_policy_record(r)],fmt,"Policy, Governance & Regulation","Curated governance, standards, transparency, law and human-rights sources.")
    ws=wb.add_worksheet("Source Config"); write_source_config(ws,config,fmt)
    ws=wb.add_worksheet("Update Log"); write_update_log(ws,logs,fmt)
    ws=wb.add_worksheet("Automation Guide"); write_guide(ws,fmt)
    wb.close()

def main() -> int:
    ap=argparse.ArgumentParser()
    ap.add_argument("--offline-demo",action="store_true",help="Use local fixture candidates and do not make network calls.")
    ap.add_argument("--rebuild-only",action="store_true",help="Rebuild XLSX from saved data without fetching.")
    args=ap.parse_args()

    config=load_config()
    SESSION.headers.update(make_session_headers(config))
    library=load_json(LIBRARY_PATH,[])
    discovered=load_json(DISCOVERED_PATH,[])
    state=load_json(STATE_PATH,{"policy_bootstrapped":False,"policy_seen":[]})
    logs=load_json(LOG_PATH,[])

    fetched=[]
    errors=[]
    sources_checked=0
    since=today_utc()-timedelta(days=int(config["project"]["lookback_days"]))

    if args.offline_demo:
        fetched=offline_fixture()
        sources_checked=1
        mode="offline-demo"
    elif args.rebuild_only:
        mode="rebuild-only"
    else:
        mode="scheduled/manual"
        for query in config.get("queries",[]):
            for name,fn in [("arXiv",fetch_arxiv),("OpenAlex",fetch_openalex),("Crossref",fetch_crossref)]:
                try:
                    fetched.extend(fn(query,config,since))
                    sources_checked+=1
                except Exception as exc:
                    errors.append(f"{name}/{query['name']}: {type(exc).__name__}: {exc}")
                if name=="arXiv":
                    time.sleep(float(config["runtime"].get("arxiv_delay_seconds",3)))

        policy_all=[]
        for page in config.get("policy_pages",[]):
            try:
                policy_all.extend(watch_policy_page(page,config))
                sources_checked+=1
            except Exception as exc:
                errors.append(f"Policy/{page['name']}: {type(exc).__name__}: {exc}")

        current_seen={canonical_url(x) for x in state.get("policy_seen",[]) if x}
        policy_urls={canonical_url(x.get("source_url","")) for x in policy_all if x.get("source_url")}
        if not state.get("policy_bootstrapped",False) and not config["project"].get("emit_on_first_policy_scan",False):
            state["policy_bootstrapped"]=True
            state["policy_seen"]=sorted(policy_urls)
        else:
            new_policy=[x for x in policy_all if canonical_url(x.get("source_url","")) not in current_seen]
            fetched.extend(new_policy)
            state["policy_bootstrapped"]=True
            state["policy_seen"]=sorted(current_seen | policy_urls)

    scored=[score_candidate(x,config) for x in fetched]
    # Reject records with no meaningful topic signal or below threshold.
    min_score=int(config["project"]["minimum_review_score"])
    relevant=[x for x in scored if x.get("matched_keywords") and int(x.get("importance_score",0)) >= min_score]
    new_items=dedupe_candidates(relevant,library,discovered)

    next_id=max([int(r.get("id") or 0) for r in library] or [0])+1
    auto_promoted=0
    for item in new_items:
        official=item.get("source_adapter")=="official"
        auto_rule=bool(config["project"].get("auto_promote_official",True)) and official and int(item.get("importance_score",0)) >= int(config["project"]["auto_promote_official_score"])
        if auto_rule:
            item["decision"]="Auto-promoted"
            library.append(enrich_for_library(item,next_id))
            next_id+=1
            auto_promoted+=1
        else:
            item["decision"]="Pending"
        item.setdefault("notes_status","Automatically discovered; verify bibliographic details before formal citation.")
        discovered.append(item)

    # Preserve newest first in discovery store while keeping a deterministic order.
    discovered.sort(key=lambda x:(x.get("detected_date",""),int(x.get("importance_score",0))),reverse=True)

    week_cutoff=today_utc()-timedelta(days=int(config["project"]["new_this_week_days"]))
    new_week=[x for x in discovered if (parse_date(x.get("detected_date")) or date.min) >= week_cutoff and int(x.get("importance_score",0)) >= min_score]
    review_queue=[x for x in discovered if x.get("decision") in ("Pending","Watch")]

    log={
        "run_time":now_iso(),
        "mode":mode,
        "sources_checked":sources_checked,
        "candidates_fetched":len(fetched),
        "new_after_dedupe":len(new_items),
        "new_this_week":len(new_week),
        "review_queue":len(review_queue),
        "auto_promoted":auto_promoted,
        "notes":"; ".join(errors[:8]) if errors else "Completed without recorded source errors",
    }
    logs.append(log)

    save_json(LIBRARY_PATH,library)
    save_json(DISCOVERED_PATH,discovered)
    save_json(STATE_PATH,state)
    save_json(LOG_PATH,logs)

    output=ROOT/config["project"]["output_xlsx"]
    build_workbook(config,library,discovered,logs,output)

    print(json.dumps(log,indent=2))
    print(f"Workbook: {output}")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
