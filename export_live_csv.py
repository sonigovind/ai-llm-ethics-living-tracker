#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
from datetime import date, timedelta
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
LIVE = ROOT / "live"
LIVE.mkdir(exist_ok=True)


def load_json(name: str, default: Any):
    path = DATA / name
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def parse_date(value):
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except Exception:
        return None


def write_csv(name: str, rows: list[dict], columns: list[tuple[str, str]]):
    path = LIVE / name
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([label for _, label in columns])
        for row in rows:
            w.writerow([row.get(key, "") for key, _ in columns])
    print(f"wrote {path} ({len(rows)} rows)")


library = load_json("library.json", [])
discovered = load_json("discovered.json", [])
logs = load_json("update_log.json", [])

library_cols = [
    ("id", "ID"), ("theme", "Theme"), ("subtopic", "Subtopic"),
    ("title", "Title"), ("author_body", "Author / Body"), ("year", "Year"),
    ("source_type", "Source Type"), ("status", "Status"),
    ("what_it_covers", "What it covers"),
    ("why_reader_benefits", "Why the reader benefits"),
    ("best_use", "Best use"), ("relevance_5", "2026 Relevance Score"),
    ("read_priority", "Read Priority"), ("region_scope", "Region / Scope"),
    ("key_ethical_question", "Key ethical question"), ("source_url", "Source URL"),
    ("notes_status", "Notes / status"), ("reader_mode", "Reader mode"),
]

discovery_cols = [
    ("detected_date", "Detected"), ("publication_date", "Published"),
    ("title", "Title"), ("theme", "Theme"), ("source_name", "Source"),
    ("source_type", "Source Type"), ("importance_score", "Importance Score"),
    ("read_priority", "Priority"), ("why_flagged", "Why flagged"),
    ("source_url", "URL"), ("doi", "DOI"), ("decision", "Decision"),
    ("notes_status", "Notes"),
]

# Stable views.
curated = sorted(library, key=lambda r: (-(int(r.get("year") or 0)), str(r.get("theme", "")), str(r.get("title", ""))))
must_read = [r for r in library if str(r.get("read_priority", "")) == "Must Read"]
must_read.sort(key=lambda r: (-(int(r.get("year") or 0)), -(int(r.get("relevance_5") or 0)), str(r.get("title", ""))))
only_2026 = [r for r in library if int(r.get("year") or 0) == 2026]
only_2026.sort(key=lambda r: (str(r.get("theme", "")), str(r.get("title", ""))))

policy_words = ("policy", "regulation", "governance", "official", "human rights", "transparency", "accountability", "standard")
def is_policy(r):
    text = " ".join(str(r.get(k, "")) for k in ("theme", "source_type", "status", "subtopic")).lower()
    return any(w in text for w in policy_words)
policy = [r for r in library if is_policy(r)]
policy.sort(key=lambda r: (-(int(r.get("year") or 0)), str(r.get("title", ""))))

# Recent discovery view.
cutoff = date.today() - timedelta(days=7)
new_week = []
for r in discovered:
    d = parse_date(r.get("detected_date"))
    score = int(r.get("importance_score") or 0)
    if d and d >= cutoff and score >= 70:
        new_week.append(r)
new_week.sort(key=lambda r: (-int(r.get("importance_score") or 0), str(r.get("publication_date", ""))), reverse=False)

write_csv("curated_library.csv", curated, library_cols)
write_csv("must_read.csv", must_read, library_cols)
write_csv("2026_only.csv", only_2026, library_cols)
write_csv("policy_regulation.csv", policy, library_cols)
write_csv("new_this_week.csv", new_week, discovery_cols)

log_cols = [
    ("run_time", "Run time"), ("mode", "Mode"), ("sources_checked", "Sources checked"),
    ("candidates_fetched", "Candidates fetched"), ("new_after_dedupe", "New after dedupe"),
    ("new_this_week", "New This Week"), ("review_queue", "Review Queue"),
    ("auto_promoted", "Auto-promoted"), ("notes", "Errors / notes"),
]
write_csv("update_log.csv", logs[-100:], log_cols)

stats_cols = [("metric", "Metric"), ("value", "Value")]
stats = [
    {"metric": "Curated sources", "value": len(library)},
    {"metric": "2026 sources", "value": len(only_2026)},
    {"metric": "Must Read", "value": len(must_read)},
    {"metric": "New this week", "value": len(new_week)},
    {"metric": "Last run", "value": logs[-1].get("run_time", "") if logs else ""},
]
write_csv("dashboard_stats.csv", stats, stats_cols)
