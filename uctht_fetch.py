#!/usr/bin/env python3
"""mcsd.mn-ээс "Үнэт цаасны төвлөрлийн судалгаа"-ны хувьцааны PDF-үүдийг татна.

    python uctht_fetch.py                    # татаагүй бүх тайлан -> pdfs/YYYY-MM-DD.pdf
    python uctht_fetch.py --since 2026-06-30 # зөвхөн энэ огнооноос хойших
    python uctht_fetch.py --latest | --list
"""
import argparse
import html
import os
import re
import sys
from urllib.parse import urljoin

import requests

API = "https://mcsd.mn/graphql"
SITE = "https://mcsd.mn/"
CATEGORY_ID = "d2985eee-dbcc-49a6-8f7a-338c9153880e"  # "Үнэт цаасны төвлөрлийн судалгаа"
QUERY = """
query get_contents($category_id: String, $status: STATUS_FILTER, $page: Int, $limit: Int) {
  get_contents(category_id: $category_id, status: $status, page: $page, limit: $limit) {
    success message count
    contents { id title body createdAt }
  }
}
"""
HEADERS = {"Content-Type": "application/json", "Origin": "https://www.mcsd.mn",
           "Referer": "https://www.mcsd.mn/", "User-Agent": "Mozilla/5.0"}


def fetch_contents(limit, page=1):
    payload = {"operationName": "get_contents",
               "variables": {"category_id": CATEGORY_ID, "status": "active", "page": page, "limit": limit},
               "query": QUERY}
    r = requests.post(API, json=payload, headers=HEADERS, timeout=30)
    r.raise_for_status()
    j = r.json()
    if "errors" in j:
        raise RuntimeError(f"GraphQL алдаа: {j['errors']}")
    res = j["data"]["get_contents"]
    if not res.get("success"):
        raise RuntimeError(f"API амжилтгүй: {res.get('message')}")
    return res["contents"]


def _strip(s):
    return html.unescape(re.sub(r"<[^>]+>", " ", s))


def equity_pdf_link(body):
    """body (HTML) доторх "Хувьцааны судалгаа [харах]" -ын PDF линк (сангийн линкийг алгасна)."""
    anchors = list(re.finditer(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', body, re.S | re.I))
    pdfs = [m for m in anchors if ".pdf" in m.group(1).lower()] or anchors
    for m in pdfs:
        ctx = _strip(body[max(0, m.start() - 160):m.start()]) + _strip(m.group(2))
        ctx = ctx.strip().splitlines()[-1] if ctx.strip() else ""
        if re.search(r"Хувьцаа", ctx, re.I) and not re.search(r"Хөрөнгө", ctx, re.I):
            return urljoin(SITE, m.group(1))
    return urljoin(SITE, pdfs[0].group(1)) if pdfs else None


def report_date(title):
    m = re.search(r"(\d{4})[./-](\d{2})[./-](\d{2})", title or "")
    return "-".join(m.groups()) if m else None


def fetch_new(out="pdfs", limit=12, since=None, latest=False, list_only=False, log=print):
    """Шинэ PDF-үүдийг татаж, [(YYYY-MM-DD, зам)] буцаана."""
    items = fetch_contents(limit)
    items.sort(key=lambda it: report_date(it["title"]) or "", reverse=True)
    if latest:
        items = items[:1]
    os.makedirs(out, exist_ok=True)
    new = []
    for it in items:
        d = report_date(it["title"])
        url = equity_pdf_link(it.get("body") or "")
        log(f"{d or '????-??-??'} | {it['title']} | {url}")
        if list_only:
            continue
        if not d or not url:
            log("  алгаслаа: огноо эсвэл PDF линк олдсонгүй (body-г шалгана уу)")
            continue
        if since and d <= since:
            continue
        path = os.path.join(out, f"{d}.pdf")
        if os.path.exists(path):
            continue
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=120)
        r.raise_for_status()
        if not r.content.startswith(b"%PDF"):
            log("  PDF биш хариу ирлээ, алгаслаа")
            continue
        with open(path, "wb") as f:
            f.write(r.content)
        log(f"  татлаа -> {path} ({len(r.content) // 1024} KB)")
        new.append((d, path))
    return sorted(new)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="pdfs")
    ap.add_argument("--limit", type=int, default=12)
    ap.add_argument("--since", help="YYYY-MM-DD: энэ огнооноос хойших тайланг л татна")
    ap.add_argument("--latest", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    new = fetch_new(a.out, a.limit, a.since, a.latest, a.list)
    if not a.list:
        print("ШИНЭ: " + (", ".join(p for _, p in new) if new else "байхгүй"))
        sys.exit(10 if new else 0)


if __name__ == "__main__":
    main()
