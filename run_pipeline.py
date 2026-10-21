#!/usr/bin/env python3
"""Бүтэн урсгал: шинэ PDF татах -> Excel шинэчлэх -> инфографик (PNG) үүсгэх.

    python run_pipeline.py                 # GitHub Actions болон гараар ажиллуулахад
    python run_pipeline.py --no-fetch      # pdfs/ доторх шинэ PDF-үүдийг л боловсруулна
"""
import argparse
import os
import re
import sys

from openpyxl import load_workbook

import uctht_infographic as info
from pdf_parser import parse_pdf
from update_workbook import add_report, add_summary_column, latest_date, update_intro


def log(*a):
    print(*a, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default="data/uctht_concentration.xlsx")
    ap.add_argument("--pdf-dir", default="pdfs")
    ap.add_argument("--out", default="output")
    ap.add_argument("--no-fetch", action="store_true")
    ap.add_argument("--scale", type=float, default=3)
    a = ap.parse_args()

    wb = load_workbook(a.xlsx)
    last = latest_date(wb)
    last_s = "20%02d-%02d-%02d" % last
    log(f"Excel-ийн сүүлийн тайлан: {last_s}")

    if not a.no_fetch:
        import uctht_fetch
        uctht_fetch.fetch_new(a.pdf_dir, since=last_s, log=log)

    todo = sorted(f for f in os.listdir(a.pdf_dir) if re.fullmatch(r"\d{4}-\d{2}-\d{2}\.pdf", f)
                  and f[:10] > last_s)
    if not todo:
        log("Шинэ тайлан алга.")
        _gh_output(new="")
        return 0

    added = []
    for f in todo:
        d = tuple(int(x) for x in (f[2:4], f[5:7], f[8:10]))
        rows, pinfo = parse_pdf(os.path.join(a.pdf_dir, f))
        for w in pinfo["warnings"]:
            log(f"  АНХААР [{f}]: {w}")
        if pinfo["warnings"] and any("таараагүй" in w or "цөөн" in w for w in pinfo["warnings"]):
            sys.exit(f"{f}: задлалт баталгаажсангүй, зогсов (Excel өөрчлөгдсөнгүй)")
        ws, new_codes = add_report(wb, d, rows)
        add_summary_column(wb, d, rows)
        if new_codes:
            log(f"  Шинэ ҮЦ код: {new_codes} (нэрийг PDF-ээс авсан, шалгана уу)")
        added.append(d)
        log(f"Нэмлээ: {ws.title} ({len(rows)} мөр)")
    update_intro(wb)
    wb.save(a.xlsx)

    os.makedirs(a.out, exist_ok=True)
    for d in added:
        stem = os.path.join(a.out, "infographic_20%02d-%02d-%02d" % d)
        out, cur, base = info.generate(a.xlsx, stem, current="%02d-%02d-%02d" % d, scale=a.scale)
        log(f"Инфографик: {out}.png  ({info.fmt(base)} → {info.fmt(cur)})")
    _gh_output(new=",".join("20%02d-%02d-%02d" % d for d in added))
    return 0


def _gh_output(**kv):
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a") as f:
            for k, v in kv.items():
                f.write(f"{k}={v}\n")


if __name__ == "__main__":
    sys.exit(main())
