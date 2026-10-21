#!/usr/bin/env python3
"""Онцгой өөрчлөлтүүдийн инфографик (HTML + өндөр нягтралтай PNG).

    python uctht_infographic.py data/uctht_concentration.xlsx
    python uctht_infographic.py data/uctht_concentration.xlsx --current 26-12-31 --base 26-09-30
    python uctht_infographic.py data/uctht_concentration.xlsx --scale 4 --no-png

Харьцуулалт: хамгийн сүүлийн тайлан улирлын сүүл (3,6,9,12) бол өмнөх улирлын сүүлтэй,
эс бөгөөс өмнөх сарын тайлантай. PNG: Playwright (Chromium), өгөгдмөл 3x (3240px өргөн).
"""
import argparse
import html
import re
import sys

from openpyxl import load_workbook

BRAND_TOP = "Big Short Ganbayar"          # "Powered by ..." (баруун дээд)
BRAND_BOTTOM = ("openmind", "MSE")         # "MSE" ногоон (баруун доод)
TOP_FF, TOP_HOLDERS, MIN_FF_PP = 5, 7, 0.01
GREEN, AMBER = "#3DDC84", "#FFB020"


# ------------------------------------------------------------------ өгөгдөл
def sheet_date(name):
    m = re.fullmatch(r"Тайлан (\d\d)-(\d\d)-(\d\d)", name)
    return tuple(int(x) for x in m.groups()) if m else None


def load_sheet(ws):
    data = {}
    for r in range(5, ws.max_row + 1):
        code = ws.cell(r, 2).value
        if code is None:
            continue
        v = lambda c: ws.cell(r, c).value or 0
        data[code] = dict(name=str(ws.cell(r, 3).value).replace("\n", " "), total=v(4), und=v(5),
                          bigp=v(9), bigh=v(10), smn=v(11), smp=v(12), smh=v(13))
    return data


def pick_sheets(wb, current=None, base=None):
    dates = {sheet_date(n): n for n in wb.sheetnames if sheet_date(n)}
    if not dates:
        sys.exit("'Тайлан YY-MM-DD' хуудас олдсонгүй")
    parse = lambda s: tuple(int(x) for x in s.split("-"))
    cur = parse(current) if current else max(dates)
    if cur not in dates:
        sys.exit(f"Хуудас олдсонгүй: {current}")
    if base:
        b = parse(base)
    else:
        earlier = [d for d in dates if d < cur]
        if not earlier:
            sys.exit("Харьцуулах өмнөх тайлан алга")
        if cur[1] in (3, 6, 9, 12):
            q = [d for d in earlier if d[1] in (3, 6, 9, 12) and d[:2] != cur[:2]]
            b = max(q) if q else max(earlier)
        else:
            b = max(earlier)
    if b not in dates:
        sys.exit(f"Хуудас олдсонгүй: {base}")
    mid = sorted(d for d in dates if b < d < cur)
    return cur, b, dates, mid


def fmt(d, sep="/"):
    return f"20{d[0]:02d}{sep}{d[1]:02d}{sep}{d[2]:02d}"


def period_label(cur, base):
    if cur[1] in (3, 6, 9, 12) and (cur[0] * 12 + cur[1]) - (base[0] * 12 + base[1]) == 3:
        return f"20{cur[0]:02d} оны {(cur[1] - 1) // 3 + 1}-р улирал"
    return f"20{cur[0]:02d} оны {cur[1]}-р сар"


# ------------------------------------------------------------------ тооцоо
def analyse(cur, base, wb, dates, mid):
    C, B = load_sheet(wb[dates[cur]]), load_sheet(wb[dates[base]])
    common = [k for k in C if k in B]
    res = {"n": len(C), "hold_cur": sum(c["smh"] for c in C.values()),
           "hold_base": sum(b["smh"] for b in B.values())}
    res["big_changed"] = sum(1 for k in common if C[k]["bigh"] != B[k]["bigh"])
    ff = [dict(code=k, name=C[k]["name"], a=B[k]["smp"], b=C[k]["smp"], d=C[k]["smp"] - B[k]["smp"],
               h0=B[k]["bigh"], h1=C[k]["bigh"]) for k in common if abs(C[k]["smp"] - B[k]["smp"]) >= MIN_FF_PP]
    res["ff"] = sorted(ff, key=lambda x: -abs(x["d"]))[:TOP_FF]
    ch = [dict(code=k, name=C[k]["name"], a=B[k]["smh"], b=C[k]["smh"],
               p=(C[k]["smh"] - B[k]["smh"]) / B[k]["smh"] * 100) for k in common if B[k]["smh"] > 0]
    res["up"] = [x for x in sorted(ch, key=lambda x: -x["p"]) if x["p"] > 0][:TOP_HOLDERS]
    res["down"] = [x for x in sorted(ch, key=lambda x: x["p"]) if x["p"] < 0][:TOP_HOLDERS]
    # онцгой байдлууд
    al = []
    for k in common:
        a, b = B[k]["smp"], C[k]["smp"]
        if a >= 0.05 > b:
            al.append((0, f'{k} {C[k]["name"]}: олон нийтийн хувь {a:.2%} → {b:.2%}, 5%-иас доош орсон'))
        dh = C[k]["bigh"] - B[k]["bigh"]
        if abs(dh) >= 2:
            al.append((1, f'{k} {C[k]["name"]}: 5%+ эзэмшигч {B[k]["bigh"]} → {C[k]["bigh"]} '
                          f'({B[k]["bigp"]:.2%} → {C[k]["bigp"]:.2%})'))
    for k in C:
        if k not in B:
            al.append((2, f'{k} {C[k]["name"]}: шинээр бүртгэгдсэн'))
    for k in B:
        if k not in C:
            al.append((2, f'{k} {B[k]["name"]}: жагсаалтаас хасагдсан'))
    res["alerts"] = [t for _, t in sorted(al)][:6]
    mids = {d: load_sheet(wb[dates[d]]) for d in mid}
    issued = []
    for k in common:
        path = [B[k]["total"]] + [mids[d][k]["total"] for d in mid if k in mids[d]] + [C[k]["total"]]
        if len(set(path)) > 1:
            issued.append(dict(code=k, name=C[k]["name"], path=path))
    res["issued"] = issued[:3]
    return res


# ------------------------------------------------------------------ HTML
CSS = """
*{box-sizing:border-box;-webkit-font-smoothing:antialiased;text-rendering:geometricPrecision}
body{margin:0;background:#050B14;display:flex;justify-content:center}
.page{position:relative;width:1080px;padding:56px 56px 112px;background:#0C1829;color:#fff;
font-family:Manrope,'Noto Sans','DejaVu Sans',sans-serif;display:flex;flex-direction:column;gap:36px}
.r{display:flex;align-items:center;gap:20px;padding:14px 0;border-bottom:1px solid #2A3E5C}
.nm{width:250px;font-weight:800;font-size:20px;line-height:1.2}
.nm small{display:block;font-weight:600;font-size:16px;color:#C3D0E2;margin-top:4px}
.bw{display:flex;align-items:center;gap:10px;height:28px;margin:3px 0}
.bw i{display:block;height:20px;border-radius:3px}.bw b{font-size:17px;font-weight:800;color:#E4ECF7}
.chip{width:108px;text-align:center;font-weight:800;font-size:20px;border:2.5px solid;border-radius:22px;padding:6px 0}
.h2{font-size:30px;font-weight:800;margin:0}.sub{font-size:17px;color:#C3D0E2;margin:8px 0 12px;line-height:1.4}
.st{flex:1;background:#14253D;border-radius:14px;padding:22px}
.st b{display:block;font-size:48px;font-weight:800;color:#3DDC84}.st span{font-size:17px;color:#D5DFEC;line-height:1.4}
.col{flex:1;background:#101F35;border-radius:16px;padding:22px 22px 12px;border-top:6px solid}
.col .nm{width:196px;font-size:17px}.col .nm small{font-size:15px}.col .r{padding:11px 0;gap:12px}
.note{background:#14253D;border-radius:14px;padding:24px 26px;border-left:6px solid #FFB020}
.note div{font-size:18px;color:#E4ECF7;line-height:1.5}
"""


def e(s):
    return html.escape(str(s))


def ff_rows(items):
    mx = max([max(x["a"], x["b"]) for x in items] + [0.01])
    w = lambda p: max(p / mx * 300, 3)
    out = ""
    for x in items:
        col = GREEN if x["d"] > 0 else AMBER
        out += (f'<div class="r"><div class="nm">{x["code"]} {e(x["name"])}<small>5%+ эзэмшигч {x["h0"]} → {x["h1"]}</small></div>'
                f'<div style="flex:1"><div class="bw"><i style="width:{w(x["a"]):.0f}px;background:#5B6E8A"></i><b style="color:#C3D0E2">{x["a"]:.2%}</b></div>'
                f'<div class="bw"><i style="width:{w(x["b"]):.0f}px;background:{col}"></i><b>{x["b"]:.2%}</b></div></div>'
                f'<div class="chip" style="color:{col};border-color:{col}">{x["d"] * 100:+.1f} пп</div></div>\n')
    return out


def holder_col(title, sub, items, color, grad, sign, mx):
    rows = "".join(
        f'<div class="r"><div class="nm">{x["code"]} {e(x["name"])}<small>{x["a"]:,} → {x["b"]:,}</small></div>'
        f'<div class="bw" style="flex:1"><i style="width:{max(abs(x["p"]) / mx * 120, 4):.0f}px;background:{grad}"></i>'
        f'<b>{sign}{abs(x["p"]):.1f}%</b></div></div>\n' for x in items) or '<div class="sub">Өөрчлөлт байхгүй</div>'
    return (f'<div class="col" style="border-top-color:{color}"><h2 class="h2" style="font-size:26px;color:{color}">{title}</h2>'
            f'<div class="sub" style="font-size:15px;min-height:42px">{sub}</div>\n{rows}</div>')


def build_html(res, cur, base):
    d_cur, d_base = fmt(cur), fmt(base)
    chg = (res["hold_cur"] / res["hold_base"] - 1) * 100 if res["hold_base"] else 0
    label = period_label(cur, base)
    mx = max([abs(x["p"]) for x in res["up"] + res["down"]] + [1])
    down_sub = f"5%-аас бага эзэмшигчдийн тоо, {fmt(base)[2:]} → {fmt(cur)[2:]}"
    small = [x["name"] for x in res["down"] + res["up"] if x["a"] < 30]
    if small:
        down_sub += f". {e(', '.join(small))}-д эзэмшигч цөөн (30-аас бага) тул хувь нь огцом харагдана."
    alerts = ""
    if res["alerts"]:
        alerts = ('<div class="note" style="border-left-color:#3DDC84"><div style="font-size:22px;font-weight:800;color:#3DDC84;margin-bottom:6px">Анхаарах өөрчлөлтүүд</div>'
                  + "".join(f'<div style="margin-top:6px">• {e(t)}</div>' for t in res["alerts"]) + "</div>")
    issued = ""
    if res["issued"]:
        lines = "".join(f'<div style="margin-top:8px"><b>{x["code"]} {e(x["name"])}</b>: ' +
                        " → ".join(f"{p:,}" for p in x["path"]) + "</div>" for x in res["issued"])
        issued = ('<div class="note"><div style="font-size:22px;font-weight:800;color:#FFB020">Шалгууштай өгөгдөл: нийт гаргасан ҮЦ-ны тоо өөрчлөгдсөн</div>'
                  f'<div>Эх тайлангаас баталгаажуулах нь зүйтэй.{lines}</div></div>')
    return f"""<!doctype html>
<html lang="mn"><head><meta charset="utf-8"><title>ҮЦ төвлөрлийн өөрчлөлт {e(label)}</title>
<meta name="viewport" content="width=1080">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;800&display=swap">
<style>{CSS}</style></head><body>
<div class="page">
<div style="position:absolute;top:40px;right:56px;font-size:17px;color:#D5DFEC;font-weight:600">Powered by <span style="color:#fff;font-weight:800">{e(BRAND_TOP)}</span></div>
<div style="padding-top:36px;max-width:760px">
<div style="font-size:18px;font-weight:800;letter-spacing:3px;color:#3DDC84">ҮЦТХТ · ТӨВЛӨРЛИЙН СУДАЛГАА · {e(label.upper())}</div>
<h1 style="margin:12px 0 10px;font-size:62px;line-height:1.05;font-weight:800">Хэн хувьцаагаа авч, хэн зарав?</h1>
<div style="font-size:20px;color:#D5DFEC">{res["n"]} хувьцаа · {d_base} → {d_cur} өөрчлөлт</div></div>
<div style="display:flex;gap:16px">
<div class="st"><b>{res["hold_cur"]:,}</b><span>5%-аас бага эзэмшигч, {d_cur}. {d_base}-нд {res["hold_base"]:,} байсан</span></div>
<div class="st"><b>{chg:+.1f}%</b><span>Эзэмшигчийн нийт тооны өөрчлөлт</span></div>
<div class="st"><b>{res["big_changed"]}</b><span>Компанид 5%-аас их эзэмшигчийн тоо өөрчлөгдсөн</span></div></div>
<div><h2 class="h2">Олон нийтийн эзэмшил (free float) өөрчлөгдсөн нь</h2>
<div class="sub">Дээд мөр {fmt(base)[2:]}, доод мөр {fmt(cur)[2:]}. Ногоон нь олон нийтийн хувь өссөн, шар нь буурсан.</div>
{ff_rows(res["ff"])}</div>
<div style="display:flex;gap:20px;align-items:stretch">
{holder_col("Хамгийн их өссөн", f"5%-аас бага эзэмшигчдийн тоо, {fmt(base)[2:]} → {fmt(cur)[2:]}", res["up"], GREEN, "linear-gradient(90deg,#16A34A,#3DDC84)", "+", mx)}
{holder_col("Хамгийн их буурсан", down_sub, res["down"], AMBER, AMBER, "−", mx)}
</div>
{alerts}{issued}
<div style="font-size:15px;color:#B4C2D6;max-width:680px;line-height:1.4">Эх сурвалж: Үнэт цаасны төвлөрсөн хадгаламжийн төв (ҮЦТХТ). Хувь нь 5%-аас бага эзэмшигчдийн эзлэх хувь, пп = хувийн пункт.</div>
<div style="position:absolute;right:56px;bottom:36px;font-size:30px;font-weight:800;letter-spacing:.5px">{e(BRAND_BOTTOM[0])}<span style="color:#3DDC84">{e(BRAND_BOTTOM[1])}</span></div>
</div></body></html>"""


def render_png(html_text, png_path, scale=3):
    """Chromium-аар 1080px өргөнийг scale дахин нягтруулж зураг авна (3 → 3240px)."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(args=["--font-render-hinting=none"])
        pg = b.new_page(viewport={"width": 1080, "height": 1400}, device_scale_factor=scale)
        pg.set_content(html_text, wait_until="load")
        try:
            pg.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        pg.evaluate("document.fonts.ready")
        pg.locator(".page").screenshot(path=png_path, type="png")
        b.close()


def generate(xlsx, out_base=None, current=None, base=None, scale=3, png=True):
    wb = load_workbook(xlsx, data_only=True)
    cur, b, dates, mid = pick_sheets(wb, current, base)
    res = analyse(cur, b, wb, dates, mid)
    doc = build_html(res, cur, b)
    out_base = out_base or f"infographic_{fmt(cur, '-')}"
    with open(out_base + ".html", "w", encoding="utf-8") as f:
        f.write(doc)
    if png:
        render_png(doc, out_base + ".png", scale)
    return out_base, cur, b


def main():
    ap = argparse.ArgumentParser(description="ҮЦТХТ онцгой өөрчлөлтийн инфографик")
    ap.add_argument("xlsx")
    ap.add_argument("-o", "--out", help="гаралтын нэр (өргөтгөлгүй)")
    ap.add_argument("--current"), ap.add_argument("--base")
    ap.add_argument("--scale", type=float, default=3)
    ap.add_argument("--no-png", action="store_true")
    a = ap.parse_args()
    out, cur, b = generate(a.xlsx, a.out, a.current, a.base, a.scale, not a.no_png)
    print(f"{fmt(b)} → {fmt(cur)} | {out}.html" + ("" if a.no_png else f", {out}.png"))


if __name__ == "__main__":
    main()
