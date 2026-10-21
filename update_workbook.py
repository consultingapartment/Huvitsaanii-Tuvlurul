"""Excel нэгтгэлд шинэ тайлан нэмэх: 'Тайлан YY-MM-DD' хуудас + 3 нэгтгэл хуудасны багана."""
import re
from copy import copy

from openpyxl.utils import get_column_letter as L

SUMMARIES = {"Эзэмшигчдийн тоо": 9, "Давамгай хувь": 5, "Чөлөөт эргэлт": 8}  # v[] индекс
REPORT_RE = re.compile(r"Тайлан (\d\d)-(\d\d)-(\d\d)")


def report_sheets(wb):
    out = {}
    for n in wb.sheetnames:
        m = REPORT_RE.fullmatch(n)
        if m:
            out[tuple(int(x) for x in m.groups())] = n
    return out


def latest_date(wb):
    d = report_sheets(wb)
    return max(d) if d else None


def _copy_style(dst, src):
    dst._style = copy(src._style)


def add_report(wb, date, rows):
    """date=(yy,mm,dd); rows=parse_pdf-ийн жагсаалт. Өмнөх хуудсыг загвар болгоно."""
    sheets = report_sheets(wb)
    prev = wb[sheets[max(sheets)]]
    ws = wb.copy_worksheet(prev)
    ws.title = "Тайлан %02d-%02d-%02d" % date
    ws.freeze_panes = "D5"
    ws["A1"] = f"ҮНЭТ ЦААСНЫ ТӨВЛӨРЛИЙН СУДАЛГАА — 20{date[0]:02d}/{date[1]:02d}/{date[2]:02d}"
    last_prev = prev.max_row
    style = {}
    names = {}
    for r in range(5, last_prev + 1):
        code = prev.cell(r, 2).value
        if code is not None:
            style[code] = [copy(prev.cell(r, c)._style) for c in range(1, 14)]
            names[code] = prev.cell(r, 3).value
    default = [copy(prev.cell(last_prev, c)._style) for c in range(1, 14)]
    new_codes = []
    for i, row in enumerate(rows):
        r, code, v = 5 + i, row["code"], row["v"]
        if code not in names:
            new_codes.append(code)
        vals = [i + 1, code, names.get(code, row["name"]), int(v[0]), int(v[1]), int(v[2]), round(v[3] / 100, 4),
                int(v[4]), round(v[5] / 100, 4), int(v[6]), int(v[7]), round(v[8] / 100, 4), int(v[9])]
        for c, val in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell.value = val
            cell._style = copy(style.get(code, default)[c - 1])
    extra = ws.max_row - (4 + len(rows))
    if extra > 0:
        ws.delete_rows(5 + len(rows), extra)
    return ws, new_codes


def _ungroup(ws):
    """Бүлэглэсэн баганы өргөнийг (min..max) тус тусад нь салгана."""
    for key, dim in list(ws.column_dimensions.items()):
        if dim.min and dim.max and dim.min < dim.max:
            w, lo, hi = dim.width, dim.min, dim.max
            dim.max = lo
            for c in range(lo + 1, hi + 1):
                ws.column_dimensions[L(c)].width = w


def _last_data_row(ws):
    r = ws.max_row
    while r > 4 and ws.cell(r, 1).value is None:
        r -= 1
    return r


def add_summary_column(wb, date, rows, names_for_new=None):
    short = "%02d.%02d.%02d" % date
    by_code = {r["code"]: r for r in rows}
    for sheet, k in SUMMARIES.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        _ungroup(ws)
        last = _last_data_row(ws)
        hdr = {c: ws.cell(4, c).value for c in range(1, ws.max_column + 1)}
        chg = next((c for c, v in hdr.items() if isinstance(v, str) and v.startswith("Өөрчлөлт")), None)
        data_cols = [c for c, v in hdr.items() if c >= 3 and v is not None and c != chg]
        new_col = chg if chg else max(data_cols) + 1
        width_prev = ws.column_dimensions[L(new_col - 1)].width
        if chg:  # Өөрчлөлт % баганыг нэг баруун тийш
            chg_w = ws.column_dimensions[L(chg)].width
            for r in range(4, last + 1):
                _copy_style(ws.cell(r, chg + 1), ws.cell(r, chg))
            ws.column_dimensions[L(chg + 1)].width = chg_w
        # шинэ компани: нэгтгэлд байхгүй кодыг хамгийн доор нэмнэ
        known = {ws.cell(r, 1).value for r in range(5, last + 1)}
        for row in rows:
            if row["code"] not in known:
                last += 1
                for c in range(1, ws.max_column + 1):
                    _copy_style(ws.cell(last, c), ws.cell(last - 1, c))
                ws.cell(last, 1).value = row["code"]
                ws.cell(last, 2).value = row["name"]
                if chg:
                    _copy_style(ws.cell(last, chg), ws.cell(last - 1, chg))
        for r in range(4, last + 1):
            _copy_style(ws.cell(r, new_col), ws.cell(r, new_col - 1))
        ws.column_dimensions[L(new_col)].width = width_prev
        ws.cell(4, new_col).value = short
        for r in range(5, last + 1):
            row = by_code.get(ws.cell(r, 1).value)
            ws.cell(r, new_col).value = None if row is None else (
                int(row["v"][k]) if k == 9 else round(row["v"][k] / 100, 4))
        if chg:
            nc = L(new_col)
            ws.cell(4, chg + 1).value = f"Өөрчлөлт %\n({short} vs {_first_date_label(ws)})"
            for r in range(5, last + 1):
                ws.cell(r, chg + 1).value = f'=IF(OR(C{r}="",C{r}=0,{nc}{r}=""),"",({nc}{r}-C{r})/C{r})'
        ws.freeze_panes = "C5"


def _first_date_label(ws):
    return str(ws.cell(4, 3).value)


def update_intro(wb):
    if "Танилцуулга" not in wb.sheetnames:
        return
    t = wb["Танилцуулга"]
    d = sorted(report_sheets(wb))
    a, b = d[0], d[-1]
    rng = "20%02d/%02d/%02d – 20%02d/%02d/%02d" % (a + b)
    t["A3"] = f"Хамрах хугацаа: {rng} ({len(d)} тайлан)"
    t["A10"] = "Тайлан хуудсууд (%s)" % rng.replace(" – ", " ... ")
