"""ҮЦТХТ "Үнэт цаасны төвлөрлийн судалгаа" PDF-ийг задлах (poppler-utils-ийн pdftotext хэрэгтэй)."""
import re
import subprocess

NUM = r"-?[\d,]+(?:\.\d+)?%?"
ROW = re.compile(r"^\s*(\d+)\s+(\d+)\s*(\D.*?)?\s+((?:" + NUM + r"\s+){9}" + NUM + r")\s*$")
SKIP = ("Хувьцаа", "Нийт", "Тайлан", "Хянасан", "ҮНЭТ", "МОНГОЛ УЛС", "Эзэмшигч", "Эзлэх", "ҮЦ-ны", "Зөрүү")


def _num(x):
    return float(x.strip("%").replace(",", ""))


def parse_pdf(path):
    """-> (rows, info). row = dict(code, name, v=[10 тоо]);
    v: [нийт гаргасан, зарласан боловч гаргаагүй, төр ҮЦ, төр %, 5%+ ҮЦ, 5%+ %, 5%+ эзэмшигч,
        5%- ҮЦ, 5%- %, 5%- эзэмшигч]"""
    text = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True, check=True).stdout
    lines = text.split("\n")
    rows, total_line = [], None
    for i, ln in enumerate(lines):
        m = ROW.match(ln)
        if m:
            _, code, name, nums = m.groups()
            name = (name or "").strip()
            nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if nxt and not re.search(r"\d", nxt) and not nxt.startswith(SKIP) and len(nxt) < 40:
                name = f"{name} {nxt}".strip()  # 2 мөрөнд хуваагдсан нэр
            rows.append(dict(code=int(code), name=name, v=[_num(x) for x in nums.split()]))
        elif ln.strip().startswith("Нийт") and re.search(r"\d", ln):
            total_line = ln
    info = {"count": len(rows), "warnings": []}
    if total_line:
        nums = [_num(x) for x in re.findall(NUM, total_line)]
        sum_total = sum(r["v"][0] for r in rows)
        sum_hold = sum(r["v"][9] for r in rows)
        if nums and abs(nums[0] - sum_total) > 1:
            info["warnings"].append(f"Нийт гаргасан ҮЦ таараагүй: PDF {nums[0]:,.0f} vs нийлбэр {sum_total:,.0f}")
        if sum_hold not in nums:
            info["warnings"].append(f"Эзэмшигчийн тоо таараагүй: нийлбэр {sum_hold:,.0f} PDF-ийн Нийт мөрөнд алга")
    else:
        info["warnings"].append("PDF-ийн 'Нийт' мөр олдсонгүй")
    if len(rows) < 100:
        info["warnings"].append(f"Мөрийн тоо сэжигтэй цөөн: {len(rows)}")
    codes = [r["code"] for r in rows]
    if len(set(codes)) != len(codes):
        info["warnings"].append("Давхардсан ҮЦ код байна")
    return rows, info
