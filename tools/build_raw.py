# -*- coding: utf-8 -*-
"""
Sinh mảng RAW (danh mục vị trí khảo sát) từ file Excel nguồn và ghi đè vào index.html.

    pip install openpyxl
    python tools/build_raw.py

Nguồn : docs/DS TRẠM DỪNG XE BUÝT.xlsx
    cột 1 STT (công thức, bỏ qua – script đánh số lại)
    cột 2 Tuyến đường | cột 3 Số nhà | cột 4 Phường/Xã | cột 5 Hạ tầng
    dòng có cột 1 = I/II/III và cột 2 rỗng là dòng phân khu vực.

Đích  : dòng `const RAW = [...];` trong index.html (thay nguyên dòng).
"""
import io
import json
import os
import re
from collections import Counter

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "docs", "DS TRẠM DỪNG XE BUÝT.xlsx")
HTML = os.path.join(ROOT, "index.html")

REGION = {"I": "TP.HCM", "II": "Vũng Tàu", "III": "Bình Dương"}


def cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def norm_loai(s):
    """Thống nhất cách viết: 'Loại' -> 'loại', '4s' -> '4S'."""
    s = re.sub(r"\bLoại\b", "loại", s)
    s = re.sub(r"loại\s*(\d)\s*[sS]\b", lambda m: "loại " + m.group(1) + "S", s)
    return s


def norm_phuong(s):
    return re.sub(r"^(Phường|Xã|Thị trấn)\s+", "", s).strip()


def main():
    ws = openpyxl.load_workbook(XLSX, data_only=True).active

    raw = []
    region = ""
    for r in range(5, ws.max_row + 1):
        a, b = cell(ws.cell(r, 1).value), cell(ws.cell(r, 2).value)
        if not b:
            if a in REGION:
                region = REGION[a]
            continue
        sonha = cell(ws.cell(r, 3).value)
        raw.append({
            "stt": 0,
            "diaban": region,
            "ten": b + (" – " + sonha if sonha else ""),
            "duong": b,
            "sonha": sonha,
            "phuong": norm_phuong(cell(ws.cell(r, 4).value)),
            "loai": norm_loai(cell(ws.cell(r, 5).value)),
            "ghichu": "",
        })

    # Gom liền mạch theo khu vực -> phường/xã (giữ thứ tự gốc trong từng phường), rồi đánh lại STT.
    order = list(REGION.values())
    raw.sort(key=lambda x: (order.index(x["diaban"]), x["phuong"]))
    for i, x in enumerate(raw):
        x["stt"] = i + 1

    print("Số vị trí :", len(raw))
    print("Khu vực   :", dict(Counter(x["diaban"] for x in raw)))
    print("Phường/xã :", len(set(x["phuong"] for x in raw)))

    line = "const RAW = " + json.dumps(raw, ensure_ascii=False) + ";"
    with io.open(HTML, encoding="utf-8") as f:
        src = f.read()
    new, n = re.subn(r"(?m)^const RAW = \[.*\];$", lambda m: line, src, count=1)
    assert n == 1, "Không tìm thấy dòng 'const RAW = [...];' trong index.html"
    with io.open(HTML, "w", encoding="utf-8", newline="") as f:
        f.write(new)
    print("Đã ghi", HTML)


if __name__ == "__main__":
    main()
