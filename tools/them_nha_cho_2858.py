# -*- coding: utf-8 -*-
"""
Chèn 51 nhà chờ của công văn 2858 vào mảng RAW trong index.html.

    pip install openpyxl
    python tools/build_raw.py            # dựng lại 610 trụ dừng (GHI ĐÈ cả mảng RAW)
    python tools/them_nha_cho_2858.py    # rồi mới chèn 51 nhà chờ vào

Thứ tự đó là bắt buộc: `build_raw.py` ghi đè nguyên mảng RAW nên chạy sau sẽ xóa mất
nhà chờ. Script này chạy lại nhiều lần không nhân đôi (bỏ qua mã đã có trong RAW).

NGUỒN: 2858 KTHT DS gửi TTHTKT.xlsx — 39 nhà chờ cần chiếu sáng + 19 nhà chờ tăng
mảng xanh, trừ 7 vị trí có ở cả hai danh sách = 51 vị trí. Phần đọc và tách hai
danh sách dùng lại nguyên của tools/doi_chieu_2858.py.

STT: 610 vị trí cũ GIỮ NGUYÊN số, nhà chờ nhận STT 611 trở đi. Dữ liệu khảo sát
trong localStorage và trên Google Sheet đều khóa theo STT — đánh số lại là mất sạch
dữ liệu hiện trường. Vị trí mới vẫn được chèn đúng chỗ sắp xếp (địa bàn → phường)
để danh sách không lặp tiêu đề nhóm phường, nên STT không còn tăng dần theo thứ tự
hiển thị.
"""
import io
import json
import os
import re

import openpyxl

from build_raw import DIABAN
from doi_chieu_2858 import CONGVAN, ch, doc_19, doc_39, doc_tru_dung, nph, quan_cua

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(ROOT, "index.html")

# Công văn thiếu phường ở 4 vị trí. Điền theo suy đoán rồi ghi rõ vào `ghichu` để
# cán bộ thấy khi mở điểm — thà nói là chưa chắc còn hơn im lặng cho qua.
#   - Điện Biên Phủ: hai vị trí ĐBP còn lại trong cùng danh sách đều là Xuân Hòa.
#   - Hàm Nghi (Cục Hải quan Thành phố): thuộc Quận 1, phường Sài Gòn.
SUY_DOAN = {
    "Điện Biên Phủ": ("Quận 3", "Xuân Hòa"),
    "Hàm Nghi": ("Quận 1", "Sài Gòn"),
}
GHI_CHU_SUY_DOAN = "Phường suy đoán từ công văn 2858, cần xác nhận"
LOAI_MAC_DINH = "Nhà chờ"


def doc_51():
    """Gộp hai danh sách của công văn, khử 7 mã trùng."""
    wb = openpyxl.load_workbook(CONGVAN, data_only=True)
    ds39, tat_ca = doc_39(wb)
    ds19 = doc_19(wb, tat_ca)
    ma39 = set(x["ma"] for x in ds39 if x["ma"])
    return ds39 + [x for x in ds19 if not (x["ma"] and x["ma"] in ma39)]


def thanh_vi_tri(x):
    """Đổi một dòng công văn thành phần tử của mảng RAW (chưa có STT)."""
    duong = x["duong"]
    diaban = quan_cua(x["ma"])
    phuong = x.get("phuong") or ""
    ghichu = ""
    if not phuong or not diaban:
        db, ph = SUY_DOAN.get(duong, ("", ""))
        diaban = diaban or db
        phuong = phuong or ph
        ghichu = GHI_CHU_SUY_DOAN
    sonha = x.get("sonha") or ""
    return {
        "stt": 0,
        "ma": x["ma"],
        "diaban": diaban,
        "ten": duong + (" – " + sonha if sonha else ""),
        "duong": duong,
        "sonha": sonha,
        "phuong": phuong,
        "loai": x.get("loai") or LOAI_MAC_DINH,
        "ghichu": ghichu,
    }


def main():
    raw = doc_tru_dung()
    da_co = set(r.get("ma", "") for r in raw if r.get("ma"))

    moi = []
    for x in doc_51():
        if x["ma"] and x["ma"] in da_co:
            continue                        # chạy lại lần hai, đừng nhân đôi
        moi.append(thanh_vi_tri(x))

    if not moi:
        print("Mảng RAW đã có đủ nhà chờ của công văn 2858 —", len(raw), "vị trí. Không đổi gì.")
        return

    # Nhà chờ nhận STT nối tiếp; vị trí cũ không được đụng vào.
    stt = max(r["stt"] for r in raw) if raw else 0
    for x in moi:
        stt += 1
        x["stt"] = stt

    # Bổ sung khóa `ma` cho vị trí cũ (mảng RAW dựng trước khi có cột này).
    for r in raw:
        r.setdefault("ma", "")

    # Chèn vào đúng chỗ sắp xếp để danh sách không lặp tiêu đề nhóm phường. Sắp
    # bằng phép so chuỗi thuần như build_raw.py, đừng dùng locale — đổi cách so là
    # thứ tự 610 vị trí cũ xáo lại hết.
    order = list(DIABAN.values())
    def khoa(r):
        return (order.index(r["diaban"]) if r["diaban"] in order else len(order),
                r["phuong"], r["stt"])
    raw = sorted(raw + moi, key=khoa)

    with io.open(HTML, encoding="utf-8") as f:
        src = f.read()
    line = "const RAW = " + json.dumps(raw, ensure_ascii=False) + ";"
    new, n = re.subn(r"(?m)^const RAW = \[.*\];$", lambda m: line, src, count=1)
    assert n == 1, "Không tìm thấy dòng 'const RAW = [...];' trong index.html"
    with io.open(HTML, "w", encoding="utf-8", newline="") as f:
        f.write(new)

    co_ma = len([x for x in moi if x["ma"]])
    canh_bao = [x for x in moi if x["ghichu"]]
    print("Đã thêm    :", len(moi), "nhà chờ (STT", moi[0]["stt"], "→", moi[-1]["stt"], ")")
    print("Có mã trạm :", co_ma, "/", len(moi))
    print("Tổng danh mục:", len(raw), "vị trí")
    if canh_bao:
        print("Phường suy đoán, cần xác nhận:")
        for x in canh_bao:
            print("   STT %-4d %-9s %-16s %s / %s" %
                  (x["stt"], x["ma"] or "(không mã)", x["duong"], x["diaban"], x["phuong"]))
    print("Đã ghi", HTML)


if __name__ == "__main__":
    main()
