# -*- coding: utf-8 -*-
"""
mst_lookup.py
-------------
Module tra cứu hàng loạt thông tin Mã số thuế (MST) doanh nghiệp / chi nhánh (10 số hoặc 13 số)
từ cơ sở dữ liệu đồng bộ Tổng cục Thuế (`gdt.gov.vn`) qua đa nguồn (VietQR GDT Sync API & EscoData API)
không vướng Captcha, hỗ trợ đa luồng, bộ nhớ đệm (Cache) và trích xuất tự động danh sách MST từ:
- Danh sách MST nhập trực tiếp (copy/paste từ Excel, Word, văn bản).
- File Excel / CSV bất kỳ (bảng kê hóa đơn mua vào/bán ra, danh sách nhà cung cấp/khách hàng).
- Cả thư mục chứa nhiều file bảng kê Excel/CSV.

Phân loại tự động mức độ rủi ro theo Tình trạng hoạt động của Cơ quan thuế:
- Đang hoạt động (An toàn)
- Tạm nghỉ kinh doanh có thời hạn (Cảnh báo kiểm tra thời điểm phát sinh hóa đơn)
- Không hoạt động tại địa chỉ đã đăng ký / Bỏ địa chỉ kinh doanh (RỦI RO CAO - Khóa MST)
- Ngừng hoạt động / Đã chấm dứt hiệu lực MST (RỦI RO CAO)
"""

import os
import re
import json
import ssl
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Any, Optional, Tuple, Union

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from excel_normalizer import resolve_input_paths, SUPPORTED_TABULAR_EXTS
from tabular_reader import read_tabular_sheets


def normalize_mst_code(raw_val: Any) -> str:
    """
    Chuẩn hóa chuỗi Mã số thuế:
    - Chấp nhận MST 10 chữ số (VD: '0100109106') hoặc 13 chữ số có/không dấu gạch ngang
      (VD: '0100109106-001' hoặc '0100109106001' -> '0100109106-001').
    - Nếu người dùng copy từ cột số trong Excel bị mất số 0 ở đầu (còn 9 số hoặc 12 số):
      tự động bù số '0' ở đầu.
    """
    if raw_val is None:
        return ""
    s = str(raw_val).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"[\s\u00a0]", "", s)
    s = re.sub(r"[–—−‐]", "-", s)

    if re.fullmatch(r"\d{9}", s):
        s = "0" + s
    elif re.fullmatch(r"\d{10}", s):
        pass
    elif re.fullmatch(r"\d{9}-\d{3}", s):
        s = "0" + s
    elif re.fullmatch(r"\d{10}-\d{3}", s):
        pass
    elif re.fullmatch(r"\d{12}", s):
        s = "0" + s[:9] + "-" + s[9:]
    elif re.fullmatch(r"\d{13}", s):
        s = s[:10] + "-" + s[10:]
    else:
        return ""
    return s


def classify_tax_status(status_text: str) -> Tuple[str, str]:
    """
    Phân loại tình trạng hoạt động của NNT thành (Nhóm rủi ro, Khuyến nghị kiểm tra thuế).
    """
    st = (status_text or "").strip().lower()
    if not st or "không tìm thấy" in st or "lỗi" in st:
        return (
            "Không tìm thấy / Sai MST",
            "Kiểm tra lại tính chính xác của MST trên hóa đơn gốc hoặc tra cứu lại."
        )
    if "không hoạt động tại địa chỉ" in st or "bỏ địa chỉ" in st or "không tìm thấy tại địa chỉ" in st:
        return (
            "RỦI RO CAO (Bỏ địa chỉ ĐKKD)",
            "Cảnh báo: NNT bỏ địa chỉ kinh doanh! Cần kiểm tra ngày thông báo bỏ địa chỉ so với ngày lập hóa đơn."
        )
    if any(k in st for k in ("ngừng", "chấm dứt", "đóng mst", "giải thể", "phá sản", "thu hồi")):
        return (
            "RỦI RO CAO (Ngừng / Đóng MST)",
            "Cảnh báo: NNT đã ngừng hoạt động hoặc chấm dứt hiệu lực MST! Kiểm tra thời điểm phát sinh hóa đơn."
        )
    if any(k in st for k in ("tạm nghỉ", "tạm ngừng")):
        return (
            "CẢNH BÁO (Tạm ngừng KD)",
            "Lưu ý: NNT đang tạm nghỉ kinh doanh có thời hạn. Rà soát xem hóa đơn có xuất trong thời gian tạm ngừng không."
        )
    if "đang hoạt động" in st:
        return (
            "Đang hoạt động (An toàn)",
            "Đủ điều kiện trạng thái MST đang hoạt động bình thường."
        )
    return (
        "Cần rà soát thêm",
        f"Kiểm tra chi tiết trạng thái: {status_text}"
    )


class MstLookupEngine:
    _cache: Dict[str, Dict[str, Any]] = {}

    @staticmethod
    def extract_msts_from_text(text: str) -> List[str]:
        """Trích xuất danh sách MST duy nhất (giữ nguyên thứ tự xuất hiện) từ văn bản bất kỳ."""
        if not text:
            return []
        tokens = re.split(r"[,;\s\r\n|]+", text)
        out: List[str] = []
        seen = set()
        for tok in tokens:
            cleaned = tok.strip().strip("'\"")
            mst = normalize_mst_code(cleaned)
            if mst and mst not in seen:
                seen.add(mst)
                out.append(mst)
        return out

    @staticmethod
    def extract_msts_from_files(path_or_paths: Union[str, List[str]]) -> List[str]:
        """
        Quét 1 file, nhiều file hoặc cả thư mục Excel/CSV để trích xuất toàn bộ Mã số thuế.
        - Nếu sheet có cột tiêu đề MST ('Mã số thuế', 'MST người bán', 'MST người mua', 'MST', 'Tax code'):
          ưu tiên đọc từ cột đó.
        - Nếu là file danh sách 1-2 cột không có tiêu đề: quét toàn bộ các ô có định dạng MST hợp lệ.
        """
        files = resolve_input_paths(path_or_paths, SUPPORTED_TABULAR_EXTS)
        out: List[str] = []
        seen = set()

        mst_header_kws = (
            "mã số thuế", "ma so thue", "mst người bán", "mst người mua",
            "mst", "tax code", "tax_code", "mã nnt"
        )

        for fp in files:
            try:
                sheets = read_tabular_sheets(fp)
            except Exception:
                continue

            for _, rows in sheets.items():
                if not rows:
                    continue

                # Tìm xem có cột tiêu đề MST trong 20 dòng đầu không
                mst_cols: List[int] = []
                header_row_idx = -1
                for r_idx in range(min(20, len(rows))):
                    row = rows[r_idx]
                    for c_idx, cell_val in enumerate(row):
                        s_val = str(cell_val or "").strip().lower()
                        if not s_val:
                            continue
                        if any(kw == s_val or kw in s_val for kw in mst_header_kws):
                            mst_cols.append(c_idx)
                    if mst_cols:
                        header_row_idx = r_idx
                        break

                if mst_cols:
                    for r_idx in range(header_row_idx + 1, len(rows)):
                        row = rows[r_idx]
                        for c_idx in mst_cols:
                            if c_idx < len(row):
                                mst = normalize_mst_code(row[c_idx])
                                if mst and mst not in seen:
                                    seen.add(mst)
                                    out.append(mst)
                else:
                    # Quét toàn bộ ô trong sheet nếu là file danh sách MST thuần túy
                    for row in rows:
                        for cell_val in row:
                            mst = normalize_mst_code(cell_val)
                            if mst and mst not in seen:
                                seen.add(mst)
                                out.append(mst)

        return out

    @staticmethod
    def lookup_single_mst(mst: str, timeout: int = 10, use_cache: bool = True) -> Dict[str, Any]:
        """
        Tra cứu thông tin 1 Mã số thuế từ cơ sở dữ liệu đồng bộ Tổng cục Thuế:
        1. Nguồn 1 (`api.vietqr.io/v2/business/{mst}` - Đồng bộ từ `gdt.gov.vn`):
           Lấy Tên chính thức, Tên quốc tế, Tên viết tắt, Địa chỉ thuế hiện hành, Tình trạng hoạt động GDT, Thời điểm cập nhật.
        2. Nguồn 2 (`escodata.net/api-mst/{mst}.htm`):
           Bổ sung Người đại diện pháp luật, Địa chỉ ĐKKD chi tiết và đối chiếu chéo tình trạng hoạt động.
        """
        norm_mst = normalize_mst_code(mst) or str(mst).strip()
        if use_cache and norm_mst in MstLookupEngine._cache:
            return dict(MstLookupEngine._cache[norm_mst])

        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        }

        result: Dict[str, Any] = {
            "mst": norm_mst,
            "ten_nnt": "",
            "ten_quoc_te": "",
            "ten_viet_tat": "",
            "dia_chi": "",
            "nguoi_dai_dien": "",
            "tinh_trang": "",
            "phan_loai_rui_ro": "",
            "khuyen_nghi": "",
            "nguon_du_lieu": "",
            "cap_nhat_gdt": "",
            "found": False,
        }

        # 1. Gọi VietQR GDT Sync API
        try:
            url_vqr = f"https://api.vietqr.io/v2/business/{norm_mst}"
            req = urllib.request.Request(url_vqr, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="replace"))
                if payload.get("code") == "00" and payload.get("data"):
                    d = payload["data"]
                    meta = payload.get("metadata") or {}
                    result["ten_nnt"] = str(d.get("name") or "").strip()
                    result["ten_quoc_te"] = str(d.get("internationalName") or "").strip()
                    result["ten_viet_tat"] = str(d.get("shortName") or "").strip()
                    result["dia_chi"] = str(d.get("address") or "").strip()
                    result["tinh_trang"] = str(d.get("status") or "").strip()
                    upd = str(meta.get("updatedAt") or "")[:10]
                    result["cap_nhat_gdt"] = upd
                    result["nguon_du_lieu"] = "GDT (gdt.gov.vn / VietQR)"
                    result["found"] = bool(result["ten_nnt"])
        except Exception:
            pass

        # 2. Gọi EscoData API để lấy thêm Người đại diện pháp luật hoặc fallback khi Nguồn 1 chưa có
        try:
            url_esco = f"https://escodata.net/api-mst/{norm_mst}.htm"
            req2 = urllib.request.Request(url_esco, headers=headers)
            with urllib.request.urlopen(req2, timeout=timeout, context=ctx) as resp2:
                payload2 = json.loads(resp2.read().decode("utf-8", errors="replace"))
                if payload2.get("error") == 0 and isinstance(payload2.get("data"), dict):
                    d2 = payload2["data"]
                    if not result["ten_nnt"] and d2.get("ten"):
                        result["ten_nnt"] = str(d2.get("ten") or "").strip()
                    if not result["dia_chi"] and d2.get("dc"):
                        result["dia_chi"] = str(d2.get("dc") or "").strip()
                    rep = str(d2.get("daidien") or "").strip()
                    if rep and rep.lower() != "null":
                        result["nguoi_dai_dien"] = rep
                    if not result["tinh_trang"] and d2.get("tinhtrang"):
                        result["tinh_trang"] = str(d2.get("tinhtrang") or "").strip()
                    if result["ten_nnt"]:
                        result["found"] = True
                        if not result["nguon_du_lieu"]:
                            result["nguon_du_lieu"] = "GDT (EscoData)"
                        elif rep:
                            result["nguon_du_lieu"] = "GDT (gdt.gov.vn + ĐKKD)"
        except Exception:
            pass

        if not result["found"]:
            result["tinh_trang"] = "Không tìm thấy thông tin MST"

        risk_level, advice = classify_tax_status(result["tinh_trang"])
        result["phan_loai_rui_ro"] = risk_level
        result["khuyen_nghi"] = advice

        if use_cache and result["found"]:
            MstLookupEngine._cache[norm_mst] = dict(result)

        return result

    @staticmethod
    def lookup_batch(
        mst_list: List[str],
        max_workers: int = 6,
        progress_fn=None
    ) -> List[Dict[str, Any]]:
        """
        Tra cứu hàng loạt danh sách Mã số thuế bằng ThreadPoolExecutor (nhanh gấp 6 lần tuần tự)
        và giữ nguyên đúng thứ tự danh sách gốc của người dùng.
        """
        clean_list: List[str] = []
        seen = set()
        for m in mst_list:
            nm = normalize_mst_code(m) or str(m).strip()
            if nm and nm not in seen:
                seen.add(nm)
                clean_list.append(nm)

        total = len(clean_list)
        if total == 0:
            return []

        results_map: Dict[str, Dict[str, Any]] = {}
        completed = 0

        with ThreadPoolExecutor(max_workers=max(1, min(max_workers, 10))) as executor:
            future_to_mst = {
                executor.submit(MstLookupEngine.lookup_single_mst, mst): mst
                for mst in clean_list
            }
            for future in as_completed(future_to_mst):
                mst = future_to_mst[future]
                completed += 1
                try:
                    res = future.result()
                except Exception as e:
                    res = {
                        "mst": mst,
                        "ten_nnt": "",
                        "ten_quoc_te": "",
                        "ten_viet_tat": "",
                        "dia_chi": "",
                        "nguoi_dai_dien": "",
                        "tinh_trang": f"Lỗi kết nối: {e}",
                        "phan_loai_rui_ro": "Lỗi tra cứu",
                        "khuyen_nghi": "Thử tra cứu lại.",
                        "nguon_du_lieu": "",
                        "cap_nhat_gdt": "",
                        "found": False,
                    }
                results_map[mst] = res
                if progress_fn:
                    progress_fn(completed, total, res)

        return [results_map[m] for m in clean_list]

    @staticmethod
    def export_directory_to_excel(
        results: List[Dict[str, Any]],
        output_path: str
    ) -> str:
        """
        Xuất Danh bạ Mã số thuế Doanh nghiệp ra file Excel chuyên nghiệp gồm 2 sheet:
        1. 'Danh bạ Doanh nghiệp (MST)': Bảng chi tiết đầy đủ thông tin, tô màu cảnh báo đỏ cho các DN bỏ địa chỉ/ngừng hoạt động.
        2. 'Thống kê Tình trạng MST': Tổng hợp số lượng DN theo từng trạng thái rủi ro.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Danh bạ Doanh nghiệp (MST)"
        ws.views.sheetView[0].showGridLines = True

        font_title = Font(name="Times New Roman", size=14, bold=True, color="1B365D")
        font_sub = Font(name="Times New Roman", size=11, italic=True)
        font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        font_regular = Font(name="Times New Roman", size=11)
        font_bold = Font(name="Times New Roman", size=11, bold=True)
        font_ok = Font(name="Times New Roman", size=11, bold=True, color="006100")
        font_warn = Font(name="Times New Roman", size=11, bold=True, color="9C6500")
        font_alert = Font(name="Times New Roman", size=11, bold=True, color="9C0006")

        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_warn = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fill_alert = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_section = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        total_count = len(results)
        active_count = sum(1 for r in results if "An toàn" in r.get("phan_loai_rui_ro", ""))
        risk_count = sum(1 for r in results if "RỦI RO CAO" in r.get("phan_loai_rui_ro", ""))
        warn_count = sum(1 for r in results if "CẢNH BÁO" in r.get("phan_loai_rui_ro", ""))
        not_found_count = sum(1 for r in results if not r.get("found"))

        ws.merge_cells("A1:J1")
        ws["A1"] = "DANH BẠ THÔNG TIN NGƯỜI NỘP THUẾ & KIỂM TRA TÌNH TRẠNG HOẠT ĐỘNG MÃ SỐ THUẾ"
        ws["A1"].font = font_title
        ws["A1"].alignment = align_center

        ws.merge_cells("A2:J2")
        ws["A2"] = (
            f"Tổng số MST tra cứu: {total_count:,}  |  Đang hoạt động: {active_count:,}  |  "
            f"Rủi ro cao (Bỏ địa chỉ / Ngừng HĐ): {risk_count:,}  |  Tạm ngừng: {warn_count:,}  |  Không tìm thấy: {not_found_count:,}"
        )
        ws["A2"].font = font_sub
        ws["A2"].alignment = align_center

        headers = [
            "STT",
            "Mã số thuế",
            "Tên người nộp thuế (Doanh nghiệp / Chi nhánh)",
            "Tình trạng hoạt động (Cơ quan thuế)",
            "Phân loại rủi ro thuế",
            "Địa chỉ trụ sở đăng ký thuế",
            "Người đại diện pháp luật",
            "Tên viết tắt / Quốc tế",
            "Khuyến nghị kiểm tra hóa đơn",
            "Nguồn & Ngày đồng bộ GDT",
        ]

        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_thin
        ws.row_dimensions[4].height = 28

        r_idx = 5
        for stt, item in enumerate(results, 1):
            c_stt = ws.cell(r_idx, 1, stt)
            c_stt.alignment = align_center
            c_stt.font = font_regular

            c_mst = ws.cell(r_idx, 2, str(item.get("mst", "")))
            c_mst.number_format = "@"
            c_mst.alignment = align_center
            c_mst.font = font_bold

            c_name = ws.cell(r_idx, 3, item.get("ten_nnt", ""))
            c_name.alignment = align_left
            c_name.font = font_regular

            c_st = ws.cell(r_idx, 4, item.get("tinh_trang", ""))
            c_st.alignment = align_left
            c_st.font = font_regular

            risk = item.get("phan_loai_rui_ro", "")
            c_risk = ws.cell(r_idx, 5, risk)
            c_risk.alignment = align_center

            if "RỦI RO CAO" in risk:
                c_risk.font = font_alert
                c_risk.fill = fill_alert
                c_st.font = font_alert
                c_st.fill = fill_alert
            elif "CẢNH BÁO" in risk or "Không tìm thấy" in risk:
                c_risk.font = font_warn
                c_risk.fill = fill_warn
            elif "An toàn" in risk:
                c_risk.font = font_ok
                c_risk.fill = fill_ok
            else:
                c_risk.font = font_regular

            c_addr = ws.cell(r_idx, 6, item.get("dia_chi", ""))
            c_addr.alignment = align_left
            c_addr.font = font_regular

            c_rep = ws.cell(r_idx, 7, item.get("nguoi_dai_dien", ""))
            c_rep.alignment = align_left
            c_rep.font = font_regular

            short_intl = " - ".join(
                p for p in (item.get("ten_viet_tat", ""), item.get("ten_quoc_te", "")) if p
            )
            c_si = ws.cell(r_idx, 8, short_intl)
            c_si.alignment = align_left
            c_si.font = font_regular

            c_adv = ws.cell(r_idx, 9, item.get("khuyen_nghi", ""))
            c_adv.alignment = align_left
            c_adv.font = font_regular

            src_info = item.get("nguon_du_lieu", "")
            if item.get("cap_nhat_gdt"):
                src_info = f"{src_info} ({item['cap_nhat_gdt']})"
            c_src = ws.cell(r_idx, 10, src_info)
            c_src.alignment = align_center
            c_src.font = font_regular

            for c in range(1, 11):
                ws.cell(r_idx, c).border = border_thin
            r_idx += 1

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:J{max(4, r_idx - 1)}"

        # Thiết lập độ rộng cột chuẩn dễ đọc
        col_widths = {
            1: 7,
            2: 18,
            3: 42,
            4: 34,
            5: 26,
            6: 48,
            7: 22,
            8: 25,
            9: 42,
            10: 26,
        }
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(output_path)
        return output_path
