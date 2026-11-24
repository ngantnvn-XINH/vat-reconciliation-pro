# -*- coding: utf-8 -*-
"""
excel_normalizer.py
-------------------
Module chuẩn hóa thông minh mọi định dạng file Excel bảng kê hóa đơn / dữ liệu thuế:
- Hỗ trợ chọn 1 file, NHIỀU file cùng lúc, hoặc chọn CẢ THƯ MỤC (quét đệ quy thư mục con).
- Quét toàn bộ các sheet có dữ liệu hóa đơn trong 1 workbook (VD: bảng kê nhiều kỳ/nhiều năm chia theo từng sheet Tháng/Quý/Năm).
- Đọc đa định dạng (.xlsx, .xls BIFF8, HTML-xls, XML-xls, TSV/CSV) thông qua `tabular_reader`.
- Tự động dò dòng tiêu đề (Header auto-detection) kể cả dòng trống, tiêu đề gộp 2-3 tầng.
- Nhận diện cột linh hoạt (Số HĐ, Ký hiệu, Ngày HĐ, MST, Tên đối tác, Doanh số, Thuế suất, Tiền thuế, Diễn giải...).
- Tự động khử trùng lặp khi người dùng nạp chồng các file xuất trùng khoảng thời gian từ TCT.
"""

import os
import re
import datetime
from typing import Dict, List, Any, Optional, Tuple, Union

from tabular_reader import read_tabular_sheets

# Từ khóa nhận diện các trường thông tin chuẩn (chữ thường, không dấu/có dấu)
COLUMN_KEYWORDS = {
    'ky_hieu': [
        'ký hiệu hoá đơn', 'ký hiệu hóa đơn', 'ky hieu hoa don', 'ký hiệu hđ', 'ky hieu hd', 
        'mẫu số - ký hiệu', 'mẫu số', 'ký hiệu', 'serial'
    ],
    'so_hd': [
        'số hoá đơn', 'số hóa đơn', 'so hoa don', 'số hđ', 'so hd', 'shđ', 'shd', 'inv no', 'invoice no', 'số cừ', 'so cu'
    ],
    'ngay_hd': [
        'ngày hóa đơn', 'ngay hoa don', 'ngày hđ', 'ngay hd', 'ngày lập', 'ngay lap', 
        'ngày phát hành', 'ngay phat hanh', 'ngày chứng từ', 'ngay chung tu', 'inv date', 'date'
    ],
    'mst': [
        'mã số thuế người bán', 'ma so thue nguoi ban', 'mã số thuế người mua', 'ma so thue nguoi mua',
        'mã số thuế', 'ma so thue', 'mst người bán', 'mst người mua', 'mst', 'tax code'
    ],
    'ten_doi_tac': [
        'tên người bán', 'ten nguoi ban', 'tên người mua', 'ten nguoi mua', 
        'tên đối tác', 'ten doi tac', 'tên nnt', 'ten nnt', 'tên đơn vị', 'ten don vi',
        'tên khách hàng', 'tên nhà cung cấp', 'tên cty', 'tên công ty', 'vendor', 'customer'
    ],
    'dien_giai': [
        'tên hàng hóa dịch vụ', 'ten hang hoa dich vu', 'tên hàng hóa, dịch vụ',
        'mặt hàng', 'mat hang', 'diễn giải', 'dien giai', 'nội dung', 'noi dung',
        'hàng hóa', 'hang hoa', 'description', 'item'
    ],
    'doanh_so': [
        'giá trị hhdv mua vào chưa có thuế', 'doanh số mua chưa có thuế', 'doanh số bán chưa có thuế',
        'tổng tiền chưa thuế', 'tong tien chua thue',
        'doanh số chưa thuế', 'doanh so chua thue', 'tiền hàng', 'tien hang', 'giá trị chưa thuế',
        'giá trị hhdv', 'doanh thu', 'doanh so mua', 'doanh số mua', 'doanh so', 'thành tiền', 'thanh tien', 'amount', 'net amount'
    ],
    'thue_suat': [
        'thuế suất (%)', 'thuế suất', 'thue suat', 'ts (%)', 'ts', 'vat rate', 'rate'
    ],
    'tien_thue': [
        'tổng tiền thuế', 'tong tien thue',
        'thuế gtgt', 'thue gtgt', 'tiền thuế gtgt', 'tien thue gtgt', 'tiền thuế', 'tien thue',
        'tiền thuế đang ghi', 'vat amount', 'vat'
    ],
    'tong_tien': [
        'tổng tiền thanh toán', 'tong tien thanh toan',
        'tổng cộng', 'tong cong', 'tổng thanh toán', 'tong thanh toan', 'tổng tiền', 'tong tien', 'total'
    ],
    'ky_ke_khai': [
        'kì kê khai', 'ki ke khai', 'kỳ kê khai', 'ky ke khai', 'note tháng/năm', 'note thang/nam',
        'kỳ thuế', 'ky thue', 'tháng/năm', 'thang/nam', 'period'
    ],
    'trang_thai_hd': [
        'trạng thái hóa đơn', 'trang thai hoa don', 'trạng thái hđ', 'tình trạng hóa đơn'
    ]
}

SUPPORTED_TABULAR_EXTS = ('.xlsx', '.xls', '.csv')


def clean_str(val: Any) -> str:
    if val is None:
        return ''
    return str(val).strip()


def parse_number(val: Any) -> float:
    """Chuyển đổi số an toàn từ mọi dạng chuỗi hoặc số thực/nguyên"""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    is_neg = False
    if s.startswith('(') and s.endswith(')'):
        is_neg = True
        s = s[1:-1].strip()
    elif s.startswith('-'):
        is_neg = True
        s = s[1:].strip()
        
    s = re.sub(r'[^\d.,]', '', s)
    if not s:
        return 0.0
        
    if '.' in s and ',' in s:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) > 2:
            s = s.replace(',', '')
        elif len(parts) == 2 and len(parts[1]) == 3:
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    elif '.' in s:
        parts = s.split('.')
        if len(parts) > 2:
            s = s.replace('.', '')
        elif len(parts) == 2 and len(parts[1]) == 3 and int(parts[0]) > 0:
            s = s.replace('.', '')
            
    try:
        num = float(s)
        return -num if is_neg else num
    except ValueError:
        return 0.0


def parse_date(val: Any) -> Tuple[Optional[datetime.date], str]:
    """
    Chuẩn hóa ngày về đối tượng datetime.date và chuỗi ISO YYYY-MM-DD.
    Hỗ trợ cả Excel serial number (vd: 44176 -> 2020-12-11).
    """
    if val is None:
        return None, ''
    if isinstance(val, datetime.datetime):
        return val.date(), val.date().isoformat()
    if isinstance(val, datetime.date):
        return val, val.isoformat()
        
    if isinstance(val, (int, float)):
        try:
            n = float(val)
            if 30000 <= n <= 60000:
                dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(n)))
                return dt, dt.isoformat()
        except Exception:
            pass
            
    s = str(val).strip()
    if not s:
        return None, ''
        
    if s.isdigit() and len(s) == 5:
        n = int(s)
        if 30000 <= n <= 60000:
            dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=n)
            return dt, dt.isoformat()
            
    date_formats = [
        '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%Y/%m/%d',
        '%d/%m/%y', '%d-%m-%y', '%d.%m.%Y', '%d.%m.%y',
        '%Y%m%d'
    ]
    s_clean = s.split(' ')[0].strip()
    for fmt in date_formats:
        try:
            dt = datetime.datetime.strptime(s_clean, fmt).date()
            return dt, dt.isoformat()
        except ValueError:
            continue
            
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m:
        try:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            dt = datetime.date(year, month, day)
            return dt, dt.isoformat()
        except ValueError:
            pass
            
    return None, s


def parse_tax_rate(val: Any) -> Tuple[float, str]:
    if val is None:
        return 0.1, '10%'
    s = str(val).strip().upper()
    if not s:
        return 0.1, '10%'
    if 'KCT' in s or 'KHÔNG CHỊU THUẾ' in s or 'KHONG CHIU THUE' in s:
        return -1.0, 'KCT'
    if 'KKT' in s or 'KHÔNG KÊ KHAI' in s or 'KHÔNG TÍNH THUẾ' in s or 'KHONG KE KHAI' in s:
        return -2.0, 'KKT'
    if '0%' in s or s == '0':
        return 0.0, '0%'
        
    num = parse_number(s)
    if num >= 1.0:
        if 7.5 <= num <= 8.5:
            return 0.08, '8%'
        if 9.5 <= num <= 10.5:
            return 0.10, '10%'
        if 4.5 <= num <= 5.5:
            return 0.05, '5%'
        return num / 100.0, f'{num:.1f}%'
    else:
        if 0.075 <= num <= 0.085:
            return 0.08, '8%'
        if 0.095 <= num <= 0.105:
            return 0.10, '10%'
        if 0.045 <= num <= 0.055:
            return 0.05, '5%'
        if num == 0.0:
            return 0.0, '0%'
        return num, f'{num*100:.1f}%'


def clean_invoice_num(val: Any) -> str:
    if val is None:
        return ''
    s = str(val).strip()
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return s


def detect_headers(rows: List[List[Any]]) -> Tuple[int, Dict[str, int], int]:
    """
    Tự động dò dòng tiêu đề:
    Thử kiểm tra cả dòng đơn và ghép 2 dòng liên tiếp.
    Trả về: (data_start_row, best_mapping, best_score)
    """
    best_score = -1
    best_row_idx = 0
    best_mapping: Dict[str, int] = {}
    is_two_line = False
    
    max_scan = min(25, len(rows))
    for r in range(max_scan):
        curr_row = [clean_str(c).lower() for c in rows[r]]
        next_row = [clean_str(c).lower() for c in rows[r+1]] if r + 1 < len(rows) else []
        
        for test_mode in ['single', 'double']:
            if test_mode == 'double' and r + 1 >= len(rows):
                continue
                
            mapping: Dict[str, int] = {}
            col_assigned = set()
            
            num_cols = max(len(curr_row), len(next_row)) if test_mode == 'double' else len(curr_row)
            
            priority_fields = [
                'ky_hieu', 'so_hd', 'ngay_hd', 'doanh_so', 'tien_thue', 
                'mst', 'dien_giai', 'thue_suat', 'tong_tien', 'ten_doi_tac', 'ky_ke_khai', 'trang_thai_hd'
            ]
            
            for field in priority_fields:
                kw_list = COLUMN_KEYWORDS[field]
                for col_idx in range(num_cols):
                    if col_idx in col_assigned:
                        continue
                    c1 = curr_row[col_idx] if col_idx < len(curr_row) else ''
                    c2 = next_row[col_idx] if (test_mode == 'double' and col_idx < len(next_row)) else ''
                    combined = f"{c1} {c2}".strip()
                    
                    target = combined if test_mode == 'double' else c1
                    if not target:
                        continue
                        
                    matched = False
                    for kw in kw_list:
                        if kw == target or target.startswith(kw) or kw in target:
                            matched = True
                            break
                    if matched:
                        mapping[field] = col_idx
                        col_assigned.add(col_idx)
                        break
                        
            score = 0
            if 'so_hd' in mapping: score += 4
            if 'ngay_hd' in mapping: score += 4
            if 'doanh_so' in mapping: score += 4
            if 'tien_thue' in mapping: score += 4
            if 'ky_hieu' in mapping: score += 2
            if 'mst' in mapping: score += 2
            if 'dien_giai' in mapping: score += 2
            if 'thue_suat' in mapping: score += 1
            
            if score > best_score:
                best_score = score
                best_row_idx = r
                best_mapping = mapping
                is_two_line = (test_mode == 'double')
                
    data_start_row = best_row_idx + (2 if is_two_line else 1)
    return data_start_row, best_mapping, best_score


def resolve_input_paths(
    path_or_paths: Union[str, List[str]],
    extensions: Tuple[str, ...] = SUPPORTED_TABULAR_EXTS
) -> List[str]:
    """
    Giải phân giải đầu vào linh hoạt:
    - Chuỗi ngăn cách bởi dấu `;` hoặc xuống dòng (nhiều file/thư mục).
    - Đường dẫn tới 1 file cụ thể.
    - Đường dẫn tới 1 thư mục (tự động quét đệ quy lấy tất cả file có đuôi phù hợp, bỏ qua file tạm `~$`).
    """
    if not path_or_paths:
        return []
    if isinstance(path_or_paths, str):
        raw_items = [p.strip() for p in re.split(r'[;\n]+', path_or_paths) if p.strip()]
    else:
        raw_items = [str(p).strip() for p in path_or_paths if str(p).strip()]

    resolved: List[str] = []
    seen = set()
    exts_lower = tuple(e.lower() for e in extensions)

    for item in raw_items:
        if os.path.isfile(item):
            ap = os.path.abspath(item)
            if ap not in seen:
                seen.add(ap)
                resolved.append(ap)
        elif os.path.isdir(item):
            for root_dir, _, files in os.walk(item):
                for fn in sorted(files):
                    if fn.startswith('~$'):
                        continue
                    if fn.lower().endswith(exts_lower):
                        ap = os.path.abspath(os.path.join(root_dir, fn))
                        if ap not in seen:
                            seen.add(ap)
                            resolved.append(ap)
    return resolved


class ExcelNormalizer:
    @staticmethod
    def _normalize_rows(
        all_raw_rows: List[List[Any]],
        source_file: str = "",
        source_sheet: str = ""
    ) -> List[Dict[str, Any]]:
        if not all_raw_rows:
            return []
            
        start_row, col_map, score = detect_headers(all_raw_rows)
        # Nếu sheet không có cột Số HĐ và điểm nhận diện quá thấp -> sheet bìa/tổng hợp không phải bảng kê HĐ
        if score < 6 or ('so_hd' not in col_map and 'doanh_so' not in col_map):
            return []
        
        if 'mst' not in col_map and all_raw_rows[0]:
            for c_idx in range(len(all_raw_rows[0])):
                sample_vals = [
                    str(all_raw_rows[r][c_idx] or '').strip()
                    for r in range(start_row, min(start_row + 10, len(all_raw_rows)))
                    if c_idx < len(all_raw_rows[r])
                ]
                if any('mst' in v.lower() or (v.isdigit() and len(v) in (10, 13)) for v in sample_vals):
                    col_map['mst'] = c_idx
                    break

        normalized_records = []
        for r_idx in range(start_row, len(all_raw_rows)):
            row = all_raw_rows[r_idx]
            if not row or not any(row):
                continue
                
            def get_val(f_name):
                idx = col_map.get(f_name)
                if idx is not None and idx < len(row):
                    return row[idx]
                return None
                
            so_hd_raw = get_val('so_hd')
            ngay_hd_raw = get_val('ngay_hd')
            doanh_so_raw = get_val('doanh_so')
            tien_thue_raw = get_val('tien_thue')
            
            first_cols_text = ' '.join([str(c or '').lower() for c in row[:6]])
            if any(kw in first_cols_text for kw in ['tổng cộng', 'tong cong', 'tổng số', 'tong so', 'tổng tiền', 'cộng nhóm']):
                continue
                
            # Bỏ qua hóa đơn đã hủy nếu cột trạng thái HĐ ghi rõ đã hủy
            trang_thai_hd = clean_str(get_val('trang_thai_hd')).lower()
            if any(h in trang_thai_hd for h in ['đã hủy', 'da huy', 'hóa đơn hủy', 'hoa don huy']):
                continue

            so_hd = clean_invoice_num(so_hd_raw)
            dt_obj, dt_str = parse_date(ngay_hd_raw)
            doanh_so = parse_number(doanh_so_raw)
            tien_thue = parse_number(tien_thue_raw)
            
            if not so_hd and not dt_str and doanh_so == 0 and tien_thue == 0:
                continue
            # Nếu không có số HĐ và không có ngày HĐ (chỉ là dòng tiểu mục/tổng phụ) -> bỏ qua
            if not so_hd and not dt_str:
                continue
                
            mst = clean_str(get_val('mst'))
            if re.fullmatch(r'\d+\.0', mst):
                mst = mst[:-2]
            ten_dt = clean_str(get_val('ten_doi_tac'))
            dien_giai = clean_str(get_val('dien_giai'))
            ky_hieu = clean_str(get_val('ky_hieu'))
            
            raw_ts = get_val('thue_suat')
            ts_float, ts_text = parse_tax_rate(raw_ts)
            ky_ke_khai = clean_str(get_val('ky_ke_khai'))
            
            if not ky_ke_khai and dt_obj:
                ky_ke_khai = f"T{dt_obj.month:02d}/{dt_obj.year}"
                
            # Nếu file không có cột thuế suất nhưng có doanh số & tiền thuế -> suy luận thuế suất thực tế
            if raw_ts is None or str(raw_ts).strip() == '':
                if doanh_so > 0 and tien_thue > 0:
                    ratio = tien_thue / doanh_so
                    if 0.075 <= ratio <= 0.085:
                        ts_float, ts_text = 0.08, '8%'
                    elif 0.045 <= ratio <= 0.055:
                        ts_float, ts_text = 0.05, '5%'
                    elif 0.095 <= ratio <= 0.105:
                        ts_float, ts_text = 0.10, '10%'
                elif doanh_so > 0 and tien_thue == 0:
                    ts_float, ts_text = 0.0, '0%'
            elif tien_thue == 0 and doanh_so > 0 and ts_float > 0:
                tien_thue = round(doanh_so * ts_float)
                
            rec = {
                'row_index': r_idx + 1,
                'source_file': source_file,
                'source_sheet': source_sheet,
                'so_hd': so_hd,
                'ky_hieu': ky_hieu,
                'ngay_hd': dt_str,
                'ngay_hd_date': dt_obj,
                'mst': mst,
                'ten_doi_tac': ten_dt,
                'dien_giai': dien_giai,
                'doanh_so': doanh_so,
                'thue_suat_float': ts_float,
                'thue_suat_text': ts_text,
                'tien_thue': tien_thue,
                'tong_tien': doanh_so + tien_thue,
                'ky_ke_khai': ky_ke_khai
            }
            normalized_records.append(rec)
            
        return normalized_records

    @staticmethod
    def normalize_invoice_sheet(file_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Chuẩn hóa 1 file Excel/CSV.
        - Nếu chỉ định `sheet_name`: chỉ đọc sheet đó.
        - Nếu `sheet_name=None`: tự động quét TẤT CẢ các sheet trong workbook có bảng kê hóa đơn
          (hữu ích khi kế toán lưu Bảng kê nhiều tháng/quý/năm trên nhiều sheet của cùng 1 file).
        """
        sheets_dict = read_tabular_sheets(file_path)
        fname = os.path.basename(file_path)
        if not sheets_dict:
            return []

        if sheet_name and sheet_name in sheets_dict:
            return ExcelNormalizer._normalize_rows(sheets_dict[sheet_name], fname, sheet_name)

        all_records: List[Dict[str, Any]] = []
        for sname, rows in sheets_dict.items():
            recs = ExcelNormalizer._normalize_rows(rows, fname, sname)
            all_records.extend(recs)
        return all_records

    @staticmethod
    def normalize_multiple_sources(
        path_or_paths: Union[str, List[str]],
        deduplicate: bool = True,
        log_fn=None
    ) -> List[Dict[str, Any]]:
        """
        Chuẩn hóa và gộp hóa đơn từ:
        - 1 file đơn lẻ
        - Nhiều file Excel/CSV được chọn cùng lúc (ngăn cách bởi `;`)
        - Cả thư mục chứa hàng chục file HĐĐT tải từ TCT theo từng tháng/quý/năm hoặc bảng kê nhiều kỳ hoàn thuế.
        """
        files = resolve_input_paths(path_or_paths, SUPPORTED_TABULAR_EXTS)
        all_records: List[Dict[str, Any]] = []

        for fp in files:
            fname = os.path.basename(fp)
            try:
                recs = ExcelNormalizer.normalize_invoice_sheet(fp)
                if log_fn:
                    log_fn(f"   + File [{fname}]: nạp {len(recs):,} dòng hóa đơn.")
                all_records.extend(recs)
            except Exception as e:
                if log_fn:
                    log_fn(f"   [!] Bỏ qua file [{fname}] do lỗi đọc: {e}")

        if not deduplicate or not all_records:
            return all_records

        # Khử trùng lặp chỉ khi CÙNG một dòng hóa đơn bị xuất trùng trong 2 file tải chồng ngày từ TCT
        # Lưu ý: Một hóa đơn có thể có nhiều dòng mặt hàng/thuế suất trong cùng 1 file (khác row_index),
        # nên chỉ coi là trùng nếu trùng (ky_hieu, so_hd, ngay_hd, mst, doanh_so, tien_thue) ở KHÁC file nguồn.
        unique_records: List[Dict[str, Any]] = []
        seen_cross_file = set()
        seen_cross_file_origin: Dict[Any, Tuple[str, str]] = {}
        for rec in all_records:
            sig = (
                str(rec.get('ky_hieu', '')).strip().upper(),
                str(rec.get('so_hd', '')).strip().lstrip('0') or '0',
                str(rec.get('ngay_hd', '')).strip(),
                str(rec.get('mst', '')).strip(),
                round(float(rec.get('doanh_so', 0)), 2),
                round(float(rec.get('tien_thue', 0)), 2),
            )
            origin = (rec.get('source_file', ''), rec.get('source_sheet', ''))
            # Nếu đã gặp đúng hóa đơn này ở một file/sheet khác (do tải trùng khoảng ngày từ TCT)
            if sig in seen_cross_file and seen_cross_file_origin.get(sig) != origin:
                continue
            if sig not in seen_cross_file:
                seen_cross_file.add(sig)
                seen_cross_file_origin[sig] = origin
            unique_records.append(rec)

        return unique_records
