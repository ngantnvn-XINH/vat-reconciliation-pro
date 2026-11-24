# -*- coding: utf-8 -*-
"""
xml_parser.py
-------------
Module đọc và trích xuất dữ liệu tờ khai thuế GTGT (Mẫu 01/GTGT theo Thông tư 26/2015, TT 156/2013, TT 80/2021)
từ 1 file XML, NHIỀU file XML, hoặc CẢ THƯ MỤC chứa hồ sơ khai thuế XML (HTKK / eTax).

Tính năng đặc biệt cho kiểm tra / hoàn thuế nhiều kỳ hoặc nhiều năm:
- Tự động lọc chỉ lấy đúng tờ khai `01/GTGT` (bỏ qua tờ khai TNDN, TNCN, BCTC, LPMB nếu chọn chung cả thư mục XML của doanh nghiệp).
- Tự động nhóm theo từng kỳ tính thuế (Tháng/Quý) và chỉ lấy BẢN HỢP LỆ CUỐI CÙNG của mỗi kỳ (bản Bổ sung có `lan_ke_khai` cao nhất; nếu cùng lần bổ sung thì lấy bản có `ngay_nop` mới nhất).
- Tự động cộng gộp các chỉ tiêu phát sinh trong kỳ (`[23]`, `[24]`, `[25]`, `[26]`, `[27]`, `[28]`, `[29]`, `[30]`, `[31]`, `[32]`, `[32a]`, `[33]`, `[34]`, `[35]`, `[36]`) trên toàn bộ các kỳ trong khoảng hoàn thuế / nhiều năm để đối chiếu khớp 100% với Bảng kê nhiều kỳ.
- Trả về đồng thời cả `periods_detail` (danh sách chi tiết từng kỳ tờ khai đã chọn) để xuất bảng kê từng kỳ ra Excel.
"""

import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional, Tuple, Union

from excel_normalizer import resolve_input_paths


def clean_tag(tag: str) -> str:
    """Loại bỏ namespace XML: {http://...}tagName -> tagName"""
    return tag.split('}')[-1] if '}' in tag else tag


def to_float(val: Any) -> float:
    if val is None:
        return 0.0
    s = str(val).strip().replace(',', '')
    try:
        return float(s)
    except ValueError:
        return 0.0


def normalize_period_key(ky_tinh_thue: str, thang_quy: str = "", nam: str = "") -> Tuple[str, Tuple[int, int]]:
    """
    Chuẩn hóa kỳ tính thuế về mã thống nhất (VD: 'T08/2024' hoặc 'Q3/2024') và khóa sắp xếp (year, month_or_quarter).
    Hỗ trợ các dạng:
    - 'Tháng 08/2024', '08/2024', 'T08/2024', '8/2024' -> ('08/2024', (2024, 8))
    - 'Quý 03/2024', 'Q3/2024', 'Q03/2024' -> ('Q3/2024', (2024, 30 + 3))
    - '2024' -> ('2024', (2024, 0))
    """
    raw = str(ky_tinh_thue or '').strip()
    s = raw.upper()

    # Kiểm tra dạng Quý
    m_q = re.search(r'(?:QUÝ|QUY|Q)\s*0?([1-4])\s*[/\-]\s*(\d{4})', s)
    if m_q:
        q, y = int(m_q.group(1)), int(m_q.group(2))
        return f"Q{q}/{y}", (y, 30 + q)

    # Kiểm tra dạng Tháng
    m_t = re.search(r'(?:THÁNG|THANG|T)?\s*0?([1-9]|1[0-2])\s*[/\-]\s*(\d{4})', s)
    if m_t:
        m, y = int(m_t.group(1)), int(m_t.group(2))
        return f"{m:02d}/{y}", (y, m)

    # Dạng viết tắt YYMM hoặc YYQ# (HTKK cũ)
    m_short = re.match(r'^(\d{2})(0[1-9]|1[0-2]|Q[1-4])$', s)
    if m_short:
        y = 2000 + int(m_short.group(1))
        p = m_short.group(2)
        if p.startswith('Q'):
            q = int(p[1])
            return f"Q{q}/{y}", (y, 30 + q)
        else:
            m = int(p)
            return f"{m:02d}/{y}", (y, m)

    if thang_quy and nam and str(nam).isdigit():
        y = int(nam)
        tq = str(thang_quy).strip().upper()
        if tq.startswith('Q') or len(tq) == 1 and tq in '1234' and 'QU' in s:
            q = int(re.sub(r'\D', '', tq) or 1)
            return f"Q{q}/{y}", (y, 30 + q)
        elif tq.isdigit():
            m = int(tq)
            return f"{m:02d}/{y}", (y, m)

    m_y = re.search(r'(20\d{2}|19\d{2})', s)
    if m_y:
        y = int(m_y.group(1))
        return raw or str(y), (y, 0)

    return raw or "UNKNOWN", (0, 0)


class TaxXmlParser:
    @staticmethod
    def is_vat_01_xml(xml_file_path: str) -> bool:
        """
        Kiểm tra nhanh xem file XML có phải là Tờ khai thuế GTGT (01/GTGT) hay không.
        Giúp tự động lọc bỏ tờ khai TNDN, TNCN, BCTC, LPMB khi người dùng chọn cả thư mục XML tổng hợp.
        """
        try:
            tree = ET.parse(xml_file_path)
            root = tree.getroot()
        except Exception:
            return False

        ma_tkhai = ""
        ten_tkhai = ""
        has_vat_ct = False

        for elem in root.iter():
            tag = clean_tag(elem.tag).lower()
            txt = (elem.text or '').strip()
            if tag == 'matkhai' and txt:
                ma_tkhai = txt.upper()
            elif tag == 'tentkhai' and txt:
                ten_tkhai = txt.upper()
            elif tag in ('ct23', 'ct24', 'ct25', 'ct32', 'ct33', 'ct34', 'ct35'):
                has_vat_ct = True

        combined = f"{ma_tkhai} {ten_tkhai}"
        # Nếu có thẻ maTKhai/tenTKhai nhưng thuộc sắc thuế khác (TNCN, TNDN, BCTC, NTNN, LPMB) -> loại trừ
        if any(other in combined for other in ("TNCN", "TNDN", "BCTC", "CÂN ĐỐI", "NTNN", "NHÀ THẦU", "MÔN BÀI", "LPMB", "05/QTT", "03/TNDN")):
            return False
        if "GTGT" in combined or "GIÁ TRỊ GIA TĂNG" in combined or ma_tkhai in ("01", "842", "01/GTGT"):
            return True
        return has_vat_ct

    @staticmethod
    def parse_vat_return(xml_file_path: str) -> Dict[str, Any]:
        """
        Đọc và phân tích 1 file XML tờ khai 01/GTGT.
        Lưu ý: Chỉ quét chỉ tiêu trong nhánh <CTieuTKhaiChinh> (hoặc tờ khai chính, tránh quét nhầm vào phụ lục PL01_1_KHBS làm ghi đè chỉ tiêu bằng số chênh lệch KHBS).
        """
        if not os.path.exists(xml_file_path):
            raise FileNotFoundError(f"Không tìm thấy file XML: {xml_file_path}")
            
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        info: Dict[str, Any] = {
            'file_path': os.path.abspath(xml_file_path),
            'file_name': os.path.basename(xml_file_path),
            'ma_tkhai': '',
            'ten_tkhai': '',
            'ma_so_thue': '',
            'ten_nnt': '',
            'ky_tinh_thue': '',
            'period_key': '',
            'period_sort_key': (0, 0),
            'thang_quy': '',
            'nam': '',
            'loai_tkhai': 'C',
            'lan_ke_khai': 0,
            'ngay_nop': '',
            'ct22': 0.0,
            'ct23': 0.0,
            'ct23a': 0.0,
            'ct24': 0.0,
            'ct24a': 0.0,
            'ct25': 0.0,
            'ct26': 0.0,
            'ct27': 0.0,
            'ct28': 0.0,
            'ct29': 0.0,
            'ct30': 0.0,
            'ct31': 0.0,
            'ct32': 0.0,
            'ct33': 0.0,
            'ct32a': 0.0,
            'ct34': 0.0,
            'ct35': 0.0,
            'ct36': 0.0,
            'ct37': 0.0,
            'ct38': 0.0,
            'ct40': 0.0,
            'ct41': 0.0,
            'ct42': 0.0,
            'ct43': 0.0,
            'pl_giam_thue': []
        }
        
        ctieu_chinh_node = None
        for elem in root.iter():
            tag = clean_tag(elem.tag).lower()
            text = (elem.text or '').strip()
            
            if tag == 'ctieutkhaichinh':
                ctieu_chinh_node = elem
            elif tag == 'matkhai' and not info['ma_tkhai']:
                info['ma_tkhai'] = text
            elif tag == 'tentkhai' and not info['ten_tkhai']:
                info['ten_tkhai'] = text
            elif tag in ('mst', 'masothue') and not info['ma_so_thue']:
                info['ma_so_thue'] = text
            elif tag in ('tennnt', 'tennguoinopthue', 'ten_nnt') and not info['ten_nnt']:
                info['ten_nnt'] = text
            elif tag in ('kytinhthue', 'ky_tinh_thue', 'kykhai', 'kykkhai') and not info['ky_tinh_thue']:
                info['ky_tinh_thue'] = text
            elif tag in ('thang', 'quy') and not info['thang_quy']:
                info['thang_quy'] = text
            elif tag in ('nam', 'namke') and not info['nam']:
                info['nam'] = text
            elif tag == 'loaitkhai' and text:
                info['loai_tkhai'] = text.upper()
            elif tag in ('lankk', 'lankekhai', 'lan_bo_sung', 'solan', 'solanbosung', 'solanbsung', 'lanbosung', 'lanbsung'):
                val_lan = int(to_float(text))
                if val_lan > info['lan_ke_khai']:
                    info['lan_ke_khai'] = val_lan
            elif tag in ('ngaynop', 'ngaylap', 'ngaylaptkhai') and not info['ngay_nop']:
                info['ngay_nop'] = text

        if info['loai_tkhai'] == 'B' and info['lan_ke_khai'] == 0:
            info['lan_ke_khai'] = 1

        # Quét các chỉ tiêu tờ khai trong nhánh CTieuTKhaiChinh (nếu có), tránh quét vào PLuc (như PL01_1_KHBS)
        scan_target = ctieu_chinh_node if ctieu_chinh_node is not None else root
        for elem in scan_target.iter():
            tag = clean_tag(elem.tag).lower()
            if tag in ('pluc', 'pl01_1_khbs', 'ctieukhbs'):
                continue
            text = (elem.text or '').strip()

            ct_match = re.fullmatch(r'ct_?([0-9]{2}[a-z]?)', tag)
            if ct_match:
                ct_num = ct_match.group(1).lower()
                field_key = f'ct{ct_num}'
                if field_key in info:
                    info[field_key] = to_float(text)
                    
            if 'ct' in elem.attrib:
                ct_attr = elem.attrib['ct'].lower()
                field_key = f'ct{ct_attr}'
                if field_key in info:
                    info[field_key] = to_float(text)

        p_key, p_sort = normalize_period_key(info['ky_tinh_thue'], info['thang_quy'], info['nam'])
        info['period_key'] = p_key
        info['period_sort_key'] = p_sort
        if not info['ky_tinh_thue']:
            info['ky_tinh_thue'] = p_key

        # Tự tính toán chuẩn hóa nếu XML thiếu tổng hợp
        if info['ct27'] == 0:
            info['ct27'] = info['ct29'] + info['ct30'] + info['ct32'] + info['ct32a']
        if info['ct28'] == 0:
            info['ct28'] = info['ct31'] + info['ct33']
        if info['ct34'] == 0:
            info['ct34'] = info['ct26'] + info['ct27']
        if info['ct35'] == 0:
            info['ct35'] = info['ct28']
        if info['ct36'] == 0:
            info['ct36'] = info['ct35'] - info['ct25']
            
        return info

    @staticmethod
    def parse_multiple_vat_returns(
        path_or_paths: Union[str, List[str]],
        log_fn=None
    ) -> Optional[Dict[str, Any]]:
        """
        Đọc 1 file XML, nhiều file XML hoặc cả thư mục chứa tờ khai XML nhiều kỳ / nhiều năm:
        1. Tự động lọc chỉ lấy tờ khai 01/GTGT.
        2. Với mỗi kỳ tính thuế (`period_key`), tự động chọn tờ khai có `lan_ke_khai` lớn nhất
           (bản bổ sung cuối cùng của kỳ đó; nếu trùng lần bổ sung thì lấy `ngay_nop` mới nhất).
        3. Cộng gộp các chỉ tiêu phát sinh trong kỳ (`ct23`..`ct36`) của tất cả các kỳ hiệu lực cuối cùng.
        4. Lưu danh sách `periods_detail` đã sắp xếp theo thời gian để xuất báo cáo chi tiết từng kỳ.
        """
        xml_files = resolve_input_paths(path_or_paths, ('.xml',))
        if not xml_files:
            return None

        parsed_list: List[Dict[str, Any]] = []
        for fp in xml_files:
            if not TaxXmlParser.is_vat_01_xml(fp):
                continue
            try:
                ret = TaxXmlParser.parse_vat_return(fp)
                parsed_list.append(ret)
            except Exception as e:
                if log_fn:
                    log_fn(f"   [!] Bỏ qua XML [{os.path.basename(fp)}]: {e}")

        if not parsed_list:
            return None

        # Nếu chỉ có đúng 1 file tờ khai -> trả về cấu trúc tương thích hoàn toàn, kèm periods_detail = [ret]
        if len(parsed_list) == 1:
            single = parsed_list[0]
            single['periods_detail'] = [single]
            single['total_xml_scanned'] = 1
            single['effective_periods_count'] = 1
            return single

        # Nhóm theo period_key và chọn bản bổ sung cuối cùng cho từng kỳ
        by_period: Dict[str, List[Dict[str, Any]]] = {}
        for item in parsed_list:
            pk = item.get('period_key') or item.get('ky_tinh_thue') or item['file_name']
            by_period.setdefault(pk, []).append(item)

        effective_periods: List[Dict[str, Any]] = []
        for pk, candidates in by_period.items():
            # Sắp xếp ưu tiên: lan_ke_khai cao nhất -> ngay_nop mới nhất
            candidates.sort(key=lambda x: (int(x.get('lan_ke_khai', 0)), str(x.get('ngay_nop', ''))))
            chosen = candidates[-1]
            chosen['total_versions_in_period'] = len(candidates)
            effective_periods.append(chosen)

        # Sắp xếp các kỳ theo thứ tự thời gian
        effective_periods.sort(key=lambda x: (x.get('period_sort_key', (0, 0)), str(x.get('period_key', ''))))

        if log_fn:
            log_fn(
                f"   + Đã quét {len(parsed_list)} tờ khai 01/GTGT -> Chốt {len(effective_periods)} kỳ tính thuế hiệu lực cuối cùng "
                f"(Từ kỳ {effective_periods[0]['period_key']} đến {effective_periods[-1]['period_key']})."
            )

        # Cộng gộp các chỉ tiêu phát sinh qua các kỳ
        sum_fields = [
            'ct23', 'ct23a', 'ct24', 'ct24a', 'ct25',
            'ct26', 'ct27', 'ct28', 'ct29', 'ct30', 'ct31',
            'ct32', 'ct33', 'ct32a', 'ct34', 'ct35', 'ct36',
            'ct37', 'ct38', 'ct40', 'ct41', 'ct42'
        ]

        first_p = effective_periods[0]
        last_p = effective_periods[-1]
        period_range_label = (
            f"{first_p['period_key']} -> {last_p['period_key']} ({len(effective_periods)} kỳ)"
            if len(effective_periods) > 1
            else first_p['ky_tinh_thue']
        )

        aggregated: Dict[str, Any] = {
            'file_name': f"{len(effective_periods)} tờ khai 01/GTGT ({period_range_label})",
            'ma_so_thue': last_p.get('ma_so_thue') or first_p.get('ma_so_thue', ''),
            'ten_nnt': last_p.get('ten_nnt') or first_p.get('ten_nnt', ''),
            'ky_tinh_thue': period_range_label,
            'lan_ke_khai': f"Chốt bản cuối mỗi kỳ ({len(effective_periods)} kỳ / {len(parsed_list)} XML)",
            'ngay_nop': last_p.get('ngay_nop', ''),
            'ct22': first_p.get('ct22', 0.0),  # Kỳ đầu tiên chuyển sang
            'ct43': last_p.get('ct43', 0.0),   # Kỳ cuối cùng còn được khấu trừ chuyển kỳ sau
            'periods_detail': effective_periods,
            'total_xml_scanned': len(parsed_list),
            'effective_periods_count': len(effective_periods),
        }

        for sf in sum_fields:
            aggregated[sf] = sum(float(p.get(sf, 0.0)) for p in effective_periods)

        return aggregated
