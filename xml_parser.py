# -*- coding: utf-8 -*-
"""
xml_parser.py
-------------
Module đọc và trích xuất dữ liệu tờ khai thuế GTGT (Mẫu 01/GTGT theo Thông tư 26/2015, TT 156/2013, TT 80/2021)
từ file XML kết xuất từ phần mềm HTKK hoặc trang Thuế điện tử (thuedientu.gdt.gov.vn / eTax).

Trích xuất đầy đủ các chỉ tiêu cốt lõi:
- [22]: Thuế GTGT còn được khấu trừ kỳ trước chuyển sang
- [23]: Giá trị HHDV mua vào trong kỳ
- [24]: Thuế GTGT của HHDV mua vào trong kỳ
- [25]: Thuế GTGT đầu vào được khấu trừ kỳ này
- [26]: Doanh thu HHDV bán ra không chịu thuế GTGT
- [27]: Doanh thu HHDV bán ra chịu thuế GTGT ([27] = [29] + [30] + [32] + [32a])
- [28]: Thuế GTGT của HHDV bán ra chịu thuế GTGT ([28] = [31] + [33])
- [29]: Doanh thu HHDV thuế suất 0%
- [30]: Doanh thu HHDV thuế suất 5%
- [31]: Thuế GTGT của HHDV thuế suất 5%
- [32]: Doanh thu HHDV thuế suất 10%
- [33]: Thuế GTGT của HHDV thuế suất 10%
- [32a]: Doanh thu HHDV không phải tính nộp thuế (KKT)
- [34]: Tổng doanh thu ([34] = [26] + [27])
- [35]: Tổng thuế GTGT đầu ra ([35] = [28])
- [36]: Thuế GTGT phát sinh trong kỳ ([36] = [35] - [25])
- [40]: Thuế GTGT còn phải nộp
- [43]: Thuế GTGT còn được khấu trừ chuyển kỳ sau
"""

import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, Any, Optional

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

class TaxXmlParser:
    @staticmethod
    def parse_vat_return(xml_file_path: str) -> Dict[str, Any]:
        """
        Đọc và phân tích file XML tờ khai 01/GTGT
        """
        if not os.path.exists(xml_file_path):
            raise FileNotFoundError(f"Không tìm thấy file XML: {xml_file_path}")
            
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        info: Dict[str, Any] = {
            'file_name': os.path.basename(xml_file_path),
            'ma_so_thue': '',
            'ten_nnt': '',
            'ky_tinh_thue': '',
            'thang_quy': '',
            'nam': '',
            'lan_ke_khai': 0,
            'ngay_nop': '',
            'ct22': 0.0,
            'ct23': 0.0,
            'ct24': 0.0,
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
            'ct40': 0.0,
            'ct43': 0.0,
            'pl_giam_thue': [] # Phụ lục giảm thuế (NQ43, NQ101, NQ142, NQ204 nếu có)
        }
        
        # Duyệt cây XML bắt các node quan trọng không phân biệt namespace
        for elem in root.iter():
            tag = clean_tag(elem.tag).lower()
            text = (elem.text or '').strip()
            
            # Thông tin chung NNT
            if tag in ('mst', 'masothue'):
                if not info['ma_so_thue']: info['ma_so_thue'] = text
            elif tag in ('tennnt', 'tennguoinopthue', 'ten_nnt'):
                if not info['ten_nnt']: info['ten_nnt'] = text
            elif tag in ('kytinhthue', 'ky_tinh_thue', 'kykhai'):
                if not info['ky_tinh_thue']: info['ky_tinh_thue'] = text
            elif tag in ('thang', 'quy'):
                info['thang_quy'] = text
            elif tag in ('nam', 'namke'):
                info['nam'] = text
            elif tag in ('lankk', 'lankekhai', 'lan_bo_sung'):
                info['lan_ke_khai'] = int(to_float(text))
            elif tag in ('ngaynop', 'ngaylap'):
                info['ngay_nop'] = text
                
            # Các chỉ tiêu tờ khai 01/GTGT
            # XML HTKK thường lưu dạng <ct22> hoặc <ct_22> hoặc <ct ct="22">
            ct_match = re.search(r'ct_?([0-9]{2}[a-z]?)', tag)
            if ct_match:
                ct_num = ct_match.group(1).lower()
                field_key = f'ct{ct_num}'
                if field_key in info:
                    info[field_key] = to_float(text)
                    
            # Nếu thẻ có attribute ct="23"
            if 'ct' in elem.attrib:
                ct_attr = elem.attrib['ct'].lower()
                field_key = f'ct{ct_attr}'
                if field_key in info:
                    info[field_key] = to_float(text)
                    
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
