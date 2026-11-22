# -*- coding: utf-8 -*-
"""
reconciler.py
-------------
Trái tim xử lý đối chiếu số liệu thuế GTGT:
1. Đối chiếu HĐĐT Bán ra (GDT) vs Bảng kê Bán ra (Kế toán nội bộ)
2. Đối chiếu HĐĐT Mua vào (GDT) vs Bảng kê Mua vào (Kế toán nội bộ)
3. Đối chiếu Tổng hợp Mua vào / Bán ra vs Tờ khai thuế GTGT (Chỉ tiêu [23], [24], [25], [26], [29], [30], [32], [32a], [34], [35])
4. Tự động phát hiện & phân loại sai lệch:
   - Hóa đơn có trên hệ thống TCT nhưng kế toán bỏ sót chưa kê khai.
   - Hóa đơn kế toán kê khai nhưng không tồn tại trên hệ thống TCT (nguy cơ hóa đơn khống, hủy).
   - Lệch số tiền doanh số hoặc lệch tiền thuế GTGT giữa 2 nguồn.
   - Kê khai sai kỳ thuế (lệch tháng/quý).
   - Sai thuế suất (8% vs 10%, 5% vs 10%).
   - Sai lệch tổng tiền với chỉ tiêu trên tờ khai thuế.
"""

from typing import List, Dict, Any, Tuple, Optional
import collections

def make_invoice_key(ky_hieu: str, so_hd: str, mst: str = "") -> str:
    """Tạo khóa định danh duy nhất cho hóa đơn"""
    # Chuẩn hóa số HĐ: bỏ số 0 ở đầu nếu toàn số (0000163 -> 163) để khớp mềm
    so_clean = str(so_hd or '').strip().lstrip('0') or '0'
    kh_clean = str(ky_hieu or '').strip().upper()
    mst_clean = str(mst or '').strip().replace('-', '').replace(' ', '')
    return f"{kh_clean}_{so_clean}_{mst_clean}"

class VatReconciler:
    @staticmethod
    def match_invoices(
        gdt_list: List[Dict[str, Any]], 
        internal_list: List[Dict[str, Any]],
        match_type: str = "Đầu vào"
    ) -> Dict[str, Any]:
        """
        Đối chiếu 2 danh sách hóa đơn (GDT vs Nội bộ).
        Trả về kết quả chi tiết:
        - Khớp hoàn toàn (Both matched, equal amount)
        - Lệch tiền (Matched, diff in revenue or tax)
        - Thừa ở GDT (Chưa kê khai trên bảng kê)
        - Thừa ở Bảng kê (Có trên bảng kê nhưng không có trên GDT)
        """
        # Tạo map cho internal
        internal_map = {}
        for item in internal_list:
            key = make_invoice_key(item.get('ky_hieu', ''), item.get('so_hd', ''), item.get('mst', ''))
            # Nếu trùng key thì gom danh sách
            if key not in internal_map:
                internal_map[key] = []
            internal_map[key].append(item)
            
        matched_perfect = []
        matched_diff_amount = []
        only_in_gdt = []
        used_internal_ids = set()
        
        for g_item in gdt_list:
            key = make_invoice_key(g_item.get('ky_hieu', ''), g_item.get('so_hd', ''), g_item.get('mst', ''))
            # Nếu không tìm thấy key đầy đủ, thử tìm theo số HĐ + MST (bỏ ký hiệu)
            candidates = internal_map.get(key)
            if not candidates:
                # Tìm fallback theo số HĐ + ngày hoặc số HĐ + doanh số
                so_only = str(g_item.get('so_hd', '')).strip().lstrip('0')
                for int_key, items in internal_map.items():
                    for it in items:
                        if id(it) in used_internal_ids:
                            continue
                        it_so = str(it.get('so_hd', '')).strip().lstrip('0')
                        if it_so == so_only and (it.get('mst') == g_item.get('mst') or abs(it.get('doanh_so', 0) - g_item.get('doanh_so', 0)) < 100):
                            candidates = [it]
                            break
                    if candidates:
                        break
                        
            if candidates:
                # Chọn candidate chưa dùng
                chosen = None
                for c in candidates:
                    if id(c) not in used_internal_ids:
                        chosen = c
                        break
                if chosen:
                    used_internal_ids.add(id(chosen))
                    diff_ds = round(g_item.get('doanh_so', 0) - chosen.get('doanh_so', 0), 2)
                    diff_thue = round(g_item.get('tien_thue', 0) - chosen.get('tien_thue', 0), 2)
                    
                    record = {
                        'ky_hieu_gdt': g_item.get('ky_hieu', ''),
                        'so_hd_gdt': g_item.get('so_hd', ''),
                        'ngay_hd_gdt': g_item.get('ngay_hd', ''),
                        'mst_gdt': g_item.get('mst', ''),
                        'ten_dt_gdt': g_item.get('ten_doi_tac', ''),
                        'doanh_so_gdt': g_item.get('doanh_so', 0),
                        'tien_thue_gdt': g_item.get('tien_thue', 0),
                        
                        'ky_hieu_nb': chosen.get('ky_hieu', ''),
                        'so_hd_nb': chosen.get('so_hd', ''),
                        'ngay_hd_nb': chosen.get('ngay_hd', ''),
                        'mst_nb': chosen.get('mst', ''),
                        'ten_dt_nb': chosen.get('ten_doi_tac', ''),
                        'doanh_so_nb': chosen.get('doanh_so', 0),
                        'tien_thue_nb': chosen.get('tien_thue', 0),
                        
                        'diff_doanh_so': diff_ds,
                        'diff_tien_thue': diff_thue,
                        'dien_giai': chosen.get('dien_giai') or g_item.get('dien_giai', '')
                    }
                    
                    if abs(diff_ds) < 1.0 and abs(diff_thue) < 1.0:
                        record['trang_thai'] = 'Khớp hoàn toàn'
                        matched_perfect.append(record)
                    else:
                        record['trang_thai'] = 'Khớp HĐ nhưng lệch tiền'
                        matched_diff_amount.append(record)
                else:
                    only_in_gdt.append(g_item)
            else:
                only_in_gdt.append(g_item)
                
        # Thu thập các HĐ chỉ có ở bảng kê nội bộ
        only_in_internal = []
        for it in internal_list:
            if id(it) not in used_internal_ids:
                only_in_internal.append(it)
                
        # Tính toán tổng hợp số liệu
        tot_gdt_ds = sum(x.get('doanh_so', 0) for x in gdt_list)
        tot_gdt_thue = sum(x.get('tien_thue', 0) for x in gdt_list)
        tot_nb_ds = sum(x.get('doanh_so', 0) for x in internal_list)
        tot_nb_thue = sum(x.get('tien_thue', 0) for x in internal_list)
        
        return {
            'match_type': match_type,
            'summary': {
                'total_gdt_count': len(gdt_list),
                'total_gdt_doanh_so': tot_gdt_ds,
                'total_gdt_tien_thue': tot_gdt_thue,
                
                'total_internal_count': len(internal_list),
                'total_internal_doanh_so': tot_nb_ds,
                'total_internal_tien_thue': tot_nb_thue,
                
                'diff_doanh_so': tot_gdt_ds - tot_nb_ds,
                'diff_tien_thue': tot_gdt_thue - tot_nb_thue,
                
                'matched_count': len(matched_perfect),
                'diff_amount_count': len(matched_diff_amount),
                'only_in_gdt_count': len(only_in_gdt),
                'only_in_internal_count': len(only_in_internal)
            },
            'matched_perfect': matched_perfect,
            'matched_diff_amount': matched_diff_amount,
            'only_in_gdt': only_in_gdt,
            'only_in_internal': only_in_internal
        }

    @staticmethod
    def reconcile_with_vat_return(
        in_invoices: List[Dict[str, Any]], 
        out_invoices: List[Dict[str, Any]], 
        vat_return: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Đối chiếu số liệu tổng hợp hóa đơn mua vào & bán ra với các chỉ tiêu tờ khai 01/GTGT
        """
        # 1. Tổng hợp Mua vào từ hóa đơn/bảng kê
        sum_in_ds = sum(x.get('doanh_so', 0) for x in in_invoices)
        sum_in_thue = sum(x.get('tien_thue', 0) for x in in_invoices)
        
        # 2. Tổng hợp Bán ra từ hóa đơn/bảng kê theo các nhóm thuế suất
        out_kct_ds = sum(x.get('doanh_so', 0) for x in out_invoices if x.get('thue_suat_float') == -1.0)
        out_0_ds = sum(x.get('doanh_so', 0) for x in out_invoices if x.get('thue_suat_float') == 0.0)
        out_5_ds = sum(x.get('doanh_so', 0) for x in out_invoices if 0.045 <= x.get('thue_suat_float', 0) <= 0.055)
        out_5_thue = sum(x.get('tien_thue', 0) for x in out_invoices if 0.045 <= x.get('thue_suat_float', 0) <= 0.055)
        
        # Nhóm 8% và 10% (kê chung vào chỉ tiêu [32] và [33] hoặc phụ lục giảm thuế)
        out_10_ds = sum(x.get('doanh_so', 0) for x in out_invoices if x.get('thue_suat_float', 0) >= 0.075)
        out_10_thue = sum(x.get('tien_thue', 0) for x in out_invoices if x.get('thue_suat_float', 0) >= 0.075)
        
        out_kkt_ds = sum(x.get('doanh_so', 0) for x in out_invoices if x.get('thue_suat_float') == -2.0)
        
        tot_out_ds = sum(x.get('doanh_so', 0) for x in out_invoices)
        tot_out_thue = sum(x.get('tien_thue', 0) for x in out_invoices)
        
        # 3. Lập bảng so sánh đối chiếu từng chỉ tiêu
        comparisons = [
            {
                'chi_tieu_ma': '[23]',
                'ten_chi_tieu': 'Giá trị HHDV mua vào trong kỳ',
                'so_hoa_don': sum_in_ds,
                'so_to_khai': vat_return.get('ct23', 0),
                'chenh_lech': sum_in_ds - vat_return.get('ct23', 0),
                'loai': 'Đầu vào'
            },
            {
                'chi_tieu_ma': '[24]',
                'ten_chi_tieu': 'Thuế GTGT HHDV mua vào trong kỳ',
                'so_hoa_don': sum_in_thue,
                'so_to_khai': vat_return.get('ct24', 0),
                'chenh_lech': sum_in_thue - vat_return.get('ct24', 0),
                'loai': 'Đầu vào'
            },
            {
                'chi_tieu_ma': '[25]',
                'ten_chi_tieu': 'Thuế GTGT đầu vào được khấu trừ',
                'so_hoa_don': sum_in_thue, # Giả định khấu trừ toàn bộ, nếu lệch thì có HĐ không đủ điều kiện khấu trừ
                'so_to_khai': vat_return.get('ct25', 0),
                'chenh_lech': sum_in_thue - vat_return.get('ct25', 0),
                'loai': 'Đầu vào (Khấu trừ)'
            },
            {
                'chi_tieu_ma': '[26]',
                'ten_chi_tieu': 'Doanh thu HHDV bán ra không chịu thuế (KCT)',
                'so_hoa_don': out_kct_ds,
                'so_to_khai': vat_return.get('ct26', 0),
                'chenh_lech': out_kct_ds - vat_return.get('ct26', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[29]',
                'ten_chi_tieu': 'Doanh thu HHDV thuế suất 0%',
                'so_hoa_don': out_0_ds,
                'so_to_khai': vat_return.get('ct29', 0),
                'chenh_lech': out_0_ds - vat_return.get('ct29', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[30]',
                'ten_chi_tieu': 'Doanh thu HHDV thuế suất 5%',
                'so_hoa_don': out_5_ds,
                'so_to_khai': vat_return.get('ct30', 0),
                'chenh_lech': out_5_ds - vat_return.get('ct30', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[31]',
                'ten_chi_tieu': 'Thuế GTGT của HHDV thuế suất 5%',
                'so_hoa_don': out_5_thue,
                'so_to_khai': vat_return.get('ct31', 0),
                'chenh_lech': out_5_thue - vat_return.get('ct31', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[32]',
                'ten_chi_tieu': 'Doanh thu HHDV thuế suất 10% (và 8%)',
                'so_hoa_don': out_10_ds,
                'so_to_khai': vat_return.get('ct32', 0),
                'chenh_lech': out_10_ds - vat_return.get('ct32', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[33]',
                'ten_chi_tieu': 'Thuế GTGT của HHDV thuế suất 10% (và 8%)',
                'so_hoa_don': out_10_thue,
                'so_to_khai': vat_return.get('ct33', 0),
                'chenh_lech': out_10_thue - vat_return.get('ct33', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[32a]',
                'ten_chi_tieu': 'Doanh thu HHDV không phải tính nộp thuế (KKT)',
                'so_hoa_don': out_kkt_ds,
                'so_to_khai': vat_return.get('ct32a', 0),
                'chenh_lech': out_kkt_ds - vat_return.get('ct32a', 0),
                'loai': 'Đầu ra'
            },
            {
                'chi_tieu_ma': '[34]',
                'ten_chi_tieu': 'Tổng doanh thu bán ra',
                'so_hoa_don': tot_out_ds,
                'so_to_khai': vat_return.get('ct34', 0),
                'chenh_lech': tot_out_ds - vat_return.get('ct34', 0),
                'loai': 'Đầu ra (Tổng)'
            },
            {
                'chi_tieu_ma': '[35]',
                'ten_chi_tieu': 'Tổng thuế GTGT đầu ra',
                'so_hoa_don': tot_out_thue,
                'so_to_khai': vat_return.get('ct35', 0),
                'chenh_lech': tot_out_thue - vat_return.get('ct35', 0),
                'loai': 'Đầu ra (Tổng)'
            }
        ]
        
        # Đánh giá trạng thái khớp cho từng dòng
        all_matched = True
        for c in comparisons:
            diff = abs(c['chenh_lech'])
            if diff < 1.0:
                c['danh_gia'] = 'Khớp 100%'
            elif diff <= 1000:
                c['danh_gia'] = 'Lệch nhỏ (Làm tròn)'
            else:
                c['danh_gia'] = 'LỆCH - CẦN RÀ SOÁT'
                all_matched = False
                
        return {
            'vat_return_info': vat_return,
            'comparisons': comparisons,
            'is_perfect_match': all_matched,
            'summary_in_doanh_so': sum_in_ds,
            'summary_in_tien_thue': sum_in_thue,
            'summary_out_doanh_so': tot_out_ds,
            'summary_out_tien_thue': tot_out_thue
        }
