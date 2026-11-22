# -*- coding: utf-8 -*-
"""
test_suite.py
-------------
Bộ kiểm thử tự động toàn diện (Automated Test Suite) cho VAT Reconciliation Pro:
1. Test ExcelNormalizer: Kiểm tra khả năng chuẩn hóa 3 file Excel thực tế có format khác nhau hoàn toàn.
2. Test XML Parser: Tạo mock XML tờ khai 01/GTGT chuẩn HTKK và parse kiểm tra các chỉ tiêu.
3. Test Reconciler: Kiểm tra đối chiếu khớp HĐ, phát hiện lệch tiền, bỏ sót, và đối chiếu tờ khai.
4. Test Exporter: Xuất file Excel kết quả và xác thực XML OpenPyXL.
"""

import os
import sys
import tempfile
import openpyxl
from excel_normalizer import ExcelNormalizer
from xml_parser import TaxXmlParser
from reconciler import VatReconciler
from excel_exporter import ReconciliationReportExporter

def test_excel_normalizer():
    print("=== 1. TEST EXCEL NORMALIZER ===")
    sample_files = [
        (r'C:\Users\Admin\.gemini\antigravity\brain\b0dacb17-cdbc-41ec-9606-07d11a8ae0d6\.user_uploaded\media_1794904719222_09176e07.xlsx', 1240),
        (r'C:\Users\Admin\.gemini\antigravity\brain\b0dacb17-cdbc-41ec-9606-07d11a8ae0d6\.user_uploaded\media_1794971175774_8d0647ba.xlsx', 3375),
        (r'C:\Users\Admin\.gemini\antigravity\brain\b0dacb17-cdbc-41ec-9606-07d11a8ae0d6\.user_uploaded\media_1794885503416_00dc8492.xlsx', 7210)
    ]
    for fp, expected_rows in sample_files:
        assert os.path.exists(fp), f"File không tồn tại: {fp}"
        records = ExcelNormalizer.normalize_invoice_sheet(fp)
        print(f"[*] File {os.path.basename(fp)}: Chuẩn hóa {len(records):,} dòng (Kỳ vọng ~{expected_rows})")
        assert len(records) >= expected_rows - 5, "Số dòng chuẩn hóa bị thiếu!"
        # Kiểm tra kiểu dữ liệu
        first = records[0]
        assert isinstance(first['doanh_so'], float), "Doanh số không phải float!"
        assert isinstance(first['tien_thue'], float), "Tiền thuế không phải float!"
        assert first['so_hd'] != '', "Số hóa đơn bị rỗng!"
    print(">>> PASS: ExcelNormalizer chuẩn hóa xuất sắc mọi định dạng.\n")

def test_xml_parser():
    print("=== 2. TEST TAX XML PARSER ===")
    # Tạo file XML tờ khai 01/GTGT giả lập theo đúng cấu trúc HTKK
    mock_xml = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhai">
    <HSoKhaiThue>
        <TTinChung>
            <mst>0314066040</mst>
            <tenNNT>CÔNG TY CỔ PHẦN MANNY</tenNNT>
            <kyTinhThue>Tháng 08/2024</kyTinhThue>
            <lanKhai>0</lanKhai>
            <ngayNop>2024-09-20</ngayNop>
        </TTinChung>
        <CTietTKhai>
            <ct22>15000000</ct22>
            <ct23>1233285</ct23>
            <ct24>123329</ct24>
            <ct25>123329</ct25>
            <ct26>0</ct26>
            <ct29>0</ct29>
            <ct30>0</ct30>
            <ct31>0</ct31>
            <ct32>50000000</ct32>
            <ct33>5000000</ct33>
            <ct34>50000000</ct34>
            <ct35>5000000</ct35>
            <ct36>4876671</ct36>
            <ct40>0</ct40>
            <ct43>10123329</ct43>
        </CTietTKhai>
    </HSoKhaiThue>
</HSoThueDTu>
"""
    with tempfile.NamedTemporaryFile('w', suffix='.xml', delete=False, encoding='utf-8') as tf:
        tf.write(mock_xml)
        tmp_xml_path = tf.name
        
    try:
        ret = TaxXmlParser.parse_vat_return(tmp_xml_path)
        print(f"[*] Parse XML: NNT={ret['ten_nnt']}, MST={ret['ma_so_thue']}, [23]={ret['ct23']:,}, [34]={ret['ct34']:,}")
        assert ret['ma_so_thue'] == '0314066040'
        assert ret['ct23'] == 1233285.0
        assert ret['ct24'] == 123329.0
        assert ret['ct32'] == 50000000.0
        assert ret['ct43'] == 10123329.0
    finally:
        os.remove(tmp_xml_path)
    print(">>> PASS: TaxXmlParser phân tích XML chuẩn xác.\n")

def test_reconciler_and_export():
    print("=== 3. TEST RECONCILER & EXPORTER ===")
    # Dữ liệu giả lập 2 bên
    gdt_invoices = [
        {'ky_hieu': '1C24TAC', 'so_hd': '00004671', 'ngay_hd': '2024-08-17', 'mst': '0102900049', 'ten_doi_tac': 'CTY A', 'doanh_so': 1233285.0, 'tien_thue': 123329.0, 'dien_giai': 'Du lịch'},
        {'ky_hieu': '1K24THN', 'so_hd': '00036822', 'ngay_hd': '2024-04-22', 'mst': '0100109106', 'ten_doi_tac': 'CTY B', 'doanh_so': 909091.0, 'tien_thue': 90909.0, 'dien_giai': 'Xăng xe'},
        {'ky_hieu': '1K24THN', 'so_hd': '00099999', 'ngay_hd': '2024-08-20', 'mst': '0100109106', 'ten_doi_tac': 'CTY C', 'doanh_so': 500000.0, 'tien_thue': 50000.0, 'dien_giai': 'HĐ bỏ sót trên bảng kê'}
    ]
    
    internal_invoices = [
        {'ky_hieu': '1C24TAC', 'so_hd': '4671', 'ngay_hd': '2024-08-17', 'mst': '0102900049', 'ten_doi_tac': 'CTY A', 'doanh_so': 1233285.0, 'tien_thue': 123329.0, 'dien_giai': 'Du lịch (Khớp 100%)'},
        {'ky_hieu': '1K24THN', 'so_hd': '36822', 'ngay_hd': '2024-04-22', 'mst': '0100109106', 'ten_doi_tac': 'CTY B', 'doanh_so': 909091.0, 'tien_thue': 99999.0, 'dien_giai': 'Lệch tiền thuế nội bộ nhập nhầm'},
        {'ky_hieu': '2K24THN', 'so_hd': '00088888', 'ngay_hd': '2024-08-25', 'mst': '0109999999', 'ten_doi_tac': 'CTY D', 'doanh_so': 2000000.0, 'tien_thue': 200000.0, 'dien_giai': 'Chỉ có ở nội bộ, TCT không có'}
    ]
    
    res = VatReconciler.match_invoices(gdt_invoices, internal_invoices, match_type="Đầu vào")
    s = res['summary']
    print(f"[*] Kết quả đối chiếu HĐ: Khớp 100%={s['matched_count']}, Lệch tiền={s['diff_amount_count']}, Thừa GDT={s['only_in_gdt_count']}, Thừa Nội bộ={s['only_in_internal_count']}")
    
    assert s['matched_count'] == 1, "Số HĐ khớp 100% phải là 1"
    assert s['diff_amount_count'] == 1, "Số HĐ lệch tiền phải là 1"
    assert s['only_in_gdt_count'] == 1, "Số HĐ thừa GDT phải là 1"
    assert s['only_in_internal_count'] == 1, "Số HĐ thừa Nội bộ phải là 1"
    
    # Test đối chiếu với tờ khai
    mock_vat_return = {
        'ten_nnt': 'CÔNG TY CỔ PHẦN MANNY', 'ma_so_thue': '0314066040', 'ky_tinh_thue': 'T08/2024', 'lan_ke_khai': 0,
        'ct23': 2142376.0, 'ct24': 214238.0, 'ct25': 214238.0,
        'ct26': 0, 'ct29': 0, 'ct30': 0, 'ct31': 0, 'ct32': 0, 'ct33': 0, 'ct32a': 0, 'ct34': 0, 'ct35': 0
    }
    
    ret_rec = VatReconciler.reconcile_with_vat_return(gdt_invoices[:2], [], mock_vat_return)
    assert ret_rec['is_perfect_match'] == True, "Đối chiếu tờ khai phải khớp 100%!"
    print(f"[*] Đối chiếu tờ khai: Khớp 100% [23] ({ret_rec['summary_in_doanh_so']:,} == {mock_vat_return['ct23']:,})")
    
    # Test xuất file Excel
    out_file = r'C:\Users\Admin\.gemini\antigravity\scratch\vat_reconciliation_pro\test_output_report.xlsx'
    ReconciliationReportExporter.export_report(out_file, in_reconciliation=res, return_reconciliation=ret_rec)
    assert os.path.exists(out_file), "Không tạo được file Excel!"
    
    wb = openpyxl.load_workbook(out_file, read_only=True)
    print(f"[*] Đã xuất file Excel kiểm thử: {out_file} (Sheets: {wb.sheetnames})")
    assert 'Đối chiếu Tờ khai GTGT' in wb.sheetnames
    assert 'Đối chiếu Đầu vào' in wb.sheetnames
    
    print(">>> PASS: Thuật toán đối chiếu và xuất báo cáo Excel hoạt động hoàn hảo.\n")

if __name__ == '__main__':
    test_excel_normalizer()
    test_xml_parser()
    test_reconciler_and_export()
    print("=================================================================")
    print(" TẤT CẢ 3 TEST CASES ĐỀU PASS 100%! HỆ THỐNG SẴN SÀNG SẢN XUẤT. ")
    print("=================================================================")
