# -*- coding: utf-8 -*-
"""
test_suite.py
-------------
Bộ kiểm thử tự động toàn diện (Automated Test Suite) cho VAT Reconciliation Pro v1.1:
1. Test ExcelNormalizer: Kiểm tra chuẩn hóa nhiều file, cả thư mục, workbook nhiều sheet (bảng kê theo cả kỳ hoàn / nhiều năm) và tự động khử trùng lặp khi tải chồng file TCT.
2. Test TaxXmlParser: Kiểm tra đọc 1 file XML và đọc CẢ THƯ MỤC chứa nhiều tờ khai XML nhiều kỳ + tờ khai bổ sung + lọc bỏ tờ khai sắc thuế khác (TNDN/TNCN).
3. Test Reconciler & Exporter: Kiểm tra đối chiếu khớp HĐ, phát hiện lệch tiền, bỏ sót, đối chiếu tổng hợp nhiều kỳ tờ khai và xuất Excel có sheet 'Chi tiết Tờ khai theo Kỳ'.
"""

import os
import sys
import shutil
import tempfile
import openpyxl

from excel_normalizer import ExcelNormalizer
from xml_parser import TaxXmlParser
from reconciler import VatReconciler
from excel_exporter import ReconciliationReportExporter
from mst_lookup import MstLookupEngine, normalize_mst_code, classify_tax_status


def _create_sample_excel(file_path: str, sheets_data: dict):
    wb = openpyxl.Workbook()
    default_ws = wb.active
    for sname, rows in sheets_data.items():
        ws = wb.create_sheet(title=sname)
        for r in rows:
            ws.append(r)
    if len(wb.sheetnames) > 1 and default_ws.title in wb.sheetnames:
        wb.remove(default_ws)
    wb.save(file_path)


def test_excel_normalizer_multi_files_and_folders():
    print("=== 1. TEST EXCEL NORMALIZER (NHIỀU FILE / CẢ THƯ MỤC / NHIỀU SHEET) ===")
    tmp_dir = tempfile.mkdtemp(prefix="vat_test_excel_")
    try:
        # 1. Tạo thư mục chứa 2 file HĐĐT tải từ TCT (Tháng 07 & Tháng 08, có 1 HĐ bị trùng giữa 2 file)
        gdt_dir = os.path.join(tmp_dir, "GDT_MuaVao")
        os.makedirs(gdt_dir, exist_ok=True)

        header = ["STT", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập", "MST người bán", "Tên người bán", "Tổng tiền chưa thuế", "Tổng tiền thuế"]
        _create_sample_excel(os.path.join(gdt_dir, "TCT_T07_2024.xlsx"), {
            "DanhSach": [
                ["DANH SÁCH HÓA ĐƠN ĐIỆN TỬ MUA VÀO THÁNG 07/2024"],
                header,
                [1, "1C24TAA", "0000101", "15/07/2024", "0100000001", "CTY A", "10,000,000", "1,000,000"],
                [2, "1C24TAB", "0000102", "20/07/2024", "0100000002", "CTY B", "20,000,000", "1,600,000"],
            ]
        })
        _create_sample_excel(os.path.join(gdt_dir, "TCT_T08_2024.xlsx"), {
            "DanhSach": [
                ["DANH SÁCH HÓA ĐƠN ĐIỆN TỬ MUA VÀO THÁNG 08/2024"],
                header,
                # HĐ 102 bị tải trùng sang file tháng 8
                [1, "1C24TAB", "0000102", "20/07/2024", "0100000002", "CTY B", "20,000,000", "1,600,000"],
                [2, "1C24TAC", "0000205", "10/08/2024", "0100000003", "CTY C", "50,000,000", "5,000,000"],
            ]
        })

        gdt_records = ExcelNormalizer.normalize_multiple_sources(gdt_dir, deduplicate=True)
        print(f"[*] Quét cả thư mục TCT (2 file, 1 HĐ trùng): Thu được {len(gdt_records)} HĐ duy nhất.")
        assert len(gdt_records) == 3, f"Kỳ vọng 3 HĐ sau khử trùng lặp, nhận được {len(gdt_records)}"

        # 2. Tạo file Bảng kê nội bộ cả kỳ hoàn gồm 2 sheet (Sheet Q1_2024 và Sheet Q2_2024) + 1 sheet Bìa
        bk_path = os.path.join(tmp_dir, "BangKe_CaKyHoan_2024.xlsx")
        bk_header = ["Ký hiệu HĐ", "Số HĐ", "Ngày HĐ", "Mã số thuế", "Tên đơn vị", "Doanh số mua chưa có thuế", "Thuế suất", "Thuế GTGT"]
        _create_sample_excel(bk_path, {
            "Bia_HoSo": [["BẢNG KÊ HÓA ĐƠN ĐỀ NGHỊ HOÀN THUẾ"]],
            "Thang_07": [
                bk_header,
                ["1C24TAA", "101", "2024-07-15", "0100000001", "CTY A", 10000000, "10%", 1000000],
                ["1C24TAB", "102", "2024-07-20", "0100000002", "CTY B", 20000000, "8%", 1600000],
            ],
            "Thang_08": [
                bk_header,
                ["1C24TAC", "205", "2024-08-10", "0100000003", "CTY C", 50000000, "10%", 5000000],
            ]
        })

        bk_records = ExcelNormalizer.normalize_multiple_sources(bk_path, deduplicate=False)
        print(f"[*] Quét Bảng kê nhiều sheet (Bỏ qua sheet bìa): Thu được {len(bk_records)} HĐ.")
        assert len(bk_records) == 3, f"Kỳ vọng 3 HĐ từ 2 sheet dữ liệu, nhận được {len(bk_records)}"
        assert sum(r['doanh_so'] for r in bk_records) == 80000000.0
        assert sum(r['tien_thue'] for r in bk_records) == 7600000.0
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    print(">>> PASS: ExcelNormalizer xử lý hoàn hảo nhiều file, cả thư mục và nhiều sheet.\n")


def test_xml_parser_single_and_folder():
    print("=== 2. TEST TAX XML PARSER (1 KỲ & CẢ THƯ MỤC NHIỀU KỲ / BỔ SUNG) ===")
    tmp_dir = tempfile.mkdtemp(prefix="vat_test_xml_")
    try:
        # Kỳ T07/2024: Bản lần đầu (lanKhai=0) và Bản bổ sung lần 2 (soLan=2)
        xml_t07_L0 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
    <HSoKhaiThue>
        <TTinChung><TTinTKhaiThue>
            <TKhaiThue><maTKhai>842</maTKhai><tenTKhai>Tờ khai thuế GTGT (01/GTGT)</tenTKhai>
                <KyKKhaiThue><kyKKhai>07/2024</kyKKhai></KyKKhaiThue>
                <soLan>0</soLan><ngayLapTKhai>2024-08-15</ngayLapTKhai>
            </TKhaiThue>
            <NNT><mst>0100000001</mst><tenNNT>CÔNG TY CỔ PHẦN ABC</tenNNT></NNT>
        </TTinTKhaiThue></TTinChung>
        <CTieuTKhaiChinh>
            <ct22>5000000</ct22><ct23>10000000</ct23><ct24>1000000</ct24><ct25>1000000</ct25>
            <ct32>0</ct32><ct33>0</ct33><ct34>0</ct34><ct35>0</ct35><ct43>6000000</ct43>
        </CTieuTKhaiChinh>
    </HSoKhaiThue>
</HSoThueDTu>"""

        xml_t07_L2 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
    <HSoKhaiThue>
        <TTinChung><TTinTKhaiThue>
            <TKhaiThue><maTKhai>842</maTKhai><tenTKhai>Tờ khai thuế GTGT (01/GTGT)</tenTKhai>
                <KyKKhaiThue><kyKKhai>07/2024</kyKKhai></KyKKhaiThue>
                <soLan>2</soLan><ngayLapTKhai>2024-09-05</ngayLapTKhai>
            </TKhaiThue>
            <NNT><mst>0100000001</mst><tenNNT>CÔNG TY CỔ PHẦN ABC</tenNNT></NNT>
        </TTinTKhaiThue></TTinChung>
        <CTieuTKhaiChinh>
            <ct22>5000000</ct22><ct23>30000000</ct23><ct24>2600000</ct24><ct25>2600000</ct25>
            <ct32>15000000</ct32><ct33>1500000</ct33><ct34>15000000</ct34><ct35>1500000</ct35><ct43>6100000</ct43>
        </CTieuTKhaiChinh>
        <PLuc><PL01_1_KHBS><ct23>20000000</ct23></PL01_1_KHBS></PLuc>
    </HSoKhaiThue>
</HSoThueDTu>"""

        # Kỳ T08/2024: Bản lần đầu
        xml_t08_L0 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
    <HSoKhaiThue>
        <TTinChung><TTinTKhaiThue>
            <TKhaiThue><maTKhai>842</maTKhai><tenTKhai>Tờ khai thuế GTGT (01/GTGT)</tenTKhai>
                <KyKKhaiThue><kyKKhai>08/2024</kyKKhai></KyKKhaiThue>
                <soLan>0</soLan><ngayLapTKhai>2024-09-18</ngayLapTKhai>
            </TKhaiThue>
            <NNT><mst>0100000001</mst><tenNNT>CÔNG TY CỔ PHẦN ABC</tenNNT></NNT>
        </TTinTKhaiThue></TTinChung>
        <CTieuTKhaiChinh>
            <ct22>6100000</ct22><ct23>50000000</ct23><ct24>5000000</ct24><ct25>5000000</ct25>
            <ct32>25000000</ct32><ct33>2500000</ct33><ct34>25000000</ct34><ct35>2500000</ct35><ct43>8600000</ct43>
        </CTieuTKhaiChinh>
    </HSoKhaiThue>
</HSoThueDTu>"""

        # 1 file XML tờ khai TNDN nằm lẫn trong cùng thư mục -> phải tự động lọc bỏ
        xml_tndn = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
    <HSoKhaiThue>
        <TTinChung><TTinTKhaiThue>
            <TKhaiThue><maTKhai>892</maTKhai><tenTKhai>Tờ khai quyết toán thuế TNDN (03/TNDN)</tenTKhai></TKhaiThue>
        </TTinTKhaiThue></TTinChung>
    </HSoKhaiThue>
</HSoThueDTu>"""

        with open(os.path.join(tmp_dir, "GTGT_T07_L0.xml"), "w", encoding="utf-8") as f: f.write(xml_t07_L0)
        with open(os.path.join(tmp_dir, "GTGT_T07_L2.xml"), "w", encoding="utf-8") as f: f.write(xml_t07_L2)
        with open(os.path.join(tmp_dir, "GTGT_T08_L0.xml"), "w", encoding="utf-8") as f: f.write(xml_t08_L0)
        with open(os.path.join(tmp_dir, "TNDN_2024.xml"), "w", encoding="utf-8") as f: f.write(xml_tndn)

        agg = TaxXmlParser.parse_multiple_vat_returns(tmp_dir)
        assert agg is not None
        print(f"[*] Tổng hợp thư mục XML: Quét {agg['total_xml_scanned']} XML GTGT -> Chốt {agg['effective_periods_count']} kỳ ({agg['ky_tinh_thue']})")
        assert agg['total_xml_scanned'] == 3, "Phải lọc bỏ file TNDN, chỉ nhận 3 file GTGT"
        assert agg['effective_periods_count'] == 2, "Phải chốt đúng 2 kỳ (07/2024 và 08/2024)"
        assert agg['ct23'] == 80000000.0, f"Tổng [23] phải là 30tr (T07_L2) + 50tr (T08_L0) = 80tr, nhận được {agg['ct23']}"
        assert agg['ct24'] == 7600000.0, f"Tổng [24] phải là 2.6tr + 5tr = 7.6tr, nhận được {agg['ct24']}"
        assert agg['ct32'] == 40000000.0, f"Tổng [32] phải là 15tr + 25tr = 40tr, nhận được {agg['ct32']}"
        assert agg['ct22'] == 5000000.0, "[22] đầu giai đoạn phải là 5tr"
        assert agg['ct43'] == 8600000.0, "[43] cuối giai đoạn phải là 8.6tr"
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    print(">>> PASS: TaxXmlParser tự động chọn bản bổ sung cuối cùng và cộng gộp nhiều kỳ chuẩn xác.\n")


def test_reconciler_and_export():
    print("=== 3. TEST RECONCILER & EXPORTER (ĐA KỲ) ===")
    gdt_invoices = [
        {'ky_hieu': '1C24TAC', 'so_hd': '00004671', 'ngay_hd': '2024-07-17', 'mst': '0100000001', 'ten_doi_tac': 'CTY A', 'doanh_so': 30000000.0, 'tien_thue': 2600000.0, 'dien_giai': 'HĐ kỳ T07'},
        {'ky_hieu': '1K24THN', 'so_hd': '00036822', 'ngay_hd': '2024-08-22', 'mst': '0100000002', 'ten_doi_tac': 'CTY B', 'doanh_so': 50000000.0, 'tien_thue': 5000000.0, 'dien_giai': 'HĐ kỳ T08'},
        {'ky_hieu': '1K24THN', 'so_hd': '00099999', 'ngay_hd': '2024-08-25', 'mst': '0100000003', 'ten_doi_tac': 'CTY C', 'doanh_so': 500000.0, 'tien_thue': 50000.0, 'dien_giai': 'HĐ bỏ sót trên bảng kê'}
    ]
    
    internal_invoices = [
        {'ky_hieu': '1C24TAC', 'so_hd': '4671', 'ngay_hd': '2024-07-17', 'mst': '0100000001', 'ten_doi_tac': 'CTY A', 'doanh_so': 30000000.0, 'tien_thue': 2600000.0, 'dien_giai': 'HĐ kỳ T07 (Khớp 100%)'},
        {'ky_hieu': '1K24THN', 'so_hd': '36822', 'ngay_hd': '2024-08-22', 'mst': '0100000002', 'ten_doi_tac': 'CTY B', 'doanh_so': 50000000.0, 'tien_thue': 4900000.0, 'dien_giai': 'Lệch tiền thuế nội bộ nhập nhầm'},
        {'ky_hieu': '2K24THN', 'so_hd': '00088888', 'ngay_hd': '2024-08-28', 'mst': '0109999999', 'ten_doi_tac': 'CTY D', 'doanh_so': 2000000.0, 'tien_thue': 200000.0, 'dien_giai': 'Chỉ có ở nội bộ'}
    ]
    
    res = VatReconciler.match_invoices(gdt_invoices, internal_invoices, match_type="Đầu vào")
    s = res['summary']
    assert s['matched_count'] == 1
    assert s['diff_amount_count'] == 1
    assert s['only_in_gdt_count'] == 1
    assert s['only_in_internal_count'] == 1
    
    mock_vat_return = {
        'ten_nnt': 'CÔNG TY CỔ PHẦN ABC', 'ma_so_thue': '0100000001',
        'ky_tinh_thue': '07/2024 -> 08/2024 (2 kỳ)', 'lan_ke_khai': 'Chốt bản cuối mỗi kỳ (2 kỳ)',
        'ct23': 80000000.0, 'ct24': 7600000.0, 'ct25': 7600000.0,
        'ct26': 0, 'ct29': 0, 'ct30': 0, 'ct31': 0, 'ct32': 0, 'ct33': 0, 'ct32a': 0, 'ct34': 0, 'ct35': 0,
        'periods_detail': [
            {'period_key': '07/2024', 'lan_ke_khai': 2, 'ngay_nop': '2024-09-05', 'file_name': 'GTGT_T07_L2.xml', 'ct22': 5000000, 'ct23': 30000000, 'ct24': 2600000, 'ct25': 2600000, 'ct43': 6100000},
            {'period_key': '08/2024', 'lan_ke_khai': 0, 'ngay_nop': '2024-09-18', 'file_name': 'GTGT_T08_L0.xml', 'ct22': 6100000, 'ct23': 50000000, 'ct24': 5000000, 'ct25': 5000000, 'ct43': 8600000},
        ]
    }
    
    ret_rec = VatReconciler.reconcile_with_vat_return(gdt_invoices[:2], [], mock_vat_return)
    assert ret_rec['is_perfect_match'] is True
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        out_file = os.path.join(tmp_dir, "test_output_report.xlsx")
        ReconciliationReportExporter.export_report(out_file, in_reconciliation=res, return_reconciliation=ret_rec)
        assert os.path.exists(out_file)
        wb = openpyxl.load_workbook(out_file, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        print(f"[*] Sheets xuất ra: {sheet_names}")
        assert 'Đối chiếu Tờ khai GTGT' in sheet_names
        assert 'Chi tiết Tờ khai theo Kỳ' in sheet_names
        assert 'Đối chiếu Đầu vào' in sheet_names
    
    print(">>> PASS: Thuật toán đối chiếu và xuất báo cáo Excel đa kỳ hoạt động hoàn hảo.\n")


def test_mst_lookup_and_directory_export():
    print("=== 4. TEST MST DIRECTORY LOOKUP & EXCEL EXPORT ===")
    # 1. Kiểm tra chuẩn hóa MST 10 số, 13 số, bù số 0 ở đầu
    assert normalize_mst_code("100109106") == "0100109106"
    assert normalize_mst_code("0100109106001") == "0100109106-001"
    assert normalize_mst_code("0100109106-001") == "0100109106-001"

    # 2. Kiểm tra trích xuất MST từ văn bản và từ file Excel bảng kê
    txt = "0100109106, 0300588569; 0100109106-001\n0100109106"
    extracted = MstLookupEngine.extract_msts_from_text(txt)
    assert extracted == ["0100109106", "0300588569", "0100109106-001"]

    # 3. Kiểm tra phân loại rủi ro tình trạng hoạt động
    r_ok, _ = classify_tax_status("NNT đang hoạt động (đã được cấp GCN ĐKT)")
    assert "An toàn" in r_ok
    r_high1, _ = classify_tax_status("NNT không hoạt động tại địa chỉ đã đăng ký")
    assert "RỦI RO CAO" in r_high1
    r_high2, _ = classify_tax_status("NNT ngừng HĐ nhưng chưa hoàn thành thủ tục chấm dứt hiệu lực MST")
    assert "RỦI RO CAO" in r_high2

    # 4. Kiểm tra xuất file Excel Danh bạ MST từ kết quả mẫu
    sample_results = [
        {
            "mst": "0100000001",
            "ten_nnt": "CÔNG TY CỔ PHẦN MẪU A",
            "ten_quoc_te": "SAMPLE A CORP",
            "ten_viet_tat": "SAMPLE A",
            "dia_chi": "Số 1 Đường Mẫu, Quận Cầu Giấy, Hà Nội",
            "nguoi_dai_dien": "Nguyễn Văn A",
            "tinh_trang": "NNT đang hoạt động",
            "phan_loai_rui_ro": "Đang hoạt động (An toàn)",
            "khuyen_nghi": "Đủ điều kiện trạng thái MST đang hoạt động bình thường.",
            "nguon_du_lieu": "GDT (gdt.gov.vn)",
            "cap_nhat_gdt": "2026-11-10",
            "found": True,
        },
        {
            "mst": "0100000002-001",
            "ten_nnt": "CHI NHÁNH CÔNG TY MẪU B",
            "ten_quoc_te": "",
            "ten_viet_tat": "",
            "dia_chi": "Số 2 Đường Mẫu, TP Hồ Chí Minh",
            "nguoi_dai_dien": "",
            "tinh_trang": "NNT không hoạt động tại địa chỉ đã đăng ký",
            "phan_loai_rui_ro": "RỦI RO CAO (Bỏ địa chỉ ĐKKD)",
            "khuyen_nghi": "Cảnh báo: NNT bỏ địa chỉ kinh doanh!",
            "nguon_du_lieu": "GDT (gdt.gov.vn)",
            "cap_nhat_gdt": "2026-11-16",
            "found": True,
        },
    ]

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_path = os.path.join(tmp_dir, "DanhBa_MST_Test.xlsx")
        MstLookupEngine.export_directory_to_excel(sample_results, out_path)
        assert os.path.exists(out_path)
        wb = openpyxl.load_workbook(out_path, read_only=True)
        assert "Danh bạ Doanh nghiệp (MST)" in wb.sheetnames
        ws = wb["Danh bạ Doanh nghiệp (MST)"]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        assert len(rows) >= 6
        assert rows[4][1] == "0100000001"
        assert rows[5][1] == "0100000002-001"

    print(">>> PASS: Tra cứu Danh bạ Mã số thuế & Xuất Excel cảnh báo rủi ro hoạt động hoàn hảo.\n")


if __name__ == '__main__':
    test_excel_normalizer_multi_files_and_folders()
    test_xml_parser_single_and_folder()
    test_reconciler_and_export()
    test_mst_lookup_and_directory_export()
    print("=================================================================")
    print(" TẤT CẢ 4 TEST CASES ĐỀU PASS 100%! HỆ THỐNG SẴN SÀNG SẢN XUẤT. ")
    print("=================================================================")
