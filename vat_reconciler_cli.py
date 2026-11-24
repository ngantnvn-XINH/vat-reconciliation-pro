# -*- coding: utf-8 -*-
"""
vat_reconciler_cli.py
---------------------
Giao diện dòng lệnh (CLI) tự động hóa:
1. Chế độ Đối chiếu GTGT:
   python vat_reconciler_cli.py --in-gdt <file/folder> --in-nb <file/folder> --out-gdt <file/folder> --out-nb <file/folder> --xml <file/folder_xml> --out <ket_qua.xlsx>
2. Chế độ Tra cứu Danh bạ Mã số thuế & Tình trạng hoạt động (Cục Thuế):
   python vat_reconciler_cli.py --mst-list "0100109106,0300588569" --mst-out DanhBa_MST.xlsx
   hoặc:
   python vat_reconciler_cli.py --mst-file <file_hoac_thu_muc_excel_chua_mst> --mst-out DanhBa_MST.xlsx
"""

import os
import sys
import argparse
from excel_normalizer import ExcelNormalizer
from xml_parser import TaxXmlParser
from reconciler import VatReconciler
from excel_exporter import ReconciliationReportExporter
from mst_lookup import MstLookupEngine


def main():
    parser = argparse.ArgumentParser(description="Phần mềm Đối chiếu Thuế GTGT Đa chiều & Tra cứu Danh bạ Mã số thuế (GDT)")
    parser.add_argument('--in-gdt', help="File hoặc Thư mục Excel HĐĐT Mua vào tải từ TCT", default=None)
    parser.add_argument('--in-nb', help="File hoặc Thư mục Excel Bảng kê Mua vào nội bộ kế toán", default=None)
    parser.add_argument('--out-gdt', help="File hoặc Thư mục Excel HĐĐT Bán ra tải từ TCT", default=None)
    parser.add_argument('--out-nb', help="File hoặc Thư mục Excel Bảng kê Bán ra nội bộ kế toán", default=None)
    parser.add_argument('--xml', help="File hoặc Thư mục XML Tờ khai thuế 01/GTGT (1 kỳ, cả kỳ hoàn hoặc nhiều năm)", default=None)
    parser.add_argument('--out', help="Đường dẫn file Excel kết quả đối chiếu", default="KetQua_DoiChieu_ThueGTGT.xlsx")

    parser.add_argument('--mst-list', help="Danh sách Mã số thuế cần tra cứu (ngăn cách bởi dấu phẩy hoặc khoảng trắng)", default=None)
    parser.add_argument('--mst-file', help="File hoặc Thư mục Excel/CSV chứa cột Mã số thuế cần trích xuất và tra cứu", default=None)
    parser.add_argument('--mst-out', help="Đường dẫn file Excel xuất Danh bạ Mã số thuế", default="DanhBa_MaSoThue_DoanhNghiep.xlsx")
    
    args = parser.parse_args()
    
    print("="*70)
    print("PHẦN MỀM ĐỐI CHIẾU THUẾ GTGT & TRA CỨU DANH BẠ MÃ SỐ THUẾ (v1.2)")
    print("="*70)

    # Nếu chạy chế độ tra cứu Danh bạ Mã số thuế
    if args.mst_list or args.mst_file:
        msts = []
        seen = set()
        if args.mst_list:
            for m in MstLookupEngine.extract_msts_from_text(args.mst_list):
                if m not in seen:
                    seen.add(m)
                    msts.append(m)
        if args.mst_file:
            for m in MstLookupEngine.extract_msts_from_files(args.mst_file):
                if m not in seen:
                    seen.add(m)
                    msts.append(m)

        if not msts:
            print("[!] Không tìm thấy Mã số thuế hợp lệ nào để tra cứu.")
            return

        print(f"[*] Bắt đầu tra cứu {len(msts):,} Mã số thuế từ dữ liệu đồng bộ Tổng cục Thuế...")
        results = MstLookupEngine.lookup_batch(
            msts,
            max_workers=6,
            progress_fn=lambda d, t, r: print(f"    [{d}/{t}] {r['mst']} -> {r['ten_nnt'] or 'Không tìm thấy'} | {r['tinh_trang']}")
        )
        saved_mst = MstLookupEngine.export_directory_to_excel(results, args.mst_out)
        print(f"\n>>> ĐÃ XUẤT FILE EXCEL DANH BẠ MÃ SỐ THUẾ TẠI: {os.path.abspath(saved_mst)}")
        if not (args.in_gdt or args.in_nb or args.out_gdt or args.out_nb or args.xml):
            print("="*70)
            return
    
    in_gdt_list = []
    in_nb_list = []
    out_gdt_list = []
    out_nb_list = []
    
    if args.in_gdt:
        print(f"[*] Đang chuẩn hóa HĐĐT Mua vào (TCT): {args.in_gdt}...")
        in_gdt_list = ExcelNormalizer.normalize_multiple_sources(args.in_gdt, deduplicate=True, log_fn=print)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(in_gdt_list):,} hóa đơn")
        
    if args.in_nb:
        print(f"[*] Đang chuẩn hóa Bảng kê Mua vào (Nội bộ): {args.in_nb}...")
        in_nb_list = ExcelNormalizer.normalize_multiple_sources(args.in_nb, deduplicate=False, log_fn=print)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(in_nb_list):,} hóa đơn")
        
    if args.out_gdt:
        print(f"[*] Đang chuẩn hóa HĐĐT Bán ra (TCT): {args.out_gdt}...")
        out_gdt_list = ExcelNormalizer.normalize_multiple_sources(args.out_gdt, deduplicate=True, log_fn=print)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(out_gdt_list):,} hóa đơn")
        
    if args.out_nb:
        print(f"[*] Đang chuẩn hóa Bảng kê Bán ra (Nội bộ): {args.out_nb}...")
        out_nb_list = ExcelNormalizer.normalize_multiple_sources(args.out_nb, deduplicate=False, log_fn=print)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(out_nb_list):,} hóa đơn")
        
    in_rec = None
    if in_gdt_list or in_nb_list:
        print("[*] Đang chạy thuật toán đối chiếu Đầu vào (TCT vs Nội bộ)...")
        in_rec = VatReconciler.match_invoices(in_gdt_list, in_nb_list, match_type="Đầu vào")
        s = in_rec['summary']
        print(f"    + Khớp 100%: {s['matched_count']:,} HĐ")
        print(f"    + Khớp HĐ nhưng lệch tiền: {s['diff_amount_count']:,} HĐ")
        print(f"    + Chỉ có trên TCT (bỏ sót): {s['only_in_gdt_count']:,} HĐ")
        print(f"    + Chỉ có ở Nội bộ (nghi vấn): {s['only_in_internal_count']:,} HĐ")
        
    out_rec = None
    if out_gdt_list or out_nb_list:
        print("[*] Đang chạy thuật toán đối chiếu Đầu ra (TCT vs Nội bộ)...")
        out_rec = VatReconciler.match_invoices(out_gdt_list, out_nb_list, match_type="Đầu ra")
        s = out_rec['summary']
        print(f"    + Khớp 100%: {s['matched_count']:,} HĐ")
        print(f"    + Khớp HĐ nhưng lệch tiền: {s['diff_amount_count']:,} HĐ")
        print(f"    + Chỉ có trên TCT: {s['only_in_gdt_count']:,} HĐ")
        print(f"    + Chỉ có ở Nội bộ: {s['only_in_internal_count']:,} HĐ")
        
    ret_rec = None
    if args.xml:
        print(f"[*] Đang phân tích hồ sơ XML tờ khai 01/GTGT: {args.xml}...")
        vat_ret = TaxXmlParser.parse_multiple_vat_returns(args.xml, log_fn=print)
        if vat_ret:
            print(f"    -> NNT: {vat_ret['ten_nnt']} | Kỳ/Giai đoạn: {vat_ret['ky_tinh_thue']} | Lần nộp: {vat_ret['lan_ke_khai']}")
            
            final_in = in_nb_list if in_nb_list else in_gdt_list
            final_out = out_nb_list if out_nb_list else out_gdt_list
            
            print("[*] Đang đối chiếu tổng hợp với các chỉ tiêu tờ khai 01/GTGT...")
            ret_rec = VatReconciler.reconcile_with_vat_return(final_in, final_out, vat_ret)
            for comp in ret_rec['comparisons']:
                diff_str = f"{comp['chenh_lech']:+,.0f}" if comp['chenh_lech'] != 0 else "0"
                print(f"    - Chỉ tiêu {comp['chi_tieu_ma']:5s} | HĐ: {comp['so_hoa_don']:>14,.0f} | Tờ khai: {comp['so_to_khai']:>14,.0f} | Lệch: {diff_str:>14s} | {comp['danh_gia']}")
            
    print(f"[*] Đang xuất báo cáo Excel đối chiếu: {args.out}...")
    saved_file = ReconciliationReportExporter.export_report(args.out, in_rec, out_rec, ret_rec)
    print(f"\n>>> THÀNH CÔNG! Báo cáo đã lưu tại: {os.path.abspath(saved_file)}")
    print("="*70)

if __name__ == '__main__':
    main()
