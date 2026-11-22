# -*- coding: utf-8 -*-
"""
vat_reconciler_cli.py
---------------------
Giao diện dòng lệnh (CLI) tự động hóa:
Cho phép chạy đối chiếu tự động không cần bấm GUI (hỗ trợ tích hợp batch script, server hoặc pipeline xử lý hàng loạt).

Cách dùng:
    python vat_reconciler_cli.py --in-gdt <file_hd_gdt> --in-nb <file_bk_muavao> --out-gdt <file_hd_banra> --out-nb <file_bk_banra> --xml <to_khai.xml> --out <ket_qua.xlsx>
"""

import os
import sys
import argparse
from excel_normalizer import ExcelNormalizer
from xml_parser import TaxXmlParser
from reconciler import VatReconciler
from excel_exporter import ReconciliationReportExporter

def main():
    parser = argparse.ArgumentParser(description="Phần mềm Đối chiếu Thuế GTGT Đa chiều (HĐĐT vs Bảng kê vs Tờ khai 01/GTGT)")
    parser.add_argument('--in-gdt', help="File Excel HĐĐT Mua vào tải từ TCT", default=None)
    parser.add_argument('--in-nb', help="File Excel Bảng kê Mua vào nội bộ kế toán", default=None)
    parser.add_argument('--out-gdt', help="File Excel HĐĐT Bán ra tải từ TCT", default=None)
    parser.add_argument('--out-nb', help="File Excel Bảng kê Bán ra nội bộ kế toán", default=None)
    parser.add_argument('--xml', help="File XML Tờ khai thuế 01/GTGT (kết xuất từ HTKK/eTax)", default=None)
    parser.add_argument('--out', help="Đường dẫn file Excel kết quả đối chiếu", default="KetQua_DoiChieu_ThueGTGT.xlsx")
    
    args = parser.parse_args()
    
    print("="*70)
    print("PHẦN MỀM ĐỐI CHIẾU SỐ LIỆU THUẾ GTGT ĐA CHIỀU (VAT RECONCILIATION PRO)")
    print("="*70)
    
    in_gdt_list = []
    in_nb_list = []
    out_gdt_list = []
    out_nb_list = []
    
    if args.in_gdt:
        print(f"[*] Đang chuẩn hóa HĐĐT Mua vào (TCT): {args.in_gdt}...")
        in_gdt_list = ExcelNormalizer.normalize_invoice_sheet(args.in_gdt)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(in_gdt_list):,} hóa đơn")
        
    if args.in_nb:
        print(f"[*] Đang chuẩn hóa Bảng kê Mua vào (Nội bộ): {args.in_nb}...")
        in_nb_list = ExcelNormalizer.normalize_invoice_sheet(args.in_nb)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(in_nb_list):,} hóa đơn")
        
    if args.out_gdt:
        print(f"[*] Đang chuẩn hóa HĐĐT Bán ra (TCT): {args.out_gdt}...")
        out_gdt_list = ExcelNormalizer.normalize_invoice_sheet(args.out_gdt)
        print(f"    -> Đã nạp & chuẩn hóa thành công: {len(out_gdt_list):,} hóa đơn")
        
    if args.out_nb:
        print(f"[*] Đang chuẩn hóa Bảng kê Bán ra (Nội bộ): {args.out_nb}...")
        out_nb_list = ExcelNormalizer.normalize_invoice_sheet(args.out_nb)
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
        print(f"[*] Đang phân tích file XML tờ khai 01/GTGT: {args.xml}...")
        vat_ret = TaxXmlParser.parse_vat_return(args.xml)
        print(f"    -> NNT: {vat_ret['ten_nnt']} | Kỳ thuế: {vat_ret['ky_tinh_thue']} | Lần nộp: {vat_ret['lan_ke_khai']}")
        
        # Chọn nguồn dữ liệu để đối chiếu tờ khai: ưu tiên nội bộ (bảng kê), nếu rỗng lấy TCT
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
