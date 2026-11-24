# -*- coding: utf-8 -*-
"""
vat_reconciler_gui.py
---------------------
Giao diện đồ họa máy tính để bàn (Desktop GUI - Tkinter) - Phiên bản v1.1:
Dành riêng cho Kế toán thuế, Kiểm toán viên & Đoàn Kiểm tra / Hoàn thuế:
- 5 ô chọn dữ liệu đều hỗ trợ:
  + [Chọn nhiều file...] (chọn 1 hoặc hàng chục file Excel/XML cùng lúc)
  + [Chọn thư mục...] (chọn cả thư mục chứa hóa đơn nhiều tháng/quý/năm hoặc cả thư mục tờ khai XML GTGT của cả kỳ hoàn thuế / nhiều năm)
- Tự động gộp bảng kê nhiều kỳ/nhiều sheet, tự động lọc tờ khai 01/GTGT bản bổ sung cuối cùng của từng kỳ và cộng gộp toàn giai đoạn.
- Khung Log hiển thị tiến trình chuẩn hóa và đối chiếu thời gian thực.
- Bảng Dashboard tóm tắt nhanh kết quả đối chiếu chỉ tiêu ngay trên màn hình.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from excel_normalizer import ExcelNormalizer, resolve_input_paths, SUPPORTED_TABULAR_EXTS
from xml_parser import TaxXmlParser
from reconciler import VatReconciler
from excel_exporter import ReconciliationReportExporter


class VatReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm Đối chiếu Thuế GTGT Đa chiều (Hỗ trợ Nhiều Kỳ Hoàn / Nhiều Năm) - VAT Reconciliation Pro v1.1")
        self.root.geometry("1080x780")
        self.root.minsize(920, 680)
        
        # Biến lưu đường dẫn file / danh sách file / thư mục (ngăn cách bởi ;)
        self.var_in_gdt = tk.StringVar()
        self.var_in_nb = tk.StringVar()
        self.var_out_gdt = tk.StringVar()
        self.var_out_nb = tk.StringVar()
        self.var_xml_return = tk.StringVar()
        self.var_out_file = tk.StringVar(value=os.path.abspath("BaoCao_DoiChieu_ThueGTGT.xlsx"))
        
        self.last_output_path = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # 1. Header Frame
        header_frame = tk.Frame(self.root, bg="#1F4E79", height=64)
        header_frame.pack(fill=tk.X)
        
        lbl_title = tk.Label(
            header_frame, 
            text="HỆ THỐNG ĐỐI CHIẾU HÓA ĐƠN ĐIỆN TỬ, BẢNG KÊ & TỜ KHAI THUẾ GTGT (ĐA KỲ / CẢ KỲ HOÀN / NHIỀU NĂM)", 
            font=("Segoe UI", 12, "bold"), 
            fg="white", 
            bg="#1F4E79"
        )
        lbl_title.pack(pady=(10, 2))
        
        lbl_sub = tk.Label(
            header_frame, 
            text="Hỗ trợ chọn 1 file, nhiều file cùng lúc hoặc cả thư mục — Tự động gộp bảng kê nhiều kỳ & chốt tờ khai 01/GTGT bản bổ sung cuối cùng mỗi kỳ", 
            font=("Segoe UI", 9, "italic"), 
            fg="#D9E1F2", 
            bg="#1F4E79"
        )
        lbl_sub.pack(pady=(0, 10))
        
        # 2. Main Content Frame
        main_frame = tk.Frame(self.root, padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Group 1: Chọn dữ liệu đầu vào ---
        grp_input = tk.LabelFrame(
            main_frame,
            text="  1. Chọn hồ sơ & dữ liệu (Chọn 1 file, Nhiều file hoặc Cả thư mục GTGT theo kỳ hoàn / nhiều năm)  ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79",
            padx=10,
            pady=8
        )
        grp_input.pack(fill=tk.X, pady=(0, 10))
        
        def add_multi_source_row(parent, row, label_text, var_holder, file_types):
            lbl = tk.Label(parent, text=label_text, width=28, anchor="w", font=("Segoe UI", 9))
            lbl.grid(row=row, column=0, sticky="w", pady=3)
            
            entry = tk.Entry(parent, textvariable=var_holder, font=("Segoe UI", 9))
            entry.grid(row=row, column=1, sticky="ew", padx=5, pady=3)
            
            def browse_files():
                fns = filedialog.askopenfilenames(
                    title=f"Chọn 1 hoặc nhiều file cho: {label_text}",
                    filetypes=file_types
                )
                if fns:
                    norm_paths = [os.path.normpath(f) for f in fns]
                    cur = var_holder.get().strip()
                    if cur:
                        combined = cur + "; " + "; ".join(norm_paths)
                    else:
                        combined = "; ".join(norm_paths)
                    var_holder.set(combined)

            def browse_folder():
                folder = filedialog.askdirectory(
                    title=f"Chọn thư mục chứa dữ liệu cho: {label_text}"
                )
                if folder:
                    norm_folder = os.path.normpath(folder)
                    cur = var_holder.get().strip()
                    if cur:
                        var_holder.set(cur + "; " + norm_folder)
                    else:
                        var_holder.set(norm_folder)

            def clear_val():
                var_holder.set("")

            btn_files = tk.Button(
                parent, text="Chọn (nhiều) file...", command=browse_files,
                width=15, bg="#E1EBF5", font=("Segoe UI", 8)
            )
            btn_files.grid(row=row, column=2, padx=2, pady=3)

            btn_dir = tk.Button(
                parent, text="Chọn cả thư mục...", command=browse_folder,
                width=15, bg="#D9E1F2", font=("Segoe UI", 8, "bold"), fg="#1F4E79"
            )
            btn_dir.grid(row=row, column=3, padx=2, pady=3)

            btn_clr = tk.Button(
                parent, text="Xóa", command=clear_val,
                width=5, bg="#F2F2F2", font=("Segoe UI", 8)
            )
            btn_clr.grid(row=row, column=4, padx=(2, 0), pady=3)
            
        grp_input.columnconfigure(1, weight=1)
        excel_types = [("Excel / CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
        xml_types = [("XML Tờ khai 01/GTGT", "*.xml"), ("All Files", "*.*")]
        
        add_multi_source_row(grp_input, 0, "HĐĐT Mua vào (Tải từ TCT):", self.var_in_gdt, excel_types)
        add_multi_source_row(grp_input, 1, "Bảng kê Mua vào (Nội bộ):", self.var_in_nb, excel_types)
        add_multi_source_row(grp_input, 2, "HĐĐT Bán ra (Tải từ TCT):", self.var_out_gdt, excel_types)
        add_multi_source_row(grp_input, 3, "Bảng kê Bán ra (Nội bộ):", self.var_out_nb, excel_types)
        add_multi_source_row(grp_input, 4, "Tờ khai thuế 01/GTGT (XML):", self.var_xml_return, xml_types)
        
        # --- Output file row ---
        lbl_out = tk.Label(grp_input, text="Nơi lưu file Excel kết quả:", width=28, anchor="w", font=("Segoe UI", 9, "bold"))
        lbl_out.grid(row=5, column=0, sticky="w", pady=(8, 3))
        entry_out = tk.Entry(grp_input, textvariable=self.var_out_file, font=("Segoe UI", 9))
        entry_out.grid(row=5, column=1, sticky="ew", padx=5, pady=(8, 3))
        def browse_out():
            fn = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if fn:
                self.var_out_file.set(os.path.normpath(fn))
        btn_out = tk.Button(grp_input, text="Đổi vị trí lưu...", command=browse_out, width=15, bg="#E1EBF5", font=("Segoe UI", 8))
        btn_out.grid(row=5, column=2, padx=2, pady=(8, 3))
        
        # --- Action Buttons ---
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_run = tk.Button(
            btn_frame, 
            text="▶ THỰC HIỆN ĐỐI CHIẾU SỐ LIỆU", 
            font=("Segoe UI", 11, "bold"), 
            bg="#2E75B6", 
            fg="white", 
            padx=15, 
            pady=6,
            command=self.start_reconciliation
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_open_excel = tk.Button(
            btn_frame, 
            text="📊 Mở File Excel Báo Cáo", 
            font=("Segoe UI", 10, "bold"), 
            bg="#385623", 
            fg="white", 
            padx=15, 
            pady=6,
            state=tk.DISABLED,
            command=self.open_result_file
        )
        self.btn_open_excel.pack(side=tk.LEFT)
        
        # --- Group 2: Dashboard Tóm Tắt & Log ---
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
        # Frame Treeview tóm tắt tờ khai
        tree_frame = tk.LabelFrame(
            paned,
            text="  2. Bảng đối chiếu nhanh các chỉ tiêu Tờ khai GTGT (01/GTGT - 1 kỳ hoặc Tổng hợp cả kỳ hoàn / nhiều năm)  ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79"
        )
        paned.add(tree_frame, weight=3)
        
        cols = ("ma_ct", "ten_ct", "so_hd", "so_tk", "chenh_lech", "danh_gia")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=6)
        self.tree.heading("ma_ct", text="Chỉ tiêu")
        self.tree.heading("ten_ct", text="Tên chỉ tiêu tờ khai 01/GTGT")
        self.tree.heading("so_hd", text="Số HĐ / Bảng kê")
        self.tree.heading("so_tk", text="Số trên Tờ khai (Tổng kỳ)")
        self.tree.heading("chenh_lech", text="Chênh lệch")
        self.tree.heading("danh_gia", text="Trạng thái khớp")
        
        self.tree.column("ma_ct", width=70, anchor="center")
        self.tree.column("ten_ct", width=300, anchor="w")
        self.tree.column("so_hd", width=135, anchor="e")
        self.tree.column("so_tk", width=135, anchor="e")
        self.tree.column("chenh_lech", width=135, anchor="e")
        self.tree.column("danh_gia", width=150, anchor="center")
        
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Frame Log
        log_frame = tk.LabelFrame(paned, text="  3. Nhật ký tiến trình chuẩn hóa & kiểm tra  ", font=("Segoe UI", 10, "bold"), fg="#1F4E79")
        paned.add(log_frame, weight=2)
        
        self.txt_log = tk.Text(log_frame, font=("Consolas", 9), bg="#F8F9FA", wrap=tk.WORD)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
    def log(self, msg):
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
        
    def start_reconciliation(self):
        in_gdt = self.var_in_gdt.get().strip()
        in_nb = self.var_in_nb.get().strip()
        out_gdt = self.var_out_gdt.get().strip()
        out_nb = self.var_out_nb.get().strip()
        xml_ret = self.var_xml_return.get().strip()
        
        if not (in_gdt or in_nb or out_gdt or out_nb or xml_ret):
            messagebox.showwarning("Chưa chọn dữ liệu", "Vui lòng chọn ít nhất 1 file/thư mục hóa đơn, bảng kê hoặc tờ khai XML để đối chiếu!")
            return
            
        self.btn_run.config(state=tk.DISABLED)
        self.btn_open_excel.config(state=tk.DISABLED)
        self.txt_log.delete("1.0", tk.END)
        for row in self.tree.get_children():
            self.tree.delete(row)
            
        threading.Thread(target=self._run_process, daemon=True).start()
        
    def _run_process(self):
        try:
            self.log("[*] Bắt đầu quy trình kiểm tra đối chiếu số liệu...")
            in_gdt_list, in_nb_list = [], []
            out_gdt_list, out_nb_list = [], []
            
            in_gdt = self.var_in_gdt.get().strip()
            if in_gdt:
                files = resolve_input_paths(in_gdt, SUPPORTED_TABULAR_EXTS)
                self.log(f"-> Đang chuẩn hóa HĐĐT Mua vào TCT ({len(files)} file)...")
                in_gdt_list = ExcelNormalizer.normalize_multiple_sources(in_gdt, deduplicate=True, log_fn=self.log)
                self.log(f"   => Tổng cộng HĐĐT Mua vào TCT sau chuẩn hóa: {len(in_gdt_list):,} dòng hóa đơn.")
                
            in_nb = self.var_in_nb.get().strip()
            if in_nb:
                files = resolve_input_paths(in_nb, SUPPORTED_TABULAR_EXTS)
                self.log(f"-> Đang chuẩn hóa Bảng kê Mua vào Nội bộ ({len(files)} file)...")
                # Bảng kê nội bộ giữ nguyên toàn bộ các dòng kế toán đã kê khai (không tự ý xóa trùng)
                in_nb_list = ExcelNormalizer.normalize_multiple_sources(in_nb, deduplicate=False, log_fn=self.log)
                self.log(f"   => Tổng cộng Bảng kê Mua vào Nội bộ: {len(in_nb_list):,} dòng hóa đơn.")
                
            out_gdt = self.var_out_gdt.get().strip()
            if out_gdt:
                files = resolve_input_paths(out_gdt, SUPPORTED_TABULAR_EXTS)
                self.log(f"-> Đang chuẩn hóa HĐĐT Bán ra TCT ({len(files)} file)...")
                out_gdt_list = ExcelNormalizer.normalize_multiple_sources(out_gdt, deduplicate=True, log_fn=self.log)
                self.log(f"   => Tổng cộng HĐĐT Bán ra TCT sau chuẩn hóa: {len(out_gdt_list):,} dòng hóa đơn.")
                
            out_nb = self.var_out_nb.get().strip()
            if out_nb:
                files = resolve_input_paths(out_nb, SUPPORTED_TABULAR_EXTS)
                self.log(f"-> Đang chuẩn hóa Bảng kê Bán ra Nội bộ ({len(files)} file)...")
                out_nb_list = ExcelNormalizer.normalize_multiple_sources(out_nb, deduplicate=False, log_fn=self.log)
                self.log(f"   => Tổng cộng Bảng kê Bán ra Nội bộ: {len(out_nb_list):,} dòng hóa đơn.")
                
            in_rec = None
            if in_gdt_list or in_nb_list:
                self.log("[*] Chạy đối chiếu Mua vào (TCT vs Bảng kê nội bộ)...")
                in_rec = VatReconciler.match_invoices(in_gdt_list, in_nb_list, match_type="Đầu vào")
                s = in_rec['summary']
                self.log(f"   + Khớp 100%: {s['matched_count']:,} HĐ")
                self.log(f"   + Khớp HĐ nhưng lệch tiền: {s['diff_amount_count']:,} HĐ")
                self.log(f"   + Chỉ có trên TCT (bỏ sót kê khai): {s['only_in_gdt_count']:,} HĐ")
                self.log(f"   + Chỉ có ở Nội bộ (nghi vấn khống/hủy): {s['only_in_internal_count']:,} HĐ")
                
            out_rec = None
            if out_gdt_list or out_nb_list:
                self.log("[*] Chạy đối chiếu Bán ra (TCT vs Bảng kê nội bộ)...")
                out_rec = VatReconciler.match_invoices(out_gdt_list, out_nb_list, match_type="Đầu ra")
                s = out_rec['summary']
                self.log(f"   + Khớp 100%: {s['matched_count']:,} HĐ")
                self.log(f"   + Khớp HĐ nhưng lệch tiền: {s['diff_amount_count']:,} HĐ")
                self.log(f"   + Chỉ có trên TCT: {s['only_in_gdt_count']:,} HĐ")
                self.log(f"   + Chỉ có ở Nội bộ: {s['only_in_internal_count']:,} HĐ")
                
            ret_rec = None
            xml_ret = self.var_xml_return.get().strip()
            if xml_ret:
                self.log(f"[*] Phân tích hồ sơ Tờ khai thuế 01/GTGT (XML)...")
                vat_ret = TaxXmlParser.parse_multiple_vat_returns(xml_ret, log_fn=self.log)
                if vat_ret:
                    self.log(f"   + NNT: {vat_ret['ten_nnt']} | MST: {vat_ret['ma_so_thue']} | Kỳ/Giai đoạn: {vat_ret['ky_tinh_thue']}")
                    
                    final_in = in_nb_list if in_nb_list else in_gdt_list
                    final_out = out_nb_list if out_nb_list else out_gdt_list
                    
                    ret_rec = VatReconciler.reconcile_with_vat_return(final_in, final_out, vat_ret)
                    
                    for comp in ret_rec['comparisons']:
                        diff_val = comp['chenh_lech']
                        diff_str = f"{diff_val:+,.0f}" if diff_val != 0 else "0"
                        self.tree.insert("", tk.END, values=(
                            comp['chi_tieu_ma'],
                            comp['ten_chi_tieu'],
                            f"{comp['so_hoa_don']:,.0f}",
                            f"{comp['so_to_khai']:,.0f}",
                            diff_str,
                            comp['danh_gia']
                        ))
                else:
                    self.log("   [!] Không tìm thấy tờ khai XML 01/GTGT hợp lệ nào trong đường dẫn đã chọn.")
                    
            out_path = self.var_out_file.get().strip()
            self.log(f"[*] Xuất báo cáo Excel chuyên sâu ra: {out_path}...")
            saved = ReconciliationReportExporter.export_report(out_path, in_rec, out_rec, ret_rec)
            self.last_output_path = os.path.abspath(saved)
            
            self.log(f"\n>>> ĐỐI CHIẾU HOÀN TẤT THÀNH CÔNG!")
            self.log(f"File kết quả: {self.last_output_path}")
            messagebox.showinfo("Thành công", f"Đã hoàn thành đối chiếu!\nBáo cáo đã lưu tại:\n{self.last_output_path}")
            self.btn_open_excel.config(state=tk.NORMAL)
            
        except Exception as e:
            self.log(f"\n[!] LỖI TRONG QUÁ TRÌNH XỬ LÝ: {str(e)}")
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")
        finally:
            self.btn_run.config(state=tk.NORMAL)
            
    def open_result_file(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(self.last_output_path)
        else:
            messagebox.showwarning("Chưa có file", "Chưa có file báo cáo kết quả để mở!")


def main():
    root = tk.Tk()
    app = VatReconcilerApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
