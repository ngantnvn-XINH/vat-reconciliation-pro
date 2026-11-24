# -*- coding: utf-8 -*-
"""
vat_reconciler_gui.py
---------------------
Giao diện đồ họa máy tính để bàn (Desktop GUI - Tkinter) - Phiên bản v1.2:
Dành riêng cho Kế toán thuế, Kiểm toán viên & Đoàn Kiểm tra / Hoàn thuế:
- Tab 1: Đối chiếu Hóa đơn điện tử TCT vs Bảng kê Nội bộ vs Tờ khai 01/GTGT (1 kỳ, cả kỳ hoàn hoặc nhiều năm).
- Tab 2: Tra cứu Danh bạ Mã số thuế & Tình trạng hoạt động Doanh nghiệp từ dữ liệu Tổng cục Thuế (nhập danh sách MST hoặc quét từ file/thư mục Excel bảng kê, xuất ra file Excel Danh bạ MST kèm cảnh báo rủi ro bỏ địa chỉ/ngừng hoạt động).
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
from mst_lookup import MstLookupEngine


class VatReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm Đối chiếu Thuế GTGT & Tra cứu Danh bạ Mã Số Thuế (GDT) - VAT Reconciliation Pro v1.2")
        self.root.geometry("1120x820")
        self.root.minsize(960, 700)
        
        # Biến Tab 1: Đối chiếu GTGT
        self.var_in_gdt = tk.StringVar()
        self.var_in_nb = tk.StringVar()
        self.var_out_gdt = tk.StringVar()
        self.var_out_nb = tk.StringVar()
        self.var_xml_return = tk.StringVar()
        self.var_out_file = tk.StringVar(value=os.path.abspath("BaoCao_DoiChieu_ThueGTGT.xlsx"))
        self.last_output_path = None

        # Biến Tab 2: Tra cứu Danh bạ Mã số thuế
        self.var_mst_files = tk.StringVar()
        self.var_mst_out_file = tk.StringVar(value=os.path.abspath("DanhBa_MaSoThue_DoanhNghiep.xlsx"))
        self.last_mst_output_path = None
        self.mst_results_cache = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # 1. Header Frame
        header_frame = tk.Frame(self.root, bg="#1F4E79", height=64)
        header_frame.pack(fill=tk.X)
        
        lbl_title = tk.Label(
            header_frame, 
            text="HỆ THỐNG ĐỐI CHIẾU THUẾ GTGT ĐA KỲ & TRA CỨU DANH BẠ MÃ SỐ THUẾ (CỤC THUẾ)", 
            font=("Segoe UI", 12, "bold"), 
            fg="white", 
            bg="#1F4E79"
        )
        lbl_title.pack(pady=(10, 2))
        
        lbl_sub = tk.Label(
            header_frame, 
            text="Đối chiếu HĐĐT TCT vs Bảng kê nhiều kỳ vs 01/GTGT  |  Tra cứu hàng loạt Tình trạng hoạt động & Địa chỉ Doanh nghiệp theo MST", 
            font=("Segoe UI", 9, "italic"), 
            fg="#D9E1F2", 
            bg="#1F4E79"
        )
        lbl_sub.pack(pady=(0, 10))
        
        # 2. Notebook 2 Tabs
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[14, 6])

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        self.tab_reconcile = tk.Frame(self.notebook, padx=10, pady=8)
        self.tab_mst = tk.Frame(self.notebook, padx=10, pady=8)

        self.notebook.add(self.tab_reconcile, text="  1. Đối chiếu Hóa đơn & Tờ khai GTGT (Đa kỳ / Nhiều năm)  ")
        self.notebook.add(self.tab_mst, text="  2. Tra cứu Danh bạ Mã số thuế & Tình trạng hoạt động (Cục Thuế)  ")

        self._build_reconcile_tab(self.tab_reconcile)
        self._build_mst_tab(self.tab_mst)

    def _build_reconcile_tab(self, main_frame):
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
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_send_mst = tk.Button(
            btn_frame,
            text="🔍 Chuyển toàn bộ MST hóa đơn sang Tab Tra cứu Tình trạng HĐ",
            font=("Segoe UI", 9, "bold"),
            bg="#D9E1F2",
            fg="#1F4E79",
            padx=12,
            pady=6,
            command=self.transfer_invoices_msts_to_lookup_tab
        )
        self.btn_send_mst.pack(side=tk.LEFT)
        
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
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
        
        log_frame = tk.LabelFrame(paned, text="  3. Nhật ký tiến trình chuẩn hóa & kiểm tra  ", font=("Segoe UI", 10, "bold"), fg="#1F4E79")
        paned.add(log_frame, weight=2)
        
        self.txt_log = tk.Text(log_frame, font=("Consolas", 9), bg="#F8F9FA", wrap=tk.WORD)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

    def _build_mst_tab(self, parent):
        # Khung nhập dữ liệu MST
        top_frame = tk.LabelFrame(
            parent,
            text="  1. Nhập danh sách Mã số thuế hoặc Chọn File / Thư mục Excel Bảng kê chứa MST  ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79",
            padx=10,
            pady=8
        )
        top_frame.pack(fill=tk.X, pady=(0, 8))

        # Dòng 1: Chọn file/thư mục Excel chứa cột MST
        file_row = tk.Frame(top_frame)
        file_row.pack(fill=tk.X, pady=(0, 6))

        tk.Label(file_row, text="Trích xuất MST từ File/Thư mục Excel:", width=30, anchor="w", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        entry_mst_files = tk.Entry(file_row, textvariable=self.var_mst_files, font=("Segoe UI", 9))
        entry_mst_files.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        def browse_mst_files():
            fns = filedialog.askopenfilenames(
                title="Chọn 1 hoặc nhiều file Excel/CSV chứa Mã số thuế",
                filetypes=[("Excel / CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
            )
            if fns:
                norm_paths = [os.path.normpath(f) for f in fns]
                cur = self.var_mst_files.get().strip()
                self.var_mst_files.set((cur + "; " if cur else "") + "; ".join(norm_paths))
                self._extract_msts_from_selected_paths()

        def browse_mst_folder():
            folder = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel/CSV có Mã số thuế")
            if folder:
                norm_folder = os.path.normpath(folder)
                cur = self.var_mst_files.get().strip()
                self.var_mst_files.set((cur + "; " if cur else "") + norm_folder)
                self._extract_msts_from_selected_paths()

        tk.Button(file_row, text="Chọn (nhiều) file...", command=browse_mst_files, width=15, bg="#E1EBF5", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(file_row, text="Chọn cả thư mục...", command=browse_mst_folder, width=15, bg="#D9E1F2", font=("Segoe UI", 8, "bold"), fg="#1F4E79").pack(side=tk.LEFT, padx=2)
        tk.Button(file_row, text="Nạp MST vào ô dưới", command=self._extract_msts_from_selected_paths, width=16, bg="#E2EFDA", font=("Segoe UI", 8, "bold"), fg="#385623").pack(side=tk.LEFT, padx=2)

        # Dòng 2: Ô dán trực tiếp danh sách MST
        paste_frame = tk.Frame(top_frame)
        paste_frame.pack(fill=tk.X, pady=(2, 6))

        lbl_paste = tk.Label(
            paste_frame,
            text="Danh sách MST cần tra cứu:\n(Mỗi dòng 1 MST, hoặc cách\nnhau bởi dấu phẩy / khoảng trắng)\nHỗ trợ MST 10 số & 13 số",
            width=30,
            anchor="nw",
            justify=tk.LEFT,
            font=("Segoe UI", 9)
        )
        lbl_paste.pack(side=tk.LEFT, anchor="n")

        txt_container = tk.Frame(paste_frame)
        txt_container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.txt_mst_input = tk.Text(txt_container, height=5, font=("Consolas", 10))
        mst_in_scroll = ttk.Scrollbar(txt_container, orient=tk.VERTICAL, command=self.txt_mst_input.yview)
        self.txt_mst_input.configure(yscrollcommand=mst_in_scroll.set)
        self.txt_mst_input.pack(side=tk.LEFT, fill=tk.X, expand=True)
        mst_in_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        side_btns = tk.Frame(paste_frame)
        side_btns.pack(side=tk.LEFT, anchor="n", padx=2)
        tk.Button(
            side_btns,
            text="Xóa danh sách",
            command=lambda: (self.txt_mst_input.delete("1.0", tk.END), self.var_mst_files.set("")),
            width=15,
            bg="#F2F2F2",
            font=("Segoe UI", 8)
        ).pack(pady=(0, 4))

        # Dòng 3: Nơi lưu file Excel Danh bạ MST
        out_row = tk.Frame(top_frame)
        out_row.pack(fill=tk.X, pady=(4, 0))
        tk.Label(out_row, text="Nơi lưu file Excel Danh bạ MST:", width=30, anchor="w", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        tk.Entry(out_row, textvariable=self.var_mst_out_file, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        def browse_mst_out():
            fn = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if fn:
                self.var_mst_out_file.set(os.path.normpath(fn))

        tk.Button(out_row, text="Đổi vị trí lưu...", command=browse_mst_out, width=15, bg="#E1EBF5", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)

        # Cụm nút hành động Tab 2
        action_frame = tk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.btn_mst_run = tk.Button(
            action_frame,
            text="🔍 TRA CỨU TÌNH TRẠNG HOẠT ĐỘNG & XUẤT EXCEL DANH BẠ MST",
            font=("Segoe UI", 11, "bold"),
            bg="#2E75B6",
            fg="white",
            padx=15,
            pady=6,
            command=self.start_mst_lookup
        )
        self.btn_mst_run.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_mst_open_excel = tk.Button(
            action_frame,
            text="📊 Mở File Excel Danh Bạ MST",
            font=("Segoe UI", 10, "bold"),
            bg="#385623",
            fg="white",
            padx=15,
            pady=6,
            state=tk.DISABLED,
            command=self.open_mst_result_file
        )
        self.btn_mst_open_excel.pack(side=tk.LEFT, padx=(0, 15))

        self.lbl_mst_status = tk.Label(
            action_frame,
            text="Sẵn sàng tra cứu dữ liệu đồng bộ từ Tổng cục Thuế (gdt.gov.vn)",
            font=("Segoe UI", 9, "italic"),
            fg="#1F4E79"
        )
        self.lbl_mst_status.pack(side=tk.LEFT)

        # Bảng kết quả trực tiếp
        res_frame = tk.LabelFrame(
            parent,
            text="  2. Kết quả tra cứu Danh bạ Người nộp thuế (Tình trạng hoạt động, Phân loại rủi ro, Địa chỉ, Người đại diện)  ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79"
        )
        res_frame.pack(fill=tk.BOTH, expand=True)

        mst_cols = ("stt", "mst", "ten_nnt", "tinh_trang", "rui_ro", "dia_chi", "dai_dien")
        self.tree_mst = ttk.Treeview(res_frame, columns=mst_cols, show="headings", height=14)
        self.tree_mst.heading("stt", text="STT")
        self.tree_mst.heading("mst", text="Mã số thuế")
        self.tree_mst.heading("ten_nnt", text="Tên doanh nghiệp / Người nộp thuế")
        self.tree_mst.heading("tinh_trang", text="Tình trạng hoạt động (Cục Thuế)")
        self.tree_mst.heading("rui_ro", text="Đánh giá rủi ro")
        self.tree_mst.heading("dia_chi", text="Địa chỉ trụ sở đăng ký thuế")
        self.tree_mst.heading("dai_dien", text="Người đại diện")

        self.tree_mst.column("stt", width=45, anchor="center")
        self.tree_mst.column("mst", width=110, anchor="center")
        self.tree_mst.column("ten_nnt", width=250, anchor="w")
        self.tree_mst.column("tinh_trang", width=210, anchor="w")
        self.tree_mst.column("rui_ro", width=165, anchor="center")
        self.tree_mst.column("dia_chi", width=260, anchor="w")
        self.tree_mst.column("dai_dien", width=120, anchor="w")

        self.tree_mst.tag_configure("risk_high", background="#FFC7CE", foreground="#9C0006")
        self.tree_mst.tag_configure("risk_warn", background="#FFEB9C", foreground="#9C6500")
        self.tree_mst.tag_configure("risk_ok", background="#C6EFCE", foreground="#006100")

        mst_yscroll = ttk.Scrollbar(res_frame, orient=tk.VERTICAL, command=self.tree_mst.yview)
        mst_xscroll = ttk.Scrollbar(res_frame, orient=tk.HORIZONTAL, command=self.tree_mst.xview)
        self.tree_mst.configure(yscrollcommand=mst_yscroll.set, xscrollcommand=mst_xscroll.set)
        self.tree_mst.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(5, 0), pady=(5, 0))
        mst_xscroll.pack(side=tk.BOTTOM, fill=tk.X, padx=(5, 0))
        mst_yscroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

    def _extract_msts_from_selected_paths(self):
        paths = self.var_mst_files.get().strip()
        if not paths:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn file hoặc thư mục Excel/CSV trước!")
            return
        msts = MstLookupEngine.extract_msts_from_files(paths)
        if not msts:
            messagebox.showinfo("Không tìm thấy MST", "Không tìm thấy Mã số thuế hợp lệ (10 số hoặc 13 số) trong các file đã chọn.")
            return
        cur_txt = self.txt_mst_input.get("1.0", tk.END).strip()
        existing = MstLookupEngine.extract_msts_from_text(cur_txt)
        merged = list(existing)
        seen = set(existing)
        for m in msts:
            if m not in seen:
                seen.add(m)
                merged.append(m)
        self.txt_mst_input.delete("1.0", tk.END)
        self.txt_mst_input.insert(tk.END, "\n".join(merged))
        self.lbl_mst_status.config(text=f"Đã trích xuất và nạp {len(merged):,} Mã số thuế duy nhất vào danh sách.")

    def transfer_invoices_msts_to_lookup_tab(self):
        sources = [
            self.var_in_gdt.get().strip(),
            self.var_in_nb.get().strip(),
            self.var_out_gdt.get().strip(),
            self.var_out_nb.get().strip(),
        ]
        valid_sources = [s for s in sources if s]
        if not valid_sources:
            messagebox.showwarning("Chưa có file ở Tab 1", "Vui lòng chọn file/thư mục hóa đơn hoặc bảng kê ở Tab 1 trước!")
            return
        combined_paths = "; ".join(valid_sources)
        self.var_mst_files.set(combined_paths)
        self.notebook.select(self.tab_mst)
        self._extract_msts_from_selected_paths()

    def start_mst_lookup(self):
        raw_text = self.txt_mst_input.get("1.0", tk.END).strip()
        mst_list = MstLookupEngine.extract_msts_from_text(raw_text)

        if not mst_list and self.var_mst_files.get().strip():
            mst_list = MstLookupEngine.extract_msts_from_files(self.var_mst_files.get().strip())
            if mst_list:
                self.txt_mst_input.delete("1.0", tk.END)
                self.txt_mst_input.insert(tk.END, "\n".join(mst_list))

        if not mst_list:
            messagebox.showwarning(
                "Chưa có Mã số thuế",
                "Vui lòng dán danh sách Mã số thuế vào ô nhập hoặc chọn file/thư mục Excel chứa MST!"
            )
            return

        self.btn_mst_run.config(state=tk.DISABLED)
        self.btn_mst_open_excel.config(state=tk.DISABLED)
        for row in self.tree_mst.get_children():
            self.tree_mst.delete(row)

        self.lbl_mst_status.config(text=f"Đang tra cứu 0/{len(mst_list):,} Mã số thuế từ dữ liệu Cục Thuế...")
        threading.Thread(target=self._run_mst_lookup_thread, args=(mst_list,), daemon=True).start()

    def _run_mst_lookup_thread(self, mst_list):
        try:
            def on_progress(done_cnt, total_cnt, last_res):
                self.root.after(
                    0,
                    lambda d=done_cnt, t=total_cnt, r=last_res: self.lbl_mst_status.config(
                        text=f"Đang tra cứu {d:,}/{t:,} MST... ({r.get('mst')}: {r.get('ten_nnt') or r.get('tinh_trang')})"
                    )
                )

            results = MstLookupEngine.lookup_batch(mst_list, max_workers=6, progress_fn=on_progress)
            self.mst_results_cache = results

            def populate_table():
                for idx, item in enumerate(results, 1):
                    risk = item.get("phan_loai_rui_ro", "")
                    if "RỦI RO CAO" in risk:
                        tag = "risk_high"
                    elif "CẢNH BÁO" in risk or "Không tìm thấy" in risk:
                        tag = "risk_warn"
                    elif "An toàn" in risk:
                        tag = "risk_ok"
                    else:
                        tag = ""
                    self.tree_mst.insert(
                        "",
                        tk.END,
                        values=(
                            idx,
                            item.get("mst", ""),
                            item.get("ten_nnt", ""),
                            item.get("tinh_trang", ""),
                            risk,
                            item.get("dia_chi", ""),
                            item.get("nguoi_dai_dien", ""),
                        ),
                        tags=(tag,) if tag else ()
                    )

            self.root.after(0, populate_table)

            out_path = self.var_mst_out_file.get().strip() or os.path.abspath("DanhBa_MaSoThue_DoanhNghiep.xlsx")
            saved = MstLookupEngine.export_directory_to_excel(results, out_path)
            self.last_mst_output_path = os.path.abspath(saved)

            active_cnt = sum(1 for r in results if "An toàn" in r.get("phan_loai_rui_ro", ""))
            risk_cnt = sum(1 for r in results if "RỦI RO CAO" in r.get("phan_loai_rui_ro", ""))

            def on_done():
                self.lbl_mst_status.config(
                    text=f"Hoàn tất {len(results):,} MST (Đang hoạt động: {active_cnt:,} | Rủi ro cao: {risk_cnt:,}). Đã xuất Excel: {self.last_mst_output_path}"
                )
                self.btn_mst_open_excel.config(state=tk.NORMAL)
                messagebox.showinfo(
                    "Tra cứu & Xuất Danh bạ MST thành công",
                    f"Đã tra cứu hoàn tất {len(results):,} Mã số thuế!\n"
                    f"- Đang hoạt động (An toàn): {active_cnt:,}\n"
                    f"- Rủi ro cao (Bỏ địa chỉ / Ngừng HĐ): {risk_cnt:,}\n\n"
                    f"File Excel Danh bạ đã lưu tại:\n{self.last_mst_output_path}"
                )

            self.root.after(0, on_done)
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda m=err_msg: messagebox.showerror("Lỗi tra cứu MST", f"Đã xảy ra lỗi:\n{m}"))
        finally:
            self.root.after(0, lambda: self.btn_mst_run.config(state=tk.NORMAL))

    def open_mst_result_file(self):
        if self.last_mst_output_path and os.path.exists(self.last_mst_output_path):
            os.startfile(self.last_mst_output_path)
        else:
            messagebox.showwarning("Chưa có file", "Chưa có file Excel Danh bạ MST để mở!")

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
            err_msg = str(e)
            self.log(f"\n[!] LỖI TRONG QUÁ TRÌNH XỬ LÝ: {err_msg}")
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{err_msg}")
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
