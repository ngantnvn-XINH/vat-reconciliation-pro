# -*- coding: utf-8 -*-
"""
excel_exporter.py
-----------------
Module xuất báo cáo Excel đối chiếu thuế GTGT đa chiều:
1. Sheet 'Tổng hợp đối chiếu Tờ khai': So sánh chỉ tiêu [23], [24], [25], [26], [29], [30], [32], [34], [35] giữa Bảng kê và Tờ khai thuế.
2. Sheet 'Đối chiếu Đầu vào (GDT vs NB)': Danh sách khớp, lệch tiền, thừa bên GDT hoặc thừa bên Nội bộ.
3. Sheet 'Đối chiếu Đầu ra (GDT vs NB)': Danh sách chi tiết hóa đơn đầu ra 2 bên.
4. Sheet 'Hóa đơn rủi ro & Lệch': Tổng hợp các hóa đơn nghi vấn (chênh lệch tiền, chỉ có 1 bên, sai thuế suất).

Định dạng chuyên nghiệp theo quy chuẩn thanh tra thuế:
- Cột số định dạng #,##0 hoặc #,##0.00
- Highlight màu đỏ các ô chênh lệch, màu xanh các ô khớp
- Auto-filter, đóng băng tiêu đề (Freeze Panes), độ rộng cột tự căn chỉnh
"""

import os
from typing import Dict, Any, List, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class ReconciliationReportExporter:
    @staticmethod
    def export_report(
        output_path: str,
        in_reconciliation: Optional[Dict[str, Any]] = None,
        out_reconciliation: Optional[Dict[str, Any]] = None,
        return_reconciliation: Optional[Dict[str, Any]] = None
    ) -> str:
        wb = openpyxl.Workbook()
        # Xóa sheet mặc định
        default_sheet = wb.active
        
        # Style định dạng
        font_title = Font(name='Times New Roman', size=14, bold=True, color='1B365D')
        font_sub = Font(name='Times New Roman', size=11, italic=True)
        font_header = Font(name='Times New Roman', size=11, bold=True, color='FFFFFF')
        font_bold = Font(name='Times New Roman', size=11, bold=True)
        font_regular = Font(name='Times New Roman', size=11)
        font_alert = Font(name='Times New Roman', size=11, bold=True, color='9C0006')
        font_ok = Font(name='Times New Roman', size=11, bold=True, color='006100')
        
        fill_header = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        fill_section = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        fill_alert = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        fill_ok = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        fill_warn = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        
        border_thin = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        border_double_bottom = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='double', color='000000')
        )
        
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_left = Alignment(horizontal='left', vertical='center')
        align_right = Alignment(horizontal='right', vertical='center')
        
        # ==========================================
        # 1. SHEET TỔNG HỢP ĐỐI CHIẾU TỜ KHAI THUẾ
        # ==========================================
        if return_reconciliation:
            ws_summary = wb.create_sheet(title='Đối chiếu Tờ khai GTGT')
            ws_summary.views.sheetView[0].showGridLines = True
            
            # Header title
            ws_summary.merge_cells('A1:G1')
            ws_summary['A1'] = "BẢNG ĐỐI CHIẾU TỔNG HỢP SỐ LIỆU HÓA ĐƠN VỚI TỜ KHAI THUẾ GTGT (01/GTGT)"
            ws_summary['A1'].font = font_title
            ws_summary['A1'].alignment = align_center
            
            ret_info = return_reconciliation.get('vat_return_info', {})
            ws_summary.merge_cells('A2:G2')
            ws_summary['A2'] = f"Kỳ thuế: {ret_info.get('ky_tinh_thue', '')} - Lần nộp: {ret_info.get('lan_ke_khai', 0)} - NNT: {ret_info.get('ten_nnt', '')} (MST: {ret_info.get('ma_so_thue', '')})"
            ws_summary['A2'].font = font_sub
            ws_summary['A2'].alignment = align_center
            
            headers_ret = ['Mã chỉ tiêu', 'Tên chỉ tiêu tờ khai 01/GTGT', 'Phân loại', 'Số theo Hóa đơn / Bảng kê', 'Số kê khai trên Tờ khai', 'Chênh lệch (HĐ - Tờ khai)', 'Đánh giá trạng thái']
            for c_idx, h in enumerate(headers_ret, 1):
                cell = ws_summary.cell(4, c_idx, h)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = align_center
                cell.border = border_thin
            ws_summary.row_dimensions[4].height = 28
            
            curr_row = 5
            for comp in return_reconciliation.get('comparisons', []):
                ws_summary.cell(curr_row, 1, comp['chi_tieu_ma']).alignment = align_center
                ws_summary.cell(curr_row, 2, comp['ten_chi_tieu']).alignment = align_left
                ws_summary.cell(curr_row, 3, comp['loai']).alignment = align_center
                
                c_hd = ws_summary.cell(curr_row, 4, comp['so_hoa_don'])
                c_hd.number_format = '#,##0'
                c_hd.alignment = align_right
                
                c_tk = ws_summary.cell(curr_row, 5, comp['so_to_khai'])
                c_tk.number_format = '#,##0'
                c_tk.alignment = align_right
                
                c_diff = ws_summary.cell(curr_row, 6, comp['chenh_lech'])
                c_diff.number_format = '#,##0'
                c_diff.alignment = align_right
                
                c_dg = ws_summary.cell(curr_row, 7, comp['danh_gia'])
                c_dg.alignment = align_center
                
                # Highlight kết quả
                if comp['danh_gia'] == 'Khớp 100%':
                    c_dg.font = font_ok
                    c_dg.fill = fill_ok
                elif 'Lệch nhỏ' in comp['danh_gia']:
                    c_dg.font = font_bold
                    c_dg.fill = fill_warn
                else:
                    c_dg.font = font_alert
                    c_dg.fill = fill_alert
                    c_diff.font = font_alert
                    
                for c in range(1, 8):
                    ws_summary.cell(curr_row, c).border = border_thin
                curr_row += 1
                
            # Đóng băng tiêu đề
            ws_summary.freeze_panes = 'A5'
            
        # ==========================================
        # 2. SHEET ĐỐI CHIẾU ĐẦU VÀO (GDT vs NỘI BỘ)
        # ==========================================
        def create_match_sheet(sheet_title, rec_data):
            if not rec_data:
                return
            ws = wb.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True
            
            # Title
            ws.merge_cells('A1:N1')
            ws['A1'] = f"KẾT QUẢ ĐỐI CHIẾU HÓA ĐƠN {rec_data.get('match_type', '').upper()} GIỮA HỆ THỐNG TCT (HĐĐT) VÀ BẢNG KÊ NỘI BỘ"
            ws['A1'].font = font_title
            ws['A1'].alignment = align_center
            
            sum_info = rec_data.get('summary', {})
            ws.merge_cells('A2:N2')
            ws['A2'] = (f"Tổng HĐ TCT: {sum_info.get('total_gdt_count',0):,} (DS: {sum_info.get('total_gdt_doanh_so',0):,.0f} | Thuế: {sum_info.get('total_gdt_tien_thue',0):,.0f})  ---  "
                        f"Tổng HĐ Nội bộ: {sum_info.get('total_internal_count',0):,} (DS: {sum_info.get('total_internal_doanh_so',0):,.0f} | Thuế: {sum_info.get('total_internal_tien_thue',0):,.0f})")
            ws['A2'].font = font_sub
            ws['A2'].alignment = align_center
            
            headers = [
                'STT', 'Trạng thái đối chiếu', 'Số HĐ (TCT)', 'Ký hiệu (TCT)', 'Ngày HĐ (TCT)', 'MST',
                'Doanh số (TCT)', 'Tiền thuế (TCT)', 'Số HĐ (Nội bộ)', 'Ngày HĐ (Nội bộ)', 
                'Doanh số (Nội bộ)', 'Tiền thuế (Nội bộ)', 'Lệch Doanh số', 'Lệch Tiền thuế', 'Diễn giải / Ghi chú'
            ]
            for c_idx, h in enumerate(headers, 1):
                cell = ws.cell(4, c_idx, h)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = align_center
                cell.border = border_thin
            ws.row_dimensions[4].height = 26
            
            r_idx = 5
            stt = 1
            # 1. Khớp hoàn toàn
            for item in rec_data.get('matched_perfect', []):
                ws.cell(r_idx, 1, stt).alignment = align_center
                c_st = ws.cell(r_idx, 2, item['trang_thai'])
                c_st.alignment = align_center; c_st.font = font_ok; c_st.fill = fill_ok
                
                ws.cell(r_idx, 3, item['so_hd_gdt']).alignment = align_center
                ws.cell(r_idx, 4, item['ky_hieu_gdt']).alignment = align_center
                ws.cell(r_idx, 5, item['ngay_hd_gdt']).alignment = align_center
                ws.cell(r_idx, 6, item['mst_gdt']).alignment = align_center
                
                c_ds1 = ws.cell(r_idx, 7, item['doanh_so_gdt']); c_ds1.number_format = '#,##0'; c_ds1.alignment = align_right
                c_th1 = ws.cell(r_idx, 8, item['tien_thue_gdt']); c_th1.number_format = '#,##0'; c_th1.alignment = align_right
                
                ws.cell(r_idx, 9, item['so_hd_nb']).alignment = align_center
                ws.cell(r_idx, 10, item['ngay_hd_nb']).alignment = align_center
                c_ds2 = ws.cell(r_idx, 11, item['doanh_so_nb']); c_ds2.number_format = '#,##0'; c_ds2.alignment = align_right
                c_th2 = ws.cell(r_idx, 12, item['tien_thue_nb']); c_th2.number_format = '#,##0'; c_th2.alignment = align_right
                
                c_df1 = ws.cell(r_idx, 13, item['diff_doanh_so']); c_df1.number_format = '#,##0'; c_df1.alignment = align_right
                c_df2 = ws.cell(r_idx, 14, item['diff_tien_thue']); c_df2.number_format = '#,##0'; c_df2.alignment = align_right
                ws.cell(r_idx, 15, item['dien_giai']).alignment = align_left
                
                for c in range(1, 16): ws.cell(r_idx, c).border = border_thin
                r_idx += 1; stt += 1
                
            # 2. Khớp HĐ nhưng lệch tiền
            for item in rec_data.get('matched_diff_amount', []):
                ws.cell(r_idx, 1, stt).alignment = align_center
                c_st = ws.cell(r_idx, 2, item['trang_thai'])
                c_st.alignment = align_center; c_st.font = font_alert; c_st.fill = fill_alert
                
                ws.cell(r_idx, 3, item['so_hd_gdt']).alignment = align_center
                ws.cell(r_idx, 4, item['ky_hieu_gdt']).alignment = align_center
                ws.cell(r_idx, 5, item['ngay_hd_gdt']).alignment = align_center
                ws.cell(r_idx, 6, item['mst_gdt']).alignment = align_center
                
                c_ds1 = ws.cell(r_idx, 7, item['doanh_so_gdt']); c_ds1.number_format = '#,##0'; c_ds1.alignment = align_right
                c_th1 = ws.cell(r_idx, 8, item['tien_thue_gdt']); c_th1.number_format = '#,##0'; c_th1.alignment = align_right
                
                ws.cell(r_idx, 9, item['so_hd_nb']).alignment = align_center
                ws.cell(r_idx, 10, item['ngay_hd_nb']).alignment = align_center
                c_ds2 = ws.cell(r_idx, 11, item['doanh_so_nb']); c_ds2.number_format = '#,##0'; c_ds2.alignment = align_right
                c_th2 = ws.cell(r_idx, 12, item['tien_thue_nb']); c_th2.number_format = '#,##0'; c_th2.alignment = align_right
                
                c_df1 = ws.cell(r_idx, 13, item['diff_doanh_so']); c_df1.number_format = '#,##0'; c_df1.alignment = align_right; c_df1.font = font_alert
                c_df2 = ws.cell(r_idx, 14, item['diff_tien_thue']); c_df2.number_format = '#,##0'; c_df2.alignment = align_right; c_df2.font = font_alert
                ws.cell(r_idx, 15, item['dien_giai']).alignment = align_left
                
                for c in range(1, 16): ws.cell(r_idx, c).border = border_thin
                r_idx += 1; stt += 1
                
            # 3. Chỉ có ở GDT (Chưa kê khai)
            for item in rec_data.get('only_in_gdt', []):
                ws.cell(r_idx, 1, stt).alignment = align_center
                c_st = ws.cell(r_idx, 2, 'Chỉ có trên TCT (Bỏ sót)')
                c_st.alignment = align_center; c_st.font = font_alert; c_st.fill = fill_warn
                
                ws.cell(r_idx, 3, item.get('so_hd', '')).alignment = align_center
                ws.cell(r_idx, 4, item.get('ky_hieu', '')).alignment = align_center
                ws.cell(r_idx, 5, item.get('ngay_hd', '')).alignment = align_center
                ws.cell(r_idx, 6, item.get('mst', '')).alignment = align_center
                
                c_ds1 = ws.cell(r_idx, 7, item.get('doanh_so', 0)); c_ds1.number_format = '#,##0'; c_ds1.alignment = align_right
                c_th1 = ws.cell(r_idx, 8, item.get('tien_thue', 0)); c_th1.number_format = '#,##0'; c_th1.alignment = align_right
                
                ws.cell(r_idx, 9, '-').alignment = align_center
                ws.cell(r_idx, 10, '-').alignment = align_center
                ws.cell(r_idx, 11, 0).number_format = '#,##0'
                ws.cell(r_idx, 12, 0).number_format = '#,##0'
                
                c_df1 = ws.cell(r_idx, 13, item.get('doanh_so', 0)); c_df1.number_format = '#,##0'; c_df1.alignment = align_right; c_df1.font = font_alert
                c_df2 = ws.cell(r_idx, 14, item.get('tien_thue', 0)); c_df2.number_format = '#,##0'; c_df2.alignment = align_right; c_df2.font = font_alert
                ws.cell(r_idx, 15, item.get('dien_giai', '')).alignment = align_left
                
                for c in range(1, 16): ws.cell(r_idx, c).border = border_thin
                r_idx += 1; stt += 1
                
            # 4. Chỉ có ở Nội bộ (Nghi vấn khống/hủy)
            for item in rec_data.get('only_in_internal', []):
                ws.cell(r_idx, 1, stt).alignment = align_center
                c_st = ws.cell(r_idx, 2, 'Chỉ có ở Nội bộ (Không có trên TCT)')
                c_st.alignment = align_center; c_st.font = font_alert; c_st.fill = fill_alert
                
                ws.cell(r_idx, 3, '-').alignment = align_center
                ws.cell(r_idx, 4, '-').alignment = align_center
                ws.cell(r_idx, 5, '-').alignment = align_center
                ws.cell(r_idx, 6, item.get('mst', '')).alignment = align_center
                ws.cell(r_idx, 7, 0).number_format = '#,##0'
                ws.cell(r_idx, 8, 0).number_format = '#,##0'
                
                ws.cell(r_idx, 9, item.get('so_hd', '')).alignment = align_center
                ws.cell(r_idx, 10, item.get('ngay_hd', '')).alignment = align_center
                c_ds2 = ws.cell(r_idx, 11, item.get('doanh_so', 0)); c_ds2.number_format = '#,##0'; c_ds2.alignment = align_right
                c_th2 = ws.cell(r_idx, 12, item.get('tien_thue', 0)); c_th2.number_format = '#,##0'; c_th2.alignment = align_right
                
                c_df1 = ws.cell(r_idx, 13, -item.get('doanh_so', 0)); c_df1.number_format = '#,##0'; c_df1.alignment = align_right; c_df1.font = font_alert
                c_df2 = ws.cell(r_idx, 14, -item.get('tien_thue', 0)); c_df2.number_format = '#,##0'; c_df2.alignment = align_right; c_df2.font = font_alert
                ws.cell(r_idx, 15, item.get('dien_giai', '')).alignment = align_left
                
                for c in range(1, 16): ws.cell(r_idx, c).border = border_thin
                r_idx += 1; stt += 1
                
            ws.freeze_panes = 'A5'
            
        create_match_sheet('Đối chiếu Đầu vào', in_reconciliation)
        create_match_sheet('Đối chiếu Đầu ra', out_reconciliation)
        
        # Xóa default sheet nếu có sheet khác
        if len(wb.sheetnames) > 1 and default_sheet.title in wb.sheetnames:
            wb.remove(default_sheet)
            
        # Tự động canh chỉnh độ rộng cột cho tất cả các sheet
        for sheet in wb.worksheets:
            for col in sheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row in (1, 2): # Bỏ qua dòng tiêu đề merge
                        continue
                    v = str(cell.value or '')
                    if '\n' in v:
                        v = max(v.split('\n'), key=len)
                    max_len = max(max_len, len(v))
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        return output_path
