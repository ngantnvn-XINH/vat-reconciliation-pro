# VAT Reconciliation Pro — Phần Mềm Đối Chiếu Thuế GTGT Đa Chiều (v1.1)

Ứng dụng chuyên sâu độc lập hỗ trợ Kiểm toán, Thanh tra / Kiểm tra / Hoàn thuế và Kế toán doanh nghiệp tự động hóa 100% quy trình đối chiếu số liệu thuế GTGT **cho cả kỳ hoàn thuế hoặc nhiều năm liên tiếp**:
1. **Đối chiếu HĐĐT Bán ra (TCT) vs Bảng kê Bán ra (Nội bộ)** — hỗ trợ chọn 1 file, nhiều file cùng lúc hoặc chọn cả thư mục.
2. **Đối chiếu HĐĐT Mua vào (TCT) vs Bảng kê Mua vào (Nội bộ)** — tự động quét mọi sheet trong workbook (bảng kê chia theo sheet Tháng/Quý/Năm) và tự động khử trùng lặp khi tải chồng các file HĐĐT từ TCT.
3. **Đối chiếu Tổng hợp Hóa đơn vs Tờ khai thuế 01/GTGT (1 kỳ hoặc Cả thư mục GTGT nhiều kỳ/nhiều năm)** — tự động lọc tờ khai `01/GTGT`, chốt bản bổ sung cuối cùng của từng kỳ, cộng gộp chỉ tiêu `[23]`..`[36]` toàn giai đoạn và xuất thêm sheet **`Chi tiết Tờ khai theo Kỳ`**.

---

## 🌟 Tính Năng Nổi Bật

### 1. Bộ Chuẩn Hóa File Excel Thông Minh (`excel_normalizer.py`)
Giải quyết triệt để vấn đề các file Excel đầu vào **không đồng nhất về định dạng**:
- **Tự động dò tiêu đề (Header Auto-Detection)**: Bất kể bảng kê có bao nhiêu dòng trống phía trên, hay tiêu đề gộp 2-3 tầng (như *"Doanh số mua"* ở dòng 1 và *"chưa thuế"* ở dòng 2), thuật toán đều tự động nhận diện chính xác.
- **Nhận diện cột linh hoạt**: Hỗ trợ mọi biến thể tên cột (`Số HĐ`, `Số hoá đơn`, `SHĐ`, `Invoice No`, `Ký hiệu`, `MST`, `Diễn giải`, `Mặt hàng`, `Tiền hàng`, `Doanh số`, `Thuế GTGT`, `Tiền thuế`...).
- **Làm sạch số liệu tài chính**: Chuyển đổi an toàn dấu chấm/phẩy hàng nghìn, số âm trong ngoặc đơn `(1,234.56)` -> `-1234.56`, chuỗi số lẫn text.
- **Chuẩn hóa ngày tháng đa năng**: Nhận diện chuẩn cả Excel Serial Date (`44176`) và các dạng chuỗi ngày tháng (`dd/mm/yyyy`, `yyyy-mm-dd`, `dd-mm-yy`).
- **Chuẩn hóa thuế suất**: Xử lý `0.1` -> `10%`, `10` -> `10%`, `8%`, `5%`, `KCT`, `KKT`.

### 2. Thuật Toán Đối Chiếu Chéo 4 Nhóm (`reconciler.py`)
- **Khớp hoàn toàn**: Cùng số HĐ, ký hiệu, MST và số tiền doanh số, tiền thuế khớp nhau đến từng đồng.
- **Khớp HĐ nhưng lệch tiền**: Phát hiện sai lệch số tiền doanh số hoặc thuế giữa bảng kê và hóa đơn gốc TCT.
- **Có trên TCT nhưng bỏ sót trên Bảng kê**: Phát hiện các hóa đơn người bán đã xuất trên trang Thuế nhưng kế toán nội bộ chưa kê khai.
- **Có trên Bảng kê nhưng không có trên TCT**: Cảnh báo rủi ro hóa đơn giả mạo, hóa đơn khống, hoặc hóa đơn đã bị hủy trên hệ thống TCT.

### 3. Phân Tích & Đối Chiếu Tờ Khai 01/GTGT XML (`xml_parser.py`)
- Đọc trực tiếp file XML tờ khai 01/GTGT kết xuất từ HTKK hoặc trang Thuế điện tử (`eTax`).
- Tự động so sánh số tổng hợp từ hóa đơn/bảng kê với từng chỉ tiêu trên tờ khai.
- Đánh giá ngay: `Khớp 100%`, `Lệch nhỏ (Làm tròn)`, hoặc `LỆCH - CẦN RÀ SOÁT`.

### 4. Giao Diện Người Dùng Thân Thiện
- **Giao diện Desktop (GUI)**: Chạy qua `vat_reconciler_gui.py` hoặc click đúp file `run.bat`. Trực quan, dễ dùng cho kế toán viên.
- **Giao diện Dòng lệnh (CLI)**: Chạy qua `vat_reconciler_cli.py`, phù hợp chạy tự động, batch processing hàng loạt kỳ thuế.
- **Xuất Báo Cáo Excel Chuyên Sâu**: File `.xlsx` chuẩn quy thức thanh tra thuế, highlight màu trực quan (đỏ khi lệch, xanh khi khớp 100%), tự động đóng băng tiêu đề (Freeze Panes).

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng

### Yêu cầu hệ thống
- Python 3.8+ (khuyên dùng Python 3.10, 3.11 hoặc 3.12).
- Thư viện `openpyxl`.

```bash
pip install -r requirements.txt
```

### Cách 1: Chạy Giao Diện Đồ Họa (GUI)
Click đúp file `run.bat` (trên Windows) hoặc gõ:
```bash
python vat_reconciler_gui.py
```
1. Nhấp các nút **"Chọn file..."** để chọn file hóa đơn TCT, bảng kê nội bộ hoặc tờ khai XML.
2. Nhấn nút **"▶ THỰC HIỆN ĐỐI CHIẾU SỐ LIỆU"**.
3. Xem kết quả nhanh trên bảng Dashboard và nhấn **"📊 Mở File Excel Báo Cáo"** để xem chi tiết.

### Cách 2: Chạy Dòng Lệnh (CLI)
```bash
python vat_reconciler_cli.py \
    --in-gdt "Hoadon_Muavao_TCT.xlsx" \
    --in-nb "Bangke_Muavao_Noibo.xlsx" \
    --out-gdt "Hoadon_Banra_TCT.xlsx" \
    --out-nb "Bangke_Banra_Noibo.xlsx" \
    --xml "ToKhai_01_GTGT.xml" \
    --out "KetQua_DoiChieu_ThueGTGT.xlsx"
```

### Chạy Bộ Kiểm Thử Tự Động
```bash
python test_suite.py
```

---

## 📁 Cấu Trúc Dự Án

```text
vat_reconciliation_pro/
├── excel_normalizer.py       # Bộ chuẩn hóa định dạng Excel thông minh
├── xml_parser.py             # Trích xuất chỉ tiêu Tờ khai 01/GTGT XML
├── reconciler.py             # Thuật toán đối chiếu HĐ và Tờ khai
├── excel_exporter.py         # Xuất báo cáo Excel đối chiếu chuyên sâu
├── vat_reconciler_gui.py     # Giao diện Desktop GUI (Tkinter)
├── vat_reconciler_cli.py     # Giao diện dòng lệnh CLI
├── test_suite.py             # Bộ test cases tự động
├── run.bat                   # Phím tắt khởi động nhanh trên Windows
├── requirements.txt          # Danh sách thư viện phụ thuộc
├── .gitignore                # Bỏ qua file tạm Excel, cache
└── README.md                 # Tài liệu hướng dẫn sử dụng
```
