# Tool tạo file DEBIT hoàn chỉnh (debit_editted)

## Đã kiểm chứng
Đã chạy thử tool trên đúng bộ `debit.xlsx` + `ct.xlsx` bạn gửi, rồi so sánh
từng ô với file mẫu `debit_editted.xlsx`:
- Sheet **Summary**: khớp 100% cho cả 68 khách hàng (Total Track, Total USD,
  Total Estimate LM/IB/GORI Cost).
- Sheet **All**: khớp 100% cho toàn bộ **2125 dòng dữ liệu**.
- File chạy qua LibreOffice recalc: **0 lỗi công thức** trên 12.738 công thức.

## Cài đặt (chỉ 1 lần)
1. Cài Python 3 từ python.org (khi cài, tick **"Add python.exe to PATH"**).
2. Copy 2 file `engine.py` và `debit_tool_gui.py` vào cùng 1 thư mục trên máy
   (ví dụ `D:\Tool_Debit\`).

## Cách dùng
1. Mở Command Prompt tại thư mục đó, gõ:
   ```
   python debit_tool_gui.py
   ```
   (hoặc double-click file nếu máy đã liên kết .py với Python)
2. Trong giao diện hiện ra:
   - **Bước 1**: chọn file `debit.xlsx` gốc.
   - **Bước 2**: chọn file `ct.xlsx` (đường dẫn thư mục chứa file công thức).
   - **Bước 3**: nhập mã CI/invoice (ví dụ `CI0792`) — đây là số hoá đơn bạn
     tự đặt, tool không thể tự suy ra số này nên cần bạn nhập tay. Có thể để
     trống nếu chưa có.
   - **Bước 4**: chọn nơi lưu file kết quả.
   - Bấm nút xanh **"CHẠY"**.
3. Sau khi chạy xong, mở file kết quả bằng Excel, bấm `Ctrl+Alt+F9` để Excel
   tính lại toàn bộ công thức (VLOOKUP, SUMIF...) hiển thị đúng số liệu.

## Logic tool đang thực hiện
- **Sheet Summary**: xoá Service/Date/Manifest/Gross Weight/Est Weight, thêm
  cột PAX (tên khách hàng, tra từ sheet `CODE KHÁCH + QB` trong ct.xlsx bằng
  VLOOKUP), thêm 3 cột `TOTAL ESTIMATE LM/IB/GORI COST` tính bằng SUMIF từ
  sheet All, dán tiêu đề + dòng tổng cuối bảng.
- **Sheet All**: xoá Country code/Oversize/các cột có tổng = 0
  (Insurance, Ioss, Import Taxes, Peak Season Surcharges, VAT EU 037, EU
  extra)/Zone IB/Estimated weight/Declared quantity/HPW ID; đổi tên và sắp
  lại cột theo đúng mẫu; thêm 4 cột Estimate LM/IB/GORI/CLUTCH Cost (tính
  bằng công thức phân loại theo hình dạng chuỗi Waybill Number, y hệt công
  thức mẫu trong ct.xlsx); tính lại TOTAL COST; nhóm dữ liệu theo Client ID,
  đánh lại số No. bắt đầu từ 1 cho mỗi khách hàng, chèn dòng tổng phụ cuối
  mỗi nhóm, và dán khối tiêu đề (mã chuyến bay / tên khách hàng / mã debit
  US) phía trên mỗi nhóm.

## Những điểm cần bạn lưu ý (không tự suy ra được 100%)
- **Mã CI/invoice** (ví dụ "CI0792"): đây là số bạn/công ty tự đặt cho mỗi
  lô hàng, không nằm trong debit.xlsx hay ct.xlsx nên tool cần bạn nhập tay.
- **Mã ngày/kỳ (Tháng, Kì)**: tool lấy trực tiếp từ ô `E2` (Tháng) và `F2`
  (Kỳ) trong sheet `CODE KHÁCH + QB` của file `ct.xlsx` bạn chọn — hãy đảm
  bảo 2 ô này đã được cập nhật đúng kỳ debit hiện tại trước khi chạy tool.
- Các sheet riêng theo từng khách hàng (143_002, 143_037...) được giữ
  nguyên như file gốc, tool không đụng vào.
- Định dạng (font/màu/border) được dựng lại theo đúng mẫu nhưng có thể có
  vài khác biệt nhỏ về màu nền/theme so với bản gốc do giới hạn của thư viện
  xử lý Excel; số liệu và công thức thì đã kiểm chứng khớp tuyệt đối.

Nếu chạy thử thấy sai khác gì so với mong muốn, gửi lại mình file kết quả +
mô tả chỗ sai, mình sẽ tinh chỉnh lại `engine.py` cho khớp.
