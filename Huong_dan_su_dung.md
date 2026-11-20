# Worklog Report Pro (v2.0 — Standalone Windows `.exe`)

Ứng dụng Windows Desktop chạy độc lập (**không cần cài Python, không mở cửa sổ lệnh**), thiết kế tối ưu cho thao tác tổng hợp báo cáo `Work_log` (+ Excel Notes) & `Inspection/Rework` hằng ngày.

## Cách mở ứng dụng
Nhấp đúp trực tiếp vào file **`WorklogReportPro.exe`** (đặt sẵn trong `Worklog_Report_Tool\dist\WorklogReportPro.exe` và bản sao ngay tại thư mục `Hoa\WorklogReportPro.exe`).

## Tính năng nổi bật (v2.0)
1. **Kéo thả file Excel trực tiếp (`Drag & Drop`)**:
   - Mở lên tự động nhận diện `Work_log_*.xlsx` và `Inspection_And_Rework_*.xlsx` cùng thư mục (và ghi nhớ file đã chọn lần trước).
   - Kéo thả file `.xlsx` tháng mới vào bất kỳ vị trí nào trên cửa sổ app để nạp ngay.
2. **Giao diện Sáng / Tối (`Light / Dark Mode`) & 5 thẻ KPI trực quan**:
   - Hiển thị nhanh Nhân sự, Work Load (Frame + TB frame/h), Work BBOX (+ TB bbox/h), Work Time, và Tỷ lệ lỗi + Trạng thái đối chiếu 2 file (bấm trực tiếp vào thẻ Lỗi để nhảy sang tab Đối chiếu).
3. **Lọc linh hoạt 1 người hoặc nhiều người**:
   - Lọc khoảng thời gian tùy ý + nút chọn nhanh (`Tuần trước`, `Tuần này`, `Tháng này`, `Tất cả`).
   - Chọn **TASK** bằng thẻ pill trực quan (`OD`, `SSEG`, `PFE`...).
   - Chọn **Annotator** theo danh sách có hiện sẵn sản lượng trong kỳ; **nhấp đúp** vào 1 tên để xem riêng 1 người (có thanh thông báo + nút quay lại tất cả).
4. **Chuyển đổi nguồn & Đối chiếu chi tiết tới từng ngày lệch**:
   - Chuyển nhanh nguồn Frame/BBOX giữa **Rework `#IMAGE`** và **Work_log `+ Note`**.
   - Chuyển nhanh cột ngày Rework giữa **Ngày làm việc (Cột V)** và **DATE Inspect (Cột A)**.
   - Tự động đối chiếu 2 file và **trỏ đích danh từng ngày bị lệch** (`Ngày bị lệch`, `Ô Work_log`, `Dòng Rework`, `Gợi ý nguyên nhân`):
     - Rê chuột vào dòng `⚠` trên Mẫu 1 / Mẫu 2 để xem nhanh các ngày lệch.
     - Nhấp đúp dòng `⚠` (hoặc bấm vào thẻ KPI Lỗi) để nhảy thẳng tới danh sách **Ngày bị lệch**.
     - Nhấp đúp 1 dòng ngày lệch để mở đúng **Excel Note** của ngày đó.
5. **Xuất Excel chuẩn 2 mẫu + Công thức sống**:
   - Xuất đồng thời `Report_By_Task` (Mẫu 1), `Report_By_Annotator` (Mẫu 2), và `Audit_Compare_&_Notes` với đầy đủ merge cell, viền đậm phân tách, chữ đỏ cho Error Rate/Average và công thức Excel tự tính.

## Phím tắt tiện dụng
- `F5`: Nạp lại 2 file Excel sau khi vừa chỉnh sửa xong
- `Ctrl + F`: Tìm nhanh tên Annotator
- `Ctrl + C`: Copy bảng đang xem (dán thẳng vào Excel bằng `Ctrl+V`)
- `Ctrl + E`: Xuất file Excel (`.xlsx`)
