# KAROFI – Báo cáo giao hàng NCC

Trang: https://khovanhcm.github.io/report-ncc/karofi/

- Nguồn yêu cầu: Google Sheets "KAROFI" (sheet 1.RULE BÁO CÁO).
- Data NCC: sheet "NCC KAROFI" / Trang tính1, gộp theo DO, lấy booking đến 15/11/2026.
- Trạng thái + ngày hoàn tất: đối chiếu bill giao hàng với 2.DATA BILL (cột AH, AI).
- Leadtime: 4.LEADTIME, tỉnh gửi Long An; đúng cam kết khi (ngày hoàn tất − ngày xe lấy) ≤ leadtime tối đa.
- Leadtime trả chứng từ = ngày trả chứng từ NCC − ngày hoàn tất.
- Dữ liệu nhúng sẵn trong index.html; cập nhật bằng cách nhờ Claude đọc lại 2 sheet và dựng lại trang.
