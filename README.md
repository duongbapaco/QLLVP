# QLLVP
# Hệ Thống Quản Trị Lương & Nhân Sự - Tổng Công Ty Giấy Việt Nam (VINAPACO)
Hệ thống quản lý, tính toán và đồng bộ bảng lương tự động cho các đơn vị trực thuộc VINAPACO, kết hợp giữa Web App giao diện trực quan, GitHub (quản lý mã nguồn) và Google Drive/Sheets (lưu trữ dữ liệu đám mây).
## 🚀 Các Tính Năng Chính
- **Quản lý đa kỳ/đa năm:** Hỗ trợ chọn năm lưu trữ (`2024`, `2025`, `2026`,...) và linh hoạt chuyển đổi các tháng trong năm.
- **Tự động hóa toàn diện 5 Tab chuyên nghiệp trên Google Sheets:**
  1. `Lương CB`: Tính lương cơ bản, ngày công thực tế, phép, lễ, bảo hiểm (10.5%).
  2. `Lương CV`: Tính lương công việc, hệ số hoàn thành công việc (HS HTCV), tạm ứng, quỹ PCTT và tổng thực lĩnh NET liên kết tự động với tab Lương CB.
  3. `Tạm ứng`: Danh sách chi tiết tiền tạm ứng kèm số tài khoản ngân hàng.
  4. `ĂnCa`: Thống kê tiền ăn ca theo ngày công thực tế.
  5. `Ngan hang`: Bảng lương thanh toán qua ngân hàng sẵn sàng xuất file cho kế toán.
- **Đồng bộ một chạm:** Đẩy toàn bộ số liệu trực tiếp từ web lên thư mục phân loại theo `Năm/Tháng` trên Google Drive doanh nghiệp.

## ⚙️ Hướng Dẫn Triển Khai & Cập Nhật Code
1. Clone hoặc tải mã nguồn từ repository GitHub này lên Google Apps Script (`Code.gs` và `Index.html`).
2. Trên giao diện Google Apps Script, chọn **Deploy > Manage deployments > Edit > New version > Deploy**.
3. Mở ứng dụng Web App và vận hành bình thường.
