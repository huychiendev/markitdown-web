# System Architecture & Code Map

## Giới thiệu

Ứng dụng `MarkItDown Web UI` là một wrapper web gọn nhẹ dành cho thư viện `markitdown` của Microsoft. Ứng dụng giải quyết triệt để các hạn chế của `markitdown` nguyên bản:
- **PPTX**: Trích xuất ảnh và chèn lại link hình ảnh vào Markdown (markitdown gốc mất ảnh).
- **Excel (.xlsx)**: Trích xuất cả giá trị lẫn công thức gốc. Hiển thị dạng `value (=FORMULA)` thay vì chỉ giá trị.

## Cấu trúc thư mục

- `/app.py`:
  - **Backend API (FastAPI)**. Điểm khởi đầu của ứng dụng.
  - Chứa route `/` để trả về giao diện frontend.
  - Chứa route `/api/convert` và `/api/convert_batch`: nhận file upload, xử lý convert. File `.xlsx` dùng `openpyxl` trực tiếp (hỗ trợ formula), các file khác dùng `markitdown`.

- `/static/index.html`:
  - **Frontend UI**.
  - Xử lý kéo thả file, tương tác người dùng, gọi API chuyển đổi, preview Markdown, tải xuống.
  - _Xem thêm chi tiết tại `structure/UI_TREE.md`._

- `/static/conversions/`:
  - Thư mục động lưu file được convert và ảnh extract ra từ file PPTX.

- `Dockerfile` & `docker-compose.yml`:
  - Phục vụ môi trường chạy production trên VPS thông qua Docker. Giữ máy tính gốc hoàn toàn sạch sẽ.
