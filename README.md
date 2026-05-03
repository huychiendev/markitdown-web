# System Architecture & Code Map

## Giới thiệu

Ứng dụng `MarkItDown Web UI` là một wrapper web gọn nhẹ dành cho thư viện `markitdown` của Microsoft. Ứng dụng giải quyết triệt để hạn chế của `markitdown` nguyên bản khi chuyển đổi file PPTX (chỉ trích xuất text, mất ảnh), bằng cách tự động can thiệp vào mã nguồn file để trích xuất ảnh và chèn lại link hình ảnh vào mã Markdown.

## Cấu trúc thư mục

- `/app.py`:
  - **Backend API (FastAPI)**. Điểm khởi đầu của ứng dụng.
  - Chứa route `/` để trả về giao diện frontend.
  - Chứa route `/api/convert` thực hiện logic nhận file upload, gọi `markitdown`, extract ảnh bằng `python-pptx` và trả về kết quả.

- `/static/index.html`:
  - **Frontend UI**.
  - Xử lý kéo thả file, tương tác người dùng, gọi API chuyển đổi, preview Markdown, tải xuống.
  - _Xem thêm chi tiết tại `structure/UI_TREE.md`._

- `/static/conversions/`:
  - Thư mục động lưu file được convert và ảnh extract ra từ file PPTX.

- `Dockerfile` & `docker-compose.yml`:
  - Phục vụ môi trường chạy production trên VPS thông qua Docker. Giữ máy tính gốc hoàn toàn sạch sẽ.
