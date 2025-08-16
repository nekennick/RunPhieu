---
trigger: manual
---

# QLVT Detailed Design

## Kiến trúc
- Mô hình MVC đơn giản
- Main window (WordProcessorApp) là controller chính
- Các worker threads xử lý tác vụ nặng

## Các module chính
1. ActivationManager
   - Kiểm tra trạng thái kích hoạt qua GitHub Gist
   - Xử lý các trường hợp lỗi kết nối

2. WordProcessorApp
   - Giao diện chính của ứng dụng
   - Quản lý các chức năng xử lý văn bản

3. ReplaceWorker
   - Thread xử lý thay thế nội dung
   - Tránh treo giao diện

## Quy tắc code
1. Error Handling
   - Bắt tất cả exceptions có thể xảy ra
   - Log lỗi chi tiết với prefix [DEBUG]
   - Hiển thị thông báo lỗi thân thiện cho user

2. Đặt tên
   - Tên biến/hàm: tiếng Việt không dấu
   - Prefix qt_ cho các widget PyQt
   - Suffix _worker cho các thread

3. Comment
   - Docstring tiếng Việt cho các class/method chính
   - Comment chi tiết các xử lý phức tạp
   - Debug log với prefix [DEBUG]