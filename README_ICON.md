# Icon cho ứng dụng QLVT

## Mô tả
Ứng dụng QLVT đã được thêm icon chuyên nghiệp với các đặc điểm sau:

### Thiết kế icon
- **Hình dạng**: Hình tròn với nền màu xanh dương
- **Chữ**: "QLVT" màu trắng ở giữa
- **Chi tiết**: 3 dấu chấm nhỏ màu trắng ở phía trên
- **Kích thước**: Hỗ trợ đa dạng kích thước (16x16 đến 256x256 pixels)

### File icon
- **Tên file**: `icon.ico`
- **Định dạng**: ICO (Windows Icon)
- **Kích thước file**: ~26KB

### Tích hợp
1. **Trong ứng dụng**: Icon hiển thị trên thanh tiêu đề cửa sổ
2. **Trong file .exe**: Icon được nhúng vào file thực thi khi build
3. **Trên Desktop**: Icon hiển thị khi tạo shortcut

### Cách sử dụng
- Icon tự động được tải khi khởi động ứng dụng
- Không cần cấu hình thêm
- Icon sẽ hiển thị trong:
  - Thanh tiêu đề cửa sổ
  - Taskbar
  - Alt+Tab switcher
  - File Explorer

### Build với icon
Khi build file .exe, PyInstaller sẽ tự động sử dụng icon:
```bash
pyinstaller --onefile --windowed --icon=icon.ico qlvt.py
```

### Lưu ý
- Icon được tạo tự động bằng Python PIL
- Hỗ trợ đầy đủ cho Windows
- Tương thích với hệ thống auto-update 