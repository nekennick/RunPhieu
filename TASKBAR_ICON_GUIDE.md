# Hướng dẫn hiển thị icon trên Taskbar

## Vấn đề
Icon không hiển thị trên taskbar Windows có thể do nhiều nguyên nhân. Dưới đây là các bước khắc phục:

## ✅ Đã thực hiện trong code

### 1. Thiết lập icon cho QApplication
```python
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Thiết lập icon cho toàn bộ ứng dụng
    icon = QIcon("icon.ico")
    app.setWindowIcon(icon)
    
    # Thiết lập tên ứng dụng cho taskbar
    app.setApplicationName("QLVT Processor")
    app.setApplicationDisplayName("QLVT Processor")
```

### 2. Thiết lập icon cho cửa sổ chính
```python
class WordProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # Thiết lập icon cho ứng dụng
        icon = QIcon("icon.ico")
        self.setWindowIcon(icon)
        
        # Thiết lập icon cho taskbar (Windows)
        if hasattr(self, 'setWindowIcon'):
            self.setWindowIcon(icon)
            
        # Thiết lập thuộc tính cửa sổ để hiển thị icon tốt hơn
        self.setWindowFlags(self.windowFlags() | Qt.Window)
```

### 3. Cập nhật PyInstaller spec
```python
exe = EXE(
    # ... other options ...
    icon='icon.ico',
)
```

## 🔧 Khắc phục thủ công (nếu cần)

### Bước 1: Kiểm tra file icon
- Đảm bảo file `icon.ico` tồn tại trong thư mục
- File phải có định dạng ICO hợp lệ
- Kích thước file khoảng 26KB

### Bước 2: Xóa cache Windows
1. Mở Task Manager
2. Tìm và kết thúc tất cả tiến trình Python
3. Xóa cache icon Windows:
   ```
   ie4uinit.exe -show
   ie4uinit.exe -ClearIconCache
   ```

### Bước 3: Kiểm tra Windows Explorer
1. Mở File Explorer
2. Đi đến thư mục chứa file .exe
3. Kiểm tra xem icon có hiển thị đúng không
4. Nếu không, click chuột phải → Properties → Change Icon

### Bước 4: Rebuild ứng dụng
```bash
# Xóa các file build cũ
rmdir /s build
rmdir /s dist
del *.spec

# Build lại với icon
pyinstaller --onefile --windowed --icon=icon.ico qlvt.py
```

## 🧪 Test icon

### Chạy test script
```bash
python test_icon.py
```

### Kiểm tra các vị trí:
1. **Thanh tiêu đề cửa sổ**: Icon phải hiển thị bên trái tiêu đề
2. **Taskbar**: Icon phải hiển thị khi ứng dụng đang chạy
3. **Alt+Tab**: Icon phải hiển thị khi chuyển đổi ứng dụng
4. **File Explorer**: Icon phải hiển thị cho file .exe

## 📋 Checklist

- [ ] File `icon.ico` tồn tại và hợp lệ
- [ ] Code thiết lập icon cho QApplication
- [ ] Code thiết lập icon cho cửa sổ chính
- [ ] PyInstaller spec có `icon='icon.ico'`
- [ ] GitHub Actions build với `--icon=icon.ico`
- [ ] Test script chạy thành công
- [ ] Icon hiển thị trên taskbar

## 🚨 Lưu ý quan trọng

1. **Windows 10/11**: Icon có thể mất vài giây để hiển thị
2. **Cache**: Windows cache icon, có thể cần restart Explorer
3. **DPI Scaling**: Icon có thể bị mờ trên màn hình độ phân giải cao
4. **File .exe**: Icon chỉ hiển thị đầy đủ sau khi build thành file .exe

## 🔄 Troubleshooting

### Icon không hiển thị trên taskbar:
1. Kiểm tra `app.setWindowIcon(icon)` đã được gọi
2. Kiểm tra `self.setWindowIcon(icon)` trong class chính
3. Thử restart Windows Explorer
4. Rebuild ứng dụng với icon mới

### Icon hiển thị mờ:
1. Tạo icon với độ phân giải cao hơn (256x256)
2. Sử dụng PNG thay vì ICO
3. Kiểm tra DPI scaling settings

### Icon không hiển thị trong Alt+Tab:
1. Đảm bảo `Qt.Window` flag được set
2. Kiểm tra window không bị minimize
3. Thử `self.setWindowState(Qt.WindowActive)` 