# 🔧 TROUBLESHOOTING AUTO-UPDATE

## 🚨 VẤN ĐỀ: Ứng dụng dừng không hoạt động ở bước cài đặt

### **🔍 NGUYÊN NHÂN CÓ THỂ:**

#### **1. File Locking Issues:**
- File .exe đang được sử dụng bởi ứng dụng
- Antivirus đang quét file
- Windows Explorer đang truy cập file

#### **2. Permission Issues:**
- Không có quyền Administrator
- Thư mục đích bị bảo vệ
- User Account Control (UAC) chặn

#### **3. Antivirus Interference:**
- Antivirus chặn thay thế file
- Real-time protection bật
- Quarantine file tạm thời

#### **4. Network/Download Issues:**
- File download không hoàn chỉnh
- Checksum mismatch
- Corrupted download

## 🛠️ GIẢI PHÁP ĐÃ IMPLEMENT (v1.0.4):

### **✅ Cải thiện Batch Script:**
```batch
@echo off
echo ========================================
echo    CÀI ĐẶT BẢN CẬP NHẬT QLVT
echo ========================================
echo.
echo Đang chuẩn bị cài đặt...
timeout /t 3 /nobreak >nul

echo Kiểm tra file nguồn...
if not exist "{new_exe_path}" (
    echo LỖI: File nguồn không tồn tại!
    pause
    exit /b 1
)

echo Kiểm tra file đích...
if not exist "{current_exe_path}" (
    echo LỖI: File đích không tồn tại!
    pause
    exit /b 1
)

echo Đang thay thế file...
copy "{new_exe_path}" "{current_exe_path}" /Y
if %errorlevel% equ 0 (
    echo CÀI ĐẶT THÀNH CÔNG!
    echo Khởi động lại ứng dụng...
    start "" "{current_exe_path}"
    del "{new_exe_path}"
    del "%~f0"
    exit /b 0
) else (
    echo LỖI CÀI ĐẶT!
    echo Mã lỗi: %errorlevel%
    pause
    exit /b 1
)
```

### **✅ Error Handling Cải thiện:**
- **File existence check**: Kiểm tra file nguồn và đích
- **Timeout handling**: 30 giây timeout cho batch script
- **Detailed error messages**: Hiển thị mã lỗi và hướng dẫn
- **Progress tracking**: Log chi tiết quá trình cài đặt

### **✅ User Experience Cải thiện:**
- **Success dialog**: Thông báo thành công rõ ràng
- **Error dialog**: Hiển thị lỗi chi tiết với hướng dẫn
- **Detailed error info**: Nút "Show Details" cho lỗi

## 🔧 CÁCH KHẮC PHỤC THỦ CÔNG:

### **Bước 1: Kiểm tra quyền Administrator**
```bash
# Chạy ứng dụng với quyền Administrator
# Right-click → "Run as administrator"
```

### **Bước 2: Tắt Antivirus tạm thời**
1. Mở Antivirus settings
2. Tắt Real-time protection
3. Thêm thư mục vào whitelist
4. Thử cài đặt lại

### **Bước 3: Kiểm tra file lock**
```bash
# Kiểm tra file có bị lock không
tasklist /fi "imagename eq QLVT_Processor_v1.0.1.exe"

# Kill process nếu cần
taskkill /f /im QLVT_Processor_v1.0.1.exe
```

### **Bước 4: Manual installation**
```bash
# Copy file thủ công
copy "QLVT_Processor_v1.0.4.exe" "QLVT_Processor_v1.0.1.exe" /Y

# Chạy file mới
start QLVT_Processor_v1.0.4.exe
```

## 📊 LOGS DEBUG:

### **Khi cài đặt thành công:**
```
[UPDATE] Tạo batch script: C:\Temp\QLVT_Update\update_qlvt.bat
[UPDATE] Chạy batch script...
[UPDATE] Batch script chạy thành công
[UPDATE] Output: CÀI ĐẶT THÀNH CÔNG!
[UPDATE] Cài đặt thành công, chuẩn bị restart
```

### **Khi cài đặt thất bại:**
```
[UPDATE] Tạo batch script: C:\Temp\QLVT_Update\update_qlvt.bat
[UPDATE] Chạy batch script...
[UPDATE] Batch script lỗi với mã: 1
[UPDATE] Error: LỖI CÀI ĐẶT!
[UPDATE] Cài đặt thất bại
```

## 🎯 KIỂM TRA SAU KHI FIX:

### **Test Case 1: Normal Update**
1. Chạy `QLVT_Processor_v1.0.1.exe`
2. Đợi auto-check (3 giây)
3. Click "Yes" để update
4. Quan sát progress bar
5. Kiểm tra batch script window
6. Verify restart thành công

### **Test Case 2: Error Handling**
1. Chạy `QLVT_Processor_v1.0.1.exe`
2. Mở file trong Notepad (tạo lock)
3. Thử update
4. Kiểm tra error message
5. Verify ứng dụng không bị crash

### **Test Case 3: Timeout Test**
1. Chạy `QLVT_Processor_v1.0.1.exe`
2. Thử update với network chậm
3. Kiểm tra timeout handling
4. Verify error message

## 🔍 DEBUGGING TOOLS:

### **Kiểm tra temp folder:**
```bash
# Xem file tạm
dir %TEMP%\QLVT_Update\

# Xem batch script
type %TEMP%\QLVT_Update\update_qlvt.bat
```

### **Kiểm tra process:**
```bash
# Xem process đang chạy
tasklist | findstr QLVT

# Xem file handles
handle.exe QLVT_Processor
```

### **Kiểm tra logs:**
```bash
# Xem Windows Event Log
eventvwr.msc

# Filter: Application errors
```

## 🎉 KẾT QUẢ MONG ĐỢI SAU FIX:

### **✅ Thành công:**
- Batch script chạy với giao diện rõ ràng
- Error messages chi tiết và hữu ích
- Timeout handling đúng cách
- Ứng dụng restart thành công
- File cleanup tự động

### **❌ Vẫn có thể gặp:**
- Antivirus chặn (cần tắt tạm thời)
- File permission (cần Administrator)
- Network issues (cần retry)

## 🚀 NEXT STEPS:

1. **Test với v1.0.4**: Chạy version cũ và test update
2. **Monitor logs**: Quan sát debug output
3. **User feedback**: Thu thập phản hồi từ user
4. **Iterative improvement**: Cải thiện dựa trên feedback

**Fix đã được implement trong v1.0.4 - hãy test ngay!** 🔧 