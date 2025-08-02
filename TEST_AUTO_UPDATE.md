# 🧪 HƯỚNG DẪN TEST HỆ THỐNG AUTO-UPDATE

## 🎯 MỤC TIÊU TEST
Kiểm tra hệ thống auto-update hoạt động từ version 1.0.1 lên 1.0.2

## 📁 FILES CẦN THIẾT
- `QLVT_Processor_v1.0.1.exe` (version cũ)
- `QLVT_Processor_v1.0.2.exe` (version mới - đã được release trên GitHub)

## 🚀 CÁC BƯỚC TEST

### **Bước 1: Chạy version cũ**
1. Chạy file `QLVT_Processor_v1.0.1.exe`
2. Ứng dụng sẽ hiển thị title: "Xử lý phiếu hàng loạt v1.0.1"
3. Sau 3 giây, ứng dụng sẽ tự động check update

### **Bước 2: Kiểm tra auto-update**
1. Nếu có version mới → Hiển thị dialog "Cập nhật mới"
2. Dialog sẽ hiển thị: "Có phiên bản mới: v1.0.2"
3. Click "Yes" để bắt đầu update

### **Bước 3: Quá trình update**
1. Hiển thị progress dialog "Đang cập nhật..."
2. Progress bar sẽ hiển thị % download
3. Khi 100% → "Đang cài đặt cập nhật..."
4. Thông báo "Cập nhật thành công! Ứng dụng sẽ khởi động lại."

### **Bước 4: Test manual check**
1. Click nút "🧪 Test Auto-Update"
2. Nếu không có version mới → "Không có phiên bản mới để cập nhật."
3. Nếu có version mới → Hiển thị dialog update

## 🔍 LOGS DEBUG

### **Khi check update:**
```
[UPDATE] Đang kiểm tra cập nhật từ nekennick/RunPhieu
[UPDATE] Phiên bản hiện tại: 1.0.1
[UPDATE] Phiên bản mới nhất: 1.0.2
[UPDATE] Có phiên bản mới: 1.0.2
```

### **Khi download:**
```
[UPDATE] Tìm thấy file: QLVT_Processor_v1.0.2.exe
[UPDATE] Bắt đầu tải xuống: https://github.com/nekennick/RunPhieu/releases/download/v1.0.2/QLVT_Processor_v1.0.2.exe
[UPDATE] Tải xuống hoàn tất: C:\Users\...\AppData\Local\Temp\QLVT_Update\QLVT_Processor_v1.0.2.exe
```

### **Khi cài đặt:**
```
[UPDATE] Cài đặt từ: C:\Users\...\AppData\Local\Temp\QLVT_Update\QLVT_Processor_v1.0.2.exe
[UPDATE] Cài đặt đến: D:\Python\QLVT\dist\QLVT_Processor_v1.0.1.exe
[UPDATE] Chạy batch script: C:\Users\...\AppData\Local\Temp\QLVT_Update\update_qlvt.bat
```

## 🎯 KẾT QUẢ MONG ĐỢI

### **✅ Thành công:**
- Ứng dụng tự động phát hiện version mới
- Download và cài đặt thành công
- Restart với version mới (v1.0.2)
- Title bar hiển thị "v1.0.2"
- Nút "🧪 Test Auto-Update" có sẵn

### **❌ Lỗi có thể gặp:**
- Network timeout → "Timeout khi kiểm tra cập nhật"
- File permission → "Lỗi cài đặt cập nhật"
- GitHub API error → "Lỗi API: 404/403"

## 🔧 TROUBLESHOOTING

### **Nếu không check được update:**
1. Kiểm tra internet connection
2. Kiểm tra GitHub repository: https://github.com/nekennick/RunPhieu/releases
3. Kiểm tra tag v1.0.2 đã được tạo

### **Nếu download bị lỗi:**
1. Kiểm tra file size (khoảng 42MB)
2. Kiểm tra thư mục temp: `%TEMP%\QLVT_Update\`
3. Kiểm tra antivirus có block không

### **Nếu cài đặt bị lỗi:**
1. Chạy với quyền Administrator
2. Kiểm tra file gốc có bị lock không
3. Kiểm tra disk space

## 📊 METRICS TEST

- **Auto-check time**: ~3 giây sau khởi động
- **Download time**: ~30-60 giây (tùy internet)
- **Install time**: ~5-10 giây
- **Total update time**: ~1-2 phút

## 🎉 HOÀN THÀNH TEST

Khi test thành công:
1. ✅ Version cũ (1.0.1) → Version mới (1.0.2)
2. ✅ Auto-update hoạt động hoàn hảo
3. ✅ User experience mượt mà
4. ✅ Error handling đầy đủ

**Hệ thống auto-update đã sẵn sàng cho production!** 🚀 