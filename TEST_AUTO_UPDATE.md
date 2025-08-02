# 🧪 HƯỚNG DẪN TEST HỆ THỐNG AUTO-UPDATE

## 🎯 MỤC TIÊU TEST
Kiểm tra hệ thống auto-update hoạt động qua nhiều phiên bản:
- **Test 1**: v1.0.1 → v1.0.2
- **Test 2**: v1.0.2 → v1.0.3
- **Test 3**: v1.0.1 → v1.0.3 (skip version)

## 📁 FILES CẦN THIẾT
- `QLVT_Processor_v1.0.1.exe` (version cũ nhất)
- `QLVT_Processor_v1.0.2.exe` (version trung gian)
- `QLVT_Processor_v1.0.3.exe` (version mới nhất - đã được release trên GitHub)

## 🚀 CÁC BƯỚC TEST

### **🧪 TEST 1: v1.0.1 → v1.0.2**

#### **Bước 1: Chạy version cũ**
1. Chạy file `QLVT_Processor_v1.0.1.exe`
2. Ứng dụng sẽ hiển thị title: "Xử lý phiếu hàng loạt v1.0.1"
3. Sau 3 giây, ứng dụng sẽ tự động check update

#### **Bước 2: Kiểm tra auto-update**
1. Nếu có version mới → Hiển thị dialog "Cập nhật mới"
2. Dialog sẽ hiển thị: "Có phiên bản mới: v1.0.2"
3. Click "Yes" để bắt đầu update

#### **Bước 3: Quá trình update**
1. Hiển thị progress dialog "Đang cập nhật..."
2. Progress bar sẽ hiển thị % download
3. Khi 100% → "Đang cài đặt cập nhật..."
4. Thông báo "Cập nhật thành công! Ứng dụng sẽ khởi động lại."

#### **Bước 4: Verify kết quả**
- ✅ Title bar hiển thị "v1.0.2"
- ✅ Nút "🧪 Test Auto-Update" có sẵn
- ✅ Nút "📊 Version Info" có sẵn

---

### **🧪 TEST 2: v1.0.2 → v1.0.3**

#### **Bước 1: Chạy version trung gian**
1. Chạy file `QLVT_Processor_v1.0.2.exe`
2. Ứng dụng sẽ hiển thị title: "Xử lý phiếu hàng loạt v1.0.2"
3. Sau 3 giây, ứng dụng sẽ tự động check update

#### **Bước 2: Kiểm tra auto-update**
1. Nếu có version mới → Hiển thị dialog "Cập nhật mới"
2. Dialog sẽ hiển thị: "Có phiên bản mới: v1.0.3"
3. Click "Yes" để bắt đầu update

#### **Bước 3: Verify kết quả**
- ✅ Title bar hiển thị "v1.0.3"
- ✅ Nút "🧪 Test Auto-Update" có sẵn
- ✅ Nút "📊 Version Info" có sẵn
- ✅ Click "📊 Version Info" → Hiển thị "Phiên bản hiện tại: v1.0.3"

---

### **🧪 TEST 3: v1.0.1 → v1.0.3 (Skip Version)**

#### **Bước 1: Chạy version cũ nhất**
1. Chạy file `QLVT_Processor_v1.0.1.exe`
2. Ứng dụng sẽ hiển thị title: "Xử lý phiếu hàng loạt v1.0.1"
3. Sau 3 giây, ứng dụng sẽ tự động check update

#### **Bước 2: Kiểm tra auto-update**
1. Nếu có version mới → Hiển thị dialog "Cập nhật mới"
2. Dialog sẽ hiển thị: "Có phiên bản mới: v1.0.3" (không phải v1.0.2)
3. Click "Yes" để bắt đầu update

#### **Bước 3: Verify kết quả**
- ✅ Title bar hiển thị "v1.0.3" (skip qua v1.0.2)
- ✅ Semantic versioning hoạt động đúng

---

### **🧪 TEST 4: Manual Check**

#### **Bước 1: Chạy bất kỳ version nào**
1. Chạy file `QLVT_Processor_v1.0.1.exe` hoặc `v1.0.2.exe`
2. Click nút "🧪 Test Auto-Update"

#### **Bước 2: Kiểm tra kết quả**
1. Nếu có version mới → Hiển thị dialog update
2. Nếu không có version mới → "Không có phiên bản mới để cập nhật."

---

### **🧪 TEST 5: Version Info**

#### **Bước 1: Chạy version mới nhất**
1. Chạy file `QLVT_Processor_v1.0.3.exe`
2. Click nút "📊 Version Info"

#### **Bước 2: Kiểm tra kết quả**
- ✅ Hiển thị dialog "Thông tin phiên bản"
- ✅ Text: "Phiên bản hiện tại: v1.0.3"
- ✅ Informative text: "Bạn đang sử dụng phiên bản này để xử lý phiếu hàng loạt."

## 🔍 LOGS DEBUG

### **Khi check update:**
```
[UPDATE] Đang kiểm tra cập nhật từ nekennick/RunPhieu
[UPDATE] Phiên bản hiện tại: 1.0.1
[UPDATE] Phiên bản mới nhất: 1.0.3
[UPDATE] Có phiên bản mới: 1.0.3
```

### **Khi download:**
```
[UPDATE] Tìm thấy file: QLVT_Processor_v1.0.3.exe
[UPDATE] Bắt đầu tải xuống: https://github.com/nekennick/RunPhieu/releases/download/v1.0.3/QLVT_Processor_v1.0.3.exe
[UPDATE] Tải xuống hoàn tất: C:\Users\...\AppData\Local\Temp\QLVT_Update\QLVT_Processor_v1.0.3.exe
```

### **Khi cài đặt:**
```
[UPDATE] Cài đặt từ: C:\Users\...\AppData\Local\Temp\QLVT_Update\QLVT_Processor_v1.0.3.exe
[UPDATE] Cài đặt đến: D:\Python\QLVT\dist\QLVT_Processor_v1.0.1.exe
[UPDATE] Chạy batch script: C:\Users\...\AppData\Local\Temp\QLVT_Update\update_qlvt.bat
```

## 🎯 KẾT QUẢ MONG ĐỢI

### **✅ Thành công:**
- Ứng dụng tự động phát hiện version mới nhất
- Download và cài đặt thành công
- Restart với version mới nhất
- Title bar hiển thị version chính xác
- Tất cả nút chức năng hoạt động
- Semantic versioning hoạt động đúng

### **❌ Lỗi có thể gặp:**
- Network timeout → "Timeout khi kiểm tra cập nhật"
- File permission → "Lỗi cài đặt cập nhật"
- GitHub API error → "Lỗi API: 404/403"

## 🔧 TROUBLESHOOTING

### **Nếu không check được update:**
1. Kiểm tra internet connection
2. Kiểm tra GitHub repository: https://github.com/nekennick/RunPhieu/releases
3. Kiểm tra tag v1.0.3 đã được tạo

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
1. ✅ Version cũ → Version mới nhất (skip version trung gian)
2. ✅ Auto-update hoạt động hoàn hảo
3. ✅ User experience mượt mà
4. ✅ Error handling đầy đủ
5. ✅ Semantic versioning chính xác
6. ✅ Tất cả tính năng mới hoạt động

## 🆕 TÍNH NĂNG MỚI TRONG V1.0.3

- ✅ **Nút "📊 Version Info"**: Hiển thị thông tin phiên bản
- ✅ **Auto-update system**: Hoàn chỉnh và ổn định
- ✅ **Semantic versioning**: So sánh version chính xác
- ✅ **Error handling**: Xử lý lỗi chi tiết

**Hệ thống auto-update đã sẵn sàng cho production!** 🚀 