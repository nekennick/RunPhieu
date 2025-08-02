# Hệ thống Remote Activation/Deactivation cho QLVT

## Tổng quan

Hệ thống này cho phép bạn kích hoạt và vô hiệu hóa ứng dụng QLVT từ xa thông qua GitHub Gist. Khi ứng dụng bị vô hiệu hóa, nó sẽ hiển thị thông báo và tự động thoát.

## Tính năng chính

- ✅ **Kiểm tra activation khi khởi động**: Ứng dụng tự động kiểm tra trạng thái khi khởi động
- ✅ **Cache kết quả**: Giảm số lượng API calls bằng cách cache kết quả trong 1 giờ
- ✅ **Thông báo chi tiết**: Hiển thị thông báo rõ ràng khi bị deactivate
- ✅ **Nút kiểm tra thủ công**: Có thể kiểm tra trạng thái bất cứ lúc nào
- ✅ **Fallback mechanism**: Nếu có lỗi, ứng dụng vẫn chạy (fail-safe)
- ✅ **Timeout protection**: Bảo vệ khỏi network timeout

## Cách sử dụng

### 1. Setup ban đầu

#### Cách 1: Tự động (Khuyến nghị)
```bash
python setup_activation.py
```
- Chạy script và làm theo hướng dẫn
- Script sẽ tạo Gist và cung cấp Gist ID

#### Cách 2: Thủ công
1. Tạo GitHub Gist tại https://gist.github.com/
2. Thêm file `activation_status.json` với nội dung:
```json
{
  "activated": true,
  "expiry_date": "2025-12-31",
  "message": "Ứng dụng đang hoạt động bình thường",
  "last_updated": "2024-01-15T10:30:00Z"
}
```
3. Copy Gist ID từ URL
4. Cập nhật `self.gist_id` trong `qlvt.py`

### 2. Quản lý từ xa

#### Vô hiệu hóa ứng dụng:
1. Vào Gist đã tạo
2. Edit file `activation_status.json`
3. Thay đổi `"activated": false`
4. Thêm thông báo vào `"message"`
5. Cập nhật `"last_updated"`

#### Kích hoạt lại:
1. Edit Gist
2. Thay đổi `"activated": true`
3. Cập nhật thông tin khác

### 3. Kiểm tra trạng thái

- **Tự động**: Khi khởi động ứng dụng
- **Thủ công**: Nhấn nút "🔐 Kiểm tra trạng thái" trong ứng dụng

## Cấu trúc JSON

```json
{
  "activated": true,                    // true = kích hoạt, false = vô hiệu hóa
  "expiry_date": "2025-12-31",         // Ngày hết hạn (tùy chọn)
  "message": "Thông báo cho user",     // Thông báo hiển thị cho user
  "last_updated": "2024-01-15T10:30:00Z" // Thời gian cập nhật cuối
}
```

## Files quan trọng

- `qlvt.py`: File chính chứa code ứng dụng
- `setup_activation.py`: Script tự động setup Gist
- `SETUP_ACTIVATION.md`: Hướng dẫn chi tiết
- `activation_cache.json`: Cache trạng thái (tự động tạo)
- `activation_config.json`: Config sau khi setup (tự động tạo)

## Bảo mật

- Gist nên được set là "secret" (không public)
- Chỉ admin mới có quyền edit Gist
- Cache được lưu local để tránh spam API calls
- Timeout 10 giây cho network requests

## Troubleshooting

### Lỗi thường gặp

1. **"Lỗi API: 404"**
   - Gist ID không đúng hoặc Gist không tồn tại
   - Kiểm tra lại Gist ID trong code

2. **"Timeout khi kiểm tra activation"**
   - Kết nối internet chậm
   - Ứng dụng sẽ fallback về trạng thái mặc định

3. **"Không thể kiểm tra trạng thái activation"**
   - Lỗi network hoặc GitHub API
   - Xóa file `activation_cache.json` để force check lại

### Debug

Để debug, kiểm tra console output:
```
[ACTIVATION] Đang kiểm tra trạng thái activation...
[ACTIVATION] Trạng thái: {'activated': True, ...}
```

## Ví dụ sử dụng

### Tạm thời vô hiệu hóa do bảo trì:
```json
{
  "activated": false,
  "expiry_date": "2025-12-31",
  "message": "Ứng dụng tạm thời bị vô hiệu hóa do bảo trì hệ thống. Vui lòng thử lại sau 2 giờ.",
  "last_updated": "2024-01-15T14:30:00Z"
}
```

### Kích hoạt lại:
```json
{
  "activated": true,
  "expiry_date": "2025-12-31",
  "message": "Ứng dụng đã được kích hoạt lại. Cảm ơn sự kiên nhẫn của bạn.",
  "last_updated": "2024-01-15T16:30:00Z"
}
```

## Lưu ý

- Ứng dụng sẽ cache kết quả trong 1 giờ để giảm API calls
- Nếu có lỗi network, ứng dụng sẽ fallback về trạng thái mặc định (activated)
- File cache có thể bị xóa để force check lại trạng thái
- Hệ thống được thiết kế để fail-safe (không block ứng dụng khi có lỗi) 