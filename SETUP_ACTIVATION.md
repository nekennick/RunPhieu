# Hướng dẫn thiết lập hệ thống Remote Activation

## Bước 1: Tạo GitHub Gist

1. Truy cập https://gist.github.com/
2. Đăng nhập vào tài khoản GitHub
3. Tạo Gist mới với nội dung:

```json
{
  "activated": true,
  "expiry_date": "2025-12-31",
  "message": "Ứng dụng đang hoạt động bình thường",
  "last_updated": "2024-01-15T10:30:00Z"
}
```

4. Đặt tên file là `activation_status.json`
5. Chọn "Create secret gist" (không public)
6. Click "Create secret gist"

## Bước 2: Lấy Gist ID

1. Sau khi tạo Gist, URL sẽ có dạng: `https://gist.github.com/username/GIST_ID`
2. Copy GIST_ID (chuỗi ký tự dài)
3. Thay thế `YOUR_GIST_ID_HERE` trong code bằng GIST_ID thực

## Bước 3: Cập nhật code

Trong file `qlvt.py`, tìm dòng:
```python
self.gist_id = "YOUR_GIST_ID_HERE"
```

Thay thế bằng GIST_ID thực:
```python
self.gist_id = "abc123def456ghi789"  # GIST_ID thực của bạn
```

## Bước 4: Test hệ thống

1. Chạy ứng dụng - nó sẽ kiểm tra activation
2. Nếu `activated: true` - ứng dụng chạy bình thường
3. Nếu `activated: false` - ứng dụng hiển thị thông báo và thoát

## Bước 5: Quản lý từ xa

### Để vô hiệu hóa ứng dụng:
1. Vào Gist đã tạo
2. Edit file `activation_status.json`
3. Thay đổi:
```json
{
  "activated": false,
  "expiry_date": "2025-12-31",
  "message": "Ứng dụng tạm thời bị vô hiệu hóa do bảo trì",
  "last_updated": "2024-01-15T10:30:00Z"
}
```

### Để kích hoạt lại:
1. Edit Gist
2. Thay đổi `"activated": true`
3. Cập nhật `message` và `last_updated`

## Lưu ý bảo mật

- Gist nên được set là "secret" (không public)
- Chỉ admin mới có quyền edit Gist
- Có thể thêm authentication nếu cần bảo mật cao hơn
- Cache được lưu local để tránh spam API calls

## Troubleshooting

### Lỗi "Không thể kiểm tra trạng thái activation"
- Kiểm tra kết nối internet
- Kiểm tra Gist ID có đúng không
- Kiểm tra Gist có tồn tại và accessible không

### Ứng dụng không chạy
- Kiểm tra `activated` có phải `true` không
- Kiểm tra `expiry_date` có quá hạn không
- Xóa file `activation_cache.json` để force check lại

## Cấu trúc JSON

```json
{
  "activated": true,                    // true = kích hoạt, false = vô hiệu hóa
  "expiry_date": "2025-12-31",         // Ngày hết hạn (tùy chọn)
  "message": "Thông báo cho user",     // Thông báo hiển thị cho user
  "last_updated": "2024-01-15T10:30:00Z" // Thời gian cập nhật cuối
}
```

## Tính năng

- ✅ Kiểm tra activation khi khởi động
- ✅ Cache kết quả để giảm API calls
- ✅ Thông báo chi tiết khi bị deactivate
- ✅ Nút kiểm tra trạng thái thủ công
- ✅ Fallback mechanism khi có lỗi
- ✅ Timeout cho network requests 