# Script để xóa và tạo lại tag v1.0.21

Write-Host "=== Rebuild Tag v1.0.21 ===" -ForegroundColor Cyan
Write-Host ""

# Bước 1: Xóa tag local
Write-Host "Bước 1: Xóa tag local v1.0.21..." -ForegroundColor Yellow
git tag -d v1.0.21
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã xóa tag local" -ForegroundColor Green
} else {
    Write-Host "⚠ Tag local không tồn tại hoặc đã bị xóa" -ForegroundColor Yellow
}
Write-Host ""

# Bước 2: Xóa tag trên GitHub
Write-Host "Bước 2: Xóa tag trên GitHub..." -ForegroundColor Yellow
git push origin :refs/tags/v1.0.21
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã xóa tag trên GitHub" -ForegroundColor Green
} else {
    Write-Host "⚠ Tag trên GitHub không tồn tại hoặc đã bị xóa" -ForegroundColor Yellow
}
Write-Host ""

# Bước 3: Commit code mới
Write-Host "Bước 3: Commit code mới..." -ForegroundColor Yellow
git add .
git commit -m "Fix: Sửa lỗi thay thế text không hoạt động và in trang đầu tiên"
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã commit code mới" -ForegroundColor Green
} else {
    Write-Host "⚠ Không có thay đổi để commit hoặc đã commit rồi" -ForegroundColor Yellow
}
Write-Host ""

# Bước 4: Push code
Write-Host "Bước 4: Push code lên GitHub..." -ForegroundColor Yellow
git push
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã push code" -ForegroundColor Green
} else {
    Write-Host "✗ Lỗi khi push code" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Bước 5: Tạo lại tag v1.0.21
Write-Host "Bước 5: Tạo lại tag v1.0.21..." -ForegroundColor Yellow
git tag v1.0.21
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã tạo tag local" -ForegroundColor Green
} else {
    Write-Host "✗ Lỗi khi tạo tag local" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Bước 6: Push tag lên GitHub
Write-Host "Bước 6: Push tag lên GitHub..." -ForegroundColor Yellow
git push origin v1.0.21
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Đã push tag lên GitHub" -ForegroundColor Green
} else {
    Write-Host "✗ Lỗi khi push tag" -ForegroundColor Red
    exit 1
}
Write-Host ""

Write-Host "=== HOÀN TẤT ===" -ForegroundColor Cyan
Write-Host "Tag v1.0.21 đã được tạo lại với code mới!" -ForegroundColor Green
Write-Host "GitHub Actions sẽ tự động build file exe." -ForegroundColor Green
Write-Host ""
