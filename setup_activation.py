import requests
import json
import os

def create_activation_gist(github_token):
    """Tạo Gist cho hệ thống activation"""
    
    # Nội dung Gist
    gist_content = {
        "activated": True,
        "expiry_date": "2025-12-31",
        "message": "Ứng dụng đang hoạt động bình thường",
        "last_updated": "2024-01-15T10:30:00Z"
    }
    
    # Tạo Gist
    gist_data = {
        "description": "QLVT Activation Status",
        "public": False,  # Secret gist
        "files": {
            "activation_status.json": {
                "content": json.dumps(gist_content, ensure_ascii=False, indent=2)
            }
        }
    }
    
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    try:
        response = requests.post(
            "https://api.github.com/gists",
            headers=headers,
            json=gist_data
        )
        
        if response.status_code == 201:
            gist_info = response.json()
            gist_id = gist_info['id']
            gist_url = gist_info['html_url']
            
            print(f"✅ Tạo Gist thành công!")
            print(f"Gist ID: {gist_id}")
            print(f"Gist URL: {gist_url}")
            print(f"\nCập nhật code với Gist ID: {gist_id}")
            
            return gist_id
        else:
            print(f"❌ Lỗi tạo Gist: {response.status_code}")
            print(f"Response: {response.text}")
            return None
            
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

if __name__ == "__main__":
    print("=== Setup Remote Activation System ===")
    print("Bạn cần GitHub Personal Access Token để tạo Gist.")
    print("Tạo token tại: https://github.com/settings/tokens")
    print("Token cần quyền: gist")
    print()
    
    token = input("Nhập GitHub Personal Access Token: ").strip()
    
    if not token:
        print("❌ Token không được để trống!")
        exit(1)
    
    gist_id = create_activation_gist(token)
    
    if gist_id:
        print(f"\n=== Hướng dẫn tiếp theo ===")
        print(f"1. Mở file qlvt.py")
        print(f"2. Tìm dòng: self.gist_id = \"YOUR_GIST_ID_HERE\"")
        print(f"3. Thay thế bằng: self.gist_id = \"{gist_id}\"")
        print(f"4. Lưu file và chạy lại ứng dụng")
        print(f"\nGist URL: https://gist.github.com/YOUR_USERNAME/{gist_id}")
        print("(Thay YOUR_USERNAME bằng username GitHub của bạn)")
        
        # Tạo file config
        config = {
            "gist_id": gist_id,
            "setup_date": "2024-01-15T10:30:00Z"
        }
        
        with open("activation_config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ Đã lưu config vào file: activation_config.json")
    else:
        print("❌ Không thể tạo Gist. Vui lòng kiểm tra token và thử lại.") 