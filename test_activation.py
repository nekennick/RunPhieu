#!/usr/bin/env python3
"""
Test script cho hệ thống Remote Activation
"""

import json
import requests
from datetime import datetime

def test_activation_status(gist_id):
    """Test kiểm tra trạng thái activation"""
    
    print("=== Test Remote Activation System ===")
    print(f"Gist ID: {gist_id}")
    print()
    
    # Test 1: Kiểm tra Gist có tồn tại không
    print("1. Kiểm tra Gist có tồn tại...")
    api_url = f"https://api.github.com/gists/{gist_id}"
    
    try:
        response = requests.get(api_url, timeout=10)
        if response.status_code == 200:
            print("✅ Gist tồn tại")
            gist_data = response.json()
            
            # Test 2: Kiểm tra file activation_status.json
            print("\n2. Kiểm tra file activation_status.json...")
            files = gist_data.get('files', {})
            
            if 'activation_status.json' in files:
                print("✅ File activation_status.json tồn tại")
                
                # Test 3: Parse JSON content
                print("\n3. Parse JSON content...")
                content = files['activation_status.json']['content']
                
                try:
                    status_data = json.loads(content)
                    print("✅ JSON hợp lệ")
                    print(f"   - activated: {status_data.get('activated')}")
                    print(f"   - expiry_date: {status_data.get('expiry_date')}")
                    print(f"   - message: {status_data.get('message')}")
                    print(f"   - last_updated: {status_data.get('last_updated')}")
                    
                    # Test 4: Kiểm tra trạng thái
                    print("\n4. Kiểm tra trạng thái...")
                    if status_data.get('activated', True):
                        print("✅ Ứng dụng đang được kích hoạt")
                    else:
                        print("❌ Ứng dụng đã bị vô hiệu hóa")
                        print(f"   Lý do: {status_data.get('message', 'Không có thông tin')}")
                    
                    return True
                    
                except json.JSONDecodeError as e:
                    print(f"❌ JSON không hợp lệ: {e}")
                    return False
            else:
                print("❌ File activation_status.json không tồn tại")
                print("   Files có sẵn:", list(files.keys()))
                return False
        else:
            print(f"❌ Gist không tồn tại hoặc không accessible (Status: {response.status_code})")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ Timeout khi kiểm tra Gist")
        return False
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return False

def simulate_deactivation(gist_id, github_token):
    """Simulate việc deactivate ứng dụng"""
    
    print("\n=== Simulate Deactivation ===")
    
    # Tạo nội dung deactivated
    deactivated_content = {
        "activated": False,
        "expiry_date": "2025-12-31",
        "message": "Ứng dụng tạm thời bị vô hiệu hóa do bảo trì hệ thống",
        "last_updated": datetime.now().isoformat()
    }
    
    # Cập nhật Gist
    gist_data = {
        "files": {
            "activation_status.json": {
                "content": json.dumps(deactivated_content, ensure_ascii=False, indent=2)
            }
        }
    }
    
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    try:
        response = requests.patch(
            f"https://api.github.com/gists/{gist_id}",
            headers=headers,
            json=gist_data
        )
        
        if response.status_code == 200:
            print("✅ Đã deactivate ứng dụng thành công")
            return True
        else:
            print(f"❌ Lỗi deactivate: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return False

def simulate_activation(gist_id, github_token):
    """Simulate việc activate lại ứng dụng"""
    
    print("\n=== Simulate Activation ===")
    
    # Tạo nội dung activated
    activated_content = {
        "activated": True,
        "expiry_date": "2025-12-31",
        "message": "Ứng dụng đã được kích hoạt lại",
        "last_updated": datetime.now().isoformat()
    }
    
    # Cập nhật Gist
    gist_data = {
        "files": {
            "activation_status.json": {
                "content": json.dumps(activated_content, ensure_ascii=False, indent=2)
            }
        }
    }
    
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    try:
        response = requests.patch(
            f"https://api.github.com/gists/{gist_id}",
            headers=headers,
            json=gist_data
        )
        
        if response.status_code == 200:
            print("✅ Đã activate ứng dụng thành công")
            return True
        else:
            print(f"❌ Lỗi activate: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return False

if __name__ == "__main__":
    print("Test script cho hệ thống Remote Activation")
    print("=" * 50)
    
    # Nhập thông tin
    gist_id = input("Nhập Gist ID: ").strip()
    
    if not gist_id:
        print("❌ Gist ID không được để trống!")
        exit(1)
    
    # Test cơ bản
    if test_activation_status(gist_id):
        print("\n✅ Test cơ bản thành công!")
        
        # Hỏi có muốn test simulate không
        choice = input("\nBạn có muốn test simulate deactivation/activation không? (y/n): ").strip().lower()
        
        if choice == 'y':
            github_token = input("Nhập GitHub Personal Access Token: ").strip()
            
            if github_token:
                # Test deactivation
                if simulate_deactivation(gist_id, github_token):
                    print("\n--- Sau khi deactivate ---")
                    test_activation_status(gist_id)
                
                # Test activation
                if simulate_activation(gist_id, github_token):
                    print("\n--- Sau khi activate ---")
                    test_activation_status(gist_id)
            else:
                print("❌ Token không được để trống!")
    else:
        print("\n❌ Test cơ bản thất bại!")
        print("Vui lòng kiểm tra Gist ID và thử lại.") 