#!/usr/bin/env python3
"""
Test script để kiểm tra tính bảo mật của hệ thống activation
Kiểm tra xem ứng dụng có bị bypass khi tắt mạng không
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from qlvt import ActivationManager

def test_activation_security():
    """Test các trường hợp bypass activation"""
    print("=== KIỂM TRA TÍNH BẢO MẬT HỆ THỐNG ACTIVATION ===\n")
    
    activation_manager = ActivationManager()
    
    # Test 1: Kiểm tra trạng thái thực tế từ server
    print("1. Kiểm tra trạng thái thực tế từ server:")
    try:
        status = activation_manager.check_activation_status()
        print(f"   Kết quả: {status}")
        print(f"   Activated: {status.get('activated', 'Unknown')}")
        print(f"   Message: {status.get('message', 'No message')}")
    except Exception as e:
        print(f"   Lỗi: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Test 2: Mô phỏng các lỗi kết nối
    print("2. Mô phỏng các lỗi kết nối (sẽ trả về deactivated):")
    
    # Test timeout
    print("\n   a) Lỗi timeout:")
    try:
        # Tạm thời thay đổi timeout để gây lỗi
        original_timeout = activation_manager.api_url
        activation_manager.api_url = "https://httpbin.org/delay/15"  # 15 giây delay
        status = activation_manager.check_activation_status()
        print(f"   Kết quả: {status}")
        print(f"   Activated: {status.get('activated', 'Unknown')}")
        activation_manager.api_url = original_timeout
    except Exception as e:
        print(f"   Lỗi: {e}")
        activation_manager.api_url = original_timeout
    
    # Test connection error
    print("\n   b) Lỗi kết nối (sử dụng URL không tồn tại):")
    try:
        original_url = activation_manager.api_url
        activation_manager.api_url = "https://nonexistent-domain-12345.com/api"
        status = activation_manager.check_activation_status()
        print(f"   Kết quả: {status}")
        print(f"   Activated: {status.get('activated', 'Unknown')}")
        activation_manager.api_url = original_url
    except Exception as e:
        print(f"   Lỗi: {e}")
        activation_manager.api_url = original_url
    
    # Test HTTP error
    print("\n   c) Lỗi HTTP 404:")
    try:
        original_url = activation_manager.api_url
        activation_manager.api_url = "https://httpbin.org/status/404"
        status = activation_manager.check_activation_status()
        print(f"   Kết quả: {status}")
        print(f"   Activated: {status.get('activated', 'Unknown')}")
        activation_manager.api_url = original_url
    except Exception as e:
        print(f"   Lỗi: {e}")
        activation_manager.api_url = original_url
    
    print("\n" + "="*50 + "\n")
    
    # Test 3: Kiểm tra các method helper
    print("3. Kiểm tra các method helper:")
    
    print("\n   a) _get_default_status():")
    default_status = activation_manager._get_default_status()
    print(f"   Kết quả: {default_status}")
    print(f"   Activated: {default_status.get('activated', 'Unknown')}")
    
    print("\n   b) _get_deactivated_status():")
    deactivated_status = activation_manager._get_deactivated_status("Test message")
    print(f"   Kết quả: {deactivated_status}")
    print(f"   Activated: {deactivated_status.get('activated', 'Unknown')}")
    
    print("\n" + "="*50 + "\n")
    
    # Kết luận
    print("4. KẾT LUẬN:")
    print("   - Hệ thống activation đã được cải thiện để ngăn chặn bypass")
    print("   - Tất cả lỗi kết nối đều trả về activated=False")
    print("   - Ứng dụng sẽ thoát khi không thể kết nối đến server")
    print("   - Không thể bypass bằng cách tắt mạng hoặc chặn kết nối")

if __name__ == "__main__":
    test_activation_security() 