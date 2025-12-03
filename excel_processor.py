import pandas as pd
import re
import os
import sys
import warnings
from datetime import datetime
from abc import ABC, abstractmethod

# Tắt cảnh báo từ pandas
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)
warnings.filterwarnings('ignore', category=UserWarning)
pd.options.mode.chained_assignment = None


# ============================================================================
# SHARED UTILITIES
# ============================================================================

def list_excel_files():
    """Lấy danh sách các file Excel trong thư mục hiện tại và cho phép người dùng chọn"""
    excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        print('Không tìm thấy file Excel nào trong thư mục!')
        sys.exit(1)
    
    # Hiển thị danh sách file
    print('\nDanh sách file Excel:')
    for idx, file in enumerate(excel_files, 1):
        print(f'{idx}. {file}')
    
    # Cho phép người dùng chọn file
    while True:
        try:
            choice = int(input('\nVui lòng chọn số thứ tự file Excel (1-' + str(len(excel_files)) + '): '))
            if 1 <= choice <= len(excel_files):
                return excel_files[choice - 1]
            else:
                print('Số thứ tự không hợp lệ!')
        except ValueError:
            print('Vui lòng nhập một số!')


def format_date(date):
    """Định dạng lại cột ngày thành dd/mm/yyyy"""
    if pd.isna(date):
        return ''
    try:
        return pd.to_datetime(date).strftime('%d/%m/%Y')
    except:
        return ''


def format_quantity(qty):
    """Định dạng lại cột số lượng"""
    if pd.isna(qty):
        return ''
    if isinstance(qty, str):
        if qty == 'Yêu cầu':
            return qty
        try:
            # Chuyển đổi chuỗi số sang số thực
            qty = float(qty.replace(',', '.'))
        except:
            return qty
    try:
        # Chuyển đổi sang số thực và làm tròn đến 1 chữ số thập phân
        qty = float(qty)
        if qty.is_integer():
            # Nếu là số nguyên, thêm .0
            return f"{int(qty)}.0"
        else:
            # Làm tròn đến 1 chữ số thập phân
            return f"{qty:.1f}"
    except:
        return qty


# ============================================================================
# BASE PROCESSOR CLASS
# ============================================================================

class ExcelProcessor(ABC):
    """Base class cho các processor xử lý file Excel"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.df_ketqua = None
    
    def read_file(self):
        """Đọc file Excel"""
        try:
            print(f'\nĐang đọc file {self.file_path}...')
            self.df = pd.read_excel(self.file_path)
            return True
        except Exception as e:
            print('\n' + '='*40)
            print('✗ Lỗi khi đọc file!')
            print(f'✗ Chi tiết lỗi: {str(e)}')
            print('='*40)
            return False
    
    @abstractmethod
    def process(self):
        """Xử lý dữ liệu - phương thức trừu tượng phải được implement bởi subclass"""
        pass
    
    def export(self):
        """Xuất file kết quả"""
        try:
            # Tạo tên file với timestamp để tránh trùng lắp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'Ket_qua_xu_ly_{timestamp}.xlsx'
            self.df_ketqua.to_excel(output_file, index=False)
            print('\n' + '='*40)
            print(f'✓ Xuất file thành công!')
            print(f'✓ Tên file kết quả: {output_file}')
            print('='*40)
            return True
        except Exception as e:
            print('\n' + '='*40)
            print('✗ Lỗi khi xuất file!')
            print(f'✗ Chi tiết lỗi: {str(e)}')
            print('='*40)
            return False
    
    def run(self):
        """Chạy toàn bộ quy trình xử lý"""
        if not self.read_file():
            return False
        if not self.process():
            return False
        if not self.export():
            return False
        return True


# ============================================================================
# SCTX PROCESSOR
# ============================================================================

class SCTXProcessor(ExcelProcessor):
    """Processor cho file Excel loại SCTX"""
    
    def __init__(self, file_path):
        super().__init__(file_path)
        # Pattern mã phiếu: 02.O09.42.xxxx hoặc 03.O09.42.xxxx
        self.pattern = re.compile(r'^(02|03)\.O09\.42\.\d{4}$')
    
    def is_ma_phieu(self, s):
        """Kiểm tra xem chuỗi có phải là mã phiếu không"""
        if isinstance(s, str):
            return bool(self.pattern.match(s.strip()))
        return False
    
    def process(self):
        """Xử lý dữ liệu theo logic SCTX"""
        try:
            # Duyệt cột 2 để gán mã phiếu cho từng dòng
            ma_phieu_current = None
            ma_phieu_list = []
            
            for val in self.df.iloc[:, 1]:  # cột 2 trong DataFrame
                if self.is_ma_phieu(val):
                    ma_phieu_current = val.strip()
                ma_phieu_list.append(ma_phieu_current)
            
            # Thêm cột "Mã phiếu" vào DataFrame
            self.df['Mã phiếu'] = ma_phieu_list
            
            # Tạo DataFrame kết quả với các cột cần thiết
            self.df_ketqua = self.df[[
                'Mã phiếu',      # Mã phiếu
                'Ngày',          # Ngày
                'Diễn giải',     # Diễn giải
                'Mã vật tư',     # Mã vật tư
                'Tên vật tư',    # Tên vật tư
                'Đvt',           # Đơn vị tính
                'Số lượng'       # Số lượng
            ]]
            
            # Áp dụng định dạng ngày và số lượng
            self.df_ketqua['Ngày'] = self.df_ketqua['Ngày'].apply(format_date)
            self.df_ketqua['Số lượng'] = self.df_ketqua['Số lượng'].apply(format_quantity)
            
            return True
        except Exception as e:
            print('\n' + '='*40)
            print('✗ Lỗi khi xử lý dữ liệu!')
            print(f'✗ Chi tiết lỗi: {str(e)}')
            print('='*40)
            return False


# ============================================================================
# NTVTDD PROCESSOR
# ============================================================================

class NTVTDDProcessor(ExcelProcessor):
    """Processor cho file Excel loại NTVTDD"""
    
    def __init__(self, file_path):
        super().__init__(file_path)
        # Pattern mã phiếu: XX.YYY.ZZ.NNNN (linh hoạt hơn)
        self.pattern = re.compile(r'^\d{2}\.[A-Z0-9]{3}\.[0-9]{2,3}\.\d{4}$', re.IGNORECASE)
        # Pattern mã vật tư
        self.ma_vattu_pattern = re.compile(r'^\d+(?:\.[A-Z0-9]+){3,}$', re.IGNORECASE)
    
    def is_ma_phieu(self, s):
        """Kiểm tra xem chuỗi có phải là mã phiếu không"""
        if isinstance(s, str):
            return bool(self.pattern.match(s.strip()))
        return False
    
    def is_ma_vattu(self, value):
        """Kiểm tra xem giá trị có phải là mã vật tư không"""
        if isinstance(value, str):
            return bool(self.ma_vattu_pattern.match(value.strip()))
        return False
    
    def process(self):
        """Xử lý dữ liệu theo logic NTVTDD"""
        try:
            # Duyệt cột 2 để gán mã phiếu cho từng dòng
            ma_phieu_current = None
            ma_phieu_list = []
            
            for val in self.df.iloc[:, 1]:  # cột 2 trong DataFrame
                if self.is_ma_phieu(val):
                    ma_phieu_current = val.strip()
                ma_phieu_list.append(ma_phieu_current)
            
            # Thêm cột "Mã phiếu" vào DataFrame
            self.df['Mã phiếu'] = ma_phieu_list
            
            # Phát hiện các mã vật tư đang nằm trong cột Diễn giải và chuyển về đúng cột
            ma_vattu_mask = self.df['Diễn giải'].apply(self.is_ma_vattu)
            
            self.df.loc[ma_vattu_mask & self.df['Mã vật tư'].isna(), 'Mã vật tư'] = (
                self.df.loc[ma_vattu_mask, 'Diễn giải'].str.strip()
            )
            self.df.loc[ma_vattu_mask, 'Diễn giải'] = ''
            
            # Tạo DataFrame kết quả với các cột cần thiết
            self.df_ketqua = self.df[[
                'Mã phiếu',      # Mã phiếu
                'Ngày viết',     # Ngày
                'Diễn giải',     # Diễn giải
                'Mã vật tư',     # Mã vật tư
                'Tên vật tư',    # Tên vật tư
                'ĐVT',           # Đơn vị tính
                'Số lượng'       # Số lượng
            ]]
            
            # Áp dụng định dạng ngày và số lượng
            self.df_ketqua['Ngày viết'] = self.df_ketqua['Ngày viết'].apply(format_date)
            self.df_ketqua['Số lượng'] = self.df_ketqua['Số lượng'].apply(format_quantity)
            
            return True
        except Exception as e:
            print('\n' + '='*40)
            print('✗ Lỗi khi xử lý dữ liệu!')
            print(f'✗ Chi tiết lỗi: {str(e)}')
            print('='*40)
            return False


# ============================================================================
# MAIN MENU
# ============================================================================

def show_menu():
    """Hiển thị menu chính"""
    print('\n' + '='*50)
    print('CHƯƠNG TRÌNH XỬ LÝ DỮ LIỆU EXCEL')
    print('='*50)
    print('\nChọn loại file Excel cần xử lý:')
    print('1. File loại SCTX (Mã phiếu: 02.O09.42.xxxx hoặc 03.O09.42.xxxx)')
    print('2. File loại NTVTDD (Mã phiếu linh hoạt, có xử lý mã vật tư)')
    print('0. Thoát')
    print('='*50)


def main():
    """Hàm main chạy chương trình"""
    while True:
        show_menu()
        
        try:
            choice = input('\nNhập lựa chọn của bạn (0-2): ').strip()
            
            if choice == '0':
                print('\nCảm ơn bạn đã sử dụng chương trình!')
                sys.exit(0)
            
            elif choice == '1':
                print('\n--- XỬ LÝ FILE LOẠI SCTX ---')
                file_path = list_excel_files()
                processor = SCTXProcessor(file_path)
                processor.run()
                input('\nNhấn Enter để tiếp tục...')
            
            elif choice == '2':
                print('\n--- XỬ LÝ FILE LOẠI NTVTDD ---')
                file_path = list_excel_files()
                processor = NTVTDDProcessor(file_path)
                processor.run()
                input('\nNhấn Enter để tiếp tục...')
            
            else:
                print('\n✗ Lựa chọn không hợp lệ! Vui lòng chọn 0, 1 hoặc 2.')
                input('\nNhấn Enter để tiếp tục...')
        
        except KeyboardInterrupt:
            print('\n\nChương trình bị ngắt bởi người dùng.')
            sys.exit(0)
        except Exception as e:
            print(f'\n✗ Lỗi không xác định: {str(e)}')
            input('\nNhấn Enter để tiếp tục...')


if __name__ == '__main__':
    main()
