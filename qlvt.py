import sys
import pythoncom
import win32com.client
import win32print
import requests
import subprocess
import ctypes
import json
import os
import time
import webbrowser
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, 
    QWidget, QFileDialog, QListWidget, QCheckBox, QLabel, 
    QHBoxLayout, QMessageBox, QProgressBar, QListWidgetItem, 
    QInputDialog, QLineEdit, QDialog, QDialogButtonBox, QFormLayout,
    QComboBox, QScrollArea, QTextEdit, QTabWidget, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon
import os

# Import Excel processors
from excel_processor import SCTXProcessor, NTVTDDProcessor

REPLACEMENT_FILE = "replacements.txt"

def is_admin():
    """Kiểm tra xem ứng dụng có chạy với quyền admin không"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class Logger:
    def __init__(self):
        self.log_entries = []
        self.summary = {
            "processed": 0,
            "failed": 0,
            "total": 0
        }
    
    def clear(self):
        """Xóa log cho thao tác mới"""
        self.log_entries = []
        self.summary = {
            "processed": 0,
            "failed": 0,
            "total": 0
        }
    
    def log(self, message, status="INFO"):
        """Ghi log với timestamp"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        entry = f"[{timestamp}] [{status}] {message}"
        print(entry)
        self.log_entries.append(entry)
    
    def add_to_summary(self, processed=0, failed=0, total=0):
        """Cập nhật summary"""
        self.summary["processed"] += processed
        self.summary["failed"] += failed
        self.summary["total"] += total
    
    def get_summary(self):
        """Lấy thông tin tổng hợp"""
        return (f"✓ Đã xử lý: {self.summary['processed']}/{self.summary['total']} file\n"
                f"✗ Lỗi: {self.summary['failed']} file")

# Thêm class ActivationManager
class ActivationManager:
    def __init__(self):
        # Gist ID sẽ được tạo và cập nhật sau
        self.gist_id = "0a9de72209b228810b5feee5af13005e"  # Sẽ thay thế bằng Gist ID thực
        self.api_url = f"https://api.github.com/gists/{self.gist_id}"
    
    def check_activation_status(self):
        """Kiểm tra trạng thái activation từ GitHub Gist"""
        try:
            print(f"[ACTIVATION] Đang kiểm tra trạng thái activation...")
            response = requests.get(self.api_url, timeout=10)
            
            if response.status_code == 200:
                gist_data = response.json()
                files = gist_data.get('files', {})
                
                # Tìm file activation_status.json
                activation_file = None
                for filename, file_data in files.items():
                    if filename == 'activation_status.json':
                        activation_file = file_data
                        break
                
                if activation_file:
                    content = activation_file.get('content', '{}')
                    try:
                        status_data = json.loads(content)
                        print(f"[ACTIVATION] Trạng thái: {status_data}")
                        return status_data
                    except json.JSONDecodeError as e:
                        print(f"[ACTIVATION] Lỗi parse JSON: {e}")
                        return self._get_deactivated_status("Lỗi định dạng dữ liệu từ server")
                else:
                    print(f"[ACTIVATION] Không tìm thấy file activation_status.json")
                    return self._get_deactivated_status("Không tìm thấy thông tin kích hoạt trên server")
            else:
                print(f"[ACTIVATION] Lỗi API: {response.status_code}")
                return self._get_deactivated_status(f"Lỗi kết nối đến server (HTTP {response.status_code})")
                
        except requests.exceptions.Timeout:
            print(f"[ACTIVATION] Timeout khi kiểm tra activation")
            return self._get_deactivated_status("Không thể kết nối đến server (timeout)")
        except requests.exceptions.ConnectionError:
            print(f"[ACTIVATION] Lỗi kết nối khi kiểm tra activation")
            return self._get_deactivated_status("Không có kết nối mạng đến server")
        except Exception as e:
            print(f"[ACTIVATION] Lỗi kiểm tra activation: {e}")
            return self._get_deactivated_status(f"Lỗi không xác định: {str(e)}")
    
    def _get_default_status(self):
        """Trả về trạng thái mặc định (activated) - chỉ dùng khi server trả về activated=True"""
        return {
            "activated": True,
            "expiry_date": "2025-12-31",
            "message": "Ứng dụng đang hoạt động bình thường",
            "last_updated": "2024-01-15T10:30:00Z"
        }
    
    def _get_deactivated_status(self, message):
        """Trả về trạng thái deactivated cho các lỗi kết nối"""
        return {
            "activated": False,
            "expiry_date": None,
            "message": message,
            "last_updated": "2024-01-15T10:30:00Z"
        }

class CombinedWorker(QThread):
    finished = pyqtSignal(str)
    progress = pyqtSignal(int)
    
    def __init__(self, doc_names, replacements, parent=None):
        super().__init__(parent)
        self.doc_names = doc_names
        self.replacements = replacements

    def extract_ho_ten(self, text):
        """Trích xuất họ tên từ text, loại bỏ các thông tin khác"""
        try:
            # Loại bỏ các thông tin phía sau họ tên
            # Cắt đến dấu xuống dòng đầu tiên
            if '\r' in text:
                text = text.split('\r')[0].strip()
            elif '\n' in text:
                text = text.split('\n')[0].strip()
            
            # Loại bỏ các thông tin như "Đơn vị nhập:", "Đơn vị xuất:", v.v.
            # Tìm các từ khóa có thể xuất hiện sau họ tên
            keywords_to_remove = [
                "Đơn vị nhập:"
            ]
            
            for keyword in keywords_to_remove:
                if keyword in text:
                    text = text.split(keyword)[0].strip()
                    break
            
            # Loại bỏ các ký tự đặc biệt cuối
            text = text.rstrip('.,;:!?')
            
            return text if text else None
        except Exception as e:
            print(f"[DEBUG] Lỗi trích xuất họ tên: {e}")
            return None

    def find_ho_ten_nguoi_hang(self, doc):
        """Tìm họ tên người nhận/giao hàng trong document"""
        try:
            print(f"[DEBUG] Bắt đầu tìm họ tên người nhận/giao hàng...")
            # Tìm trong tất cả các bảng
            for table_idx, table in enumerate(doc.Tables):
                try:
                    # Sử dụng Range.Cells để tránh lỗi với merged cells
                    for cell_idx, cell in enumerate(table.Range.Cells):
                        cell_text = cell.Range.Text.strip()
                        
                        # Tìm "Họ và tên người nhận hàng:"
                        if "Họ và tên người nhận hàng:" in cell_text:
                            parts = cell_text.split("Họ và tên người nhận hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người nhận: '{ho_ten}'")
                                    return ho_ten
                        # Tìm "Họ và tên người giao hàng:"
                        elif "Họ và tên người giao hàng:" in cell_text:
                            parts = cell_text.split("Họ và tên người giao hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người giao: '{ho_ten}'")
                                    return ho_ten
                except Exception as e:
                    print(f"[DEBUG] Lỗi xử lý bảng {table_idx+1}: {e}")
                    # Fallback: thử cách khác nếu có lỗi
                    try:
                        table_range = table.Range
                        table_text = table_range.Text
                        
                        # Tìm trong toàn bộ text của bảng
                        if "Họ và tên người nhận hàng:" in table_text:
                            parts = table_text.split("Họ và tên người nhận hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    return ho_ten
                        elif "Họ và tên người giao hàng:" in table_text:
                            parts = table_text.split("Họ và tên người giao hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    return ho_ten
                    except Exception as e2:
                        print(f"[DEBUG] Fallback cũng thất bại cho bảng {table_idx+1}: {e2}")
            
            print(f"[DEBUG] Không tìm thấy họ tên người nhận/giao hàng trong bất kỳ bảng nào")
            return None
        except Exception as e:
            print(f"[DEBUG] Lỗi tìm họ tên: {e}")
            return None

    def modify_document(self, doc):
        """Xử lý khung tên: thêm dòng, điền tên"""
        try:
            # Xoá ký tự xuống dòng ở đầu tài liệu nếu có
            start_range = doc.Range(0, 1)
            if start_range.Text == '\r':
                start_range.Delete()

            # Lọc ra tất cả các bảng nằm ở trang đầu tiên (page 1)
            tables_on_first_page = [table for table in doc.Tables if table.Range.Information(3) == 1]
            if tables_on_first_page:
                # Chỉ lấy bảng CUỐI CÙNG ở trang đầu tiên (bảng ký tên)
                table = tables_on_first_page[-1]
                rows = table.Rows.Count
                if rows == 4:
                    # ⚠️ CHÈN 1 DÒNG vào giữa dòng 3 và 4
                    table.Rows.Add(BeforeRow=table.Rows(4))
                
                # ✅ Tiếp tục xử lý nội dung sau khi thêm dòng
                try:
                    # Tìm ô chứa "VÕ THANH ĐIỀN" ở hàng cuối cùng
                    last_row = table.Rows.Count
                    target_cell = None
                    for col in range(1, table.Columns.Count + 1):
                        cell_text = table.Cell(last_row, col).Range.Text.strip()
                        if "VÕ THANH ĐIỀN" in cell_text:
                            # Lưu lại ô bên phải để điền họ tên
                            if col < table.Columns.Count:
                                target_cell = table.Cell(last_row, col + 1)
                            break
                    
                    # # Tìm và xóa "PHAN CÔNG HUY" trong cùng hàng cuối
                    # for col in range(1, table.Columns.Count + 1):
                    #     cell_text = table.Cell(last_row, col).Range.Text.strip()
                    #     if "PHAN CÔNG HUY" in cell_text:
                    #         # Xóa nội dung "PHAN CÔNG HUY" khỏi ô
                    #         cell = table.Cell(last_row, col)
                    #         cell.Range.Text = ""
                    #         break
                    
                    # Tìm họ tên người nhận/giao hàng và điền vào ô bên phải của "VÕ THANH ĐIỀN"
                    if target_cell:
                        ho_ten = self.find_ho_ten_nguoi_hang(doc)
                        if ho_ten:
                            target_cell.Range.Text = ho_ten
                            print(f"[DEBUG] Đã điền họ tên: {ho_ten}")
                except:
                    pass
        except Exception as e:
            print(f"[DEBUG] Exception in modify_document: {e}")

    def replace_text(self, doc):
        """Thay thế văn bản trong trang đầu tiên"""
        try:
            # Lấy range của trang đầu tiên
            try:
                page2_start = doc.GoTo(What=1, Which=1, Count=2)
                first_page_end = page2_start.Start
            except:
                first_page_end = doc.Content.End
            
            # Thay thế text trong range của trang đầu tiên
            for old, new in self.replacements:
                # Thay thế bằng vòng lặp
                count = 0
                max_iterations = 1000
                
                while count < max_iterations:
                    search_range = doc.Range(0, first_page_end)
                    search_range.Find.ClearFormatting()
                    search_range.Find.Text = old
                    search_range.Find.Forward = True
                    search_range.Find.Wrap = 0  # wdFindStop
                    search_range.Find.MatchCase = False
                    search_range.Find.MatchWholeWord = False
                    
                    if search_range.Find.Execute():
                        search_range.Text = new
                        count += 1
                    else:
                        break
        except Exception as e:
            print(f"[DEBUG] Exception in replace_text: {e}")

    def run(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            total_files = len(self.doc_names)
            processed_count = 0
            
            for i in range(word_app.Documents.Count):
                doc = word_app.Documents.Item(i + 1)
                if doc.Name in self.doc_names:
                    try:
                        print(f"[DEBUG] ===== Đang xử lý tài liệu: {doc.Name} =====")
                        
                        # 1. Xử lý khung tên (Process Title Block)
                        self.modify_document(doc)
                        
                        # 2. Thay thế văn bản (Replace Name)
                        if self.replacements:
                            self.replace_text(doc)
                        
                        processed_count += 1
                        self.progress.emit(processed_count)
                        
                    except Exception as e:
                        print(f"[DEBUG] Lỗi xử lý file {doc.Name}: {e}")
                        import traceback
                        traceback.print_exc()
            
            self.finished.emit(f"✅ Đã xử lý xong {processed_count}/{total_files} tài liệu.")
        except Exception as e:
            self.finished.emit(f"Lỗi xử lý: {e}")
        finally:
            pythoncom.CoUninitialize()


class WordProcessorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.current_version = "1.0.21"
        
        # Khởi tạo progress bar
        self.progress_bar = None

        self.setWindowTitle(f"Công cụ xử lý và lưu trữ phiếu nhập xuất kho {self.current_version} | www.khoatran.io.vn")
        self.setGeometry(200, 200, 600, 400)  # Tăng kích thước cửa sổ mặc định
        
        # Thiết lập icon cho ứng dụng
        icon = QIcon("icon.ico")
        self.setWindowIcon(icon)
        
        # Thiết lập icon cho taskbar (Windows)
        if hasattr(self, 'setWindowIcon'):
            # Đảm bảo icon hiển thị trên taskbar
            self.setWindowIcon(icon)
            
        # Thiết lập thuộc tính cửa sổ để hiển thị icon tốt hơn
        self.setWindowFlags(self.windowFlags() | Qt.Window)

        # Khởi tạo ActivationManager
        self.activation_manager = ActivationManager()
        
        # Kiểm tra activation trước khi khởi tạo UI
        if not self._check_activation():
            return  # Thoát nếu không được kích hoạt

        # Khởi tạo AutoUpdater
        self.updater = AutoUpdater("nekennick/RunPhieu")
        
        # Auto-check updates sau 3 giây
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.auto_check_updates)
        self.update_timer.start(3000)  # 3 giây

        self.layout = QVBoxLayout()

        self.status_label = QLabel("Danh sách phiếu đang mở:")
        self.layout.addWidget(self.status_label)

        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.toggle_item_check_state)
        self.layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()
        self.refresh_button = QPushButton("Load DS phiếu")
        self.refresh_button.clicked.connect(self.load_open_documents)
        button_layout.addWidget(self.refresh_button)

        # Nút Xử lý (Gộp tính năng Xử lý khung tên và Thay tên)
        self.combined_button = QPushButton("Xử lý khung tên")
        self.combined_button.clicked.connect(self.process_and_replace)
        button_layout.addWidget(self.combined_button)

        # Thêm nút In trang đầu
        self.print_button = QPushButton("In phiếu đã chọn")
        self.print_button.clicked.connect(self.print_first_pages)
        button_layout.addWidget(self.print_button)
        
        # Thêm dòng hiển thị thông tin máy in
        printer_info_layout = QHBoxLayout()
        printer_info_layout.addStretch()
        
        # Label hiển thị tên máy in
        self.printer_label = QLabel()
        self.printer_label.setStyleSheet("color: gray;")
        self.update_printer_info()
        
        # Nút chọn máy in
        select_printer_btn = QPushButton("🖨️")
        select_printer_btn.setToolTip("Chọn máy in")
        select_printer_btn.setFixedWidth(30)
        select_printer_btn.setStyleSheet("QPushButton { font-size: 14px; }")
        select_printer_btn.clicked.connect(self.select_printer)
        
        printer_info_layout.addWidget(QLabel("Máy in:"))
        printer_info_layout.addWidget(self.printer_label)
        printer_info_layout.addWidget(select_printer_btn)
        
        # Thêm dòng thông tin máy in vào layout chính
        self.layout.addLayout(printer_info_layout)

        self.save_as_button = QPushButton("Lưu tất cả file")
        self.save_as_button.clicked.connect(self.save_all_files_as)
        button_layout.addWidget(self.save_as_button)

        # Thêm nút đóng toàn bộ phiếu
        self.close_all_button = QPushButton("Đóng tất cả phiếu")
        self.close_all_button.clicked.connect(self.close_all_documents)
        button_layout.addWidget(self.close_all_button)

        self.layout.addLayout(button_layout)
        self.setLayout(self.layout)

        # Biến trạng thái để xử lý lần tải đầu tiên
        self.is_initial_load = True

        # 🔄 GỌI NGAY khi khởi động để tự động tải danh sách tài liệu đang mở
        self.load_open_documents()

        # Sau lần tải đầu tiên, các lần nhấn nút sau sẽ bỏ chọn
        self.is_initial_load = False

        # Trạng thái để bật/tắt chọn tất cả, bắt đầu bằng bỏ chọn (vì lần đầu đã chọn)
        self.select_all_enabled = False
    
    def setup_progress_bar(self):
        """Tạo và cấu hình progress bar"""
        if not self.progress_bar:
            self.progress_bar = QProgressBar()
            self.layout.insertWidget(self.layout.count() - 1, self.progress_bar)
    
    def cleanup_progress_bar(self):
        """Xóa progress bar"""
        if self.progress_bar:
            self.progress_bar.deleteLater()
            self.progress_bar = None
            
    def update_progress(self, value):
        """Cập nhật giá trị progress bar"""
        if self.progress_bar:
            self.progress_bar.setValue(value)

    def _check_activation(self):
        """Kiểm tra trạng thái activation khi khởi động"""
        # Luôn trả về True để bỏ qua kiểm tra kết nối mạng
        return True

    def select_printer(self):
        """Hiển thị hộp thoại chọn máy in"""
        try:
            # Lấy danh sách tất cả các máy in đã cài đặt
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            
            if not printers:
                QMessageBox.warning(self, "Cảnh báo", "Không tìm thấy máy in nào!")
                return
            
            # Lấy tên máy in hiện tại
            current_printer = win32print.GetDefaultPrinter()
            
            # Tìm chỉ số của máy in hiện tại trong danh sách
            current_index = 0
            if current_printer in printers:
                current_index = printers.index(current_printer)
                
            # Tạo hộp thoại chọn máy in
            printer, ok = QInputDialog.getItem(
                self, 
                "Chọn máy in", 
                "Chọn máy in mặc định:", 
                printers, 
                current=current_index,  # Chọn máy in hiện tại làm mặc định
                editable=False
            )
            
            if ok and printer:
                # Chỉ cập nhật nếu chọn máy in khác
                if printer != current_printer:
                    # Đặt máy in đã chọn làm mặc định
                    win32print.SetDefaultPrinter(printer)
                    # Cập nhật thông tin hiển thị
                    self.update_printer_info()
                    QMessageBox.information(self, "Thành công", f"Đã chọn máy in: {printer}")
                
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể chọn máy in: {str(e)}")
    
    def update_printer_info(self):
        """Cập nhật thông tin máy in mặc định"""
        try:
            # Lấy tên máy in mặc định
            default_printer = win32print.GetDefaultPrinter()
            
            # Lấy thông tin chi tiết về máy in
            printer_info = win32print.GetPrinter(win32print.OpenPrinter(default_printer), 2)
            printer_status = printer_info.get('Status', 0)
            
            # Xác định trạng thái máy in
            status_text = ""
            if printer_status == 0:
                status_text = "(Sẵn sàng)"
            else:
                status_text = "(Đang bận)"
                
            # Cập nhật giao diện
            self.printer_label.setText(f"{default_printer} {status_text}")
            
            # Đổi màu dựa trên trạng thái
            if printer_status == 0:
                self.printer_label.setStyleSheet("color: green;")
            else:
                self.printer_label.setStyleSheet("color: orange;")
                
        except Exception as e:
            self.printer_label.setText("Không thể lấy thông tin máy in")
            self.printer_label.setStyleSheet("color: red;")
            print(f"Lỗi khi lấy thông tin máy in: {e}")
    
    def show_activation_status(self):
        """Hiển thị thông tin trạng thái activation"""
        # Hiển thị thông báo đơn giản, không kiểm tra kết nối mạng
        QMessageBox.information(
            self,
            "Trạng thái",
            "✅ Ứng dụng đã sẵn sàng sử dụng"
        )

    def load_open_documents(self):
        self.file_list.clear()

        # Quyết định trạng thái check
        if self.is_initial_load:
            check_state = Qt.Checked
        else:
            check_state = Qt.Checked if self.select_all_enabled else Qt.Unchecked
            self.select_all_enabled = not self.select_all_enabled

        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            docs = word_app.Documents
            for i in range(docs.Count):
                doc = docs.Item(i + 1)
                item_text = doc.Name
                item = QListWidgetItem(item_text)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(check_state)
                self.file_list.addItem(item)
        except pythoncom.com_error as e:
            # Lỗi -2147221021 (MK_E_UNAVAILABLE) có nghĩa là Word chưa được mở
            if e.hresult == -2147221021:
                self.status_label.setText("Chưa tìm thấy file word nào đang mở")
            else:
                self.status_label.setText(f"Lỗi COM: {e}")
        except Exception as e:
            self.status_label.setText(f"Lỗi: {e}")
        finally:
            pythoncom.CoUninitialize()

    def toggle_item_check_state(self, item):
        """Đảo ngược trạng thái check của item khi được click"""
        if item.checkState() == Qt.Checked:
            item.setCheckState(Qt.Unchecked)
        else:
            item.setCheckState(Qt.Checked)

    def process_and_replace(self):
        """Xử lý gộp: Thay thế văn bản -> Xử lý khung tên"""
        # 1. Hiển thị dialog thay thế trước
        dialog = ReplaceDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            replacements = dialog.get_replacement_pairs()
            
            # 2. Lấy danh sách file được chọn
            selected_files = []
            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                if item.checkState() == Qt.Checked:
                    selected_files.append(item.text())
            
            if not selected_files:
                self.status_label.setText("⚠️ Bạn chưa chọn tài liệu nào để xử lý.")
                return
            
            # 3. Khởi chạy worker gộp
            self.setup_progress_bar()
            self.progress_bar.setMaximum(len(selected_files))
            self.status_label.setText("⏳ Đang xử lý và thay thế, vui lòng chờ...")
            
            self.combined_thread = CombinedWorker(selected_files, replacements)
            self.combined_thread.progress.connect(self.update_progress)
            self.combined_thread.finished.connect(self.on_combined_finished)
            self.combined_thread.start()

    def on_combined_finished(self, message):
        self.status_label.setText(message)
        self.cleanup_progress_bar()

    def save_all_files_as(self):
        # Chọn thư mục đích
        folder_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu file")
        if not folder_path:
            return

        selected_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_files.append(item.text())

        if not selected_files:
            self.status_label.setText("⚠️ Bạn chưa chọn tài liệu nào để lưu.")
            return

        self.status_label.setText("⏳ Đang lưu file, vui lòng chờ...")
        self.save_thread = SaveAsWorker(selected_files, folder_path)
        self.save_thread.finished.connect(self.on_save_finished)
        self.save_thread.start()

    def on_save_finished(self, message):
        self.status_label.setText(message)

    def print_first_pages(self):
        selected_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_files.append(item.text())

        if not selected_files:
            self.status_label.setText("⚠️ Bạn chưa chọn tài liệu nào để in.")
            return

        # Kiểm tra xem có giữ phím Shift không (để lưu PDF)
        modifiers = QApplication.keyboardModifiers()
        if modifiers == Qt.ShiftModifier:
            # Giữ Shift = Lưu PDF
            output_folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu file PDF")
            if not output_folder:
                return
            action_mode = "save_pdf"
        else:
            # Mặc định = In trực tiếp
            output_folder = None
            action_mode = "print"

        self.setup_progress_bar()
        if action_mode == "save_pdf":
            self.status_label.setText("⏳ Đang lưu PDF trang đầu, vui lòng chờ...")
        else:
            self.status_label.setText("⏳ Đang in trang đầu, vui lòng chờ...")
        print(f"[DEBUG] Bắt đầu xử lý {len(selected_files)} tài liệu - Mode: {action_mode}")
        
        # Khởi tạo và chạy worker
        self.print_thread = PrintWorker(selected_files, output_folder=output_folder, action_mode=action_mode)
        self.print_thread.progress.connect(self.update_progress)
        self.print_thread.finished.connect(self.on_print_finished)
        self.print_thread.start()

    def on_print_finished(self, message):
        self.status_label.setText(message)
        self.cleanup_progress_bar()

    def find_ho_ten_nguoi_hang(self, doc):
        """Tìm họ tên người nhận/giao hàng trong document"""
        try:
            print(f"[DEBUG] Bắt đầu tìm họ tên người nhận/giao hàng...")
            # Tìm trong tất cả các bảng
            for table_idx, table in enumerate(doc.Tables):
                print(f"[DEBUG] Kiểm tra bảng {table_idx + 1}")
                try:
                    # Sử dụng Range.Cells để tránh lỗi với merged cells
                    for cell_idx, cell in enumerate(table.Range.Cells):
                        cell_text = cell.Range.Text.strip()
                        if cell_text:  # Chỉ in cell có nội dung
                            print(f"[DEBUG] Bảng{table_idx+1} - Cell {cell_idx+1}: '{cell_text}'")
                        
                        # Tìm "Họ và tên người nhận hàng:"
                        if "Họ và tên người nhận hàng:" in cell_text:
                            print(f"[DEBUG] Tìm thấy 'Họ và tên người nhận hàng:' trong cell {cell_idx+1}")
                            # Trích xuất họ tên sau dấu ":"
                            parts = cell_text.split("Họ và tên người nhận hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                # Cắt họ tên đến dấu xuống dòng hoặc ký tự đặc biệt
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người nhận: '{ho_ten}'")
                                    return ho_ten
                                else:
                                    print(f"[DEBUG] Họ tên người nhận trống")
                            else:
                                print(f"[DEBUG] Không thể trích xuất họ tên người nhận")
                        # Tìm "Họ và tên người giao hàng:"
                        elif "Họ và tên người giao hàng:" in cell_text:
                            print(f"[DEBUG] Tìm thấy 'Họ và tên người giao hàng:' trong cell {cell_idx+1}")
                            # Trích xuất họ tên sau dấu ":"
                            parts = cell_text.split("Họ và tên người giao hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                # Cắt họ tên đến dấu xuống dòng hoặc ký tự đặc biệt
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người giao: '{ho_ten}'")
                                    return ho_ten
                                else:
                                    print(f"[DEBUG] Họ tên người giao trống")
                            else:
                                print(f"[DEBUG] Không thể trích xuất họ tên người giao")
                except Exception as e:
                    print(f"[DEBUG] Lỗi xử lý bảng {table_idx+1}: {e}")
                    # Fallback: thử cách khác nếu có lỗi
                    try:
                        table_range = table.Range
                        table_text = table_range.Text
                        print(f"[DEBUG] Bảng{table_idx+1} - Toàn bộ nội dung: '{table_text}'")
                        
                        # Tìm trong toàn bộ text của bảng
                        if "Họ và tên người nhận hàng:" in table_text:
                            print(f"[DEBUG] Tìm thấy 'Họ và tên người nhận hàng:' trong bảng {table_idx+1}")
                            parts = table_text.split("Họ và tên người nhận hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người nhận: '{ho_ten}'")
                                    return ho_ten
                        elif "Họ và tên người giao hàng:" in table_text:
                            print(f"[DEBUG] Tìm thấy 'Họ và tên người giao hàng:' trong bảng {table_idx+1}")
                            parts = table_text.split("Họ và tên người giao hàng:")
                            if len(parts) > 1:
                                ho_ten_part = parts[1].strip()
                                ho_ten = self.extract_ho_ten(ho_ten_part)
                                if ho_ten:
                                    print(f"[DEBUG] Trích xuất được họ tên người giao: '{ho_ten}'")
                                    return ho_ten
                    except Exception as e2:
                        print(f"[DEBUG] Fallback cũng thất bại cho bảng {table_idx+1}: {e2}")
            
            print(f"[DEBUG] Không tìm thấy trong bảng, kiểm tra paragraphs...")
            # Tìm trong paragraphs nếu không tìm thấy trong bảng
            for para_idx, para in enumerate(doc.Paragraphs):
                para_text = para.Range.Text.strip()
                if para_text:  # Chỉ in paragraph có nội dung
                    print(f"[DEBUG] Paragraph {para_idx + 1}: '{para_text}'")
                
                if "Họ và tên người nhận hàng:" in para_text:
                    print(f"[DEBUG] Tìm thấy 'Họ và tên người nhận hàng:' trong paragraph {para_idx + 1}")
                    parts = para_text.split("Họ và tên người nhận hàng:")
                    if len(parts) > 1:
                        ho_ten_part = parts[1].strip()
                        ho_ten = self.extract_ho_ten(ho_ten_part)
                        if ho_ten:
                            print(f"[DEBUG] Trích xuất được họ tên người nhận từ paragraph: '{ho_ten}'")
                            return ho_ten
                elif "Họ và tên người giao hàng:" in para_text:
                    print(f"[DEBUG] Tìm thấy 'Họ và tên người giao hàng:' trong paragraph {para_idx + 1}")
                    parts = para_text.split("Họ và tên người giao hàng:")
                    if len(parts) > 1:
                        ho_ten_part = parts[1].strip()
                        ho_ten = self.extract_ho_ten(ho_ten_part)
                        if ho_ten:
                            print(f"[DEBUG] Trích xuất được họ tên người giao từ paragraph: '{ho_ten}'")
                            return ho_ten
            
            print(f"[DEBUG] Không tìm thấy họ tên người nhận/giao hàng trong toàn bộ document")
            return None
        except Exception as e:
            print(f"[DEBUG] Lỗi tìm họ tên: {e}")
            return None

    def extract_ho_ten(self, text):
        """Trích xuất họ tên từ text, loại bỏ các thông tin khác"""
        try:
            # Loại bỏ các thông tin phía sau họ tên
            # Cắt đến dấu xuống dòng đầu tiên
            if '\r' in text:
                text = text.split('\r')[0].strip()
            elif '\n' in text:
                text = text.split('\n')[0].strip()
            
            # Loại bỏ các thông tin như "Đơn vị nhập:", "Đơn vị xuất:", v.v.
            # Tìm các từ khóa có thể xuất hiện sau họ tên
            keywords_to_remove = [
                "Đơn vị nhập:"
            ]
            
            for keyword in keywords_to_remove:
                if keyword in text:
                    text = text.split(keyword)[0].strip()
                    break
            
            # Loại bỏ các ký tự đặc biệt cuối
            text = text.rstrip('.,;:!?')
            
            return text if text else None
        except Exception as e:
            print(f"[DEBUG] Lỗi trích xuất họ tên: {e}")
            return None

    def auto_check_updates(self):
        """Tự động kiểm tra cập nhật khi khởi động"""
        self.update_timer.stop()  # Chỉ check 1 lần
        try:
            has_update, release_info = self.updater.check_for_updates(self.current_version)
            if has_update:
                self.show_update_dialog(release_info)
        except Exception as e:
            print(f"[UPDATE] Lỗi auto-check: {e}")
    
    def show_update_dialog(self, release_info):
        """Hiển thị dialog xác nhận cập nhật - bắt buộc phải cập nhật"""
        latest_version = release_info['tag_name'].lstrip('v')
        
        # Sử dụng QDialog để có thể xử lý sự kiện đóng
        dialog = QDialog(self)
        dialog.setWindowTitle("⚠️ Cập nhật bắt buộc")
        dialog.setModal(True)
        dialog.setFixedSize(400, 200)
        
        # Layout
        layout = QVBoxLayout()
        
        # Icon và tiêu đề
        title_label = QLabel(f"⚠️ Có phiên bản mới: v{latest_version}")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #d32f2f;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # Nội dung
        content_label = QLabel("Phiên bản hiện tại đã không còn khả dụng.\n\nBạn PHẢI cập nhật để tiếp tục sử dụng ứng dụng.\n\nNhấn 'Cập nhật ngay' để mở trang tải về.")
        content_label.setAlignment(Qt.AlignCenter)
        content_label.setWordWrap(True)
        layout.addWidget(content_label)
        
        # Nút cập nhật
        update_button = QPushButton("Cập nhật ngay")
        update_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        update_button.clicked.connect(lambda: self._handle_update_click(dialog, release_info))
        layout.addWidget(update_button)
        
        dialog.setLayout(layout)
        
        # Xử lý sự kiện đóng dialog (nhấn nút X)
        dialog.closeEvent = lambda event: self._handle_dialog_close(event, release_info)
        
        # Hiển thị dialog
        dialog.exec_()
    
    def _handle_update_click(self, dialog, release_info):
        """Xử lý khi người dùng nhấn nút cập nhật"""
        dialog.accept()
        self.perform_update(release_info)
    
    def _handle_dialog_close(self, event, release_info):
        """Xử lý khi người dùng đóng dialog (nhấn nút X)"""
        # Ngay cả khi đóng dialog cũng phải cập nhật
        self.perform_update(release_info)
        event.accept()

    def perform_update(self, release_info):
        """Thực hiện cập nhật - hướng dẫn người dùng đến trang tải về và đóng ứng dụng"""
        try:
            if release_info:
                # Tạo URL trực tiếp đến release mới nhất
                latest_version = release_info['tag_name']
                release_url = f"https://khoatran.io.vn/#QLVT"
                
                # Mở trực tiếp trình duyệt với URL release cụ thể
                webbrowser.open(release_url)
                
                # Hiển thị thông báo cuối cùng và đóng ứng dụng
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setWindowTitle("Cập nhật bắt buộc")
                msg.setText("Trình duyệt đã được mở!")
                msg.setInformativeText(f"Vui lòng tải phiên bản mới v{latest_version} và cài đặt.\n\nỨng dụng sẽ đóng lại sau khi bạn nhấn OK.")
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
                
                # Đóng ứng dụng
                QApplication.quit()
            else:
                QMessageBox.information(self, "Thông báo", "Không có phiên bản mới để cập nhật.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi khi cập nhật: {str(e)}")
            # Ngay cả khi có lỗi cũng đóng ứng dụng
            QApplication.quit()

    def close_all_documents(self):
        """Đóng tất cả các tài liệu Word đang mở"""
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            doc_count = word_app.Documents.Count
            
            if doc_count > 0:
                from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
                
                dialog = QDialog(self)
                dialog.setWindowTitle("Xác nhận đóng tất cả phiếu")
                dialog.setModal(True)
                dialog.setFixedSize(400, 180)
                
                layout = QVBoxLayout()
                
                message_label = QLabel(
                    f"Hiện có {doc_count} phiếu trong danh sách.\n\n"
                    f"Bạn đã in các phiếu này chưa?\n"
                    f"Bạn có chắc chắn muốn đóng tất cả?"
                )
                message_label.setWordWrap(True)
                message_label.setStyleSheet("font-size: 11pt; padding: 10px;")
                layout.addWidget(message_label)
                
                button_layout = QHBoxLayout()
                
                yes_btn = QPushButton("Đã in, đóng tất cả")
                yes_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #4CAF50;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        border-radius: 4px;
                        font-weight: bold;
                        font-size: 10pt;
                    }
                    QPushButton:hover {
                        background-color: #45a049;
                    }
                """)
                yes_btn.clicked.connect(dialog.accept)
                
                no_btn = QPushButton("Hủy")
                no_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #9E9E9E;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        border-radius: 4px;
                        font-weight: bold;
                        font-size: 10pt;
                    }
                    QPushButton:hover {
                        background-color: #757575;
                    }
                """)
                no_btn.clicked.connect(dialog.reject)
                
                button_layout.addWidget(yes_btn)
                button_layout.addWidget(no_btn)
                layout.addLayout(button_layout)
                
                dialog.setLayout(layout)
                
                result = dialog.exec_()
                
                if result != QDialog.Accepted:
                    self.status_label.setText("⚠️ Đã hủy đóng phiếu.")
                    return
                
                while word_app.Documents.Count > 0:
                    doc = word_app.Documents.Item(1)
                    doc_name = doc.Name
                    doc.Close(SaveChanges=False)
                    print(f"[DEBUG] Đã đóng tài liệu: {doc_name}")
                
                word_app.Quit()
                print("[DEBUG] Đã thoát ứng dụng Word.")
                self.status_label.setText(f"✅ Đã đóng {doc_count} phiếu và thoát Word.")
            else:
                self.status_label.setText("⚠️ Không có tài liệu Word nào đang mở để đóng.")
        except Exception as e:
            self.status_label.setText(f"Lỗi đóng tài liệu: {e}")


class ReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Xử lý khung tên")
        self.setModal(True)
        self.resize(500, 400)
        
        # Danh sách các cặp từ thay thế
        self.replacement_pairs = []
        
        # Layout chính
        layout = QVBoxLayout()
        
        # Tiêu đề
        title_label = QLabel("Nhập các cặp từ cần thay thế:")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        # Scroll area cho danh sách các cặp từ
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMaximumHeight(250)
        
        # Widget chứa danh sách
        self.pairs_widget = QWidget()
        self.pairs_layout = QVBoxLayout(self.pairs_widget)
        self.pairs_layout.setSpacing(5)
        
        scroll_area.setWidget(self.pairs_widget)
        layout.addWidget(scroll_area)
        
        # Nút thêm cặp từ mới
        add_button = QPushButton("➕ Thêm cặp từ mới")
        add_button.clicked.connect(self.add_pair)
        add_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(add_button)
        
        # Nút xóa tất cả
        clear_button = QPushButton("🗑️ Xóa tất cả")
        clear_button.clicked.connect(self.clear_all_pairs)
        clear_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        layout.addWidget(clear_button)
        
        # Nút OK và Cancel
        button_layout = QHBoxLayout()
        
        ok_button = QPushButton("Bắt đầu xử lý")
        ok_button.clicked.connect(self.accept)
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        
        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #9E9E9E;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #757575;
            }
        """)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # Tải các cặp từ đã lưu
        self.load_pairs_from_file()
        
        # Thêm ít nhất 1 cặp từ mặc định nếu chưa có
        if not self.replacement_pairs:
            self.add_pair()
    
    def add_pair(self):
        """Thêm một cặp từ thay thế mới"""
        pair_widget = self.create_pair_widget()
        self.pairs_layout.addWidget(pair_widget)
        self.replacement_pairs.append(pair_widget)
    
    def create_pair_widget(self):
        """Tạo widget cho một cặp từ thay thế"""
        pair_widget = QWidget()
        pair_layout = QHBoxLayout(pair_widget)
        pair_layout.setContentsMargins(5, 5, 5, 5)
        
        # Label số thứ tự
        index_label = QLabel(f"{len(self.replacement_pairs) + 1}.")
        index_label.setMinimumWidth(30)
        index_label.setStyleSheet("font-weight: bold; color: #666;")
        pair_layout.addWidget(index_label)
        
        # Ô nhập từ cũ
        old_edit = QLineEdit()
        old_edit.setPlaceholderText("Từ cần thay thế...")
        old_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #ddd;
                border-radius: 4px;
                background-color: #fff8dc;
            }
            QLineEdit:focus {
                border-color: #ff9800;
            }
        """)
        pair_layout.addWidget(old_edit)
        
        # Mũi tên
        arrow_label = QLabel("→")
        arrow_label.setStyleSheet("font-weight: bold; font-size: 16px; color: #666; margin: 0 10px;")
        pair_layout.addWidget(arrow_label)
        
        # Ô nhập từ mới
        new_edit = QLineEdit()
        new_edit.setPlaceholderText("Từ thay thế...")
        new_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #ddd;
                border-radius: 4px;
                background-color: #f0fff0;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
            }
        """)
        pair_layout.addWidget(new_edit)
        
        # Nút xóa
        delete_button = QPushButton("❌")
        delete_button.setMaximumWidth(30)
        delete_button.clicked.connect(lambda: self.remove_pair(pair_widget))
        delete_button.setStyleSheet("""
            QPushButton {
                background-color: #ff4444;
                color: white;
                border: none;
                border-radius: 15px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #cc0000;
            }
        """)
        pair_layout.addWidget(delete_button)
        
        return pair_widget
    
    def remove_pair(self, pair_widget):
        """Xóa một cặp từ thay thế"""
        if len(self.replacement_pairs) > 1:  # Giữ lại ít nhất 1 cặp
            self.pairs_layout.removeWidget(pair_widget)
            self.replacement_pairs.remove(pair_widget)
            pair_widget.deleteLater()
            self.update_index_labels()
        else:
            QMessageBox.information(self, "Thông báo", "Phải có ít nhất 1 cặp từ thay thế!")
    
    def clear_all_pairs(self):
        """Xóa tất cả các cặp từ thay thế"""
        reply = QMessageBox.question(self, "Xác nhận", 
                                   "Bạn có chắc muốn xóa tất cả các cặp từ thay thế?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # Xóa tất cả widget
            for pair_widget in self.replacement_pairs:
                self.pairs_layout.removeWidget(pair_widget)
                pair_widget.deleteLater()
            
            self.replacement_pairs.clear()
            
            # Thêm lại 1 cặp mặc định
            self.add_pair()
    
    def update_index_labels(self):
        """Cập nhật số thứ tự cho các cặp từ"""
        for i, pair_widget in enumerate(self.replacement_pairs):
            index_label = pair_widget.layout().itemAt(0).widget()
            index_label.setText(f"{i + 1}.")
    
    def get_replacement_pairs(self):
        """Lấy danh sách các cặp từ thay thế"""
        pairs = []
        for pair_widget in self.replacement_pairs:
            old_edit = pair_widget.layout().itemAt(1).widget()
            new_edit = pair_widget.layout().itemAt(3).widget()
            
            old_text = old_edit.text().strip()
            new_text = new_edit.text().strip()
            
            if old_text:
                pairs.append((old_text, new_text))
        
        return pairs
    
    def load_pairs_from_file(self):
        """Tải các cặp từ từ file"""
        try:
            if os.path.exists(REPLACEMENT_FILE):
                with open(REPLACEMENT_FILE, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                # Xóa các cặp hiện tại
                for pair_widget in self.replacement_pairs:
                    self.pairs_layout.removeWidget(pair_widget)
                    pair_widget.deleteLater()
                self.replacement_pairs.clear()
                
                # Thêm các cặp từ file
                for line in lines:
                    line = line.strip()
                    if '|' in line:
                        old_text, new_text = line.split('|', 1)
                        pair_widget = self.create_pair_widget()
                        self.pairs_layout.addWidget(pair_widget)
                        self.replacement_pairs.append(pair_widget)
                        
                        # Điền dữ liệu
                        old_edit = pair_widget.layout().itemAt(1).widget()
                        new_edit = pair_widget.layout().itemAt(3).widget()
                        old_edit.setText(old_text.strip())
                        new_edit.setText(new_text.strip())
                
                self.update_index_labels()
        except Exception as e:
            print(f"Lỗi tải file replacements: {e}")
    
    def save_pairs_to_file(self):
        """Lưu các cặp từ vào file"""
        try:
            pairs = self.get_replacement_pairs()
            with open(REPLACEMENT_FILE, 'w', encoding='utf-8') as f:
                for old_text, new_text in pairs:
                    f.write(f"{old_text}|{new_text}\n")
        except Exception as e:
            print(f"Lỗi lưu file replacements: {e}")
    
    def accept(self):
        """Khi nhấn OK"""
        pairs = self.get_replacement_pairs()
        if not pairs:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng nhập ít nhất 1 cặp từ thay thế!")
            return
        
        # Lưu vào file
        self.save_pairs_to_file()
        super().accept()

    

   


class SaveAsWorker(QThread):
    finished = pyqtSignal(str)
    def __init__(self, doc_names, folder_path, parent=None):
        super().__init__(parent)
        self.doc_names = doc_names
        self.folder_path = folder_path

    def find_so_phieu(self, doc):
        """Tìm số phiếu trong document"""
        import re
        try:
            # Tìm pattern "Số: XX.OXX.XX.XXXX"
            pattern = r'Số:\s*(\d{2}\.O\d{2}\.\d{2}\.\d{4})'
            for para in doc.Paragraphs:
                match = re.search(pattern, para.Range.Text)
                if match:
                    return match.group(1)  # Trả về số phiếu
            # Tìm trong bảng
            for table in doc.Tables:
                for row in table.Rows:
                    for cell in row.Cells:
                        match = re.search(pattern, cell.Range.Text)
                        if match:
                            return match.group(1)
        except Exception as e:
            print(f"[DEBUG] Exception finding so phieu: {e}")
        return None

    def run(self):
        import pythoncom
        import win32com.client
        import os
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            saved_count = 0
            for i in range(word_app.Documents.Count):
                doc = word_app.Documents.Item(i + 1)
                if doc.Name in self.doc_names:
                    try:
                        # Tìm số phiếu trong document
                        so_phieu = self.find_so_phieu(doc)
                        if so_phieu:
                            # Chuyển đổi định dạng số phiếu: XX.OXX.XX.XXXX -> XX.XXXX-XX
                            parts = so_phieu.split('.')
                            if len(parts) == 4:
                                # parts[0] = XX, parts[1] = OXX, parts[2] = XX, parts[3] = XXXX
                                new_format = f"{parts[2]}.{parts[3]}-{parts[0]}"
                                file_name = f"{new_format}{os.path.splitext(doc.Name)[1]}"
                            else:
                                # Nếu format không đúng, dùng số phiếu gốc
                                file_name = f"Phieu_{so_phieu}{os.path.splitext(doc.Name)[1]}"
                        else:
                            # Nếu không tìm thấy số phiếu, dùng tên gốc
                            file_name = os.path.splitext(doc.Name)[0] + "_saved" + os.path.splitext(doc.Name)[1]
                        
                        file_path = os.path.join(self.folder_path, file_name)
                        # Lưu file với định dạng gốc
                        doc.SaveAs(file_path)
                        saved_count += 1
                        print(f"[DEBUG] Saved: {file_name}")
                    except Exception as e:
                        print(f"[DEBUG] Exception saving {doc.Name}: {e}")
            self.finished.emit(f"✅ Đã lưu {saved_count} file vào thư mục đã chọn.")
        except Exception as e:
            self.finished.emit(f"Lỗi lưu file: {e}")
        finally:
            pythoncom.CoUninitialize()





class AutoUpdater:
    def __init__(self, github_repo):
        self.github_repo = github_repo
        self.api_url = f"https://api.github.com/repos/{github_repo}/releases/latest"
        self.temp_dir = os.path.join(os.environ.get('TEMP'), 'QLVT_Update')
        
        # Tạo thư mục temp nếu chưa có
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
    
    def check_for_updates(self, current_version):
        """Kiểm tra phiên bản mới từ GitHub"""
        try:
            print(f"[UPDATE] Đang kiểm tra cập nhật từ {self.github_repo}")
            response = requests.get(self.api_url, timeout=10)
            if response.status_code == 200:
                release_info = response.json()
                latest_version = release_info['tag_name'].lstrip('v')
                print(f"[UPDATE] Phiên bản hiện tại: {current_version}")
                print(f"[UPDATE] Phiên bản mới nhất: {latest_version}")
                
                if self.compare_versions(current_version, latest_version):
                    print(f"[UPDATE] Có phiên bản mới: {latest_version}")
                    return True, release_info
                else:
                    print(f"[UPDATE] Đã là phiên bản mới nhất")
                    return False, None
            else:
                print(f"[UPDATE] Lỗi API: {response.status_code}")
                return False, None
        except requests.exceptions.Timeout:
            print(f"[UPDATE] Timeout khi kiểm tra cập nhật")
            return False, None
        except Exception as e:
            print(f"[UPDATE] Lỗi kiểm tra cập nhật: {e}")
            return False, None
    
    def compare_versions(self, current, latest):
        """So sánh phiên bản theo semantic versioning"""
        try:
            current_parts = [int(x) for x in current.split('.')]
            latest_parts = [int(x) for x in latest.split('.')]
            
            # Đảm bảo cùng độ dài
            while len(current_parts) < len(latest_parts):
                current_parts.append(0)
            while len(latest_parts) < len(current_parts):
                latest_parts.append(0)
                
            return latest_parts > current_parts
        except Exception as e:
            print(f"[UPDATE] Lỗi so sánh version: {e}")
            return False
    
    def get_download_url(self):
        """Lấy URL download file .exe"""
        try:
            # Tạo một dialog để yêu cầu người dùng chọn file .exe
            file_path, _ = QFileDialog.getOpenFileName(
                None, "Chọn file cập nhật", "", "Executable Files (*.exe)"
            )
            if file_path:
                print(f"[UPDATE] Chọn file cập nhật: {file_path}")
                return file_path
            else:
                print(f"[UPDATE] Không chọn được file cập nhật.")
                return None
        except Exception as e:
            print(f"[UPDATE] Lỗi lấy download URL: {e}")
            return None
    
    def download_update(self, download_url, progress_callback=None):
        """Tải xuống file cập nhật với progress"""
        try:
            print(f"[UPDATE] Bắt đầu tải xuống: {download_url}")
            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            # Lấy tên file từ URL
            filename = download_url.split('/')[-1]
            temp_path = os.path.join(self.temp_dir, filename)
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(temp_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress = int((downloaded / total_size) * 100)
                            progress_callback(progress)
            
            print(f"[UPDATE] Tải xuống hoàn tất: {temp_path}")
            return temp_path
        except Exception as e:
            print(f"[UPDATE] Lỗi tải xuống: {e}")
            return None
    
    def check_admin_privileges(self):
        """Kiểm tra quyền Administrator"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def install_update(self, new_exe_path):
        """Cài đặt bản cập nhật"""
        try:
            current_exe_path = sys.argv[0]
            print(f"[UPDATE] Cài đặt từ: {new_exe_path}")
            print(f"[UPDATE] Cài đặt đến: {current_exe_path}")
            
            # Kiểm tra file có tồn tại không
            if not os.path.exists(new_exe_path):
                print(f"[UPDATE] Lỗi: File nguồn không tồn tại: {new_exe_path}")
                return False
            
            # Kiểm tra file đích có tồn tại không
            if not os.path.exists(current_exe_path):
                print(f"[UPDATE] Lỗi: File đích không tồn tại: {current_exe_path}")
                return False
            
            # Tạo batch script để thay thế file với cải tiến
            batch_content = f'''@echo off
setlocal enabledelayedexpansion

echo [UPDATE] ========================================
echo [UPDATE] Bắt đầu cài đặt bản cập nhật...
echo [UPDATE] Thời gian: %date% %time%
echo [UPDATE] ========================================

echo [UPDATE] Kiểm tra file nguồn...
if not exist "{new_exe_path}" (
    echo [UPDATE] LỖI: Không tìm thấy file nguồn {new_exe_path}
    pause
    exit /b 1
)

echo [UPDATE] Kiểm tra file đích...
if not exist "{current_exe_path}" (
    echo [UPDATE] LỖI: Không tìm thấy file đích {current_exe_path}
    pause
    exit /b 1
)

echo [UPDATE] Đang đóng ứng dụng hiện tại...
echo [UPDATE] Tên process: {os.path.basename(current_exe_path)}

REM Đợi ứng dụng đóng hoàn toàn
timeout /t 5 /nobreak >nul

REM Kiểm tra xem process có còn chạy không
:check_lock
echo [UPDATE] Kiểm tra process...
tasklist /FI "IMAGENAME eq {os.path.basename(current_exe_path)}" 2>NUL | find /I /N "{os.path.basename(current_exe_path)}">NUL
if "%ERRORLEVEL%"=="0" (
    echo [UPDATE] Ứng dụng vẫn đang chạy, đợi thêm...
    timeout /t 3 /nobreak >nul
    goto check_lock
)

echo [UPDATE] Ứng dụng đã đóng hoàn toàn!
echo [UPDATE] Bắt đầu cài đặt...

REM Tạo backup trước khi cài đặt
echo [UPDATE] Tạo backup...
copy "{current_exe_path}" "{current_exe_path}.backup" /Y >nul 2>&1

REM Thử copy với retry
set retry_count=0
:copy_retry
echo [UPDATE] Thử copy lần !retry_count!...
copy "{new_exe_path}" "{current_exe_path}" /Y
if %errorlevel% equ 0 (
    echo [UPDATE] ========================================
    echo [UPDATE] CÀI ĐẶT THÀNH CÔNG!
    echo [UPDATE] ========================================
    
    echo [UPDATE] Kiểm tra file mới...
    if exist "{current_exe_path}" (
        echo [UPDATE] File mới đã được tạo thành công
    ) else (
        echo [UPDATE] LỖI: File mới không tồn tại
        pause
        exit /b 1
    )
    
    echo [UPDATE] Khởi động lại ứng dụng...
    timeout /t 2 /nobreak >nul
    
    REM Khởi động ứng dụng mới
    start "" "{current_exe_path}"
    
    echo [UPDATE] Dọn dẹp file tạm...
    del "{new_exe_path}" 2>nul
    del "{current_exe_path}.backup" 2>nul
    del "%~f0" 2>nul
    
    echo [UPDATE] ========================================
    echo [UPDATE] HOÀN TẤT CÀI ĐẶT!
    echo [UPDATE] ========================================
    timeout /t 3 /nobreak >nul
    exit /b 0
) else (
    set /a retry_count+=1
    echo [UPDATE] Lỗi copy (lần !retry_count!), errorlevel: %errorlevel%
    if !retry_count! lss 5 (
        echo [UPDATE] Thử lại sau 3 giây...
        timeout /t 3 /nobreak >nul
        goto copy_retry
    ) else (
        echo [UPDATE] ========================================
        echo [UPDATE] LỖI CÀI ĐẶT SAU 5 LẦN THỬ!
        echo [UPDATE] ========================================
        echo [UPDATE] Chi tiết lỗi:
        echo [UPDATE] - File nguồn: {new_exe_path}
        echo [UPDATE] - File đích: {current_exe_path}
        echo [UPDATE] - Error level cuối: %errorlevel%
        echo [UPDATE] 
        echo [UPDATE] Vui lòng thử cài đặt thủ công hoặc liên hệ hỗ trợ.
        pause
        exit /b 1
    )
)'''
            
            batch_path = os.path.join(self.temp_dir, 'update_qlvt.bat')
            with open(batch_path, 'w', encoding='utf-8') as f:
                f.write(batch_content)
            
            print(f"[UPDATE] Tạo batch script: {batch_path}")
            
            # Chạy batch script với elevated privileges nếu cần
            try:
                print(f"[UPDATE] Chạy batch script với timeout 120 giây...")
                
                # Kiểm tra quyền admin
                if not is_admin():
                    print("[UPDATE] Không có quyền admin, thử chạy với elevated privileges...")
                    # Thử chạy với elevated privileges - sửa cách truyền argument
                    powershell_cmd = f'Start-Process cmd -ArgumentList "/c", "{batch_path}" -Verb RunAs -Wait'
                    result = subprocess.run(['powershell', '-Command', powershell_cmd],
                                          shell=True, 
                                          capture_output=True, 
                                          text=True, 
                                          timeout=120)
                else:
                    # Chạy bình thường nếu đã có quyền admin
                    result = subprocess.run(['cmd', '/c', batch_path], 
                                          shell=True, 
                                          capture_output=True, 
                                          text=True, 
                                          timeout=120)
                
                print(f"[UPDATE] Batch script return code: {result.returncode}")
                print(f"[UPDATE] Batch script output: {result.stdout}")
                if result.stderr:
                    print(f"[UPDATE] Batch script errors: {result.stderr}")
                
                # Kiểm tra kết quả chi tiết
                if result.returncode == 0:
                    print("[UPDATE] Batch script hoàn thành thành công")
                    return True
                else:
                    print(f"[UPDATE] Batch script thất bại với return code: {result.returncode}")
                    return False
                    
            except subprocess.TimeoutExpired:
                print(f"[UPDATE] Batch script timeout sau 120 giây")
                return False
            except Exception as e:
                print(f"[UPDATE] Lỗi chạy batch script: {e}")
                return False
                
        except Exception as e:
            print(f"[UPDATE] Lỗi cài đặt: {e}")
            return False


class PrintWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    
    def __init__(self, doc_names, output_folder=None, action_mode="print", batch_size=5):
        super().__init__()
        self.doc_names = doc_names
        self.output_folder = output_folder
        self.action_mode = action_mode  # "print" hoặc "save_pdf"
        self.batch_size = batch_size
        
    def reconnect_word(self, max_retries=3):
        """Thử kết nối lại Word application với retry"""
        for i in range(max_retries):
            try:
                pythoncom.CoUninitialize()  # Giải phóng kết nối cũ
                time.sleep(1)  # Đợi 1 giây
                pythoncom.CoInitialize()
                word_app = win32com.client.GetActiveObject("Word.Application")
                if word_app:
                    print(f"[DEBUG] ✓ Kết nối lại Word thành công (lần thử {i + 1})")
                    # Thiết lập lại DisplayAlerts = False
                    word_app.DisplayAlerts = False
                    return word_app
            except:
                if i < max_retries - 1:
                    print(f"[DEBUG] Không thể kết nối Word, thử lại lần {i + 2}")
                    time.sleep(2)  # Tăng thời gian đợi
        return None
    
    def get_document_by_name(self, word_app, doc_name, retries=3):
        """Tìm document theo tên với số lần thử lại"""
        for attempt in range(retries):
            try:
                # Làm mới danh sách documents
                docs_count = word_app.Documents.Count
                for j in range(docs_count):
                    try:
                        doc = word_app.Documents.Item(j + 1)
                        if doc and doc.Name == doc_name:
                            return doc
                    except:
                        continue
                        
                if attempt < retries - 1:
                    print(f"[DEBUG] Không tìm thấy {doc_name}, thử lại lần {attempt + 2}")
                    time.sleep(1)  # Đợi 1 giây trước khi thử lại
                    
            except:
                if attempt < retries - 1:
                    print(f"[DEBUG] Lỗi truy cập Documents, thử lại lần {attempt + 2}")
                    word_app = self.reconnect_word()
                    if not word_app:
                        return None
                    time.sleep(1)
                    
        return None
    
    def refresh_word_documents(self, word_app):
        """Làm mới và lấy danh sách documents hiện tại"""
        try:
            return {doc.Name: doc for i in range(word_app.Documents.Count) 
                   for doc in [word_app.Documents.Item(i + 1)]}
        except:
            return {}
    
    def run(self):
        try:
            total_docs = len(self.doc_names)
            processed = 0
            failed = 0
            skipped = []  # Danh sách file bị bỏ qua
            
            # Xử lý theo batch
            for i in range(0, total_docs, self.batch_size):
                batch = self.doc_names[i:i + self.batch_size]
                print(f"[DEBUG] Xử lý batch {i//self.batch_size + 1}/{(total_docs-1)//self.batch_size + 1}")
                
                # Khởi tạo COM mới cho mỗi batch
                pythoncom.CoInitialize()
                word_app = None
                
                try:
                    word_app = win32com.client.GetActiveObject("Word.Application")
                    if not word_app:
                        raise Exception("Không thể kết nối Word")
                    
                    # Refresh và lấy danh sách documents hiện tại
                    docs_dict = self.refresh_word_documents(word_app)
                    
                    # Xử lý từng file trong batch
                    for doc_name in batch:
                        try:
                            # Kiểm tra document có tồn tại không
                            doc = docs_dict.get(doc_name)
                            if not doc:
                                print(f"[DEBUG] Không tìm thấy file: {doc_name}")
                                skipped.append(doc_name)
                                continue
                            
                            if doc:
                                print(f"[DEBUG] Đang xử lý file: {doc_name}")
                                
                                # Kiểm tra số trang
                                total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
                                print(f"[DEBUG] Tổng số trang: {total_pages}")
                                
                                if total_pages > 0:
                                    try:
                                        # Kích hoạt document
                                        doc.Activate()
                                        time.sleep(0.5)  # Chờ một chút để đảm bảo document đã sẵn sàng
                                        
                                        # Lấy máy in mặc định
                                        default_printer = win32print.GetDefaultPrinter()
                                        word_app.ActivePrinter = default_printer
                                        
                                        # In chỉ trang đầu tiên
                                        # Tắt cảnh báo của Word để tránh popup "margins pretty small"
                                        # wdAlertsNone = 0, wdAlertsAll = -1
                                        try:
                                            word_app.DisplayAlerts = 0
                                        except:
                                            pass

                                        try:
                                            if self.action_mode == "save_pdf" and self.output_folder:
                                                # Chế độ lưu PDF - export ra PDF
                                                import os as _os
                                                safe_name = ''.join(c for c in doc_name if c.isalnum() or c in (' ', '.', '_')).rstrip()
                                                if safe_name.lower().endswith('.docx'):
                                                    safe_name = safe_name[:-5]
                                                elif safe_name.lower().endswith('.doc'):
                                                    safe_name = safe_name[:-4]
                                                elif safe_name.lower().endswith('.rtf'):
                                                    safe_name = safe_name[:-4]
                                                
                                                pdf_path = _os.path.join(self.output_folder, f"{safe_name}_trang1.pdf")
                                                
                                                print(f"[DEBUG] Export trang 1 sang PDF: {pdf_path}")
                                                doc.ExportAsFixedFormat(
                                                    OutputFileName=pdf_path,
                                                    ExportFormat=17,  # wdExportFormatPDF
                                                    OpenAfterExport=False,
                                                    From=1,
                                                    To=1,
                                                    OptimizeFor=0,
                                                    Range=3  # wdExportFromTo
                                                )
                                                print(f"[DEBUG] Đã lưu PDF: {pdf_path}")
                                            else:
                                                # Chế độ in - in trực tiếp ra máy in
                                                print(f"[DEBUG] In trực tiếp trang đầu tiên ra máy in...")
                                                
                                                # Lấy máy in mặc định
                                                default_printer = win32print.GetDefaultPrinter()
                                                print(f"[DEBUG] Máy in: {default_printer}")
                                                
                                                # Đặt máy in cho document
                                                word_app.ActivePrinter = default_printer
                                                
                                                # In chỉ trang 1 - giống VBA
                                                # PrintOut(Background, Append, Range, OutputFileName, From, To, ...)
                                                # Range=3: wdPrintFromTo
                                                print(f"[DEBUG] Gọi PrintOut với Range=3, From=1, To=1")
                                                doc.PrintOut(
                                                    False,  # Background
                                                    False,  # Append  
                                                    3,      # Range = wdPrintFromTo
                                                    "",     # OutputFileName
                                                    "1",    # From
                                                    "1"     # To
                                                )
                                                print(f"[DEBUG] Đã gửi lệnh in trang 1 ra máy in")
                                            
                                            processed += 1
                                            print(f"[DEBUG] ✓ Đã xử lý thành công: {doc_name}")
                                            
                                        finally:
                                            # Khôi phục cảnh báo
                                            try:
                                                word_app.DisplayAlerts = -1  # wdAlertsAll
                                            except:
                                                pass

                                    except Exception as print_error:
                                        print(f"[DEBUG] Lỗi khi in: {str(print_error)}")
                                        failed += 1
                                        raise
                                else:
                                    print(f"[DEBUG] Tài liệu không có nội dung: {doc_name}")
                                    skipped.append(doc_name)
                                
                        except Exception as e:
                            failed += 1
                            print(f"[DEBUG] ✗ Lỗi in file {doc_name}: {str(e)}")
                        finally:
                            if doc:
                                doc = None  # Giải phóng document
                        
                        # Cập nhật progress
                        progress = int((processed + failed) * 100 / total_docs)
                        self.progress.emit(progress)
                
                except Exception as e:
                    print(f"[DEBUG] Lỗi xử lý batch: {str(e)}")
                    # Đánh dấu các file còn lại trong batch là lỗi
                    remaining = len([x for x in batch if x not in [doc.Name for doc in word_app.Documents]])
                    failed += remaining
                
                finally:
                    # Giải phóng COM sau mỗi batch
                    pythoncom.CoUninitialize()
            
            # Tổng kết chi tiết
            action_text = "LƯU PDF" if self.action_mode == "save_pdf" else "IN PHIẾU"
            print(f"\n=== TỔNG KẾT {action_text} ===")
            print(f"Tổng số file: {total_docs}")
            print(f"✓ Đã xử lý thành công: {processed}")
            print(f"✗ Lỗi khi xử lý: {failed}")
            if skipped:
                print(f"⚠️ Không tìm thấy {len(skipped)} file:")
                for doc_name in skipped:
                    print(f"  - {doc_name}")
            
            # Thông báo tổng kết
            if processed > 0:
                if self.action_mode == "save_pdf":
                    msg = f"✅ Đã lưu PDF trang đầu của {processed}/{total_docs} tài liệu"
                    if self.output_folder:
                        msg += f"\nThư mục: {self.output_folder}"
                else:
                    msg = f"✅ Đã in xong {processed}/{total_docs} tài liệu"
                
                if failed > 0:
                    msg += f" ({failed} lỗi)"
                if skipped:
                    msg += f" ({len(skipped)} file không tìm thấy)"
                self.finished.emit(msg)
            else:
                if self.action_mode == "save_pdf":
                    self.finished.emit(f"❌ Không lưu được tài liệu nào")
                else:
                    self.finished.emit(f"❌ Không in được tài liệu nào")
            
        except Exception as e:
            self.finished.emit(f"❌ Lỗi hệ thống: {str(e)}")


# ============================================================================
# EXCEL PROCESSOR WORKER THREAD
# ============================================================================

class ExcelProcessorWorker(QThread):
    """Worker thread để xử lý Excel trong background"""
    status_update = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    progress_start = pyqtSignal()
    progress_stop = pyqtSignal()
    
    def __init__(self, file_path, processor_type):
        super().__init__()
        self.file_path = file_path
        self.processor_type = processor_type
    
    def run(self):
        try:
            self.progress_start.emit()
            
            # Chọn processor
            if self.processor_type == "sctx":
                self.status_update.emit("Khởi tạo SCTX Processor...\n")
                processor = SCTXProcessor(self.file_path)
            else:
                self.status_update.emit("Khởi tạo NTVTDD Processor...\n")
                processor = NTVTDDProcessor(self.file_path)
            
            # Đọc file
            self.status_update.emit("Đang đọc file Excel...\n")
            if not processor.read_file():
                self.finished_signal.emit(False, "Không thể đọc file Excel!")
                return
            
            self.status_update.emit("✓ Đọc file thành công!\n")
            
            # Xử lý dữ liệu
            self.status_update.emit("Đang xử lý dữ liệu...\n")
            if not processor.process():
                self.finished_signal.emit(False, "Lỗi khi xử lý dữ liệu!")
                return
            
            self.status_update.emit("✓ Xử lý dữ liệu thành công!\n")
            
            # Xuất file
            self.status_update.emit("Đang xuất file kết quả...\n")
            if not processor.export():
                self.finished_signal.emit(False, "Lỗi khi xuất file!")
                return
            
            # Tạo tên file output
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'Ket_qua_xu_ly_{timestamp}.xlsx'
            
            self.status_update.emit("✓ Xuất file thành công!\n")
            self.status_update.emit("-" * 60 + "\n")
            self.status_update.emit(f"✓ HOÀN THÀNH!\n")
            self.status_update.emit(f"✓ File kết quả: {output_file}\n")
            
            self.finished_signal.emit(True, f"Xử lý file thành công!\n\nFile kết quả: {output_file}")
            
        except Exception as e:
            self.status_update.emit(f"\n✗ LỖI: {str(e)}\n")
            self.finished_signal.emit(False, f"Đã xảy ra lỗi:\n{str(e)}")
        
        finally:
            self.progress_stop.emit()


# ============================================================================
# EXCEL PROCESSOR TAB
# ============================================================================

class ExcelProcessorTab(QWidget):
    """Tab xử lý Excel trong ứng dụng chính"""
    
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.is_processing = False
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        # title_label = QLabel("CHƯƠNG TRÌNH XỬ LÝ DỮ LIỆU EXCEL")
        # title_label.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px;")
        # title_label.setAlignment(Qt.AlignCenter)
        # layout.addWidget(title_label)
        
        # Radio buttons frame
        radio_group_box = QLabel("Chọn loại file Excel:")
        radio_group_box.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(radio_group_box)
        
        # Radio buttons
        self.processor_type = "sctx"
        self.button_group = QButtonGroup()
        
        self.sctx_radio = QRadioButton("File loại SCTX (Mã phiếu: 02.O09.42.xxxx hoặc 03.O09.42.xxxx)")
        self.sctx_radio.setChecked(True)
        self.sctx_radio.toggled.connect(lambda: self.set_processor_type("sctx"))
        self.button_group.addButton(self.sctx_radio)
        layout.addWidget(self.sctx_radio)
        
        self.ntvtdd_radio = QRadioButton("File loại NTVTDD (Mã phiếu linh hoạt, có xử lý mã vật tư)")
        self.ntvtdd_radio.toggled.connect(lambda: self.set_processor_type("ntvtdd"))
        self.button_group.addButton(self.ntvtdd_radio)
        layout.addWidget(self.ntvtdd_radio)
        
        # File selection
        file_label = QLabel("Chọn file:")
        file_label.setStyleSheet("font-weight: bold; margin-top: 20px;")
        layout.addWidget(file_label)
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel("Chưa chọn file")
        self.file_label.setStyleSheet("color: gray;")
        file_layout.addWidget(self.file_label)
        
        choose_btn = QPushButton("📁 Chọn File Excel")
        choose_btn.clicked.connect(self.choose_file)
        file_layout.addWidget(choose_btn)
        layout.addLayout(file_layout)
        
        # Process button
        self.process_btn = QPushButton("▶ Xử lý File")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setStyleSheet("padding: 10px; font-size: 14px; margin-top: 10px;")
        layout.addWidget(self.process_btn)
        
        # Progress bar
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # Indeterminate mode
        self.progress.setVisible(False)
        layout.addWidget(self.progress)
        
        # Status text
        status_label = QLabel("Trạng thái:")
        status_label.setStyleSheet("font-weight: bold; margin-top: 20px;")
        layout.addWidget(status_label)
        
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMinimumHeight(200)
        self.status_text.setStyleSheet("font-family: Consolas; font-size: 9pt;")
        layout.addWidget(self.status_text)
        
        # Initial status
        self.update_status("Sẵn sàng xử lý. Vui lòng chọn file Excel...\n")
        
        layout.addStretch()
        self.setLayout(layout)
    
    def set_processor_type(self, ptype):
        self.processor_type = ptype
    
    def choose_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.setText(filename)
            self.file_label.setStyleSheet("color: black;")
            self.process_btn.setEnabled(True)
            self.update_status(f"✓ Đã chọn file: {filename}\n")
    
    def process_file(self):
        if not self.file_path:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn file Excel trước!")
            return
        
        if self.is_processing:
            QMessageBox.information(self, "Thông báo", "Đang xử lý file, vui lòng đợi...")
            return
        
        # Disable button và start progress
        self.process_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.is_processing = True
        
        # Clear status
        self.status_text.clear()
        self.update_status(f"Bắt đầu xử lý file: {os.path.basename(self.file_path)}\n")
        self.update_status(f"Loại xử lý: {self.processor_type.upper()}\n")
        self.update_status("-" * 60 + "\n")
        
        # Run processor in thread
        self.worker = ExcelProcessorWorker(self.file_path, self.processor_type)
        self.worker.status_update.connect(self.update_status)
        self.worker.finished_signal.connect(self.on_processing_finished)
        self.worker.progress_start.connect(lambda: self.progress.setVisible(True))
        self.worker.progress_stop.connect(lambda: self.progress.setVisible(False))
        self.worker.start()
    
    def update_status(self, message):
        self.status_text.append(message.rstrip())
        self.status_text.verticalScrollBar().setValue(
            self.status_text.verticalScrollBar().maximum()
        )
    
    def on_processing_finished(self, success, message):
        self.progress.setVisible(False)
        self.process_btn.setEnabled(True)
        self.is_processing = False
        
        if success:
            QMessageBox.information(self, "Thành công", message)
        else:
            QMessageBox.critical(self, "Lỗi", message)


# ============================================================================
# MAIN WINDOW WITH TABS
# ============================================================================

class MainWindow(QWidget):
    """Cửa sổ chính với tab cho Word và Excel processor"""
    
    def __init__(self):
        super().__init__()
        self.current_version = "1.0.21"
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle(f"Công cụ xử lý phiếu nhập xuất kho {self.current_version} | www.khoatran.io.vn")
        self.setGeometry(200, 200, 600, 400)
        
        # Thiết lập icon
        icon = QIcon("icon.ico")
        self.setWindowIcon(icon)
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        
        # Main layout
        layout = QVBoxLayout()
        
        # Create tab widget
        self.tabs = QTabWidget()
        
        # Add Word Processor tab
        self.word_tab = WordProcessorApp()
        self.tabs.addTab(self.word_tab, "📄 Xử lý Word")
        
        # Add Excel Processor tab
        self.excel_tab = ExcelProcessorTab()
        self.tabs.addTab(self.excel_tab, "📊 Xử lý Excel")
        
        layout.addWidget(self.tabs)
        self.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
