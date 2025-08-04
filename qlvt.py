import sys
import pythoncom
import win32com.client
import requests
import subprocess
import ctypes
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QCheckBox, QHBoxLayout,
    QLineEdit, QFormLayout, QDialog, QDialogButtonBox, QFileDialog,
    QScrollArea, QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
import os

REPLACEMENT_FILE = "replacements.txt"

def is_admin():
    """Kiểm tra xem ứng dụng có chạy với quyền admin không"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

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

class ReplaceWorker(QThread):
    finished = pyqtSignal(str)
    def __init__(self, doc_names, replacements, parent=None):
        super().__init__(parent)
        self.doc_names = doc_names
        self.replacements = replacements

    def run(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            for i in range(word_app.Documents.Count):
                doc = word_app.Documents.Item(i + 1)
                if doc.Name in self.doc_names:
                    try:
                        # Lọc tất cả các bảng ở trang đầu tiên
                        tables_on_first_page = [table for table in doc.Tables if table.Range.Information(3) == 1]
                        print(f"[DEBUG] Tổng số bảng trên trang đầu: {len(tables_on_first_page)}")
                        if tables_on_first_page:
                            # Xử lý tất cả các bảng trên trang đầu tiên
                            for table_idx, table in enumerate(tables_on_first_page):
                                print(f"[DEBUG] ===== Đang xử lý Bảng {table_idx + 1} =====")
                                print(f"[DEBUG] Số row trong bảng {table_idx + 1}: {table.Rows.Count}")
                                print(f"[DEBUG] Số column trong bảng {table_idx + 1}: {table.Columns.Count}")
                                
                                try:
                                    # Sử dụng Range.Cells để tránh lỗi với merged cells
                                    for cell_idx, cell in enumerate(table.Range.Cells):
                                        cell_text = cell.Range.Text.strip()
                                        if cell_text:  # Chỉ in cell có nội dung
                                            print(f"[DEBUG] Bảng{table_idx+1} - Cell {cell_idx+1}: '{cell_text}'")
                                            
                                            for old, new in self.replacements:
                                                if old in cell_text:
                                                    print(f"[DEBUG] ✓ Found '{old}' in Bảng{table_idx+1} - Cell {cell_idx+1}!")
                                                    try:
                                                        # Thay thế bằng cách tìm vị trí và thay thế trực tiếp
                                                        cell_range = cell.Range
                                                        start_pos = cell_range.Start
                                                        end_pos = cell_range.End
                                                        search_range = doc.Range(start_pos, end_pos)
                                                        search_range.Find.Text = old
                                                        if search_range.Find.Execute():
                                                            search_range.Text = new
                                                            print(f"[DEBUG] ✓ Replaced '{old}' with '{new}' in Bảng{table_idx+1} - Cell {cell_idx+1}")
                                                        else:
                                                            print(f"[DEBUG] ✗ Find.Execute() failed for '{old}' in Bảng{table_idx+1} - Cell {cell_idx+1}")
                                                    except Exception as e:
                                                        print(f"[DEBUG] ✗ Exception replacing '{old}' in Bảng{table_idx+1} - Cell {cell_idx+1}: {e}")
                                                else:
                                                    print(f"[DEBUG] - NOT found '{old}' in Bảng{table_idx+1} - Cell {cell_idx+1}")
                                except Exception as e:
                                    print(f"[DEBUG] Exception processing Bảng{table_idx+1}: {e}")
                                    # Fallback: thử cách khác nếu có lỗi
                                    try:
                                        for old, new in self.replacements:
                                            # Thay thế trong toàn bộ Range của bảng
                                            table_range = table.Range
                                            table_range.Find.Text = old
                                            table_range.Find.Replacement.Text = new
                                            if table_range.Find.Execute(Replace=2, Forward=True):
                                                print(f"[DEBUG] ✓ Replaced '{old}' with '{new}' in Bảng{table_idx+1} (fallback method)")
                                    except Exception as e2:
                                        print(f"[DEBUG] Fallback also failed for Bảng{table_idx+1}: {e2}")
                    except Exception as e:
                        print(f"[DEBUG] Exception in replace: {e}")
            self.finished.emit("✅ Đã thay thế xong các tài liệu được chọn.")
        except Exception as e:
            self.finished.emit(f"Lỗi thay thế: {e}")
        finally:
            pythoncom.CoUninitialize()

class WordProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.current_version = "1.0.13"
        self.setWindowTitle(f"Xử lý phiếu hàng loạt v{self.current_version} | www.khoatran.io.vn")
        self.setGeometry(200, 200, 600, 400)  # Tăng kích thước cửa sổ mặc định

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
        self.layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()
        self.refresh_button = QPushButton("Load DS phiếu")
        self.refresh_button.clicked.connect(self.load_open_documents)
        button_layout.addWidget(self.refresh_button)

        self.process_button = QPushButton("Thay khung ký tên")
        self.process_button.clicked.connect(self.process_selected_files)
        button_layout.addWidget(self.process_button)

        # Thêm nút Replace
        self.replace_button = QPushButton("Thay tên")
        self.replace_button.clicked.connect(self.replace_selected_files)
        button_layout.addWidget(self.replace_button)

        # Thêm nút In trang đầu
        self.print_button = QPushButton("In phiếu đã chọn")
        self.print_button.clicked.connect(self.print_first_pages)
        button_layout.addWidget(self.print_button)

        # Thêm nút Save As (cuối cùng)
        self.save_as_button = QPushButton("Lưu tất cả file")
        self.save_as_button.clicked.connect(self.save_all_files_as)
        button_layout.addWidget(self.save_as_button)

    
       

        self.layout.addLayout(button_layout)
        self.setLayout(self.layout)

        # 🔄 GỌI NGAY khi khởi động để tự động tải danh sách tài liệu đang mở
        self.load_open_documents()

    def _check_activation(self):
        """Kiểm tra trạng thái activation khi khởi động"""
        try:
            status = self.activation_manager.check_activation_status()
            
            if not status.get('activated', True):
                # Hiển thị thông báo deactivated
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setWindowTitle("Lỗi")
                msg.setText("❌ Không có kết nối đến server")
                msg.setInformativeText(status.get('message', 'Không có thông tin chi tiết'))
                
                # Thêm thông tin expiry date nếu có
                expiry_date = status.get('expiry_date')
                if expiry_date:
                    msg.setDetailedText(f"Ngày hết hạn: {expiry_date}\n\nLiên hệ admin để được hỗ trợ.")
                
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
                
                # Thoát ứng dụng
                QApplication.quit()
                return False
            
            return True
            
        except Exception as e:
            print(f"[ACTIVATION] Lỗi kiểm tra activation: {e}")
            # Nếu có lỗi, cũng thoát ứng dụng để tránh bypass
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Lỗi")
            msg.setText("❌ Không thể kiểm tra trạng thái kích hoạt")
            msg.setInformativeText("Ứng dụng sẽ thoát để đảm bảo an toàn.")
            msg.setDetailedText(f"Chi tiết lỗi: {str(e)}")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
            
            QApplication.quit()
            return False

    def show_activation_status(self):
        """Hiển thị thông tin trạng thái activation"""
        try:
            status = self.activation_manager.check_activation_status()
            
            msg = QMessageBox()
            if status.get('activated', True):
                msg.setIcon(QMessageBox.Information)
                msg.setWindowTitle("Trạng thái")
                msg.setText("✅ Ứng dụng đang được kích hoạt")
            else:
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Trạng thái")
                msg.setText("❌ Lỗi kết nối đến server")
            
            # Thông tin chi tiết
            details = []
            if 'expiry_date' in status:
                details.append(f"Ngày hết hạn: {status['expiry_date']}")
            if 'message' in status:
                details.append(f"Thông báo: {status['message']}")
            if 'last_updated' in status:
                details.append(f"Cập nhật lần cuối: {status['last_updated']}")
            
            if details:
                msg.setInformativeText('\n'.join(details))
            
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
            
        except Exception as e:
            QMessageBox.warning(self, "Lỗi", f"Không thể kiểm tra trạng thái activation: {e}")

    def load_open_documents(self):
        self.file_list.clear()
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            docs = word_app.Documents
            for i in range(docs.Count):
                doc = docs.Item(i + 1)
                item_text = doc.Name
                item = QListWidgetItem(item_text)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)  # Tự động check vào tất cả file
                self.file_list.addItem(item)
        except Exception as e:
            self.status_label.setText(f"Lỗi: {e}")
        finally:
            pythoncom.CoUninitialize()

    def process_selected_files(self):
        selected_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_files.append(item.text())

        if not selected_files:
            self.status_label.setText("⚠️ Bạn chưa chọn tài liệu nào để xử lý.")
            return

        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            for i in range(word_app.Documents.Count):
                doc = word_app.Documents.Item(i + 1)
                if doc.Name in selected_files:
                    self.modify_document(doc)
            self.status_label.setText("✅ Đã xử lý xong các tài liệu được chọn.")
        except Exception as e:
            self.status_label.setText(f"Lỗi xử lý: {e}")
        finally:
            pythoncom.CoUninitialize()

    def replace_in_first_page(self, doc, replacements):
        try:
            for para in doc.Paragraphs:
                if para.Range.Information(3) == 1:  # Trang đầu tiên
                    for old, new in replacements:
                        para.Range.Text = para.Range.Text.replace(old, new)
            # Thay thế trong bảng ở trang đầu tiên (nếu có)
            for table in doc.Tables:
                if table.Range.Information(3) == 1:
                    for row in table.Rows:
                        for cell in row.Cells:
                            for old, new in replacements:
                                cell.Range.Text = cell.Range.Text.replace(old, new)
        except Exception as e:
            pass

    def modify_document(self, doc):
        try:
            # Lọc ra tất cả các bảng nằm ở trang đầu tiên (page 1)
            tables_on_first_page = [table for table in doc.Tables if table.Range.Information(3) == 1]
            print(f"[DEBUG] Số bảng trên trang đầu: {len(tables_on_first_page)}")
            if tables_on_first_page:
                # Chỉ lấy bảng CUỐI CÙNG ở trang đầu tiên (bảng ký tên)
                table = tables_on_first_page[-1]
                rows = table.Rows.Count
                print(f"[DEBUG] Số row trước khi chèn: {rows}")
                if rows == 4:
                    # ⚠️ CHÈN 1 DÒNG vào giữa dòng 3 và 4
                    table.Rows.Add(BeforeRow=table.Rows(4))
                    print(f"[DEBUG] Đã chèn 1 row, số row sau khi chèn: {table.Rows.Count}")
                # ✅ Tiếp tục xử lý nội dung sau khi thêm dòng
                try:
                    table.Cell(1, 3).Range.Text = ""  # Xoá "NGƯỜI LẬP PHIẾU"
                    table.Cell(1, 3).Merge(table.Cell(1, 4))  # Gộp ô (1,3) và (1,4)
                    
                    # Tìm ô chứa "VÕ THANH ĐIỀN" ở hàng cuối cùng
                    last_row = table.Rows.Count
                    target_cell = None
                    for col in range(1, table.Columns.Count + 1):
                        cell_text = table.Cell(last_row, col).Range.Text.strip()
                        if "VÕ THANH ĐIỀN" in cell_text:
                            # Gộp ô chứa "VÕ THANH ĐIỀN" với ô bên phải
                            if col < table.Columns.Count:
                                table.Cell(last_row, col).Merge(table.Cell(last_row, col + 1))
                                target_cell = table.Cell(last_row, col)  # Lưu ô đích
                                target_cell.Range.Text = ""  # Xoá "VÕ THANH ĐIỀN" sau khi đã gộp
                            break
                    
                    # Tìm họ tên người nhận/giao hàng và điền vào ô đích
                    if target_cell:
                        print(f"[DEBUG] Đã tìm thấy ô đích để điền họ tên")
                        ho_ten = self.find_ho_ten_nguoi_hang(doc)
                        if ho_ten:
                            target_cell.Range.Text = ho_ten
                            print(f"[DEBUG] Đã điền họ tên: {ho_ten}")
                        else:
                            print(f"[DEBUG] Không tìm thấy họ tên người nhận/giao hàng")
                    else:
                        print(f"[DEBUG] Không tìm thấy ô đích (ô chứa VÕ THANH ĐIỀN)")
                except:
                    pass
        except Exception as e:
            print(f"[DEBUG] Exception in modify_document: {e}")

    def replace_selected_files(self):
        dialog = ReplaceDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            replacements = dialog.get_replacement_pairs()
            selected_files = []
            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                if item.checkState() == Qt.Checked:
                    selected_files.append(item.text())
            if not selected_files:
                self.status_label.setText("⚠️ Bạn chưa chọn tài liệu nào để thay thế.")
                return
            self.status_label.setText("⏳ Đang thay thế, vui lòng chờ...")
            self.replace_thread = ReplaceWorker(selected_files, replacements)
            self.replace_thread.finished.connect(self.on_replace_finished)
            self.replace_thread.start()

    def on_replace_finished(self, message):
        self.status_label.setText(message)

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

        self.status_label.setText("⏳ Đang in trang đầu, vui lòng chờ...")
        self.print_thread = PrintWorker(selected_files)
        self.print_thread.finished.connect(self.on_print_finished)
        self.print_thread.start()

    def on_print_finished(self, message):
        self.status_label.setText(message)

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
        """Hiển thị dialog xác nhận cập nhật"""
        latest_version = release_info['tag_name'].lstrip('v')
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Cập nhật mới")
        msg.setText(f"Có phiên bản mới: v{latest_version}")
        msg.setInformativeText("Bạn có muốn cập nhật ngay bây giờ không?")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.Yes)
        
        if msg.exec_() == QMessageBox.Yes:
            self.perform_update(release_info)
    
    def perform_update(self, release_info):
        """Thực hiện cập nhật"""
        download_url = self.updater.get_download_url(release_info)
        if download_url:
            # Hiển thị progress dialog
            self.progress_dialog = QDialog(self)
            self.progress_dialog.setWindowTitle("Đang cập nhật...")
            self.progress_dialog.setFixedSize(400, 150)
            
            layout = QVBoxLayout()
            
            self.progress_label = QLabel("Đang tải xuống cập nhật...")
            layout.addWidget(self.progress_label)
            
            self.progress_bar = QProgressBar()
            layout.addWidget(self.progress_bar)
            
            self.progress_dialog.setLayout(layout)
            self.progress_dialog.show()
            
            # Bắt đầu update worker
            self.update_worker = UpdateWorker(self.updater, download_url, self)
            self.update_worker.progress.connect(self.update_progress)
            self.update_worker.finished.connect(self.on_update_finished)
            self.update_worker.start()
        else:
            QMessageBox.warning(self, "Lỗi", "Không thể tải xuống cập nhật.")
    
    def update_progress(self, value):
        """Cập nhật progress bar"""
        self.progress_bar.setValue(value)
        if value == 100:
            self.progress_label.setText("Đang cài đặt cập nhật...")
    
    def on_update_finished(self, message):
        """Xử lý khi update hoàn tất"""
        self.progress_dialog.close()
        
        if "thành công" in message:
            # Hiển thị thông báo và đóng ứng dụng
            QMessageBox.information(self, "Cập nhật", 
                f"{message}\n\nỨng dụng sẽ đóng để hoàn tất cài đặt.\n\nNếu ứng dụng không khởi động lại tự động, vui lòng chạy lại file .exe.")
            
            # Đóng ứng dụng ngay lập tức để batch script có thể thay thế file
            print("[UPDATE] Đóng ứng dụng để hoàn tất cài đặt...")
            QApplication.quit()
        else:
            # Hiển thị lỗi chi tiết hơn
            error_message = f"Lỗi cập nhật:\n{message}\n\nVui lòng thử lại hoặc liên hệ hỗ trợ."
            QMessageBox.critical(self, "Lỗi cập nhật", error_message)



    
            

class ReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Thay thế cụm từ")
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
        
        ok_button = QPushButton("OK")
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
            
            if old_text and new_text:  # Chỉ lấy các cặp có đủ cả 2 từ
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


class PrintWorker(QThread):
    finished = pyqtSignal(str)
    def __init__(self, doc_names, parent=None):
        super().__init__(parent)
        self.doc_names = doc_names

    def run(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            printed_count = 0
            for i in range(word_app.Documents.Count):
                doc = word_app.Documents.Item(i + 1)
                if doc.Name in self.doc_names:
                    try:
                        # In trang đầu tiên - xóa các trang khác, in
                        print(f"[DEBUG] Document name: {doc.Name}")
                        total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
                        print(f"[DEBUG] Total pages: {total_pages}")
                        
                        if total_pages > 1:
                            # Kích hoạt document này
                            doc.Activate()
                            
                            # Xóa từ trang 2 trở đi
                            word_app.Selection.GoTo(What=1, Which=1, Count=2)  # Đi đến trang 2
                            start_pos = word_app.Selection.Start
                            delete_range = doc.Range(start_pos, doc.Content.End)
                            delete_range.Delete()
                            
                            print(f"[DEBUG] Deleted pages 2-{total_pages}, now only page 1 remains")
                        
                        # In toàn bộ document (giờ chỉ còn trang 1)
                        doc.PrintOut()
                        
                        printed_count += 1
                        print(f"[DEBUG] Printed: {doc.Name}")
                    except Exception as e:
                        print(f"[DEBUG] Exception printing {doc.Name}: {e}")
            self.finished.emit(f"✅ Đã in trang đầu của {printed_count} file.")
        except Exception as e:
            self.finished.emit(f"Lỗi in file: {e}")
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
    
    def get_download_url(self, release_info):
        """Lấy URL download file .exe"""
        try:
            for asset in release_info.get('assets', []):
                if asset['name'].endswith('.exe') and 'QLVT_Processor' in asset['name']:
                    print(f"[UPDATE] Tìm thấy file: {asset['name']}")
                    return asset['browser_download_url']
            print(f"[UPDATE] Không tìm thấy file .exe trong release")
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
    
    def install_update(self, new_exe_path):
        """Cài đặt bản cập nhật"""
        try:
            current_exe_path = sys.argv[0]
            print(f"[UPDATE] Cài đặt từ: {new_exe_path}")
            print(f"[UPDATE] Cài đặt đến: {current_exe_path}")
            
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
            
            print(f"[UPDATE] Chạy batch script: {batch_path}")
            
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


class UpdateWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    
    def __init__(self, updater, download_url, parent=None):
        super().__init__(parent)
        self.updater = updater
        self.download_url = download_url
    
    def run(self):
        try:
            # Tải xuống với progress
            new_exe_path = self.updater.download_update(
                self.download_url, 
                self.progress.emit
            )
            
            if new_exe_path:
                # Cài đặt
                if self.updater.install_update(new_exe_path):
                    self.finished.emit("✅ Cập nhật thành công! Ứng dụng sẽ khởi động lại.")
                else:
                    self.finished.emit("❌ Lỗi cài đặt cập nhật.")
            else:
                self.finished.emit("❌ Lỗi tải xuống cập nhật.")
        except Exception as e:
            self.finished.emit(f"❌ Lỗi cập nhật: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WordProcessorApp()
    window.show()
    sys.exit(app.exec_())
