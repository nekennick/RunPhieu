import sys
import pythoncom
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QCheckBox, QHBoxLayout,
    QLineEdit, QFormLayout, QDialog, QDialogButtonBox, QFileDialog,
    QScrollArea, QMessageBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import os

REPLACEMENT_FILE = "replacements.txt"

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
        self.current_version = "1.0.0"
        self.setWindowTitle(f"Xử lý phiếu hàng loạt v{self.current_version} | www.khoatran.io.vn")
        self.setGeometry(200, 200, 600, 400)  # Tăng kích thước cửa sổ mặc định

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
        self.print_button = QPushButton("In tất cả phiếu")
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
                        # In trang đầu tiên
                        doc.PrintOut(From=1, To=1)
                        printed_count += 1
                        print(f"[DEBUG] Printed: {doc.Name}")
                    except Exception as e:
                        print(f"[DEBUG] Exception printing {doc.Name}: {e}")
            self.finished.emit(f"✅ Đã in trang đầu của {printed_count} file.")
        except Exception as e:
            self.finished.emit(f"Lỗi in file: {e}")
        finally:
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WordProcessorApp()
    window.show()
    sys.exit(app.exec_())
