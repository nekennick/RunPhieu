import sys
import pythoncom
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QCheckBox, QHBoxLayout,
    QLineEdit, QFormLayout, QDialog, QDialogButtonBox, QFileDialog
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
        self.setWindowTitle("Xử lý phiếu xuất nhập kho hàng loạt | www.khoatran.io.vn")
        self.setGeometry(200, 200, 600, 400)  # Tăng kích thước cửa sổ mặc định

        self.layout = QVBoxLayout()

        self.status_label = QLabel("Danh sách tài liệu Word đang mở:")
        self.layout.addWidget(self.status_label)

        self.file_list = QListWidget()
        self.layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()
        self.refresh_button = QPushButton("Tải lại danh sách")
        self.refresh_button.clicked.connect(self.load_open_documents)
        button_layout.addWidget(self.refresh_button)

        self.process_button = QPushButton("Xử lý các file đã chọn")
        self.process_button.clicked.connect(self.process_selected_files)
        button_layout.addWidget(self.process_button)

        # Thêm nút Replace
        self.replace_button = QPushButton("Thay thế cụm từ")
        self.replace_button.clicked.connect(self.replace_selected_files)
        button_layout.addWidget(self.replace_button)

        # Thêm nút In trang đầu
        self.print_button = QPushButton("In trang đầu")
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
                    for col in range(1, table.Columns.Count + 1):
                        cell_text = table.Cell(last_row, col).Range.Text.strip()
                        if "VÕ THANH ĐIỀN" in cell_text:
                            # Gộp ô chứa "VÕ THANH ĐIỀN" với ô bên phải
                            if col < table.Columns.Count:
                                table.Cell(last_row, col).Merge(table.Cell(last_row, col + 1))
                                table.Cell(last_row, col).Range.Text = ""  # Xoá "VÕ THANH ĐIỀN" sau khi đã gộp
                            break
                except:
                    pass
        except Exception as e:
            print(f"[DEBUG] Exception in modify_document: {e}")

    def replace_selected_files(self):
        dialog = ReplaceDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            replacements = dialog.get_pairs()
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

    
            

class ReplaceDialog(QDialog):
    def __init__(self, parent=None, replace_callback=None):
        super().__init__(parent)
        self.setWindowTitle("Nhập các cặp cụm từ cần thay thế")
        self.layout = QVBoxLayout()
        self.form_layout = QFormLayout()
        self.pair_edits = []
        for i in range(5):  # Cho phép nhập tối đa 5 cặp
            old_edit = QLineEdit()
            new_edit = QLineEdit()
            self.form_layout.addRow(f"Từ cũ {i+1}", old_edit)
            self.form_layout.addRow(f"Từ mới {i+1}", new_edit)
            self.pair_edits.append((old_edit, new_edit))
        self.layout.addLayout(self.form_layout)
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.on_ok_clicked)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
        # Thêm nút Thay thế
        self.replace_button = QPushButton("Thay thế")
        self.replace_button.clicked.connect(self.on_replace_clicked)
        self.layout.addWidget(self.replace_button)
        self.setLayout(self.layout)
        self.load_pairs_from_file()
        self.replace_callback = replace_callback

    def load_pairs_from_file(self):
        if os.path.exists(REPLACEMENT_FILE):
            try:
                with open(REPLACEMENT_FILE, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                for i, line in enumerate(lines):
                    if i >= 5:
                        break
                    parts = line.strip().split('=>', 1)
                    if len(parts) == 2:
                        self.pair_edits[i][0].setText(parts[0])
                        self.pair_edits[i][1].setText(parts[1])
            except Exception:
                pass

    def save_pairs_to_file(self):
        pairs = self.get_pairs()
        try:
            with open(REPLACEMENT_FILE, 'w', encoding='utf-8') as f:
                for old, new in pairs:
                    f.write(f"{old}=>{new}\n")
        except Exception:
            pass

    def on_ok_clicked(self):
        self.save_pairs_to_file()
        self.accept()

    def on_replace_clicked(self):
        self.save_pairs_to_file()
        if self.replace_callback:
            self.replace_callback(self.get_pairs())
        self.accept()

    def get_pairs(self):
        pairs = []
        for old_edit, new_edit in self.pair_edits:
            old = old_edit.text().strip()
            new = new_edit.text().strip()
            if old:
                pairs.append((old, new))
        return pairs

    

   


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
