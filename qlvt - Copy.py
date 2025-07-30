import sys
import pythoncom
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QCheckBox, QHBoxLayout
)
from PyQt5.QtCore import Qt

class WordProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word File Processor")
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
                item.setCheckState(Qt.Unchecked)
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

    def modify_document(self, doc):
        try:
            # Lọc ra tất cả các bảng nằm ở trang đầu tiên (page 1)
            tables_on_first_page = [table for table in doc.Tables if table.Range.Information(3) == 1]
            if tables_on_first_page:
                # Chỉ lấy bảng CUỐI CÙNG ở trang đầu tiên (bảng ký tên)
                table = tables_on_first_page[-1]
                rows = table.Rows.Count
                if rows == 4:
                    # ⚠️ THÊM 3 dòng vào giữa dòng 2 và 3
                    for _ in range(3):
                        table.Rows.Add(BeforeRow=table.Rows(3))
                # ✅ Tiếp tục xử lý nội dung sau khi thêm dòng
                try:
                    table.Cell(1, 3).Range.Text = ""  # Xoá "NGƯỜI LẬP PHIẾU"
                    table.Cell(7, 3).Range.Text = ""  # Xoá tên "VÕ THANH ĐIỀN"
                    table.Cell(1, 3).Merge(table.Cell(1, 4))  # Gộp ô
                except:
                    pass
        except:
            pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WordProcessorApp()
    window.show()
    sys.exit(app.exec_())
