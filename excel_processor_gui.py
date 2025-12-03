import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime

# Import processors từ excel_processor.py
from excel_processor import SCTXProcessor, NTVTDDProcessor


class ExcelProcessorGUI:
    """Giao diện tkinter cho chương trình xử lý Excel"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Chương trình xử lý dữ liệu Excel")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Biến lưu trữ
        self.file_path = None
        self.processor_type = tk.StringVar(value="sctx")
        self.is_processing = False
        
        # Tạo giao diện
        self.create_widgets()
        
        # Center window
        self.center_window()
    
    def center_window(self):
        """Căn giữa cửa sổ trên màn hình"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Tạo các widget cho giao diện"""
        
        # Main frame với padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="CHƯƠNG TRÌNH XỬ LÝ DỮ LIỆU EXCEL",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(
            row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20)
        )
        
        # Frame cho radio buttons
        radio_frame = ttk.LabelFrame(main_frame, text="Chọn loại file Excel", padding="10")
        radio_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        radio_frame.columnconfigure(0, weight=1)
        
        # Radio buttons
        sctx_radio = ttk.Radiobutton(
            radio_frame,
            text="File loại SCTX (Mã phiếu: 02.O09.42.xxxx hoặc 03.O09.42.xxxx)",
            variable=self.processor_type,
            value="sctx"
        )
        sctx_radio.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        ntvtdd_radio = ttk.Radiobutton(
            radio_frame,
            text="File loại NTVTDD (Mã phiếu linh hoạt, có xử lý mã vật tư)",
            variable=self.processor_type,
            value="ntvtdd"
        )
        ntvtdd_radio.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Frame cho file selection
        file_frame = ttk.LabelFrame(main_frame, text="Chọn file", padding="10")
        file_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # File label
        ttk.Label(file_frame, text="File đã chọn:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.file_label = ttk.Label(file_frame, text="Chưa chọn file", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # Choose file button
        choose_btn = ttk.Button(
            file_frame,
            text="📁 Chọn File Excel",
            command=self.choose_file
        )
        choose_btn.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Process button
        self.process_btn = ttk.Button(
            main_frame,
            text="▶ Xử lý File",
            command=self.process_file,
            state=tk.DISABLED
        )
        self.process_btn.grid(row=4, column=0, pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.grid(row=5, column=0, pady=(0, 15))
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Trạng thái", padding="10")
        status_frame.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Status text area
        self.status_text = scrolledtext.ScrolledText(
            status_frame,
            height=12,
            width=70,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initial status message
        self.update_status("Sẵn sàng xử lý. Vui lòng chọn file Excel...\n")
    
    def choose_file(self):
        """Mở dialog để chọn file Excel"""
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename, foreground="black")
            self.process_btn.config(state=tk.NORMAL)
            self.update_status(f"✓ Đã chọn file: {filename}\n")
    
    def process_file(self):
        """Xử lý file trong thread riêng"""
        if not self.file_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel trước!")
            return
        
        if self.is_processing:
            messagebox.showinfo("Thông báo", "Đang xử lý file, vui lòng đợi...")
            return
        
        # Disable button và start progress
        self.process_btn.config(state=tk.DISABLED)
        self.progress.start(10)
        self.is_processing = True
        
        # Clear status
        self.status_text.delete(1.0, tk.END)
        self.update_status(f"Bắt đầu xử lý file: {os.path.basename(self.file_path)}\n")
        self.update_status(f"Loại xử lý: {self.processor_type.get().upper()}\n")
        self.update_status("-" * 60 + "\n")
        
        # Run processor in thread
        thread = threading.Thread(target=self.run_processor, daemon=True)
        thread.start()
    
    def run_processor(self):
        """Chạy processor tương ứng"""
        try:
            # Chọn processor
            if self.processor_type.get() == "sctx":
                self.update_status("Khởi tạo SCTX Processor...\n")
                processor = SCTXProcessor(self.file_path)
            else:
                self.update_status("Khởi tạo NTVTDD Processor...\n")
                processor = NTVTDDProcessor(self.file_path)
            
            # Đọc file
            self.update_status("Đang đọc file Excel...\n")
            if not processor.read_file():
                self.root.after(0, lambda: messagebox.showerror(
                    "Lỗi", "Không thể đọc file Excel!"
                ))
                return
            
            self.update_status("✓ Đọc file thành công!\n")
            
            # Xử lý dữ liệu
            self.update_status("Đang xử lý dữ liệu...\n")
            if not processor.process():
                self.root.after(0, lambda: messagebox.showerror(
                    "Lỗi", "Lỗi khi xử lý dữ liệu!"
                ))
                return
            
            self.update_status("✓ Xử lý dữ liệu thành công!\n")
            
            # Xuất file
            self.update_status("Đang xuất file kết quả...\n")
            if not processor.export():
                self.root.after(0, lambda: messagebox.showerror(
                    "Lỗi", "Lỗi khi xuất file!"
                ))
                return
            
            # Tạo tên file output
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'Ket_qua_xu_ly_{timestamp}.xlsx'
            
            self.update_status("✓ Xuất file thành công!\n")
            self.update_status("-" * 60 + "\n")
            self.update_status(f"✓ HOÀN THÀNH!\n")
            self.update_status(f"✓ File kết quả: {output_file}\n")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Thành công",
                f"Xử lý file thành công!\n\nFile kết quả: {output_file}"
            ))
            
        except Exception as e:
            self.update_status(f"\n✗ LỖI: {str(e)}\n")
            self.root.after(0, lambda: messagebox.showerror(
                "Lỗi",
                f"Đã xảy ra lỗi:\n{str(e)}"
            ))
        
        finally:
            # Stop progress và enable button
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.is_processing = False
    
    def update_status(self, message):
        """Cập nhật status text (thread-safe)"""
        def _update():
            self.status_text.insert(tk.END, message)
            self.status_text.see(tk.END)
        
        self.root.after(0, _update)


def main():
    """Hàm main chạy ứng dụng GUI"""
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
