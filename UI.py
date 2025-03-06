"""用户界面模块

此模块提供了Cadence BOM转换工具的图形用户界面实现。
包含文件选择、进度显示、状态更新等功能。
"""

from typing import Callable, Dict, Any, Optional
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pandas import DataFrame
from datetime import datetime

# 定义版本号和日期常量
VERSION = "1.0.3"
version_date = "2025-02-28"

class CadenceBOMConversionUI:
    """Cadence BOM转换工具的主界面类
    
    提供图形用户界面，包括文件选择、进度显示、状态更新等功能。
    
    Attributes:
        root (tk.Tk): 主窗口实例
        on_file_open (Callable[[str], None]): 文件打开回调函数
        status_label (tk.Label): 状态显示标签
        progress_bar (ttk.Progressbar): 进度条
        input_file_entry (tk.Entry): 文件路径输入框
        version_label (tk.Label): 版本信息显示标签
    """
    
    def __init__(self, root: tk.Tk, on_file_open: Callable[[str], None]):
        self.root = root
        self.on_file_open = on_file_open
        self.root.title("Cadence BOM转换")
        
        # 设置窗口大小和属性
        width = 500
        height = 250
        resizable = False
        
        # 设置支持的文件类型
        self.file_types = [
            {
                "description": "Excel Files",
                "extensions": ["*.xls", "*.xlsx"]
            },
            {
                "description": "HTML Files",
                "extensions": ["*.html", "*.htm"]
            }
        ]
        
        # 设置窗口大小
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(resizable, resizable)
        
        # 设置窗口居中显示
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # 禁用最大化按钮，但保留最小化功能
        if not resizable:
            self.root.maxsize(width, height)
            self.root.minsize(width, height)
        self.status_label = tk.Label(self.root, text="等待文件选择...", width=40)
        self.status_label.grid(row=0, column=0, columnspan=3, padx=10, pady=20, sticky='nsew')
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=3)
        self.root.grid_columnconfigure(2, weight=1)
        
        # 创建并放置进度条
        progress_label = tk.Label(self.root, text="转换进度:", width=10, anchor="w")
        progress_label.grid(row=1, column=0, padx=10, pady=20, sticky='w')
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", mode="determinate", length=300)
        self.progress_bar.grid(row=1, column=1, columnspan=2, padx=10, pady=20, sticky='ew')
        
        # 创建并放置输入文件路径显示框和"打开文件"按钮
        input_file_label = tk.Label(self.root, text="文件路径:", width=10, anchor="w")
        input_file_label.grid(row=2, column=0, padx=10, pady=20, sticky='w')
        self.input_file_entry = tk.Entry(self.root, state='readonly', width=50)
        self.input_file_entry.grid(row=2, column=1, padx=10, pady=20, sticky='ew')
        open_file_button = tk.Button(self.root, text="打开文件", command=self.open_file)
        open_file_button.grid(row=2, column=2, padx=10, pady=20, sticky='e')

        # 创建并放置版本号和日期标签
        version_label = tk.Label(self.root, text=f"版本 {VERSION} ({version_date})", fg="gray", font=("Arial", 8), width=30, anchor="e")
        version_label.grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="se")
        
    def open_file(self) -> None:
        if not self.file_types:
            messagebox.showerror("错误", "No file types specified in config.json")
            return

        file_path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[(item['description'], ' '.join(item['extensions'])) for item in self.file_types]
        )
        if file_path:
            self._update_file_entry(file_path)
            self.on_file_open(file_path)

    def _update_file_entry(self, file_path: str) -> None:
        self.input_file_entry.config(state='normal')
        self.input_file_entry.delete(0, tk.END)
        self.input_file_entry.insert(0, file_path)
        self.input_file_entry.config(state='readonly')

    def update_status(self, text: str) -> None:
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def update_progress(self, value: float) -> None:
        self.progress_bar['value'] = value
        self.root.update_idletasks()

    def save_file(self, result_df: DataFrame) -> None:
        output_file_path = filedialog.asksaveasfilename(
            title="选择文件保存路径",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not output_file_path:
            messagebox.showwarning("警告", "用户取消了文件保存。")
            self.update_status("等待文件选择...")
            return

        try:
            result_df.to_excel(output_file_path, index=False)
            messagebox.showinfo("成功", f"文件已成功保存到: {output_file_path}")
            self.update_status("文件保存成功")
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时发生错误: {e}")
            self.update_status("等待文件选择...")

    def show_error(self, message: str) -> None:
        messagebox.showerror("错误", message)
        self.update_status("等待文件选择...")