"""主程序模块

此模块是Cadence BOM转换工具的入口点。
负责初始化用户界面并处理文件转换的线程管理。
"""

from typing import Optional
import threading
import tkinter as tk
from Format_data import convert_file
from UI import CadenceBOMConversionUI

def start_conversion(file_path: str, ui: CadenceBOMConversionUI) -> None:
    """启动文件转换过程
    
    在单独的线程中执行文件转换，以避免阻塞UI。
    
    Args:
        file_path: 要转换的文件路径
        ui: UI实例，用于更新转换进度和状态
    """
    def conversion_thread():
        try:
            ui.update_status("文件读取成功，开始转换...")
            result_df = convert_file(file_path)

            total_rows = len(result_df)
            for i, _ in enumerate(result_df.iterrows()):
                ui.update_progress((i + 1) / total_rows * 100)

            ui.update_status("转换完成，请选择保存路径...")
            ui.save_file(result_df)

        except Exception as e:
            ui.show_error(str(e))

    threading.Thread(target=conversion_thread).start()

def main() -> None:
    """程序入口点
    
    初始化主窗口和UI组件，启动事件循环。
    """
    root = tk.Tk()
    ui = CadenceBOMConversionUI(root, lambda file_path: start_conversion(file_path, ui))
    root.mainloop()

if __name__ == "__main__":
    main()