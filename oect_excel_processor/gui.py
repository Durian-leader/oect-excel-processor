#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
OECT Excel Processor - 图形用户界面
提供友好的可视化界面用于处理OECT性能测试后的Excel数据
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Optional
import queue

from .excel_processor import ExcelProcessor
from .batch_processor import BatchExcelProcessor


class ModernStyle:
    """现代化深色主题样式配置"""
    
    # 颜色配置
    BG_PRIMARY = "#1a1a2e"      # 主背景色
    BG_SECONDARY = "#16213e"    # 次要背景色
    BG_TERTIARY = "#0f3460"     # 第三背景色
    ACCENT = "#e94560"          # 强调色
    ACCENT_HOVER = "#ff6b6b"    # 悬停强调色
    TEXT_PRIMARY = "#ffffff"    # 主文字颜色
    TEXT_SECONDARY = "#a0a0a0"  # 次要文字颜色
    SUCCESS = "#00d26a"         # 成功颜色
    ERROR = "#ff4757"           # 错误颜色
    WARNING = "#ffa502"         # 警告颜色
    
    # 字体配置
    FONT_FAMILY = "Microsoft YaHei UI"
    FONT_SIZE_TITLE = 16
    FONT_SIZE_NORMAL = 10
    FONT_SIZE_SMALL = 9


class OECTProcessorGUI:
    """OECT Excel处理器图形界面"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OECT Excel 转 CSV 工具")
        self.root.geometry("700x600")
        self.root.minsize(600, 500)
        self.root.configure(bg=ModernStyle.BG_PRIMARY)
        
        # 状态变量
        self.selected_path = tk.StringVar(value="")
        self.is_batch_mode = tk.BooleanVar(value=False)
        self.transfer_enabled = tk.BooleanVar(value=True)
        self.transient_enabled = tk.BooleanVar(value=True)
        self.output_prefix = tk.StringVar(value="processed_")
        self.is_processing = False
        
        # 消息队列用于线程间通信
        self.msg_queue = queue.Queue()
        
        # 配置样式
        self._setup_styles()
        
        # 创建UI
        self._create_ui()
        
        # 启动消息处理
        self._process_queue()
    
    def _setup_styles(self):
        """设置ttk样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置进度条样式
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=ModernStyle.BG_SECONDARY,
            background=ModernStyle.ACCENT,
            lightcolor=ModernStyle.ACCENT,
            darkcolor=ModernStyle.ACCENT,
            bordercolor=ModernStyle.BG_TERTIARY,
            thickness=20
        )
        
        # 配置复选框样式
        style.configure(
            "Custom.TCheckbutton",
            background=ModernStyle.BG_PRIMARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL)
        )
        style.map("Custom.TCheckbutton",
                  background=[('active', ModernStyle.BG_PRIMARY)])
    
    def _create_ui(self):
        """创建用户界面"""
        # 主容器
        main_frame = tk.Frame(self.root, bg=ModernStyle.BG_PRIMARY)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        self._create_header(main_frame)
        
        # 文件选择区域
        self._create_file_section(main_frame)
        
        # 选项区域
        self._create_options_section(main_frame)
        
        # 处理按钮
        self._create_action_section(main_frame)
        
        # 日志区域
        self._create_log_section(main_frame)
    
    def _create_header(self, parent):
        """创建标题区域"""
        header_frame = tk.Frame(parent, bg=ModernStyle.BG_PRIMARY)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = tk.Label(
            header_frame,
            text="🔬 OECT Excel 转 CSV 工具",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_TITLE, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_PRIMARY
        )
        title_label.pack(side=tk.LEFT)
        
        subtitle_label = tk.Label(
            header_frame,
            text="LabExpress 数据处理",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_SMALL),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_PRIMARY
        )
        subtitle_label.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))
    
    def _create_file_section(self, parent):
        """创建文件选择区域"""
        file_frame = tk.Frame(parent, bg=ModernStyle.BG_SECONDARY, padx=15, pady=15)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 模式选择行
        mode_frame = tk.Frame(file_frame, bg=ModernStyle.BG_SECONDARY)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        mode_label = tk.Label(
            mode_frame,
            text="处理模式:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY
        )
        mode_label.pack(side=tk.LEFT)
        
        single_rb = tk.Radiobutton(
            mode_frame,
            text="单文件",
            variable=self.is_batch_mode,
            value=False,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            selectcolor=ModernStyle.BG_TERTIARY,
            activebackground=ModernStyle.BG_SECONDARY,
            activeforeground=ModernStyle.TEXT_PRIMARY,
            command=self._on_mode_change
        )
        single_rb.pack(side=tk.LEFT, padx=(15, 5))
        
        batch_rb = tk.Radiobutton(
            mode_frame,
            text="批量处理",
            variable=self.is_batch_mode,
            value=True,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            selectcolor=ModernStyle.BG_TERTIARY,
            activebackground=ModernStyle.BG_SECONDARY,
            activeforeground=ModernStyle.TEXT_PRIMARY,
            command=self._on_mode_change
        )
        batch_rb.pack(side=tk.LEFT, padx=5)
        
        # 文件选择行
        select_frame = tk.Frame(file_frame, bg=ModernStyle.BG_SECONDARY)
        select_frame.pack(fill=tk.X)
        
        self.select_btn = tk.Button(
            select_frame,
            text="📁 选择文件",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_TERTIARY,
            activebackground=ModernStyle.ACCENT,
            activeforeground=ModernStyle.TEXT_PRIMARY,
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2",
            command=self._select_path
        )
        self.select_btn.pack(side=tk.LEFT)
        
        self.path_label = tk.Label(
            select_frame,
            textvariable=self.selected_path,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_SMALL),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            anchor="w"
        )
        self.path_label.pack(side=tk.LEFT, padx=(15, 0), fill=tk.X, expand=True)
    
    def _create_options_section(self, parent):
        """创建选项区域"""
        options_frame = tk.Frame(parent, bg=ModernStyle.BG_SECONDARY, padx=15, pady=15)
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 工作表类型选择
        sheet_frame = tk.Frame(options_frame, bg=ModernStyle.BG_SECONDARY)
        sheet_frame.pack(fill=tk.X, pady=(0, 10))
        
        sheet_label = tk.Label(
            sheet_frame,
            text="工作表类型:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY
        )
        sheet_label.pack(side=tk.LEFT)
        
        transfer_cb = tk.Checkbutton(
            sheet_frame,
            text="Transfer",
            variable=self.transfer_enabled,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            selectcolor=ModernStyle.BG_TERTIARY,
            activebackground=ModernStyle.BG_SECONDARY,
            activeforeground=ModernStyle.TEXT_PRIMARY
        )
        transfer_cb.pack(side=tk.LEFT, padx=(15, 5))
        
        transient_cb = tk.Checkbutton(
            sheet_frame,
            text="Transient",
            variable=self.transient_enabled,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            selectcolor=ModernStyle.BG_TERTIARY,
            activebackground=ModernStyle.BG_SECONDARY,
            activeforeground=ModernStyle.TEXT_PRIMARY
        )
        transient_cb.pack(side=tk.LEFT, padx=5)
        
        # 输出前缀
        prefix_frame = tk.Frame(options_frame, bg=ModernStyle.BG_SECONDARY)
        prefix_frame.pack(fill=tk.X)
        
        prefix_label = tk.Label(
            prefix_frame,
            text="输出前缀:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY
        )
        prefix_label.pack(side=tk.LEFT)
        
        prefix_entry = tk.Entry(
            prefix_frame,
            textvariable=self.output_prefix,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_TERTIARY,
            insertbackground=ModernStyle.TEXT_PRIMARY,
            relief=tk.FLAT,
            width=25
        )
        prefix_entry.pack(side=tk.LEFT, padx=(15, 0), ipady=5)
    
    def _create_action_section(self, parent):
        """创建操作按钮区域"""
        action_frame = tk.Frame(parent, bg=ModernStyle.BG_PRIMARY)
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_btn = tk.Button(
            action_frame,
            text="⚡ 开始处理",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.ACCENT,
            activebackground=ModernStyle.ACCENT_HOVER,
            activeforeground=ModernStyle.TEXT_PRIMARY,
            relief=tk.FLAT,
            padx=30,
            pady=12,
            cursor="hand2",
            command=self._start_processing
        )
        self.process_btn.pack(side=tk.LEFT)
        
        # 进度条
        self.progress = ttk.Progressbar(
            action_frame,
            style="Custom.Horizontal.TProgressbar",
            mode='indeterminate',
            length=200
        )
        self.progress.pack(side=tk.LEFT, padx=(20, 0), fill=tk.X, expand=True)
    
    def _create_log_section(self, parent):
        """创建日志显示区域"""
        log_frame = tk.Frame(parent, bg=ModernStyle.BG_SECONDARY)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_header = tk.Label(
            log_frame,
            text="📋 处理日志",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_NORMAL, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            anchor="w"
        )
        log_header.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # 日志文本框
        log_container = tk.Frame(log_frame, bg=ModernStyle.BG_TERTIARY)
        log_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.log_text = tk.Text(
            log_container,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.FONT_SIZE_SMALL),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_TERTIARY,
            relief=tk.FLAT,
            wrap=tk.WORD,
            state=tk.DISABLED,
            padx=10,
            pady=10
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(log_container, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 配置日志标签颜色
        self.log_text.tag_configure("success", foreground=ModernStyle.SUCCESS)
        self.log_text.tag_configure("error", foreground=ModernStyle.ERROR)
        self.log_text.tag_configure("warning", foreground=ModernStyle.WARNING)
        self.log_text.tag_configure("info", foreground=ModernStyle.TEXT_SECONDARY)
    
    def _on_mode_change(self):
        """处理模式切换"""
        if self.is_batch_mode.get():
            self.select_btn.config(text="📁 选择文件夹")
        else:
            self.select_btn.config(text="📁 选择文件")
        self.selected_path.set("")
    
    def _select_path(self):
        """选择文件或文件夹"""
        if self.is_batch_mode.get():
            path = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        else:
            path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
            )
        
        if path:
            self.selected_path.set(path)
            self._log(f"已选择: {path}", "info")
    
    def _get_sheet_types(self) -> List[str]:
        """获取选中的工作表类型"""
        types = []
        if self.transfer_enabled.get():
            types.append("transfer")
        if self.transient_enabled.get():
            types.append("transient")
        return types
    
    def _log(self, message: str, tag: str = None):
        """添加日志消息"""
        self.log_text.config(state=tk.NORMAL)
        
        prefix = ""
        if tag == "success":
            prefix = "✓ "
        elif tag == "error":
            prefix = "✗ "
        elif tag == "warning":
            prefix = "⚠ "
        elif tag == "info":
            prefix = "ℹ "
        
        self.log_text.insert(tk.END, f"{prefix}{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _start_processing(self):
        """开始处理"""
        if self.is_processing:
            return
        
        path = self.selected_path.get()
        if not path:
            messagebox.showwarning("警告", "请先选择文件或文件夹!")
            return
        
        sheet_types = self._get_sheet_types()
        if not sheet_types:
            messagebox.showwarning("警告", "请至少选择一种工作表类型!")
            return
        
        # 清空日志并开始
        self._clear_log()
        self._log("开始处理...", "info")
        
        # 启动处理线程
        self.is_processing = True
        self.progress.start(10)
        self.process_btn.config(state=tk.DISABLED, text="处理中...")
        
        thread = threading.Thread(
            target=self._process_thread,
            args=(path, sheet_types, self.output_prefix.get()),
            daemon=True
        )
        thread.start()
    
    def _process_thread(self, path: str, sheet_types: List[str], prefix: str):
        """处理线程"""
        try:
            if self.is_batch_mode.get():
                self._process_batch(path, sheet_types, prefix)
            else:
                self._process_single(path, sheet_types, prefix)
            
            self.msg_queue.put(("done", None))
        except Exception as e:
            self.msg_queue.put(("error", str(e)))
    
    def _process_single(self, file_path: str, sheet_types: List[str], prefix: str):
        """处理单个文件"""
        self.msg_queue.put(("log", (f"处理文件: {os.path.basename(file_path)}", "info")))
        
        processor = ExcelProcessor(
            file_path=file_path,
            sheet_types=sheet_types,
            output_prefix=prefix
        )
        
        saved_files = processor.process_and_save()
        
        for f in saved_files:
            self.msg_queue.put(("log", (f"生成: {os.path.basename(f)}", "success")))
        
        self.msg_queue.put(("log", (f"共生成 {len(saved_files)} 个CSV文件", "success")))
    
    def _process_batch(self, directory: str, sheet_types: List[str], prefix: str):
        """批量处理文件"""
        processor = BatchExcelProcessor(
            directory=directory,
            file_pattern="*.xls",
            sheet_types=sheet_types,
            output_prefix=prefix
        )
        
        excel_files = processor.get_excel_files()
        self.msg_queue.put(("log", (f"找到 {len(excel_files)} 个Excel文件", "info")))
        
        if not excel_files:
            self.msg_queue.put(("log", ("未找到Excel文件", "warning")))
            return
        
        results = processor.process_all_files(use_multiprocessing=True)
        summary = processor.get_processing_summary(results)
        
        self.msg_queue.put(("log", (f"成功: {summary['successful_files']}, 失败: {summary['failed_files']}", 
                                    "success" if summary['failed_files'] == 0 else "warning")))
        self.msg_queue.put(("log", (f"共生成 {summary['total_csv_files']} 个CSV文件", "success")))
    
    def _process_queue(self):
        """处理消息队列"""
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()
                
                if msg_type == "log":
                    message, tag = data
                    self._log(message, tag)
                elif msg_type == "done":
                    self._processing_complete()
                elif msg_type == "error":
                    self._log(f"处理出错: {data}", "error")
                    self._processing_complete()
        except queue.Empty:
            pass
        
        self.root.after(100, self._process_queue)
    
    def _processing_complete(self):
        """处理完成"""
        self.is_processing = False
        self.progress.stop()
        self.process_btn.config(state=tk.NORMAL, text="⚡ 开始处理")
        self._log("处理完成!", "success")


def main():
    """GUI主入口"""
    root = tk.Tk()
    
    # 设置DPI感知（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = OECTProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
