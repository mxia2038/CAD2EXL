#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P&ID管道数据提取工具 - GUI版本
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import logging
import os
import sys
import json
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageTk

from extractor_core import (
    SUPPORTED_PROJECT_TYPES,
    PROJECT_FORMAT_EXAMPLES,
    extract_text_from_dwg,
    find_pipeline_numbers,
    load_medium_codes,
    parse_pipeline_number,
    create_excel_output,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PIDExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("P&ID管道数据提取工具")
        self.root.geometry("850x650")
        self.root.minsize(850, 650)

        self.dwg_file  = tk.StringVar()
        self.code_file = tk.StringVar()
        self.output_file = tk.StringVar()

        self.project_type = tk.StringVar(value=SUPPORTED_PROJECT_TYPES[0])

        self.code_file.set("test/code.xlsx")
        self.output_file.set("pipeline_data.xlsx")

        self.config_file = Path.home() / ".pid_extractor_config.json"
        self.load_recent_files()

        self.create_widgets()
        self.root.after(100, self.setup_drag_drop)

    # ------------------------------------------------------------------ layout

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.paned_window = tk.PanedWindow(self.root, orient=tk.VERTICAL, sashrelief=tk.RAISED)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # 上部：可滚动设置区
        self.setup_scroll_container = ttk.Frame(self.paned_window)
        self.setup_scroll_container.columnconfigure(0, weight=1)
        self.setup_scroll_container.rowconfigure(0, weight=1)

        self.setup_canvas = tk.Canvas(self.setup_scroll_container, highlightthickness=0)
        self.setup_scrollbar = ttk.Scrollbar(self.setup_scroll_container, orient=tk.VERTICAL,
                                             command=self.setup_canvas.yview)
        self.setup_inner_frame = ttk.Frame(self.setup_canvas)

        self.inner_window_id = self.setup_canvas.create_window((0, 0), window=self.setup_inner_frame, anchor="nw")
        self.setup_canvas.configure(yscrollcommand=self.setup_scrollbar.set)

        self.setup_inner_frame.bind("<Configure>",
            lambda e: self.setup_canvas.configure(scrollregion=self.setup_canvas.bbox("all")))
        self.setup_canvas.bind("<Configure>",
            lambda e: self.setup_canvas.itemconfigure(self.inner_window_id, width=e.width))

        self.setup_canvas.grid(row=0, column=0, sticky="nsew")
        self.setup_scrollbar.grid(row=0, column=1, sticky="ns")

        # 下部：固定结果区
        self.results_container = ttk.Frame(self.paned_window)

        self.paned_window.add(self.setup_scroll_container, minsize=450)
        self.paned_window.add(self.results_container, minsize=60)

        def _place_sash():
            total = self.paned_window.winfo_height()
            if total:
                self.paned_window.sash_place(0, int(total * 0.90), 0)
        self.root.after_idle(_place_sash)

        self.bind_mousewheel()
        self.create_setup_widgets()
        self.create_results_widgets()
        self.root.after(200, lambda: self.setup_canvas.yview_moveto(0.001))

    def bind_mousewheel(self):
        def on_mousewheel(event):
            self.setup_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.setup_canvas.bind('<Enter>',
            lambda e: self.setup_canvas.bind_all("<MouseWheel>", on_mousewheel))
        self.setup_canvas.bind('<Leave>',
            lambda e: self.setup_canvas.unbind_all("<MouseWheel>"))

    def create_setup_widgets(self):
        self.setup_inner_frame.columnconfigure(0, weight=1)

        self.setup_logo()

        ttk.Label(self.setup_inner_frame, text="P&ID管道数据提取工具",
                  font=("Microsoft YaHei", 14, "bold")).grid(row=1, column=0, pady=(0, 10))

        # 项目类型
        proj_section = ttk.LabelFrame(self.setup_inner_frame, text="🏗️ 项目类型", padding="8")
        proj_section.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        proj_section.columnconfigure(0, weight=1)

        proj_frame = ttk.Frame(proj_section)
        proj_frame.grid(row=0, column=0, sticky="ew")
        proj_frame.columnconfigure(1, weight=1)

        ttk.Label(proj_frame, text="选择项目编号标准:",
                  font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 10))

        cb = ttk.Combobox(proj_frame, textvariable=self.project_type,
                          values=SUPPORTED_PROJECT_TYPES, state="readonly", width=15)
        cb.grid(row=0, column=1, sticky="w", padx=(0, 10))

        self.format_example_label = ttk.Label(proj_frame, text="",
                                              font=("Microsoft YaHei", 9), foreground="gray")
        self.format_example_label.grid(row=0, column=2, sticky="w", padx=(15, 0))

        cb.bind('<<ComboboxSelected>>', self._on_project_changed)
        self.project_type.trace('w', self._on_project_changed)
        self._update_format_example()

        # 输入文件（两列）
        input_section = ttk.LabelFrame(self.setup_inner_frame, text="📁 输入文件", padding="8")
        input_section.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        input_section.columnconfigure(0, weight=1)
        input_section.columnconfigure(1, weight=1)

        self.dwg_drop_frame  = self._make_drop_zone(input_section, col=0,
            title="DWG源文件", hint="拖拽 .dwg 文件到下方", icon="📋", label="拖拽 DWG 文件",
            entry_var=self.dwg_file, browse_cmd=self.select_dwg_file)
        self.code_drop_frame = self._make_drop_zone(input_section, col=1,
            title="介质代码数据文件", hint="拖拽 .xlsx 文件到下方", icon="📊", label="拖拽 Excel 文件",
            entry_var=self.code_file, browse_cmd=self.select_code_file)

        # 输出设置
        out_section = ttk.LabelFrame(self.setup_inner_frame, text="💾 输出设置", padding="8")
        out_section.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        out_section.columnconfigure(0, weight=1)

        ttk.Label(out_section, text="输出文件路径",
                  font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 3))

        out_ctrl = ttk.Frame(out_section)
        out_ctrl.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        out_ctrl.columnconfigure(0, weight=1)

        ttk.Entry(out_ctrl, textvariable=self.output_file, width=60).grid(
            row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(out_ctrl, text="选择路径", command=self.select_output_file).grid(row=0, column=1)

        # 操作按钮
        ttk.Button(self.setup_inner_frame, text="🚀 开始提取数据",
                   command=self.start_extraction,
                   style="Accent.TButton").grid(row=5, column=0, pady=(10, 0), ipadx=20, ipady=5)

        # 进度 & 状态
        status_frame = ttk.Frame(self.setup_inner_frame)
        status_frame.grid(row=6, column=0, sticky="ew", pady=(8, 8))
        status_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.status_label = ttk.Label(status_frame, text="请选择DWG文件开始提取")
        self.status_label.grid(row=1, column=0)

    def _make_drop_zone(self, parent, col, title, hint, icon, label, entry_var, browse_cmd):
        """创建统一的拖拽区域，返回拖放框 Frame"""
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=col, sticky="nsew", padx=(0, 4) if col == 0 else (4, 0))
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text=title, font=("Microsoft YaHei", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 3))
        ttk.Label(frame, text=hint, font=("Microsoft YaHei", 8),
                  foreground="gray").grid(row=1, column=0, sticky="w", pady=(0, 5))

        drop_frame = tk.Frame(frame, relief="solid", borderwidth=2, bg="#f8f9fa", height=80)
        drop_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        drop_frame.pack_propagate(False)

        tk.Label(drop_frame, text=icon, font=("Microsoft YaHei", 14),
                 bg="#f8f9fa", fg="#6c757d").place(relx=0.5, rely=0.3, anchor="center")
        tk.Label(drop_frame, text=label, font=("Microsoft YaHei", 8),
                 bg="#f8f9fa", fg="#6c757d").place(relx=0.5, rely=0.7, anchor="center")

        ctrl = ttk.Frame(frame)
        ctrl.grid(row=3, column=0, sticky="ew")
        ctrl.columnconfigure(0, weight=1)

        ttk.Entry(ctrl, textvariable=entry_var, width=30).grid(
            row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(ctrl, text="浏览", command=browse_cmd).grid(row=0, column=1)

        return drop_frame

    def create_results_widgets(self):
        self.results_container.columnconfigure(0, weight=1)
        self.results_container.rowconfigure(1, weight=1)

        ttk.Label(self.results_container, text="📋 提取结果",
                  font=("Microsoft YaHei", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 8))

        result_frame = ttk.Frame(self.results_container)
        result_frame.grid(row=1, column=0, sticky="nsew")
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        self.result_text = tk.Text(result_frame, height=3, width=70, wrap="word",
                                   font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL,
                                  command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

    def setup_logo(self):
        try:
            base_path = Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent
            logo_path = base_path / "fig" / "logo.jpg"
            if logo_path.exists():
                img = Image.open(logo_path)
                w, h = img.size
                th = 50
                img = img.resize((int(th * w / h), th), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(img)
                tk.Label(self.setup_inner_frame, image=self.logo_photo).grid(
                    row=0, column=0, pady=(0, 6))
        except Exception as e:
            print(f"无法加载logo: {e}")

    # -------------------------------------------------------- project selector

    def _on_project_changed(self, *_):
        self._update_format_example()

    def _update_format_example(self):
        self.format_example_label.config(
            text=PROJECT_FORMAT_EXAMPLES.get(self.project_type.get(), ""))

    # ------------------------------------------------------ recent files / cfg

    def load_recent_files(self):
        self.recent_files = {'dwg': [], 'code': [], 'output': []}
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.recent_files = json.load(f).get('recent_files', self.recent_files)
        except Exception as e:
            print(f"无法加载配置文件: {e}")

    def save_recent_files(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({'recent_files': self.recent_files}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"无法保存配置文件: {e}")

    def add_recent_file(self, file_type, file_path):
        if file_path and file_path not in self.recent_files[file_type]:
            self.recent_files[file_type].insert(0, file_path)
            self.recent_files[file_type] = self.recent_files[file_type][:5]
            self.save_recent_files()

    # ------------------------------------------------------------ drag & drop

    def setup_drag_drop(self):
        try:
            from tkinterdnd2 import DND_FILES

            def make_handler(target_var, file_type, exts):
                def on_drop(event):
                    data = event.data.strip()
                    path = data[1:-1] if data.startswith('{') and data.endswith('}') else data
                    path = path.split('\n')[0].strip()
                    if path:
                        if any(path.lower().endswith(e) for e in exts):
                            target_var.set(path)
                            self.add_recent_file(file_type, path)
                            self._drop_feedback(event.widget, "success")
                        else:
                            self._drop_feedback(event.widget, "error")
                return on_drop

            for drop_frame, var, ftype, exts, browse in [
                (self.dwg_drop_frame,  self.dwg_file,  'dwg',  ['.dwg'],         self.select_dwg_file),
                (self.code_drop_frame, self.code_file, 'code', ['.xlsx', '.xls'], self.select_code_file),
            ]:
                drop_frame.drop_target_register(DND_FILES)
                drop_frame.dnd_bind('<<Drop>>',      make_handler(var, ftype, exts))
                drop_frame.dnd_bind('<<DragEnter>>', lambda e: self._drop_feedback(e.widget, "hover"))
                drop_frame.dnd_bind('<<DragLeave>>', lambda e: self._drop_feedback(e.widget, "normal"))
                drop_frame.bind('<Button-1>', lambda e, b=browse: b())

            print("拖拽功能已启用")
        except ImportError:
            print("拖拽功能需要安装 tkinterdnd2 库")

    def _drop_feedback(self, widget, state):
        styles = {
            "hover":   ("#e3f2fd", 3),
            "success": ("#e8f5e8", 3),
            "error":   ("#ffebee", 3),
            "normal":  ("#f8f9fa", 2),
        }
        bg, bw = styles.get(state, ("#f8f9fa", 2))
        try:
            widget.configure(bg=bg, borderwidth=bw)
            if state in ("success", "error"):
                self.root.after(1000, lambda: widget.configure(bg="#f8f9fa", borderwidth=2))
        except Exception:
            pass

    # --------------------------------------------------------- file dialogs

    def _initial_dir(self, key):
        recent = self.recent_files.get(key, [])
        return os.path.dirname(recent[0]) if recent else None

    def select_dwg_file(self):
        path = filedialog.askopenfilename(
            title="选择DWG文件",
            filetypes=[("DWG files", "*.dwg"), ("All files", "*.*")],
            initialdir=self._initial_dir('dwg'))
        if path:
            self.dwg_file.set(path)
            self.add_recent_file('dwg', path)

    def select_code_file(self):
        path = filedialog.askopenfilename(
            title="选择介质代码文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self._initial_dir('code'))
        if path:
            self.code_file.set(path)
            self.add_recent_file('code', path)

    def select_output_file(self):
        default = f"pipeline_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择数据保存位置和文件名",
            initialfile=default,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=self._initial_dir('output'))
        if path:
            self.output_file.set(path)
            self.add_recent_file('output', path)

    # ----------------------------------------------------------- extraction

    def log_message(self, message):
        """线程安全地向结果区追加日志"""
        ts = datetime.now().strftime('%H:%M:%S')
        self.root.after(0, lambda: (
            self.result_text.insert(tk.END, f"{ts} - {message}\n"),
            self.result_text.see(tk.END),
        ))

    def start_extraction(self):
        if not self.dwg_file.get():
            messagebox.showerror("错误", "请选择DWG文件")
            return
        if not self.code_file.get():
            messagebox.showerror("错误", "请选择介质代码文件")
            return
        if not self.output_file.get():
            messagebox.showerror("错误", "请选择输出文件")
            return

        self.progress.start()
        self.status_label.config(text="正在提取数据...")
        self.result_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self._run_extraction, daemon=True)
        thread.start()

    def _run_extraction(self):
        try:
            project_type = self.project_type.get()
            self.log_message(f"开始提取P&ID管道数据... (项目类型: {project_type})")

            text_entities = extract_text_from_dwg(self.dwg_file.get(), self.log_message)
            if not text_entities:
                self.log_message("未能提取到任何文本")
                self._finish(False)
                return

            self.log_message(f"提取了 {len(text_entities)} 个文本实体")

            pipeline_numbers = find_pipeline_numbers(text_entities, project_type, self.log_message)
            self.log_message(f"找到 {len(pipeline_numbers)} 个管道号")

            medium_codes = load_medium_codes(self.code_file.get())
            self.log_message(f"加载了 {len(medium_codes)} 个介质代码")

            pipeline_data = [
                parse_pipeline_number(pn, medium_codes, project_type)
                for pn in pipeline_numbers
            ]
            pipeline_data = [d for d in pipeline_data if d]
            self.log_message(f"成功解析 {len(pipeline_data)} 个管道号")

            df = create_excel_output(pipeline_data, self.output_file.get())

            for phase, count in df['相态'].value_counts().items():
                self.log_message(f"  {phase}: {count}个")

            self.log_message(f"提取完成！结果已保存到: {self.output_file.get()}")
            self._finish(True)
        except Exception as e:
            self.log_message(f"提取过程中发生错误: {e}")
            self._finish(False)

    def _finish(self, success):
        self.root.after(0, self.progress.stop)
        if success:
            self.root.after(0, lambda: self.status_label.config(text="提取完成！"))
            self.root.after(0, lambda: messagebox.showinfo("成功", "数据提取完成！"))
        else:
            self.root.after(0, lambda: self.status_label.config(text="提取失败"))
            self.root.after(0, lambda: messagebox.showerror("错误", "数据提取失败，请查看日志"))


def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        print("tkinterdnd2 不可用，拖拽功能将被禁用")
    PIDExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
