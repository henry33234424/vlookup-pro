"""界面1：选文件 + 阈值 + 开始匹配"""

import os
import threading
import queue
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar


class PageImport(ctk.CTkFrame):
    def __init__(self, master, state, show_page):
        super().__init__(master)
        self.state = state
        self.show_page = show_page
        self._msg_queue = queue.Queue()
        self.configure(fg_color="transparent")

        self._last_dir = os.path.expanduser("~")

        self._build_ui()

    def _build_ui(self):
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=50)

        # 标题
        ctk.CTkLabel(
            content,
            text="VLookup Pro",
            font=ctk.CTkFont(size=28, weight="bold"),
        ).pack(pady=(20, 5))

        ctk.CTkLabel(
            content,
            text="导入两个 Excel 文件，自动进行精确匹配 + AI 语义匹配",
            font=ctk.CTkFont(size=14),
            text_color="gray",
        ).pack(pady=(0, 30))

        # --- Excel A ---
        file_a_frame = ctk.CTkFrame(content)
        file_a_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            file_a_frame, text="Excel A（主表）：",
            font=ctk.CTkFont(size=14), width=160, anchor="w",
        ).pack(side="left", padx=10)

        self._file_a_var = StringVar(value="")
        self._file_a_entry = ctk.CTkEntry(file_a_frame, width=500, textvariable=self._file_a_var)
        self._file_a_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(
            file_a_frame, text="选择文件", width=100,
            command=lambda: self._pick_file(self._file_a_var, "a"),
        ).pack(side="right", padx=10)

        # --- Excel B ---
        file_b_frame = ctk.CTkFrame(content)
        file_b_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            file_b_frame, text="Excel B（匹配表）：",
            font=ctk.CTkFont(size=14), width=160, anchor="w",
        ).pack(side="left", padx=10)

        self._file_b_var = StringVar(value="")
        self._file_b_entry = ctk.CTkEntry(file_b_frame, width=500, textvariable=self._file_b_var)
        self._file_b_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(
            file_b_frame, text="选择文件", width=100,
            command=lambda: self._pick_file(self._file_b_var, "b"),
        ).pack(side="right", padx=10)

        # --- 阈值 ---
        threshold_frame = ctk.CTkFrame(content)
        threshold_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            threshold_frame, text="相似度阈值：",
            font=ctk.CTkFont(size=14), width=160, anchor="w",
        ).pack(side="left", padx=10)

        self._threshold_var = StringVar(value="0.75")
        ctk.CTkEntry(
            threshold_frame, width=100, textvariable=self._threshold_var,
        ).pack(side="left", padx=5)

        ctk.CTkLabel(
            threshold_frame, text="AI 模糊匹配的最低相似度（0~1）",
            font=ctk.CTkFont(size=12), text_color="gray",
        ).pack(side="left", padx=10)

        # --- 开始匹配按钮 ---
        self._match_btn = ctk.CTkButton(
            content,
            text="开始匹配",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=45,
            command=self._on_start_match,
        )
        self._match_btn.pack(pady=(30, 10), fill="x")

        # --- 进度区域（初始隐藏）---
        self._progress_frame = ctk.CTkFrame(content)

        self._progress_label = ctk.CTkLabel(
            self._progress_frame, text="正在准备...",
            font=ctk.CTkFont(size=13),
        )
        self._progress_label.pack(pady=(10, 8))

        self._progress_bar = ctk.CTkProgressBar(self._progress_frame, mode="indeterminate")
        self._progress_bar.pack(fill="x", padx=20, pady=(0, 10))

    def _pick_file(self, string_var, key):
        """打开文件选择对话框"""
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir=self._last_dir,
        )
        if path:
            string_var.set(path)
            self.state[f"file_{key}"] = path
            self._last_dir = os.path.dirname(path)

    def _validate(self):
        """验证输入"""
        self.state["file_a"] = self._file_a_var.get().strip()
        self.state["file_b"] = self._file_b_var.get().strip()

        if not self.state["file_a"]:
            messagebox.showwarning("提示", "请选择 Excel A 文件")
            return False
        if not self.state["file_b"]:
            messagebox.showwarning("提示", "请选择 Excel B 文件")
            return False
        if not os.path.isfile(self.state["file_a"]):
            messagebox.showerror("错误", f"文件不存在: {self.state['file_a']}")
            return False
        if not os.path.isfile(self.state["file_b"]):
            messagebox.showerror("错误", f"文件不存在: {self.state['file_b']}")
            return False

        try:
            val = float(self._threshold_var.get())
            if not (0 <= val <= 1):
                raise ValueError
            self.state["threshold"] = val
        except ValueError:
            messagebox.showwarning("提示", "阈值必须是 0~1 之间的数字")
            return False

        return True

    def _on_start_match(self):
        """点击开始匹配"""
        if not self._validate():
            return

        self._match_btn.configure(state="disabled")
        self._progress_label.configure(text="正在准备...")
        self._progress_frame.pack(pady=(10, 0), fill="x")
        self._progress_bar.start()

        threading.Thread(target=self._run_match_thread, daemon=True).start()
        self._poll_queue()

    def _send_progress(self, message):
        """线程安全：把进度消息放入队列"""
        self._msg_queue.put(("progress", message))

    def _poll_queue(self):
        """主线程轮询队列"""
        try:
            while True:
                msg_type, payload = self._msg_queue.get_nowait()
                if msg_type == "progress":
                    self._progress_label.configure(text=payload)
                elif msg_type == "done":
                    self._on_match_complete()
                    return
                elif msg_type == "error":
                    self._on_match_error(payload)
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _run_match_thread(self):
        """后台线程执行匹配"""
        try:
            from ..core.matcher import run_match

            result = run_match(
                self.state["file_a"],
                self.state["file_b"],
                threshold=self.state["threshold"],
                progress_callback=self._send_progress,
            )
            self.state["result"] = result
            self._msg_queue.put(("done", None))
        except Exception as e:
            self._msg_queue.put(("error", str(e)))

    def _on_match_complete(self):
        """匹配完成"""
        self._progress_bar.stop()
        self._progress_frame.pack_forget()
        self._match_btn.configure(state="normal")
        self.show_page("result")

    def _on_match_error(self, error_msg):
        """匹配出错"""
        self._progress_bar.stop()
        self._progress_frame.pack_forget()
        self._match_btn.configure(state="normal")
        messagebox.showerror("匹配失败", f"匹配过程中出错:\n{error_msg}")
