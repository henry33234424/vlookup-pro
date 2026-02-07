"""主窗口 - 页面切换 + 共享状态"""

import customtkinter as ctk

from .ui.page_import import PageImport
from .ui.page_result import PageResult


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("VLookup Pro - 智能表格匹配工具")

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # 根据屏幕分辨率设置窗口大小
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = min(1000, int(screen_w * 0.75))
        win_h = min(700, int(screen_h * 0.75))
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.minsize(750, 500)

        # 共享状态（不能用 self.state，会覆盖 tkinter 内置的 state() 方法）
        self.shared_state = {
            "file_a": None,
            "file_b": None,
            "threshold": 0.75,
            "result": None,
        }

        # 页面容器：所有页面叠放在同一位置，用 lift() 切换
        self._container = ctk.CTkFrame(self, fg_color="transparent")
        self._container.pack(fill="both", expand=True)
        self._container.grid_rowconfigure(0, weight=1)
        self._container.grid_columnconfigure(0, weight=1)

        # 创建页面
        self.pages = {}
        self.pages["import"] = PageImport(self._container, self.shared_state, self.show_page)
        self.pages["import"].grid(row=0, column=0, sticky="nsew")

        # 显示初始页面
        self.current_page = None
        self.show_page("import")

    def show_page(self, page_name):
        """切换页面（lift 方式，无布局重算）"""
        # 延迟创建 result 页面
        if page_name == "result" and page_name not in self.pages:
            self.pages["result"] = PageResult(self._container, self.shared_state, self.show_page)
            self.pages["result"].grid(row=0, column=0, sticky="nsew")

        self.current_page = page_name
        page = self.pages[page_name]

        if hasattr(page, "on_show"):
            page.on_show()

        page.tkraise()
