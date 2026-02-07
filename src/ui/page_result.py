"""界面2：tksheet 结果表 + 导出"""

import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox

from tksheet import Sheet
import pandas as pd


# 行颜色定义
COLOR_EXACT = "#C8E6C9"      # 绿色 - 精确匹配
COLOR_FUZZY = "#FFE0B2"      # 橙色 - 模糊匹配
COLOR_UNMATCHED = "#FFCDD2"  # 粉色 - 未匹配
COLOR_B_UNUSED = "#E0E0E0"   # 灰色 - B表未使用
COLOR_SEPARATOR = "#BDBDBD"  # 分隔行


class PageResult(ctk.CTkFrame):
    def __init__(self, master, state, show_page):
        super().__init__(master)
        self.state = state
        self.show_page = show_page
        self.sheet = None

        # 平台相关字体和行高（参考 word_table_filler_v0.2）
        if sys.platform == 'win32':
            self._font_size = 18
            self._row_height = 44
            self._font_name = "Microsoft YaHei"
        else:
            self._font_size = 11
            self._row_height = 25
            self._font_name = "Arial"

        self._build_ui()

    def _build_ui(self):
        # 顶部工具栏
        toolbar = ctk.CTkFrame(self)
        toolbar.pack(fill="x", padx=10, pady=(10, 5))

        self._stats_label = ctk.CTkLabel(
            toolbar,
            text="",
            font=ctk.CTkFont(size=13),
        )
        self._stats_label.pack(side="left", padx=10)

        back_btn = ctk.CTkButton(
            toolbar,
            text="返回",
            width=80,
            command=lambda: self.show_page("import"),
        )
        back_btn.pack(side="right", padx=(5, 10))

        export_btn = ctk.CTkButton(
            toolbar,
            text="导出 Excel",
            width=120,
            command=self._on_export,
        )
        export_btn.pack(side="right", padx=5)

        # 颜色图例栏
        legend_frame = ctk.CTkFrame(self, fg_color="transparent")
        legend_frame.pack(fill="x", padx=15, pady=(0, 5))

        legends = [
            (COLOR_EXACT, "精确匹配"),
            (COLOR_FUZZY, "模糊匹配"),
            (COLOR_UNMATCHED, "未匹配"),
            (COLOR_B_UNUSED, "B表未使用"),
        ]
        for color, text in legends:
            chip = ctk.CTkFrame(legend_frame, fg_color="transparent")
            chip.pack(side="left", padx=(0, 18))

            swatch = ctk.CTkFrame(chip, width=16, height=16, fg_color=color, corner_radius=3)
            swatch.pack(side="left", padx=(0, 5))
            swatch.pack_propagate(False)

            ctk.CTkLabel(chip, text=text, font=ctk.CTkFont(size=12)).pack(side="left")

        # tksheet 表格容器
        self._sheet_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._sheet_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def on_show(self):
        """页面显示时刷新数据"""
        result = self.state.get("result")
        if result is None:
            return
        self._populate_sheet(result)

    def _populate_sheet(self, result):
        """填充 tksheet 数据"""
        # 如果已有 sheet 则销毁重建
        if self.sheet is not None:
            self.sheet.destroy()

        matches = result["matches"]
        unmatched_b = result["unmatched_b"]

        # 构建表格数据
        headers = ["A表项目", "B表匹配项", "相似度", "匹配状态"]
        data = []
        row_colors = []

        # 统计
        n_exact = 0
        n_fuzzy = 0
        n_unmatched = 0

        for a_text, b_text, sim, status in matches:
            sim_str = f"{sim:.3f}" if sim > 0 else ""
            data.append([a_text, b_text, sim_str, status])
            if status == "精确匹配":
                row_colors.append(COLOR_EXACT)
                n_exact += 1
            elif status == "模糊匹配":
                row_colors.append(COLOR_FUZZY)
                n_fuzzy += 1
            else:
                row_colors.append(COLOR_UNMATCHED)
                n_unmatched += 1

        # 分隔行
        if unmatched_b:
            data.append(["━" * 10, "━" * 10, "", ""])
            row_colors.append(COLOR_SEPARATOR)

            # B 表未使用项
            for b_text in unmatched_b:
                data.append(["", b_text, "", "B表未使用"])
                row_colors.append(COLOR_B_UNUSED)

        # 更新统计信息
        self._stats_label.configure(
            text=(
                f"A表: {len(result['a_items'])} 项  |  "
                f"B表: {len(result['b_items'])} 项  |  "
                f"精确: {n_exact}  模糊: {n_fuzzy}  "
                f"未匹配: {n_unmatched}  B表未使用: {len(unmatched_b)}"
            )
        )

        # 创建 tksheet（参考 word_table_filler_v0.2 的配置方式）
        self.sheet = Sheet(
            self._sheet_frame,
            data=data,
            headers=headers,
            show_row_index=True,
            row_index_width=50,
            font=(self._font_name, self._font_size, "normal"),
            header_font=(self._font_name, self._font_size, "bold"),
            default_row_height=self._row_height,
        )
        self.sheet.pack(fill="both", expand=True)

        # 平台相关列宽
        if sys.platform == 'win32':
            self.sheet.column_width(column=0, width=500)
            self.sheet.column_width(column=1, width=500)
            self.sheet.column_width(column=2, width=140)
            self.sheet.column_width(column=3, width=160)
        else:
            self.sheet.column_width(column=0, width=350)
            self.sheet.column_width(column=1, width=350)
            self.sheet.column_width(column=2, width=100)
            self.sheet.column_width(column=3, width=120)

        # 启用编辑和快捷键
        self.sheet.enable_bindings((
            "single_select",
            "row_select",
            "column_select",
            "drag_select",
            "select_all",
            "copy",
            "paste",
            "delete",
            "undo",
            "edit_cell",
            "column_width_resize",
            "row_height_resize",
            "arrowkeys",
        ))

        # 设置行颜色
        for row_idx, color in enumerate(row_colors):
            self.sheet.highlight_rows(rows=[row_idx], bg=color)

    def _on_export(self):
        """导出 Excel"""
        if self.sheet is None:
            messagebox.showwarning("提示", "没有数据可导出")
            return

        path = filedialog.asksaveasfilename(
            title="导出结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="vlookup_result.xlsx",
        )
        if not path:
            return

        try:
            # 从 sheet 读取当前数据（可能已被编辑）
            sheet_data = self.sheet.get_sheet_data()
            headers = self.sheet.headers()

            df = pd.DataFrame(sheet_data, columns=headers)
            df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("成功", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出时出错:\n{str(e)}")
