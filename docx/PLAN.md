# VLOOKUP Pro 实现计划

## Context

构建一个桌面工具，导入两个 Excel 文件，用第一列进行精确+AI语义匹配，结果展示在可编辑表格中，支持导出。复用 word_table_filler 项目的 UI 库（CustomTkinter + tksheet）和 AI 模型（BAAI/bge-base-zh-v1.5）。

## 技术栈

| 组件 | 选型 | 来源 |
|------|------|------|
| UI 框架 | customtkinter | 复用 v0.2 |
| 可编辑表格 | tksheet | 复用 v0.2 |
| AI 模型 | BAAI/bge-base-zh-v1.5 | 复用 v0.3 |
| 数据处理 | pandas + openpyxl | 复用 |

## 目录结构

```
vlookup_pro/
├── src/
│   ├── __init__.py
│   ├── main.py              # 入口
│   ├── app.py               # 主窗口 + 页面切换 + 共享状态
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── page_import.py   # 界面1：选文件 + 阈值 + 开始匹配
│   │   └── page_result.py   # 界面2：tksheet结果表 + 导出
│   └── core/
│       ├── __init__.py
│       ├── matcher.py       # 匹配引擎（编排精确+AI）
│       ├── ai_matcher.py    # AI语义匹配（改编自v0.3）
│       └── model_manager.py # 模型缓存管理（改编自v0.3）
└── requirements.txt
```

## 文件实现清单（共 10 个文件）

### 1. `requirements.txt`
```
customtkinter>=5.2.0
pandas>=2.0.0
openpyxl>=3.1.0
numpy>=1.24.0
tksheet>=7.0.0
sentence-transformers>=2.2.0
```

### 2. `src/main.py` — 入口
- 设置 sys.path，创建 App 实例并运行

### 3. `src/app.py` — 主窗口
- 窗口 1200x900，minsize 900x700
- 主题：system + blue
- **共享 state 字典**：存储文件路径、阈值、匹配结果
- **页面切换**：`pack_forget()` / `pack()` 切换两个页面
- `show_page('import')` / `show_page('result')`

### 4. `src/ui/page_import.py` — 界面1
- 标题 + 说明
- Excel A 文件选择器（`filedialog.askopenfilename`）
- Excel B 文件选择器
- 相似度阈值输入框（默认 0.85，验证 0~1）
- "开始匹配"按钮
- 点击后弹出**进度对话框**（`CTkToplevel` 模态），后台线程执行匹配
- 匹配完成后自动切换到界面2

### 5. `src/ui/page_result.py` — 界面2
- 顶部工具栏：统计信息 + "导出Excel"按钮 + "返回"按钮
- 颜色图例栏
- **tksheet 表格**：
  - 列：`A表项目 | B表匹配项 | 相似度 | 匹配状态`
  - 启用 edit_cell、copy、paste、undo、arrowkeys 等
  - 行颜色：绿(精确) / 橙(模糊) / 粉(未匹配) / 灰(B表未使用)
  - 分隔行区分 A 表结果和 B 表未使用项
- 导出：读取 sheet 数据 → pandas DataFrame → `.xlsx`

### 6. `src/core/matcher.py` — 匹配引擎
核心流程：
1. `pd.read_excel` 加载两表第一列
2. **精确匹配**：大小写不敏感 + strip，O(n) 哈希匹配，similarity=1.0
3. **AI 模糊匹配**：仅对未匹配项调用 ai_matcher
4. **贪心一对一**：按相似度降序，每项最多匹配一次
5. 返回 `{matches, a_items, b_items}`

### 7. `src/core/ai_matcher.py` — AI 语义匹配
- 改编自 `word_table_filler_v0.3/src/core/ai_matcher.py`
- 懒加载模型（线程安全，双重检查锁）
- `model.encode(texts, normalize_embeddings=True)` 生成向量
- `np.dot(a_vectors, b_vectors.T)` 计算余弦相似度矩阵
- 贪心一对一匹配算法

### 8. `src/core/model_manager.py` — 模型管理
- 直接改编自 `word_table_filler_v0.3/src/core/model_manager.py`
- 缓存目录改为 `VLookupPro`
- 检查多个可能的缓存位置

### 9-10. `__init__.py` 文件（空文件）

## 匹配逻辑详解

```
Excel A 第一列: [a1, a2, a3, ...]
Excel B 第一列: [b1, b2, b3, ...]

Step 1 - 精确匹配:
  a1 == b3 → (a1, b3, 1.000, 精确匹配)

Step 2 - AI匹配（仅未匹配项）:
  encode([a2, a3, ...]) → vectors_a
  encode([b1, b2, ...]) → vectors_b
  sim_matrix = vectors_a @ vectors_b.T
  贪心取最高分配对:
    a2 ↔ b1, score=0.92 → (a2, b1, 0.920, 模糊匹配)
    a3 ↔ b2, score=0.60 → score < 0.85 → (a3, '', 0.000, 未匹配)

Step 3 - 结果表:
  | A表项目 | B表匹配项 | 相似度 | 匹配状态 |
  |---------|----------|--------|---------|
  | a1      | b3       | 1.000  | 精确匹配 |
  | a2      | b1       | 0.920  | 模糊匹配 |
  | a3      |          | 0.000  | 未匹配   |
  | ━━━━━━━ | ━━━━━━━━ |        |         |
  |         | b2       |        | B表未使用 |
```

## 验证方式

1. `pip install -r requirements.txt` 安装依赖
2. `python -m src.main` 启动程序
3. 选择两个测试 Excel 文件
4. 验证精确匹配项显示绿色、相似度=1.000
5. 验证模糊匹配项显示橙色、相似度在阈值以上
6. 验证未匹配项显示粉色
7. 验证 B 表未使用项显示在灰色区域
8. 双击单元格验证可编辑
9. 点击导出验证生成正确 Excel
