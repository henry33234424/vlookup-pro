# VLookup Pro

智能表格匹配工具 — 导入两个 Excel 文件，自动进行**精确匹配 + AI 语义匹配**，结果展示在可编辑表格中，支持导出。

## 功能特点

- **精确匹配**：大小写不敏感，O(n) 哈希匹配
- **AI 语义匹配**：基于 [BAAI/bge-base-zh-v1.5](https://huggingface.co/BAAI/bge-base-zh-v1.5) 中文语义模型，自动识别含义相近但文字不同的项目
- **可编辑结果表**：支持单元格编辑、复制粘贴、撤销等操作
- **一键导出**：将结果（含手动编辑）导出为 Excel 文件
- **颜色标注**：绿色(精确匹配) / 橙色(模糊匹配) / 粉色(未匹配) / 灰色(B表未使用)

## 截图

| 导入界面 | 结果界面 |
|:---:|:---:|
| 选择两个 Excel 文件，设置相似度阈值 | 匹配结果按 A 表原始顺序展示 |

## 适用场景

- 发票与入库单核对
- 银行流水与账务凭证匹配
- 供应商名称归一化
- 客户名单去重与合并
- 任何需要对比两份 Excel 名单的场景

## 快速开始

### 方式一：下载 Windows EXE（推荐）

前往 [Releases](../../releases/latest) 页面，下载 `VLookupPro-Windows.zip`，解压后双击 `VLookupPro.exe` 即可运行。

AI 模型已内置，无需联网，开箱即用。

### 方式二：从源码运行

**环境要求**：Python 3.11+

```bash
# 克隆项目
git clone https://github.com/henry33234424/vlookup-pro.git
cd vlookup-pro

# 安装依赖
pip install -r requirements.txt

# 运行
python -m src.main
```

## 匹配流程

```
Excel A (主表)          Excel B (匹配表)
    │                        │
    └──────────┬─────────────┘
               │
        1. 精确匹配
        (大小写不敏感, hash O(n))
               │
        2. AI 语义匹配
        (仅对未匹配项, 余弦相似度)
               │
        3. 贪心一对一配对
        (按相似度降序, 超过阈值才匹配)
               │
           结果展示
    (按 A 表原始顺序排列)
```

## 技术栈

| 组件 | 选型 |
|------|------|
| UI 框架 | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| 可编辑表格 | [tksheet](https://github.com/ragardner/tksheet) |
| AI 模型 | [BAAI/bge-base-zh-v1.5](https://huggingface.co/BAAI/bge-base-zh-v1.5) |
| 向量计算 | [sentence-transformers](https://www.sbert.net/) |
| 数据处理 | pandas + openpyxl + numpy |

## 项目结构

```
vlookup_pro/
├── src/
│   ├── main.py              # 入口
│   ├── app.py               # 主窗口 + 页面切换
│   ├── ui/
│   │   ├── page_import.py   # 文件选择 + 阈值设置 + 开始匹配
│   │   └── page_result.py   # 结果表格 + 导出
│   └── core/
│       ├── matcher.py       # 匹配引擎（精确 + AI）
│       ├── ai_matcher.py    # AI 语义匹配（向量编码 + 贪心配对）
│       └── model_manager.py # 模型缓存管理
├── requirements.txt
└── .github/workflows/
    └── build.yml            # GitHub Actions 自动打包 Windows EXE
```

## 注意事项

- Windows EXE 已内置 AI 模型，下载即用，无需额外配置
- 从源码运行时，首次匹配会自动下载 AI 模型（约 400MB）
- 默认相似度阈值为 0.75，可根据实际需求调整（越高越严格）

## License

MIT
