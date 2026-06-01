# 环境备份说明

本工具包依赖于 `roo_mcp` Conda 环境。

## 核心参数
- **Python 版本**: 3.10
- **环境位置**: `C:\Users\98514\.conda\envs\roo_mcp\python.exe`

## 依赖安装

依赖清单以项目根目录 `requirements.txt` 为**唯一权威源**。初始化命令：

```bash
conda create -n roo_mcp python=3.10 -y
conda activate roo_mcp
pip install -r requirements.txt
```

当前锁定的直接依赖（精确版本见 `requirements.txt`）：

| 包名 | 用途 |
|---|---|
| `mcp` | FastMCP 框架 + stdio 传输 |
| `python-docx` | DOCX 读写 + OOXML 底层操作 |
| `markdown-it-py` | 专利 MD AST Token 流解析 |
| `pymupdf` | PDF 文本提取 + 图像渲染 |
| `lxml` | XSLT 管道 (OMML→MathML→LaTeX) |
| `olefile` | MathType OLE 二进制解包 |

## 可选依赖
```bash
pip install pywin32  # Windows: Word COM 自动化 (备用)
```

## 工具模块说明

| 模块 | 功能 | 关键依赖 |
|---|---|---|
| `tools/docx_to_md.py` | DOCX → Markdown (公式 → LaTeX) | python-docx, lxml, olefile |
| `tools/mtef_fast.py` | MTEF Track 1 (TeX 元数据提取) | (stdlib only) |
| `tools/mtef_parser.py` | MTEF Track 2 (递归解析) | (stdlib only) |
| `tools/md_to_docx.py` | Markdown → DOCX (学术排版) | python-docx |
| `tools/patent_md_to_docx.py` | 专利 Markdown → CNIPA DOCX | markdown-it-py, python-docx |
| `tools/patent_style_utils.py` | 专利排版引擎 | python-docx |
| `tools/pdf_to_text.py` | PDF → 文本提取 | pymupdf |
| `tools/pdf_to_images.py` | PDF → 图片渲染 | pymupdf |

## XSLT 资源

`tools/xslt/` 目录内含公式转换所需样式表：
- `omml2mml.xsl`：OMML→MathML (TEI Consortium, BSD/CC 许可)
- `mmltex.xsl` + 依赖：MathML→LaTeX (mathconverter, MIT 许可)
