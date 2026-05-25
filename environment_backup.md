# 环境备份说明

本工具包依赖于 `roo_mcp` Conda 环境。

## 核心参数
- **Python 版本**: 3.10
- **环境位置**: `C:\Users\98514\.conda\envs\roo_mcp\python.exe`

## 依赖安装命令
```bash
conda create -n roo_mcp python=3.10 -y
conda activate roo_mcp
pip install mcp pymupdf python-docx lxml Pillow
pip install pix2tex torch  # DOCX→MD 公式 OCR (MathType→LaTeX)
```

## 可选依赖
```bash
pip install pywin32  # Windows: Word COM 自动化 (备用)
```

## 工具模块说明

| 模块 | 功能 | 关键依赖 |
|---|---|---|
| `tools/md_to_docx.py` | Markdown → DOCX | python-docx |
| `tools/docx_to_md.py` | DOCX → Markdown (MathType→LaTeX OCR) | python-docx, lxml, Pillow, pix2tex, torch (Windows) |
| `tools/pdf_to_text.py` | PDF → 文本提取 | pymupdf |
| `tools/pdf_to_images.py` | PDF → 图片渲染 | pymupdf |

## XSLT 资源

`tools/xslt/` 目录内含公式转换所需样式表：
- `omml2mml.xsl`：OMML→MathML (TEI Consortium, BSD/CC 许可)
- `mmltex.xsl` + 依赖：MathML→LaTeX (mathconverter, MIT 许可)