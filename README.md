# Roo_Code_utils MCP Server

基于 [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) 的文档处理工具集，专为 AI Agent 调用设计，覆盖 Word ↔ Markdown 互转、PDF 解析与渲染等学术/工程场景。

## 项目结构

```
Roo_Code_utils/
├── server.py              # MCP 服务入口（FastMCP）
├── tools/                   # 工具模块
│   ├── md_to_docx.py        # Markdown → DOCX（学术排版）
│   ├── docx_to_md.py        # DOCX → Markdown（公式 → LaTeX）★
│   ├── patent_md_to_docx.py # 专利 Markdown → CNIPA DOCX
│   ├── patent_style_utils.py# 专利排版引擎（字体/边距/段落构建）
│   ├── mtef_fast.py         # MTEF Track 1 解析器（TeX 元数据提取）
│   ├── mtef_parser.py       # MTEF Track 2 解析器（递归降级）
│   ├── pdf_to_text.py       # PDF → 文本提取（省 Token）
│   ├── pdf_to_images.py     # PDF → PNG 渲染（视觉后备）
│   └── xslt/                # XSLT 样式表（OMML → MathML → LaTeX）
├── CHANGELOG.md           # 更新日志
├── environment_backup.md  # 环境备份说明
├── server.py              # MCP 服务入口
└── README.md
```

## MCP 工具一览

本服务通过 MCP 协议向 Roo Code / Claude 等 AI 客户端暴露 5 个工具：

| 工具名 | 方向 | 说明 |
|---|---|---|
| `convert_docx_to_md` | DOCX → MD | Word 转 Markdown，MathType/OMML 公式 → LaTeX |
| `convert_md_to_docx` | MD → DOCX | Markdown 转 Word，中文/英文/表格学术排版 |
| `patent_md_to_docx` | 专利 MD → DOCX | 专利 Markdown 转 CNIPA 规范 Word |
| `extract_text_from_pdf` | PDF → MD | 文本块提取，省 Token（推荐首选） |
| `convert_pdf_to_images_fallback` | PDF → PNG | 逐页渲染图片（仅当文本提取不够用） |

---

## 输入防御层

所有 MCP 工具入口调用统一的 `_validate_input_file()` 校验函数，在底层引擎执行前拦截 Agent 路径幻觉：

| 校验项 | 说明 |
|---|---|
| 路径存在性 | `os.path.isfile()` 拒绝目录和不存在路径 |
| 扩展名白名单 | `convert_docx_to_md` 仅 `.docx`；`convert_md_to_docx` / `patent_md_to_docx` 仅 `.md` / `.markdown`；`extract_text_from_pdf` / `convert_pdf_to_images_fallback` 仅 `.pdf` |
| 文件大小上限 | 100MB，防止大文件导致 OOM |
| 错误传达 | 校验失败统一抛出 `ToolError`，附带自然语言纠错提示，供 AI Agent 自动修复参数后重试 |

---

## 1. convert_docx_to_md — DOCX → Markdown（公式 → LaTeX）★

将 Word 文档转换为纯 Markdown，公式转为 LaTeX 文字（非图片），适合无视觉 AI Agent 直接阅读。

### 公式处理策略

| 公式类型 | 转换路径 | 输出 |
|---|---|---|
| MathType/OLE | olefile 解包 `oleObject.bin` → MTEF Track 1（TeX 元数据） | `$...$` 或 `$$...$$` |
| MathType/OLE（兜底） | MTEF Track 2（递归解析 LINE/CHAR/TMPL） | 同上 |
| OMML 原生 | XSLT 管道（OMML → MathML → LaTeX） | 同上 |
| Inline/Display 判定 | MTEF v5 头部 Bit 0 位掩码 | 自动 |

### 架构亮点

- **零渲染依赖**：无需 MathType、Word、pix2tex、GPU，仅需 `olefile` (~100KB)
- **纯 Python MTEF**：直接读取二进制 `oleObject.bin`，彻底绕过 WMF→PNG→OCR 链路
- **双轨引擎**：Track 1（TeX 元数据）微秒级；Track 2（递归解析）单毫秒级，全量兜底
- **段落重建**：`<w:p>` → `\n\n` 逐段闭合
- **参考文献清洗**：双语锚点（`参考文献` / `References`）+ 尾部 `[N]` 独立换行
- **修订兼容**：接受 `w:ins`，跳过 `w:del`
- **表格公式降级**：块公式自动降为行内 `$...$`，维持 Markdown 表格边界
- **图片导出**：自动提取嵌入图片 → `figures/`，相对路径引用

### 输出结构

```
原名_md/
├── figures/       ← 所有图片
└── 原名.md        ← Markdown 正文（含内联 LaTeX）
```

### 调用示例

```json
{
  "convert_docx_to_md": {
    "docx_path": "D:/path/to/论文.docx"
  }
}
```

可选参数：

```json
{
  "convert_docx_to_md": {
    "docx_path": "D:/path/to/论文.docx",
    "output_md_path": "D:/output/论文.md",
    "image_dir": "D:/output/images"
  }
}
```

---

## 2. convert_md_to_docx — Markdown → DOCX

将 Markdown 文件转换为学术格式 Word 文档。

- 正文：中文宋体 + 英文 Times New Roman
- 标题：黑体，蓝色表头
- 表格：标准学术边距，蓝色表头
- 输出：默认在 `.md` 同级目录生成同名 `.docx`

```json
{
  "convert_md_to_docx": {
    "md_path": "D:/path/to/报告.md"
  }
}
```

---

## 3. patent_md_to_docx — 专利 Markdown → CNIPA 规范 DOCX

将专利 Markdown 转换为符合中国国家知识产权局 (CNIPA) 规范的 Word 文档。

- **页边距**：左 25mm / 上 25mm / 右 15mm / 下 15mm
- **正文**：宋体 12pt + Times New Roman，首行缩进 2 字符，行距 1.5
- **五大书标题**：楷体 18pt 居中，3 空格字距 + 下边框 + 自动分页
- **编号段落**：`[0001]` 编号加粗，正文标准缩进
- **权利要求**：有序列表，编号加粗，无首行缩进
- **公式占位**：居中纯黑文本保留 `$` 分隔符，兼容 MathType 宏（含定界符幂等性防御）
- **附图着陆**：自动追加"说明书附图"空白区供粘贴附图

```json
{
  "patent_md_to_docx": {
    "md_path": "D:/path/to/专利申请书.md"
  }
}
```

可选参数：

```json
{
  "patent_md_to_docx": {
    "md_path": "D:/path/to/专利申请书.md",
    "docx_path": "D:/output/专利申请书.docx"
  }
}
```

输出命名规则（未指定 `docx_path` 时自动生成）：`{MD文件名}_{年月日时分}.docx`，生成至 MD 上层的 `patent_docx/` 目录。

---

## 4. extract_text_from_pdf — PDF 文本提取（省钱方案）

从 PDF 中提取文本块，保留基础排版结构。不转换图片，仅在 Markdown 中标注图片位置供后续按需读取。

**适用场景**：AI Agent 阅读 PDF 内容、摘要、问答。极大节省 Token 消耗。

```json
{
  "extract_text_from_pdf": {
    "pdf_path": "D:/path/to/论文.pdf"
  }
}
```

可选参数：

```json
{
  "extract_text_from_pdf": {
    "pdf_path": "D:/path/to/论文.pdf",
    "output_dir": "D:/output"
  }
}
```

---

## 5. convert_pdf_to_images_fallback — PDF → 图片（视觉后备）

将 PDF 每一页渲染为高清 PNG 图片（2x 缩放）。

**适用场景**：仅当文本提取内容不完整、表格错位，或必须查看具体图表时使用。此方案成本较高，请优先使用文本提取。

```json
{
  "convert_pdf_to_images_fallback": {
    "pdf_path": "D:/path/to/图表.pdf"
  }
}
```

---

## Roo Code 配置

在 `.kilo.json` 或 Roo Code 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "Roo_Code_Utils": {
      "command": "C:/Users/98514/.conda/envs/roo_mcp/python.exe",
      "args": ["D:/FJL/Projects/Kilo_Code_Gobal_Settings/Roo_Code_utils/server.py"]
    }
  }
}
```

---

## 环境与依赖

```bash
# 创建环境
conda create -n roo_mcp python=3.10 -y
conda activate roo_mcp

# 安装依赖（requirements.txt 为唯一权威清单）
pip install -r requirements.txt
```

`requirements.txt` 内容：

```
mcp==1.27.0
python-docx==1.2.0
markdown-it-py==4.2.0
pymupdf==1.27.2.2
lxml==6.0.4
olefile==0.47
```

> 详细环境说明见 [`environment_backup.md`](environment_backup.md)

---

## 命令行调试

```bash
# DOCX → Markdown（含公式 LaTeX 转换）
python tools/docx_to_md.py "D:/path/to/论文.docx"
# 输出：D:/path/to/论文_md/（含 figures/ + 论文.md）

# Markdown → DOCX
python tools/md_to_docx.py "D:/path/to/报告.md"

# 专利 Markdown → CNIPA DOCX
python tools/patent_md_to_docx.py "D:/path/to/专利申请书.md"
```
