import sys
import os
# 锁定当前脚本所在的目录，确保能找到 tools 文件夹
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

# 导入你写好的底层处理逻辑
from tools.md_to_docx import parse_md_to_docx
from tools.patent_md_to_docx import build_patent_docx as _build_patent_docx
from tools.pdf_to_text import extract_pdf_content
from tools.pdf_to_images import pdf_to_images
from tools.docx_to_md import convert_docx_to_md as _convert_docx_to_md

# 初始化 MCP Server
# 注: mask_error_details 需独立安装 fastmcp 包 (pip install fastmcp)
#     当前环境使用 mcp.server.fastmcp 内置版本，暂不支持该参数
mcp = FastMCP("Roo_Code_Utils")

# ============================================================================
# 轻量输入防御层 — 拦截 Agent 路径幻觉，避免底层引擎崩溃
# ============================================================================

# 文件大小上限（100MB，防止 OOM）
_MAX_FILE_SIZE = 100 * 1024 * 1024

# 各工具允许的文件扩展名白名单
_ALLOWED_EXTENSIONS = {
    "convert_md_to_docx":           [".md", ".markdown"],
    "patent_md_to_docx":            [".md", ".markdown"],
    "extract_text_from_pdf":        [".pdf"],
    "convert_pdf_to_images_fallback": [".pdf"],
    "convert_docx_to_md":           [".docx"],
}


def _validate_input_file(file_path: str, tool_name: str) -> None:
    """
    在文件 I/O 前进行轻量输入防御。任何校验失败均 raise ToolError，
    提供可指导 Agent 纠错的自然语言消息。

    Args:
        file_path: 用户/Agent 传入的文件路径
        tool_name: 调用方工具名（用于查找扩展名白名单）

    Raises:
        ToolError: 路径为空、文件不存在、扩展名不合法、文件过大
    """
    # Step 1: 空值校验
    if not file_path or not isinstance(file_path, str):
        raise ToolError(
            "未提供有效的文件路径。请指定一个本地文件的绝对路径，"
            "例如 D:/documents/论文.docx"
        )

    # Step 2: 路径标准化（展开 ~ 和相对路径）
    expanded = os.path.abspath(os.path.expanduser(file_path))

    # Step 3: 文件存在性校验（isfile 比 exists 更严格，拒绝目录路径）
    if not os.path.isfile(expanded):
        raise ToolError(
            f"文件不存在或不是有效文件: {file_path}。"
            f"请检查路径拼写是否正确，并确认文件位于本地磁盘。"
        )

    # Step 4: 扩展名白名单校验（大小写不敏感）
    allowed_exts = _ALLOWED_EXTENSIONS.get(tool_name, [])
    if allowed_exts:
        actual_ext = os.path.splitext(expanded)[1].lower()
        if actual_ext not in allowed_exts:
            allowed_str = " / ".join(allowed_exts)
            raise ToolError(
                f"不支持的文件类型: {actual_ext}。"
                f"该工具仅接受 {allowed_str} 格式的文件，请传入正确的文件路径。"
            )

    # Step 5: 文件大小上限
    file_size = os.path.getsize(expanded)
    if file_size > _MAX_FILE_SIZE:
        size_mb = file_size / (1024 * 1024)
        limit_mb = _MAX_FILE_SIZE / (1024 * 1024)
        raise ToolError(
            f"文件过大: {size_mb:.1f}MB（限制 {limit_mb:.0f}MB）。"
            f"请使用较小的文件或拆分为多个部分。"
        )


@mcp.tool()
def convert_md_to_docx(md_path: str, docx_path: str = None) -> str:
    """
    将 Markdown 文件转换为 Word (DOCX) 格式。
    包含标准的中文宋体和英文 Times New Roman 格式排版，以及 Markdown 表格的转换。
    
    参数:
        md_path: 必需，Markdown 文件的绝对路径。
        docx_path: 可选，输出的 DOCX 文件绝对路径。如不填，默认在 md 同级目录生成。
    """
    _validate_input_file(md_path, "convert_md_to_docx")
    try:
        out_path = parse_md_to_docx(md_path, docx_path)
        return f"转换成功！Word 文件已保存至: {out_path}"
    except FileNotFoundError as e:
        raise ToolError(f"文件未找到: {e}")
    except PermissionError as e:
        raise ToolError(f"文件访问被拒绝: {e}")
    except Exception:
        raise

@mcp.tool()
def patent_md_to_docx(md_path: str, docx_path: str = None,
                       version: str = None) -> str:
    """
    将专利 Markdown (含 heading 层级、段落编号 [0001]、权利要求、公式占位) 转换为符合 CNIPA 规范的 DOCX。
    自动设置页边距 (左25/上25/右15/下15mm)、宋体+Times New Roman、首行缩进 2 字符。
    公式以纯黑文本保留 $ 分隔符，兼容 MathType 宏捕获。
    
    参数:
        md_path:  必需，专利 Markdown 文件的绝对路径。
        docx_path: 可选，输出的 DOCX 文件绝对路径（含 `_{时间戳}_v{版本}` 命名）。
                   如不填，自动生成为 `{原名}_{YYYYMMDD_HHMM}_v{版本}.docx`。
        version:   可选，版本号，如 "v1.0", "v1.1", "v2.0"。
                   如不填，自动从 MD 文件名提取版本号（如 `_v2.4.md` → `v2.4`）。
                   若 MD 文件名也无版本号，默认 `v1.0`。
    
    命名规则:
        `{原名}_{YYYYMMDD_HHMM}_v{大版本}.{小版本}.docx`
        版本号与输入 MD 文件名中的版本号一致，不自动递增。
    """
    import os, re, glob, datetime
    _validate_input_file(md_path, "patent_md_to_docx")
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            md_text = f.read()
        doc = _build_patent_docx(md_text)

        filename = os.path.basename(md_path)
        now = datetime.datetime.now().strftime('%Y%m%d_%H%M')

        # 从 MD 文件名提取版本号 (如 _v2.4.md → v2.4)
        # 同时剥离版本后缀得到干净的文件基名
        base_name = os.path.splitext(filename)[0]
        if version is None:
            m = re.search(r'_v(\d+)\.(\d+)\.md$', filename)
            if m:
                version = f'v{m.group(1)}.{m.group(2)}'
                base_name = filename[:m.start()]  # 切掉 _v2.4.md
            else:
                version = 'v1.0'

        if docx_path is None:
            out_dir = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                'patent_docx'
            )
            if not os.path.isdir(out_dir):
                out_dir = os.path.dirname(os.path.abspath(md_path))

            docx_path = os.path.join(
                out_dir,
                f'{base_name}_{now}_{version}.docx'
            )

        os.makedirs(os.path.dirname(os.path.abspath(docx_path)), exist_ok=True)
        doc.save(docx_path)
        return f'转换成功！专利 DOCX 已保存至: {docx_path}'

    except FileNotFoundError as e:
        raise ToolError(f"文件未找到: {e}")
    except PermissionError as e:
        raise ToolError(f"文件访问被拒绝: {e}")
    except Exception as e:
        raise ToolError(f"转换失败: {e}")

@mcp.tool()
def extract_text_from_pdf(pdf_path: str, output_dir: str = None) -> str:
    """
    【省钱高性价比方案】从 PDF 中提取文本块，并保留基础排版。
    不会转换图片，仅在 Markdown 中记录图片位置，供后续按需读取。
    
    参数:
        pdf_path: 必需，PDF 文件的绝对路径。
        output_dir: 可选，输出目录路径。若不填，工具将自动在 PDF 同级目录的 temp/ 文件夹中生成。
    """
    _validate_input_file(pdf_path, "extract_text_from_pdf")
    try:
        out_path = extract_pdf_content(pdf_path, output_dir)
        return f"PDF 文本提取成功！Markdown 文件已保存至: {out_path}。请优先阅读此文件获取信息。"
    except FileNotFoundError as e:
        raise ToolError(f"文件未找到: {e}")
    except PermissionError as e:
        raise ToolError(f"文件访问被拒绝: {e}")
    except Exception:
        raise

@mcp.tool()
def convert_pdf_to_images_fallback(pdf_path: str, output_dir: str = None) -> str:
    """
    【高成本视觉方案】将 PDF 每一页转换为 PNG 图片。
    警告：仅在文本提取内容不完整、表格错位，或必须查看具体图表时才调用此工具。
    
    参数:
        pdf_path: 必需，PDF 文件的绝对路径。
        output_dir: 可选，输出图片的目录。若不填，工具将自动在 PDF 同级目录的 temp/ 文件夹中生成。
    """
    _validate_input_file(pdf_path, "convert_pdf_to_images_fallback")
    try:
        saved_images = pdf_to_images(pdf_path, output_dir, zoom=2.0)
        return f"PDF 已转换为 {len(saved_images)} 张图片，保存在: {os.path.dirname(saved_images[0])}"
    except FileNotFoundError as e:
        raise ToolError(f"文件未找到: {e}")
    except PermissionError as e:
        raise ToolError(f"文件访问被拒绝: {e}")
    except Exception:
        raise

@mcp.tool()
def convert_docx_to_md(docx_path: str, output_md_path: str = None,
                       image_dir: str = None) -> str:
    """
    将 DOCX 文件转换为 Markdown，公式转换为 LaTeX 文字（非图片引用），适合 Agent 读取。

    公式处理策略：
    - MathType/OLE 公式 → olefile 解包 oleObject.bin → MTEF Track 1 (TeX 元数据) / Track 2 (递归解析) → $$LaTeX$$
    - OMML 原生公式 → XSLT 管道 (OMML→MathML→LaTeX) → $$LaTeX$$
    - 行内/块级自动识别：MTEF 头部 Bit 0 位掩码判定 Inline / Display

    其他特性：
    - 普通图片：导出到 {原名}_md/figures/，Markdown 中以相对路径引用
    - 表格内公式：自动降级为行内公式 $...$ 维持表格语法边界
    - 修订模式兼容：接受 w:ins 插入、跳过 w:del 删除
    - mc:AlternateContent 兼容：取 Choice 子节点
    - 参考文献清洗：双语锚点定位 (参考文献/References) + 尾部 [N] 独立换行

    参数:
        docx_path:       必需，DOCX 文件的绝对路径。
        output_md_path:  可选，输出 Markdown 文件路径。默认在 raw_md/{原名}_md/ 下生成。
        image_dir:       可选，图片导出目录。默认 {原名}_md/。
    """
    _validate_input_file(docx_path, "convert_docx_to_md")
    try:
        return _convert_docx_to_md(docx_path, output_md_path, image_dir)
    except FileNotFoundError as e:
        raise ToolError(f"文件未找到: {e}")
    except PermissionError as e:
        raise ToolError(f"文件访问被拒绝: {e}")
    except Exception:
        raise

if __name__ == "__main__":
    # 以 stdio 模式运行 MCP 服务，供 Roo Code 等客户端调用
    mcp.run(transport='stdio')