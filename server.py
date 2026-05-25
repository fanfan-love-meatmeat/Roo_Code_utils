import sys
import os
# 锁定当前脚本所在的目录，确保能找到 tools 文件夹
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server.fastmcp import FastMCP

# 导入你写好的底层处理逻辑
from tools.md_to_docx import parse_md_to_docx
from tools.pdf_to_text import extract_pdf_content
from tools.pdf_to_images import pdf_to_images
from tools.docx_to_md import convert_docx_to_md as _convert_docx_to_md

# 初始化 MCP Server
mcp = FastMCP("Roo_Code_Utils")

@mcp.tool()
def convert_md_to_docx(md_path: str, docx_path: str = None) -> str:
    """
    将 Markdown 文件转换为 Word (DOCX) 格式。
    包含标准的中文宋体和英文 Times New Roman 格式排版，以及 Markdown 表格的转换。
    
    参数:
        md_path: 必需，Markdown 文件的绝对路径。
        docx_path: 可选，输出的 DOCX 文件绝对路径。如不填，默认在 md 同级目录生成。
    """
    try:
        out_path = parse_md_to_docx(md_path, docx_path)
        return f"转换成功！Word 文件已保存至: {out_path}"
    except Exception as e:
        return f"转换失败: {str(e)}"

@mcp.tool()
def extract_text_from_pdf(pdf_path: str, output_dir: str = None) -> str:
    """
    【省钱高性价比方案】从 PDF 中提取文本块，并保留基础排版。
    不会转换图片，仅在 Markdown 中记录图片位置，供后续按需读取。
    
    参数:
        pdf_path: 必需，PDF 文件的绝对路径。
        output_dir: 可选，输出目录路径。若不填，工具将自动在 PDF 同级目录的 temp/ 文件夹中生成。
    """
    try:
        out_path = extract_pdf_content(pdf_path, output_dir)
        return f"PDF 文本提取成功！Markdown 文件已保存至: {out_path}。请优先阅读此文件获取信息。"
    except Exception as e:
        return f"PDF 提取失败: {str(e)}"

@mcp.tool()
def convert_pdf_to_images_fallback(pdf_path: str, output_dir: str = None) -> str:
    """
    【高成本视觉方案】将 PDF 每一页转换为 PNG 图片。
    警告：仅在文本提取内容不完整、表格错位，或必须查看具体图表时才调用此工具。
    
    参数:
        pdf_path: 必需，PDF 文件的绝对路径。
        output_dir: 可选，输出图片的目录。若不填，工具将自动在 PDF 同级目录的 temp/ 文件夹中生成。
    """
    try:
        saved_images = pdf_to_images(pdf_path, output_dir, zoom=2.0)
        return f"PDF 已转换为 {len(saved_images)} 张图片，保存在: {os.path.dirname(saved_images[0])}"
    except Exception as e:
        return f"图片转换失败: {str(e)}"

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
    try:
        return _convert_docx_to_md(docx_path, output_md_path, image_dir)
    except Exception as e:
        return f"转换失败: {str(e)}"

if __name__ == "__main__":
    # 以 stdio 模式运行 MCP 服务，供 Roo Code 等客户端调用
    mcp.run(transport='stdio')