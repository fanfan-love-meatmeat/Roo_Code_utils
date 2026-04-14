import fitz  # PyMuPDF
import os
import sys

def extract_pdf_content(pdf_path: str, output_dir: str = None) -> str:
    """
    极简且省钱的 PDF 提取工具。
    1. 提取文本块（保留布局）
    2. 仅记录图片位置，不自动转换图片
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

    # 智能推导输出目录：默认 -> 输入文件同级的 temp/texts/
    if output_dir is None:
        base_dir = os.path.dirname(os.path.abspath(pdf_path))
        output_dir = os.path.join(base_dir, "temp", "texts")
        
    os.makedirs(output_dir, exist_ok=True)
        
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_text_path = os.path.join(output_dir, f"{file_name}_content.md")
    
    doc = fitz.open(pdf_path)
    md_content = [f"# {file_name} 提取报告\n"]
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        md_content.append(f"## Page {page_num + 1}\n")
        
        # 1. 提取文本块 (Blocks)
        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                md_content.append(text + "\n")
        
        # 2. 检查该页是否有图片（不提取，只记录）
        images = page.get_images()
        if images:
            md_content.append(f"\n> [视觉提醒] 本页包含 {len(images)} 张图片/图表。若文本信息不足，请调用 pdf_to_images 工具提取本页图片。\n")
        
        md_content.append("\n---\n")
    
    doc.close()
    
    with open(output_text_path, "w", encoding="utf-8") as f:
        f.writelines(md_content)
    
    print(f"[pdf_to_text] 提取完成: {output_text_path}")
    return output_text_path

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python utils/pdf_to_text.py <PDF路径> [输出目录]")
        sys.exit(1)
    
    pdf = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    extract_pdf_content(pdf, out)