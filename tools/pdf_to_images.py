import sys
import os
import fitz  # PyMuPDF

def pdf_to_images(pdf_path: str, output_dir: str = None, zoom: float = 2.0) -> list:
    """
    将 PDF 转为 PNG 图片列表。
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

    # 智能推导输出目录：默认 -> 输入文件同级的 temp/images/<文件名>
    if output_dir is None:
        base_dir = os.path.dirname(os.path.abspath(pdf_path))
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.join(base_dir, "temp", "images", pdf_name)

    os.makedirs(output_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    print(f"[pdf_to_images] 共 {doc.page_count} 页: {pdf_path}")

    mat = fitz.Matrix(zoom, zoom)
    saved = []

    for i, page in enumerate(doc):
        out_path = os.path.join(output_dir, f"page_{i + 1:03d}.png")
        pix = page.get_pixmap(matrix=mat)
        pix.save(out_path)
        saved.append(out_path)

    doc.close()
    print(f"[pdf_to_images] 完成，共输出 {len(saved)} 张图片至 {output_dir}")
    return saved

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python utils/pdf_to_images.py <PDF路径> [输出目录] [缩放倍数]")
        sys.exit(1)

    pdf = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    z   = float(sys.argv[3]) if len(sys.argv) > 3 else 2.0
    pdf_to_images(pdf, out, z)