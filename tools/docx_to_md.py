# 文件用途：DOCX → Markdown 单向转换核心模块
# 上游：MCP server.py 调用 convert_docx_to_md()
# 下游：输出 .md 文件 + 图片目录
# 核心路径：python-docx 加载 → 单轨 XML 遍历 → (OMML: XSLT管道 / MathType: MTEF 双轨解析) → LaTeX → Markdown

import os
import re
import zipfile
import logging
import olefile
from io import BytesIO

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table as DocxTable

from lxml import etree

from tools.mtef_parser import mtef_to_latex
from tools.mtef_fast import is_inline_equation

logging.basicConfig(level=logging.INFO, format="[docx_to_md] %(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ============================================================================
# 全局常量
# ============================================================================

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
_XSLT_DIR = os.path.join(_MODULE_DIR, "xslt")

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

_OMML2MML_XSL = os.path.join(_XSLT_DIR, "omml2mml.xsl")
_MMLTEX_XSL = os.path.join(_XSLT_DIR, "mmltex.xsl")


# ============================================================================
# XSLT 管道（OMML → MathML → LaTeX）
# ============================================================================

_xslt_omml2mml = None
_xslt_mml2tex = None

def _get_omml_transformer():
    global _xslt_omml2mml
    if _xslt_omml2mml is None:
        if not os.path.exists(_OMML2MML_XSL):
            raise FileNotFoundError(f"XSLT 缺失: {_OMML2MML_XSL}")
        _xslt_omml2mml = etree.XSLT(etree.parse(_OMML2MML_XSL))
    return _xslt_omml2mml

def _get_mml2tex_transformer():
    global _xslt_mml2tex
    if _xslt_mml2tex is None:
        if not os.path.exists(_MMLTEX_XSL):
            raise FileNotFoundError(f"XSLT 缺失: {_MMLTEX_XSL}")
        _xslt_mml2tex = etree.XSLT(etree.parse(_MMLTEX_XSL))
    return _xslt_mml2tex

def omml_to_latex(omml_element) -> str:
    try:
        transformer = _get_omml_transformer()
        mathml_tree = transformer(omml_element)
        root = mathml_tree.getroot()
        if root is None or not len(root):
            return "[公式: OMML 转换失败]"
        tex_result = _get_mml2tex_transformer()(mathml_tree)
        latex_str = str(tex_result).strip()
        latex_str = re.sub(r"^\\[\s*", "", latex_str)
        latex_str = re.sub(r"\\]\s*$", "", latex_str)
        return latex_str.strip()
    except Exception as e:
        log.warning(f"OMML 转换异常: {e}")
        return "[公式: OMML 转换异常]"


# ============================================================================
# 图片提取与关系映射
# ============================================================================

def _build_formula_map(doc: Document, docx_path: str) -> dict:
    """
    从 word/embeddings/oleObject*.bin 提取 MathType 公式并转换为 LaTeX。
    返回 {rId_ole: {'latex': str, 'is_inline': bool}} 映射。
    """
    formula_map = {}

    # Step 1: 构建 oleObject rId → ZIP 条目名 映射
    ole_rels = {}
    for rId_val, rel in doc.part.rels.items():
        if "oleObject" in rel.reltype:
            target = rel.target_ref
            if target:
                ole_rels[rId_val] = os.path.basename(target)

    if not ole_rels:
        return formula_map

    # Step 2: 提取并解析每个 oleObject.bin
    with zipfile.ZipFile(docx_path, "r") as zf:
        for rId_ole, bin_name in ole_rels.items():
            entry_path = f"word/embeddings/{bin_name}"
            if entry_path not in zf.namelist():
                continue

            try:
                with zf.open(entry_path) as f:
                    if not olefile.isOleFile(f):
                        continue
                    ole = olefile.OleFileIO(f)
                    if not ole.exists("Equation Native"):
                        continue
                    raw = ole.openstream("Equation Native").read()
                    mtef_data = raw[28:]       # 跳过 28B OLE 头
                    latex = mtef_to_latex(mtef_data)
                    inline = is_inline_equation(mtef_data)
                    formula_map[rId_ole] = {'latex': latex, 'is_inline': inline}
            except Exception as e:
                log.warning(f"公式解析失败 ({bin_name}): {e}")
                formula_map[rId_ole] = {'latex': '[公式: 解析失败]', 'is_inline': True}
    return formula_map


def _build_image_map(doc: Document, docx_path: str, image_dir: str) -> dict:
    """
    构建 {rId: disk_path} 映射。所有图片输出到 figures/ 子目录。
    """
    figures_dir = os.path.join(image_dir, "figures")
    os.makedirs(figures_dir, exist_ok=True)
    image_map = {}

    # Step 1: 从 rels 构建 rId → 文件名映射（只取 image 类型）
    rel_to_file = {}
    for rId_val, rel in doc.part.rels.items():
        if "image" not in rel.reltype:
            continue
        target = rel.target_ref
        if target and target != "NULL":
            rel_to_file[rId_val] = os.path.basename(target)

    if not rel_to_file:
        return image_map

    # Step 2: 从 ZIP 提取图片（全部到 image_dir 扁平目录）
    with zipfile.ZipFile(docx_path, "r") as zf:
        media_files = [n for n in zf.namelist() if n.startswith("word/media/")]
        for media_path in media_files:
            media_filename = os.path.basename(media_path)
            matched_rid = None
            for rid, fname in rel_to_file.items():
                if fname == media_filename:
                    matched_rid = rid
                    break
            if matched_rid is None:
                matched_rid = media_filename

            ext = os.path.splitext(media_filename)[1].lower()

            # 跳过 WMF/EMF — MathType 公式预览图，公式已通过 MTEF 提取为 LaTeX
            if ext in (".wmf", ".emf"):
                continue

            safe_name = f"{matched_rid}{ext}"
            dest_path = os.path.join(figures_dir, safe_name)
            counter = 1
            while os.path.exists(dest_path):
                safe_name = f"{matched_rid}_{counter}{ext}"
                dest_path = os.path.join(image_dir, safe_name)
                counter += 1

            with zf.open(media_path) as src:
                with open(dest_path, "wb") as dst:
                    dst.write(src.read())

            with zf.open(media_path) as src:
                with open(dest_path, "wb") as dst:
                    dst.write(src.read())

            image_map[matched_rid] = dest_path

    return image_map


def _split_image_dirs(image_map: dict, image_dir: str, formula_rids: set,
                      md_base_dir: str = "") -> tuple[dict, dict]:
    """
    Phase 2 之后执行真分流：将文件移动到 figures/ 和 formulas/ 子目录。
    返回 (更新后的 image_map, {old_rel_path: new_rel_path}) 用于更新 MD 引用。
    """
    figures_dir = os.path.join(image_dir, "figures")
    formulas_dir = os.path.join(image_dir, "formulas")
    os.makedirs(figures_dir, exist_ok=True)
    os.makedirs(formulas_dir, exist_ok=True)

    new_map = {}
    path_updates = {}  # old_rel_path → new_rel_path

    for rid, path in image_map.items():
        if rid == "_base_dir":
            new_map[rid] = path
            continue

        fname = os.path.basename(path)
        if rid in formula_rids:
            dest_dir = formulas_dir
        else:
            dest_dir = figures_dir

        new_path = os.path.join(dest_dir, fname)
        if os.path.abspath(path) != os.path.abspath(new_path):
            old_rel = _to_relative_path(path, md_base_dir) if md_base_dir else path
            try:
                os.rename(path, new_path)
            except OSError:
                import shutil
                shutil.move(path, new_path)
            new_rel = _to_relative_path(new_path, md_base_dir) if md_base_dir else new_path
            path_updates[old_rel] = new_rel
        new_map[rid] = new_path

    return new_map, path_updates


def _cleanup_stale_files(image_dir: str):
    """删除分流后滞留在根目录的 WMF/PNG 孤立文件。"""
    for fname in os.listdir(image_dir):
        path = os.path.join(image_dir, fname)
        if os.path.isfile(path):
            ext = os.path.splitext(fname)[1].lower()
            if ext in (".wmf", ".emf", ".png", ".jpg", ".gif", ".bmp"):
                os.remove(path)


# ============================================================================
# XML 节点提取工具
# ============================================================================

def _extract_blip_rid(drawing_element) -> str | None:
    blips = drawing_element.findall(".//a:blip", NSMAP)
    for blip in blips:
        rid = blip.get(qn("r:embed"))
        if rid:
            return rid
    return None

def _extract_ole_image_rid(w_object_element) -> str | None:
    imagedata = w_object_element.find(".//v:imagedata", NSMAP)
    if imagedata is not None:
        rid = imagedata.get(qn("r:id"))
        if rid:
            return rid
    ole = w_object_element.find("o:OLEObject", NSMAP)
    if ole is not None:
        rid = ole.get(qn("r:id"))
        if rid:
            return rid
    return None


def _extract_ole_object_rid(w_object_element) -> str | None:
    """提取 OLE 对象的 oleObject rId (用于 formula_map 查询)"""
    ole = w_object_element.find("o:OLEObject", NSMAP)
    if ole is not None:
        return ole.get(qn("r:id"))
    return None

def _is_mathtype_ole(w_object_element) -> bool:
    ole = w_object_element.find("o:OLEObject", NSMAP)
    if ole is not None:
        return (ole.get("ProgID", "") or "").lower().startswith("equation")
    return False


# ============================================================================
# 格式化工具
# ============================================================================

def _format_run_text(text: str, run_element) -> str:
    if not text:
        return text
    is_bold = is_italic = False
    rpr = run_element.find("w:rPr", NSMAP)
    if rpr is not None:
        is_bold = rpr.find("w:b", NSMAP) is not None
        is_italic = rpr.find("w:i", NSMAP) is not None
    if is_bold:
        text = f"**{text}**"
    if is_italic:
        text = f"*{text}*"
    return text

def _to_relative_path(abs_path: str, base_dir: str) -> str:
    if not base_dir:
        return abs_path
    try:
        return os.path.relpath(abs_path, base_dir)
    except ValueError:
        return abs_path


# ============================================================================
# w:r 内部子节点遍历（处理 w:drawing / w:object / m:oMath 嵌套）
# ============================================================================

def _parse_run_children(run_element, image_map: dict,
                        is_in_table: bool = False) -> list:
    result = []
    text_parts = []

    for child in run_element.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "t":
            text_parts.append(child.text or "")
        elif tag in ("tab",):
            text_parts.append("\t")
        elif tag in ("br", "cr"):
            text_parts.append("\n")

        elif tag == "drawing":
            if text_parts:
                result.append(_format_run_text("".join(text_parts), run_element))
                text_parts = []
            rid = _extract_blip_rid(child)
            if rid and rid in image_map:
                rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                result.append(f"![]({rel_path})")

        elif tag == "object":
            if text_parts:
                result.append(_format_run_text("".join(text_parts), run_element))
                text_parts = []
            if _is_mathtype_ole(child):
                ole_rid = _extract_ole_object_rid(child)
                if ole_rid:
                    tag_type = "MTEF_INLINE" if is_in_table else "MTEF_BLOCK"
                    result.append(f"[{tag_type}:{ole_rid}]")
            else:
                rid = _extract_ole_image_rid(child)
                if rid and rid in image_map:
                    rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                    result.append(f"![OLE对象]({rel_path})")

        elif tag in ("oMath", "oMathPara"):
            if text_parts:
                result.append(_format_run_text("".join(text_parts), run_element))
                text_parts = []
            latex = omml_to_latex(child)
            if latex and not latex.startswith("[公式:"):
                if is_in_table:
                    latex = latex.replace("\n", " ")
                result.append(f"${latex}$" if is_in_table else f"$${latex}$$")
            else:
                result.append(latex or "[公式: 转换失败]")

    if text_parts:
        result.append(_format_run_text("".join(text_parts), run_element))
    return result


# ============================================================================
# 段落子节点遍历（含 AlternateContent / w:ins / w:del 处理）
# ============================================================================

def _parse_paragraph_children(para, image_map: dict,
                              is_in_table: bool = False) -> str:
    parts = []
    for child in para._p.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "r":
            parts.extend(_parse_run_children(child, image_map, is_in_table))

        elif tag == "ins":
            for sub in child.iterchildren():
                sub_tag = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                if sub_tag == "r":
                    parts.extend(_parse_run_children(sub, image_map, is_in_table))
                elif sub_tag in ("oMath", "oMathPara"):
                    latex = omml_to_latex(sub)
                    if latex and not latex.startswith("[公式:"):
                        parts.append(f"$${latex}$$")
                    else:
                        parts.append(latex or "[公式: 转换失败]")

        elif tag == "del":
            continue

        elif tag == "oMath":
            latex = omml_to_latex(child)
            if latex and not latex.startswith("[公式:"):
                if is_in_table:
                    latex = latex.replace("\n", " ")
                parts.append(f"${latex}$" if is_in_table else f"$${latex}$$")
            else:
                parts.append(latex or "[公式: 转换失败]")

        elif tag == "oMathPara":
            latex = omml_to_latex(child)
            if latex and not latex.startswith("[公式:"):
                if is_in_table:
                    flat = latex.replace("\n", " ").replace("\\[", "").replace("\\]", "")
                    parts.append(f"${flat}$")
                else:
                    parts.append(f"\n\n$${latex}$$\n\n")
            else:
                parts.append(latex or "[公式: 转换失败]")

        elif tag == "drawing":
            rid = _extract_blip_rid(child)
            if rid and rid in image_map:
                rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                parts.append(f"![]({rel_path})")

        elif tag == "object":
            if _is_mathtype_ole(child):
                ole_rid = _extract_ole_object_rid(child)
                if ole_rid:
                    tag_type = "MTEF_INLINE" if is_in_table else "MTEF_BLOCK"
                    parts.append(f"[{tag_type}:{ole_rid}]")
            else:
                rid = _extract_ole_image_rid(child)
                if rid and rid in image_map:
                    rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                    parts.append(f"![OLE对象]({rel_path})")

        elif tag == "AlternateContent":
            choice = child.find("mc:Choice", NSMAP)
            if choice is not None:
                for sub in choice.iterchildren():
                    sub_tag = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                    if sub_tag == "drawing":
                        rid = _extract_blip_rid(sub)
                        if rid and rid in image_map:
                            rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                            parts.append(f"![]({rel_path})")
                    elif sub_tag == "object":
                        if _is_mathtype_ole(sub):
                            ole_rid = _extract_ole_object_rid(sub)
                            if ole_rid:
                                tag_type = "MTEF_INLINE" if is_in_table else "MTEF_BLOCK"
                                parts.append(f"[{tag_type}:{ole_rid}]")
                        else:
                            rid = _extract_ole_image_rid(sub)
                            if rid and rid in image_map:
                                rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                                parts.append(f"![OLE对象]({rel_path})")

        elif tag == "AlternateContent":
            choice = child.find("mc:Choice", NSMAP)
            if choice is not None:
                for sub in choice.iterchildren():
                    sub_tag = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                    if sub_tag == "drawing":
                        rid = _extract_blip_rid(sub)
                        if rid and rid in image_map:
                            rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                            parts.append(f"![]({rel_path})")
                    elif sub_tag == "object":
                        rid = _extract_ole_image_rid(sub)
                        if rid and rid in image_map:
                            if _is_mathtype_ole(sub):
                                tag_type = "OCR_INLINE" if is_in_table else "OCR_BLOCK"
                                parts.append(f"[{tag_type}:{rid}]")
                            else:
                                rel_path = _to_relative_path(image_map[rid], image_map.get("_base_dir", ""))
                                parts.append(f"![OLE对象]({rel_path})")

    return " ".join(parts)


# ============================================================================
# 表格解析
# ============================================================================

def _parse_table(table: DocxTable, image_map: dict,
                 md_base_dir: str = "") -> str:
    rows_md = []
    header_done = False
    for row in table.rows:
        cells_md = []
        for cell in row.cells:
            cell_parts = []
            for para in cell.paragraphs:
                result = _parse_paragraph_children(para, image_map, is_in_table=True)
                if result.strip():
                    cell_parts.append(result)
            cell_text = "<br>".join(cell_parts) if cell_parts else ""
            cell_text = cell_text.replace("\n", "<br>")
            cells_md.append(cell_text)
        row_line = "| " + " | ".join(cells_md) + " |"
        rows_md.append(row_line)
        if not header_done and len(table.rows) > 0:
            rows_md.append("| " + " | ".join(["---"] * len(cells_md)) + " |")
            header_done = True
    return "\n".join(rows_md) + "\n"


# ============================================================================
# Body 遍历
# ============================================================================

def _parse_body_elements(doc, image_map: dict,
                         md_base_dir: str = "") -> str:
    image_map["_base_dir"] = md_base_dir
    body_element = doc.element.body
    body_wrapper = doc._body

    sections = []
    prev_was_empty = False

    for child in body_element.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            para = Paragraph(child, body_wrapper)
            style_name = (para.style.name or "").lower() if para.style else ""
            heading_match = re.match(r"(?:heading|标题)\s*(\d+)", style_name)
            if heading_match:
                level = min(max(int(heading_match.group(1)), 1), 6)
                text = _parse_paragraph_children(para, image_map)
                if text.strip():
                    sections.append(f"\n{'#' * level} {text.strip()}\n")
                prev_was_empty = False
                continue
            text = _parse_paragraph_children(para, image_map).strip()
            if text:
                sections.append(f"{text}\n\n")
                prev_was_empty = False
            elif not prev_was_empty:
                # 空段落保留一个段落分隔
                sections.append("\n")
                prev_was_empty = True

        elif tag == "tbl":
            sections.append("\n")
            table = DocxTable(child, body_wrapper)
            sections.append(_parse_table(table, image_map, md_base_dir) + "\n")
            prev_was_empty = False

    return "".join(sections)


# ============================================================================
# 参考文献区域清洗（作用域隔离：双语锚点定位 + 区间正则）
# ============================================================================

_REF_ANCHORS = [
    r'^# 参考文献\b',
    r'^## 参考文献\b',
    r'^# References\b',
    r'^## References\b',
    r'^# Bibliography\b',
    r'^## Bibliography\b',
]


def _normalize_reference_section(md_content: str) -> str:
    """
    仅对参考文献区域（锚点之后）执行 [N] 换行清洗。
    正文中的行内引用 [N] 完全不受影响。
    """
    for anchor_pattern in _REF_ANCHORS:
        m = re.search(anchor_pattern, md_content, re.MULTILINE)
        if m:
            cut_point = m.start()
            body = md_content[:cut_point]
            ref_section = md_content[cut_point:]
            ref_section = re.sub(
                r'(?<!\n)(\[\d+\])',
                r'\n\1',
                ref_section
            )
            return body + ref_section
    return md_content


# ============================================================================
# 主入口
# ============================================================================

def convert_docx_to_md(docx_path: str, output_md_path: str = None,
                       image_dir: str = None) -> str:
    if not os.path.exists(docx_path):
        return f"转换失败: DOCX 文件不存在: {docx_path}"

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    src_dir = os.path.dirname(os.path.abspath(docx_path))

    # 若源文件在 raw/ 下，默认输出到 raw_md/；否则输出到同目录
    if os.path.basename(src_dir) == "raw":
        parent = os.path.dirname(src_dir)
        base_dir = os.path.join(parent, "raw_md")
    else:
        base_dir = src_dir

    out_dir = os.path.join(base_dir, f"{base_name}_md")
    if output_md_path is None:
        output_md_path = os.path.join(out_dir, f"{base_name}.md")
    if image_dir is None:
        image_dir = out_dir

    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(os.path.join(out_dir, "figures"), exist_ok=True)

    try:
        # Phase 1: 加载 + 图片提取 + 公式解析
        log.info(f"加载文档: {docx_path}")
        doc = Document(docx_path)

        image_map = _build_image_map(doc, docx_path, image_dir)
        log.info(f"图片提取完成: {len([k for k in image_map if k != '_base_dir'])} 张 → {image_dir}/figures/")

        formula_map = _build_formula_map(doc, docx_path)
        log.info(f"公式解析完成: {len(formula_map)} 个")

        # Phase 2: 文档遍历（MathType 公式通过 [MTEF:rId] 占位符引用）
        log.info("文档结构遍历...")
        md_content = _parse_body_elements(doc, image_map, base_dir)

        # Phase 3: 替换 MTEF 占位符为 LaTeX
        formula_rids = set(re.findall(r'\[MTEF_(?:BLOCK|INLINE):(rId\d+)\]', md_content))
        if formula_rids:
            inline_count = 0
            block_count = 0
            for rid in formula_rids:
                info = formula_map.get(rid, {'latex': '[公式: 未解析]', 'is_inline': True})
                latex = info['latex']
                # DOM 级上下文覆盖: 公式前后有非空文本 → 强制 Inline
                ctx_inline = _is_inline_by_context(md_content, rid)
                is_inline = info['is_inline'] or ctx_inline
                if is_inline:
                    md_content = md_content.replace(f'[MTEF_INLINE:{rid}]', f'${latex}$')
                    md_content = md_content.replace(f'[MTEF_BLOCK:{rid}]', f'${latex}$')
                    inline_count += 1
                else:
                    md_content = md_content.replace(f'[MTEF_INLINE:{rid}]', f'\n\n$${latex}$$\n\n')
                    md_content = md_content.replace(f'[MTEF_BLOCK:{rid}]', f'\n\n$${latex}$$\n\n')
                    block_count += 1
            log.info(f"公式替换: {inline_count} 行内 + {block_count} 块级")

        # Phase 4: 参考文献清洗 + 输出
        md_content = _normalize_reference_section(md_content)
        with open(output_md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        log.info(f"Markdown 已保存: {output_md_path}")
        img_count = len([k for k in image_map if k != "_base_dir"])
        return f"转换成功！\n- Markdown: {output_md_path}\n- 图片: {img_count} 张\n- 公式: {len(formula_map)} 个"

    except Exception as e:
        log.error(f"转换失败: {e}", exc_info=True)
        return f"转换失败: {str(e)}"


def _is_inline_by_context(md_content: str, rid: str) -> bool:
    """DOM 级上下文判定：公式前后有非空文本 → 强制行内"""
    import re
    # 查找该 rId 的所有占位符位置
    for tag in (f'[MTEF_INLINE:{rid}]', f'[MTEF_BLOCK:{rid}]'):
        for m in re.finditer(re.escape(tag), md_content):
            pos = m.start()
            # 检查前面紧邻的是否为非空白文本（非换行符）
            before = md_content[max(0, pos-30):pos]
            after = md_content[m.end():m.end()+30]
            # 前方有中英文文本字符（非纯空格/换行）
            before_text = re.search(r'[\w\u4e00-\u9fff]', before)
            # 后方有中英文文本字符
            after_text = re.search(r'[\w\u4e00-\u9fff]', after)
            if before_text or after_text:
                return True
    return False


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python docx_to_md.py <DOCX路径> [输出MD路径] [图片目录]")
        sys.exit(1)
    result = convert_docx_to_md(
        docx_path=sys.argv[1],
        output_md_path=sys.argv[2] if len(sys.argv) > 2 else None,
        image_dir=sys.argv[3] if len(sys.argv) > 3 else None,
    )
    print(result)
