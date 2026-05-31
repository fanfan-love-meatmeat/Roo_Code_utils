# 项目排期与研发 Task: patent_md_to_docx 引擎

## Sprint 1: 铸造原子操作层 (DOCX Builder)
**目标**：封装绝对可靠的 `python-docx` 接口，只接受纯文本与参数，不涉及任何 AST 解析。
* **Task 1.1**: 实现 `init_document() -> Document`（注入 A4 纸张、25/15/25/15mm 页边距、默认字体族）。
* **Task 1.2**: 实现底层字体引擎 `set_run_font()`（确保 `w:eastAsia`, `w:ascii`, `w:hint` XML 属性安全注入）。
* **Task 1.3**: 实现通用排版构建器：
    * `add_section_title(doc, text)`
    * `add_body_paragraph(doc, text, number=None)` 
    * `add_claim_paragraph(doc, number, text)`
* **Task 1.4**: 实现视觉占位构建器：
    * `add_formula_placeholder(doc, text, is_inline=True/False)`
    * `add_image_or_placeholder(doc, img_rel_path, caption)`

## Sprint 2: 搭建 AST 状态机与主循环 (Parser & Tagger)
**目标**：解析 Markdown，识别专利语义，路由给 Sprint 1 的构建器。
* **Task 2.1**: 实现独立于 AST 的正则分类器 `classify_section(text) -> Tag`（识别五大书标题、子标题、权利要求项等）。
* **Task 2.2**: 搭建 `markdown-it-py` 主循环框架，显式管理 Token 游标（Index advancement）。
* **Task 2.3**: 核心路由分发：
    * 将 `paragraph` 结合分类器路由至对应的 Sprint 1 构建器。
    * 实现 `hr` -> `doc.add_page_break()`。
* **Task 2.4 (复杂项)**: 实现独立的表格状态机 `_parse_table_block(tokens, index)`。
* **Task 2.5 (优化项)**: 在遍历 `inline.children` 时，利用 O(1) 预读过滤 `** **` 纯空格等脏数据。

## Sprint 3: 细节补完与边界清洗 (Edge Cases)
**目标**：处理复杂的行内映射与特定边界。
* **Task 3.1**: 实现 AST 级别的 `[0001]` 编号精确提取（替代初期的正则猜测）。
* **Task 3.2**: 完成 `em` (斜体) 以及 `softbreak/hardbreak` 的映射。
* **Task 3.3**: 连续 `text` Token 的字符串合并压缩（减少 DOCX 底层的 Run 碎片化，提升最终 Word 文档的运行性能）。

## Sprint 4: 质量保障与 MCP 集成 (Production Ready)
**目标**：达到发布标准，移交 AI Agent 使用。
* **Task 4.1**: 编写 **属性断言测试 (Property Assertion Tests)**，通过解包生成的 DOCX 断言页边距、首行缩进(Pt 24)、字体绑定等 XML 属性是否达标。
* **Task 4.2**: 在 `server.py` 中注册 `@mcp.tool() patent_md_to_docx`。
* **Task 4.3**: 接入 `_validate_input_file` 等路径防御层。
* **Task 4.4**: 端到端联调：使用真实大模型生成的 `完整专利申请书.md` 跑通测试。