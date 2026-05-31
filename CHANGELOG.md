# 更新日志 (CHANGELOG)

## 2026-06-01
- **路径命名逻辑简化** (`server.py`)：移除 `patent_md_to_docx` 的 `version` 参数及版本号自动提取，改为 `{MD完整文件名}_{时间戳}.docx` 直接命名；输出目录从项目根 `patent_docx/` 调整为 MD 文件上层的 `patent_docx/`
- **MathType 幂等性防御** (`tools/patent_style_utils.py`)：在 `add_display_formula_placeholder` 和 `add_rich_paragraph` 中新增 LaTeX 定界符（`\lfloor`/`\rfloor`/`\lceil`/`\rceil`）的 `\left`/`\right` 自动包装，防止 MathType 宏因重复转换导致定界符不匹配而失败
- **开发档案迭代**：用新版 S1/S2 分层审计报告目录 + T6 设计文档替代旧版单文件审计报告；新增 `test/` 测试目录

## 2026-05-31
- **新增 MCP 工具 `patent_md_to_docx`**：将专利 Markdown 转换为符合 CNIPA 规范的 DOCX
  - 新增 `tools/patent_md_to_docx.py`：底层转换引擎，处理 heading 层级、段落编号 `[0001]`、权利要求、公式占位
  - 新增 `tools/patent_style_utils.py`：专利文档样式工具（宋体+Times New Roman、页边距 左25/上25/右15/下15mm、首行缩进 2 字符）
  - `server.py` 注册 `patent_md_to_docx` 端点，含输入校验、版本号自动提取（`_v2.4.md` → `v2.4`）、输出命名规则
- **开发档案补充**：纳入 `report/patent_md_to_docx/` 完整开发记录（CTO 设计 T1-T5 + 审记 + 实现记录）；补全 `report/docx_to_md/` 中 D2 日报、审计报告目录及 D2 实现记录

## 2026-05-26
- **开发档案归档**：纳入 `report/docx_to_md/` 完整开发记录
  - `00_CTO/`：顶层设计文档（D1-D2 开发日报 + T1-T4 技术专项）
  - `01_Auditor/`：审计报告（项目现状与工作总结）
  - `02_Implement/`：实现记录，含各阶段的详细设计说明、自检报告、回归测试脚本和调试工具
- **新增输入防御层**：`server.py` 增加 `_validate_input_file()` 统一校验
  - 路径存在性与合法性（`os.path.isfile` 拒绝目录和不存在路径）
  - 扩展名白名单（`.md` / `.pdf` / `.docx`，大小写不敏感）
  - 文件大小上限（100MB，防止 OOM）
  - 4 个 MCP 工具入口异常处理从泛化 `except Exception` 重构为精准 `ToolError` 抛出，使 AI Agent 可根据自然语言错误信息自我纠错
- **文档重构：README.md 全面扩增与规范化**
  - 新增项目目录结构图，清晰展示 tools/ 下各模块职责
  - 新增"MCP 工具一览"总览表，一目了然 4 个工具的输入输出方向
  - 各工具独立章节，均包含功能说明 + 调用示例 JSON + 可选参数
  - 补充 Roo Code MCP 配置的完整 JSON 示例
  - 命令行调试增加 `md_to_docx.py` 独立调用示例

## 2026-05-25
- **重大架构升级：新增 `convert_docx_to_md` MCP 工具，弃用 OCR，转向纯 Python MTEF 双轨解析**
  - **新增 `tools/docx_to_md.py`（~700行）**：DOCX→Markdown 核心转换引擎，承担单轨 XML 遍历、公式解析调度、段落重建、图片导出、参考文献清洗等全流程
  - **新增 `tools/mtef_fast.py`（103行）**：Track 1 快速路径 —— 从 MTEF 字节流中搜索 `TeX Input Language` 元数据标记，微秒级提取原始 TeX 明文；同时实现 MTEF v5 头部 Bit 0 位掩码解析，自动判定 Inline/Display 公式类型
  - **新增 `tools/mtef_parser.py`（674行）**：Track 2 递归兜底引擎 —— 纯 Python 状态机逐字节解析 MTEF v5 二进制流，完整覆盖 19 种记录类型（END/LINE/CHAR/TMPL/PILE/MATRIX/EMBELL/RULER/FONT_STYLE_DEF/SIZE/FULL/SUB/SUB2/SYM/SUBSYM/COLOR 及 v5 扩展 COLOR_DEF/FONT_DEF/EQN_PREFS/ENCODING_DEF），支持 Nudge 偏移量、选项位（高 4 位）和递归子结构 → LaTeX 转换
  - **新增 `tools/xslt/`（8个文件）**：OMML 原生公式转 LaTeX 的 XSLT 管道 —— `omml2mml.xsl`（OMML→MathML）+ `mmltex.xsl` 及 6 个依赖样式表（MathML→LaTeX），自动处理 `mc:AlternateContent`
  - **删除 OCR 管线**：彻底移除 `pix2tex`、`torch`、`wmf_to_png` 等 GPU 依赖，依赖精简至 `olefile`（~100KB）唯一新增
  - **MCP 服务注册**：`server.py` 新增 `convert_docx_to_md` 工具，通过 FastMCP 暴露给 AI Agent
  - **修订模式兼容**：接受 `w:ins` 插入、跳过 `w:del` 删除
  - **段落边界重建**：`<w:p>` 逐段输出 `\n\n`，保证语义完整
  - **表格内公式降级**：块公式自动降为行内 `$...$`，维持 Markdown 表格语法边界
  - **参考文献清洗**：双语锚点（`参考文献` / `References`）+ 尾部 `[N]` 独立换行
  - **图片自动导出**：嵌入图片提取到 `figures/` 子目录，相对路径引用，自动去重
  - **输出规范化**：统一输出为 `原名_md/` 目录结构（含 `figures/` + `原名.md`）
  - **实测验证**：公式 127/127 解析成功，图片 24 张导出无误
  - `environment_backup.md` 同步更新依赖清单和模块说明

## 2026-05-21
- 修正 README 中的项目路径：将 `D:/FJL/Projects/Roo_Code_utils/` 更新为重组后的 `D:/FJL/Projects/Kilo_Code_Gobal_Settings/Roo_Code_utils/`
