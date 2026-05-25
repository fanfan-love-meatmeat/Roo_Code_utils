# MTEF 公式转换调试报告

**目标文件**: `范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx`  
**测试日期**: 2026-05-25  
**调试工具**: `debug_mtef.py` / `debug_byte_trace.py`  
**测试环境**: Python 3.10 + roo_mcp conda env

---

## 1. 测试概况

| 指标 | 数值 |
|------|------|
| DOCX 中 OLE 对象总数 | **263** |
| 全部为 MathType OLE (DSMT7) | ✓ |
| Fast-Path (TeX metadata) 命中 | **0/263 (0%)** |
| Slow-Path (递归解析) 执行 | 263/263 |
| 公式输出质量 | **全部乱码** |

---

## 2. 管道结构回顾

```
DOCX (ZIP)
  └─ word/embeddings/oleObject*.bin
       └─ olefile.OleFileIO → "Equation Native" stream
            └─ raw[28:] → MTEF v5 bytes
                 ├─ Track 1 (Fast): extract_tex_metadata() → 搜索 "TeX Input Language\0"
                 └─ Track 2 (Slow): MTEFParser.parse()
                      ├─ _skip_header()   → 跳过 FONT_DEF/ENCODING_DEF/EQN_PREFS
                      └─ _parse_line()    → 递归解析 LINE/CHAR/TMPL 记录
```

---

## 3. 根因分析

### 3.1 Fast-Path 失败（预期行为）

**结论**: 非 Bug。所有 263 个公式均不含 `TeX Input Language` 标记。

**原因**: 该文档的 MathType 公式全部通过 **UI 点击方式** 创建（非 TeX 语法输入）。MathType 仅在用户使用 TeX 输入法时才会将原始 TeX 源码写入 MTEF 头部。UI 点击生成的公式中不存在该标记。

**影响**: Fast-Path 不可用，必须降级到 Slow-Path。

---

### 3.2 Slow-Path 输出乱码（核心 Bug）

#### 3.2.1 Bug 定位

**Bug 位置**: `tools/mtef_parser.py` → `MTEFParser._skip_header()` 第 293-316 行

**Bug 机理**: MTEF v5 扩展记录的 tag 值与标准记录的低 4 位冲突，导致 header 跳过逻辑提前终止。

#### 3.2.2 技术细节

MTEF v5 定义了 4 个扩展记录，其 tag 字节值范围为 `0x10-0x13`：

| Tag 值 | 扩展记录 | 低4位 (tag & 0x0F) | 冲突的标准记录 |
|--------|----------|---------------------|----------------|
| `0x10` | COLOR_DEF | `0x00` | END |
| `0x11` | FONT_DEF | `0x01` | LINE |
| `0x12` | EQN_PREFS | `0x02` | CHAR |
| `0x13` | ENCODING_DEF | `0x03` | TMPL |

`_skip_header` 中使用的分类逻辑：

```python
# mtef_parser.py line 294-296
tag = self.data[self.pos]
record_type = tag & 0x0F       # ← 仅取低4位！
options = tag >> 4

# line 298
if record_type in (LINE, CHAR, TMPL, PILE, MATRIX, ...):
    break  # 方程数据开始 ← 误触发！
```

当 MTEF 头部元数据区包含扩展记录时，流程为：

```
字节流: ... 11 06 CB CE CC E5 00  12 00 08 21 2F ...
              ↑ FONT_DEF (0x11)      ↑ ENCODING_DEF?
              ↑ 0x11 & 0x0F = 0x01 = LINE → 跳过（正确，用 elif FONT_DEF 处理）

字节流: ... 12 00 08 21 2F 45 ...
              ↑ 此处是 EQN_PREFS (0x12)
              ↑ 0x12 & 0x0F = 0x02 = CHAR → 进入 break 分支 → 提前终止！
```

#### 3.2.3 逐字节验证

**公式 #1 (oleObject233.bin) 头部 hex**:

```
000000: 05 01 00 07 04 44 53 4D 54 37 00 01  ← MTEF header + app_key + eq_options
00000C: 13 57 69 6E 41 6C 6C ...              ← 0x13 = ENCODING_DEF ← 被误判为 TMPL
00006C: E5 00 12 00 08 21 ...                 ← 0x12 = EQN_PREFS
...方程数据区在 ~0x0074 之后...
```

`_skip_header` 在 byte 12 处遇到 `0x13`：
1. `record_type = 0x13 & 0x0F = 3` → 匹配 TMPL
2. TMPL 在 break 列表中 → **立即跳出循环**
3. `self.pos` 停留在 12，但真正的方程数据在 ~0x74 之后
4. `_parse_line` 从 pos=12 开始解析，将 FONT_DEF/ENCODING_DEF 等元数据字节当作方程数据

**实际解析路径** (从错误的 pos=12 开始):

```
pos=12: tag=0x13 → 读作 TMPL(sel=0x57=87, var=105, opts=0x6E)
  → 实际是 ENCODING_DEF name="WinAllBasicCodePages\0"
  → 递归调用 _parse_line()
    → 中文字节 "Times New Roman\0" 被当作 LINE/CHAR 解析
    → 产生垃圾字符: {s\langle{}m{l{}{}...}
```

**正确的解析起点** 应该在字节 `0x74` 附近（EQN_PREFS 80字节之后）：

```
000070: ... 12 00 08 21 2F 45 8F 44 2F 41 50 F4 10 0F 47 5F ...
                 ↑ EQN_PREFS(0x12) + 80B 跳过 → 方程数据从此开始
```

#### 3.2.4 影响范围

- **文档中 263 个 MathType 公式全部受影响**
- `_skip_header` 总是在第一个扩展记录处提前终止
- 元数据区（字体名称含中文如"宋体"）被当作方程数据解析
- 输出的 LaTeX 为随机 Unicode 字符拼接的字符串

#### 3.2.5 修复方案

在 `_skip_header` 循环中，**先判断完整 tag 字节是否属于扩展记录**，再按低4位分类：

```python
# mtef_parser.py, _skip_header() 方法内，while 循环中
while self.pos < len(self.data):
    tag = self.data[self.pos]
    record_type = tag & 0x0F
    options = tag >> 4

    # [FIX] 先检查扩展记录（tag >= 0x10 时低4位不可信）
    if tag == FONT_DEF:          # 0x11
        self.pos += 1
        self._parse_font_def()
    elif tag == ENCODING_DEF:    # 0x13
        self.pos += 1
        self._parse_encoding_def()
    elif tag == EQN_PREFS:       # 0x12
        self.pos += 1
        self._parse_eqn_prefs()
    elif tag == COLOR_DEF:       # 0x10
        self.pos += 1
        self._parse_color_def()
    # [FIX END]
    elif record_type in (LINE, CHAR, TMPL, PILE, MATRIX, EMBELL, RULER,
                          SIZE, FULL, SUB, SUB2, SYM, SUBSYM):
        break  # 方程数据开始
    elif record_type == END:
        self.pos += 1
        break
    else:
        self.pos += 1  # skip unknown
```

---

## 4. 其他发现

### 4.1 字体名称编码问题

MTEF 元数据中包含 GBK 编码的中文字体名（如 `CB CE CC E5 00` = "宋体\0"），当前 `_read_null_terminated_string` 使用 UTF-8 解码：

```python
s = self.data[self.pos:end].decode('utf-8', errors='replace')
```

由于 `errors='replace'` 兜底，字体名识别不准确但不影响方程解析（仅影响字体映射）。建议后续增强为自动检测编码（GBK/Windows-1252 fallback）。

### 4.2 `_parse_line` 中的重复代码

`tools/docx_to_md.py` 第 439-479 行存在两个完全相同的 `elif tag == "AlternateContent":` 代码块。第二个块中第 475 行标签拼写为 `OCR_INLINE`/`OCR_BLOCK`（疑为 `MTEF_INLINE`/`MTEF_BLOCK` 笔误）。

### 4.3 扩展记录常量定义位置

`mtef_parser.py` 第 42-45 行定义了扩展记录常量：

```python
COLOR_DEF = 0x10
FONT_DEF = 0x11
EQN_PREFS = 0x12
ENCODING_DEF = 0x13
```

但这些常量在 `_skip_header` 中的 `elif` 分支使用了 `record_type == FONT_DEF` 比较（即 `record_type == 0x11`），而 `record_type = tag & 0x0F` 永远不可能等于 `0x11`。只有 `0x11 & 0x0F = 0x01` 时才会进入 LINE 检查。这意味着现有的 `elif record_type == FONT_DEF:` 分支 **从未被真正触发过**，`_skip_header` 中的 FONT_DEF/ENCODING_DEF 等处理逻辑是**死代码**。

---

## 5. 修复验证

### 5.1 修复效果

运行 `verify_fix.py` 对前5个公式进行修复验证：

| 公式 | 原 _skip_header 终止位 | 修复后终止位 | 差值 | 剩余数据 |
|------|------------------------|-------------|------|---------|
| #1 (oleObject233) | 12 | 191 | +179B | 40B |
| #2 (oleObject1) | 12 | 191 | +179B | 35B |
| #3 (oleObject2) | 12 | 191 | +179B | 52B |
| #4 (oleObject3) | 12 | 191 | +179B | 35B |
| #5 (oleObject4) | 12 | 191 | +179B | 52B |

修复后 `_skip_header` 正确消费了 179 字节元数据（FONT_DEF × 6 + ENCODING_DEF × 2 + EQN_PREFS × 1），方程数据从正确的边界开始。

### 5.2 残留问题

修复后方程数据区起始字节为 `0C`（SUB2 类型标签），后续字节序列 `01 00 01 00 01 02 02 02...` 的解析仍需进一步验证：

- **疑似 SIZE 表**: 该区域可能为 MTEF v5 的 SIZE 定义表（`_parse_size_record` 当前仅处理单条 SIZE 记录，未处理紧凑格式的 SIZE 表）
- **EQN_PREFS 固长问题**: `_parse_eqn_prefs` 硬编码跳过 80 字节，若 EQN_PREFS 实际长度不同将导致后续字节错位
- **建议**: 先应用 nibble 冲突修复，重新运行完整转换，根据输出质量决定是否需要进一步调整 SIZE 表解析逻辑

---

## 6. 调试产物

| 文件 | 用途 |
|------|------|
| `debug_mtef.py` | 263公式全量诊断，输出 JSONL |
| `debug_byte_trace.py` | 前5公式逐字节追踪 |
| `verify_fix.py` | 修复效果验证脚本 |
| `debug_output.jsonl` | 前10公式结构化诊断数据 |

---

## 7. 结论

| 项目 | 状态 |
|------|------|
| **根因** | `_skip_header()` 未区分 MTEF v5 扩展记录与标准记录，因低4位冲突导致 header 跳过提前终止 |
| **严重程度** | P0 — 所有公式输出为乱码，转换功能完全失效 |
| **修复复杂度** | 低 — 在 `_skip_header` 循环中增加4个 `elif tag == ...` 判断即可 |
| **Fix 影响** | 仅影响 `mtef_parser.py`，不影响其他模块 |
