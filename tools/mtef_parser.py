"""
Track 2 (Slow-Path): 纯 Python MTEF v5 二进制解析引擎
- 递归状态机逐字节读取 MTEF 记录
- 转换为 LaTeX 字符串
- 零外部依赖，纯 Python + struct
- 小端序锁死：所有 16 位值使用 '<H' 显式声明

参考：
  - Wiris MTEF v5 规范: docs.wiris.com/mathtype-mtef-v5
  - rtf2latex2e MTEF v3: rtf2latex2e.sourceforge.net/MTEF3.html
  - transpect/mathtype-extension (Ruby 参考实现)
"""
import struct
import logging

log = logging.getLogger(__name__)

# ============================================================================
# MTEF v5 记录类型常量（4-bit 低位）
# ============================================================================

END       = 0x0
LINE      = 0x1
CHAR      = 0x2
TMPL      = 0x3
PILE      = 0x4
MATRIX    = 0x5
EMBELL    = 0x6
RULER     = 0x7

# v5 新增
FONT_STYLE_DEF = 0x8
SIZE      = 0x9
FULL      = 0xA
SUB       = 0xB
SUB2      = 0xC
SYM       = 0xD
SUBSYM    = 0xE
COLOR     = 0xF              # v5 only (在 v3 中为 ENCODING_DEF)

# v5 扩展记录 (option 位区分 0x8+ 范围)
COLOR_DEF = 0x10
FONT_DEF  = 0x11
EQN_PREFS = 0x12
ENCODING_DEF = 0x13

# 选项位 (高 4 位)
# 通用
xfLMOVE   = 0x8   # nudge 跟随
# CHAR 专用
xfAUTO    = 0x1   # 函数识别候选
xfEMBELL  = 0x2   # 修饰列表跟随
# LINE 专用
xfNULL    = 0x1   # 占位空行
xfLSPACE  = 0x4   # 行间距跟随
xfRULER   = 0x2   # ruler 记录跟随

# 字符编码选项
mtefOPT_CHAR_ENC_CHAR_8 = 0x4

# ============================================================================
# 模板选择器与变体
# ============================================================================

TMPL_SELECTORS = {
    0x00: ("angle",     "angle"),
    0x01: ("paren",     "parentheses"),
    0x02: ("brace",     "curly brace"),
    0x03: ("brack",     "square bracket"),
    0x04: ("bar",       "vertical bar"),
    0x05: ("dbar",      "double bar"),
    0x06: ("floor",     "floor"),
    0x07: ("ceil",      "ceiling"),
    0x0B: ("root",      "radical"),
    0x0C: ("sqrt",      "square root"),
    0x0D: ("fract",     "fraction"),
    0x0E: ("over",      "over"),
    0x0F: ("script",    "sub/superscript"),
    0x10: ("ubar",      "underbar"),
    0x11: ("obar",      "overbar"),
    0x15: ("sum",       "summation"),
    0x16: ("prod",      "product"),
    0x17: ("coprod",    "coproduct"),
    0x18: ("union",     "union"),
    0x19: ("inter",     "intersection"),
    0x1A: ("lim",       "limit"),
    0x1B: ("ldiv",      "long division"),
    0x1C: ("slfract",   "slash fraction"),
    0x1D: ("intop",     "integral-style big op"),
    0x1E: ("sumop",     "summation-style big op"),
    0x1F: ("lscript",   "leading sub/superscript"),
    0x21: ("sint",      "single integral"),
    0x22: ("dint",      "double integral"),
    0x23: ("tint",      "triple integral"),
}

# 字体样式映射 (typeface values, v5 扩展)
# typeface > 128 → 负值 = 显式字体引用 (FONT_DEF)
# typeface <= 10 → 内置样式
STYLE_NAMES = {
    1: "mathrm", 2: "mathrm", 3: "mathit",
    4: "mathrm", 5: "mathrm", 6: "mathrm",
    7: "mathbf", 8: "mathbb",
    9: "mathsf", 10: "mathtt",
    11: "mathcal",
}


class MTEFParseError(Exception):
    pass


# ============================================================================
# MTEF 解析器
# ============================================================================

class MTEFParser:
    """MTEF v5 二进制解析器"""

    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.font_map = {}      # {typeface_index: font_name}
        self.style_map = {}     # {style_index: (font_def, bold, italic)}
        self.font_counter = 1
        self._depth = 0         # 递归深度计数器

    def _read_byte(self) -> int:
        if self.pos >= len(self.data):
            raise MTEFParseError(f"Unexpected EOF at pos {self.pos}")
        b = self.data[self.pos]
        self.pos += 1
        return b

    def _read_uint16_le(self) -> int:
        """小端序锁死"""
        if self.pos + 2 > len(self.data):
            raise MTEFParseError(f"Unexpected EOF at pos {self.pos} (need 2 bytes)")
        val = struct.unpack('<H', self.data[self.pos:self.pos+2])[0]
        self.pos += 2
        return val

    def _read_int16_le(self) -> int:
        """小端序锁死"""
        if self.pos + 2 > len(self.data):
            raise MTEFParseError(f"Unexpected EOF at pos {self.pos} (need 2 bytes)")
        val = struct.unpack('<h', self.data[self.pos:self.pos+2])[0]
        self.pos += 2
        return val

    def _read_null_terminated_string(self) -> str:
        end = self.data.index(0, self.pos) if 0 in self.data[self.pos:] else len(self.data)
        s = self.data[self.pos:end].decode('utf-8', errors='replace')
        self.pos = end + 1
        return s

    def _read_nudge(self) -> tuple[int, int]:
        """读取 nudge 偏移量 (dx, dy)"""
        dx = self._read_byte()
        dy = self._read_byte()
        if dx == 128 and dy == 128:
            dx = self._read_int16_le()
            dy = self._read_int16_le()
        return dx, dy

    def _read_char_value(self, typeface: int, options: int) -> str:
        """读取 CHAR 记录的字符值并转换为 LaTeX"""
        if typeface >= 128:
            char_val = self._read_byte()   # MTCode 单字节
        elif options & mtefOPT_CHAR_ENC_CHAR_8:
            char_val = self._read_byte()
        else:
            char_val = self._read_uint16_le()

        return self._char_to_latex(char_val, typeface)

    def _char_to_latex(self, char_val: int, typeface: int) -> str:
        """字符值 → LaTeX 字符串"""
        if typeface >= 128:
            return self._mtcode_to_latex(char_val, typeface - 128)

        # typeface < 128: 标准 Unicode / Symbol 映射
        if 0x20 <= char_val <= 0x7E:
            ch = chr(char_val)
            if ch in '\\{}_^#&$%~':
                return '\\' + ch
            return ch
        elif char_val < 0x100:
            return self._symbol_to_latex(char_val)
        else:
            return chr(char_val)

    # ========================================================================
    # MTCode 字体编码映射表
    # ========================================================================

    # 文本字体 (Times New Roman, Courier 等) → ASCII 直通
    _MTOKEN_TEXT = {c: chr(c) for c in range(0x20, 0x7F)}

    # MT Extra 字体完整映射 (基于 transpect/MT_Extra.xml)
    _MTOKEN_MT_EXTRA = {
        0x21: '\\lvert', 0x22: '\\rvert', 0x23: '\\lVert', 0x24: '\\rVert',
        0x2A: '\\times', 0x2B: '\\plus',
        0x2D: '\\minus', 0x2E: '\\div',
        0x34: '\\centernot', 0x35: '\\neg',
        0x36: '\\circledR', 0x37: '\\circledS',
        0x38: '\\copyright', 0x39: '\\textregistered',
        0x3B: '\\degree', 0x3C: '\\mathdegree',
        0x41: '\\square',
        0xB8: '\\theta_p',    # 云台俯仰角 (pitch)
        0xD5: '\\theta_y',    # 云台偏航角 (yaw)
    }

    def _mtcode_to_latex(self, char_val: int, font_idx: int) -> str:
        """MTCode 单字节值 → LaTeX 命令（基于字体路由）"""
        # Step 1: 查询字体名
        font_name = self.font_map.get(font_idx, '').lower()

        # Step 2: ASCII 直通范围 (0x20-0x7E) — 大多数文本字体
        if 0x20 <= char_val <= 0x7E:
            if 'symbol' in font_name:
                return self._symbol_to_latex(char_val)
            elif 'mt extra' in font_name:
                return self._MTOKEN_MT_EXTRA.get(char_val,
                       f'[MTCode:0x{char_val:02X}]')
            else:
                # 文本字体: 标准 ASCII
                ch = chr(char_val)
                if ch in '\\{}_^#&$%~':
                    return '\\' + ch
                return ch

        # Step 3: 扩展范围 (0x80-0xFF) — Symbol/MT Extra
        if 'symbol' in font_name:
            return self._symbol_to_latex(char_val)
        elif 'mt extra' in font_name:
            return self._MTOKEN_MT_EXTRA.get(char_val,
                   f'[MTCode:0x{char_val:02X}]')

        # Step 4: 未命中 — PUA 占位符
        log.debug(f"[MTCode] unmapped char_val=0x{char_val:02X} "
                  f"font='{font_name}' idx={font_idx}")
        return f'[MTCode:0x{char_val:02X}]'

    def _symbol_to_latex(self, code: int) -> str:
        """Symbol/MT Extra 字体编码 → LaTeX 命令"""
        symbol_map = {
            # Greek lowercase
            0x61: '\\alpha', 0x62: '\\beta', 0x63: '\\chi', 0x64: '\\delta',
            0x65: '\\epsilon', 0x66: '\\phi', 0x67: '\\gamma', 0x68: '\\eta',
            0x69: '\\iota', 0x6A: '\\phi', 0x6B: '\\kappa', 0x6C: '\\lambda',
            0x6D: '\\mu', 0x6E: '\\nu', 0x6F: 'o', 0x70: '\\pi',
            0x71: '\\theta', 0x72: '\\rho', 0x73: '\\sigma', 0x74: '\\tau',
            0x75: '\\upsilon', 0x76: '\\varpi', 0x77: '\\omega',
            0x78: '\\xi', 0x79: '\\psi', 0x7A: '\\zeta',
            # Greek uppercase
            0x41: 'A', 0x42: 'B', 0x43: '\\Chi', 0x44: '\\Delta',
            0x45: 'E', 0x46: '\\Phi', 0x47: '\\Gamma', 0x48: '\\Theta',
            0x49: 'I', 0x4A: '\\theta', 0x4B: 'K', 0x4C: '\\Lambda',
            0x4D: 'M', 0x4E: 'N', 0x4F: 'O', 0x50: '\\Pi',
            0x51: '\\Theta', 0x52: 'P', 0x53: '\\Sigma', 0x54: 'T',
            0x55: '\\Upsilon', 0x56: '\\varsigma', 0x57: '\\Omega',
            0x58: '\\Xi', 0x59: '\\Psi', 0x5A: 'Z',
            # Math operators
            0xB1: '\\pm', 0xB3: '\\ge', 0xA3: '\\le',
            0xB4: '\\times', 0xB8: '\\div', 0xD6: '\\neq',
            0xC7: '\\cap', 0xC8: '\\subset', 0xC9: '\\supset',
            0xB9: '\\propto', 0xBC: '\\ni', 0xBD: '\\not\\subset',
            0xB7: '\\bullet', 0xBB: '\\rightarrow', 0xAC: '\\leftrightarrow',
            0xC4: '\\wedge', 0xC5: '\\vee', 0xC6: '\\oplus',
            0xD8: '\\otimes', 0xA5: '\\infty', 0xA8: '\\angle',
            0xA2: '\\prime', 0xA7: '\\int', 0xA1: '\\partial',
            0xA4: '\\nabla', 0xA6: '\\exists', 0xA9: '\\forall',
            0xD1: '\\parallel', 0xD5: '\\perp', 0xD7: '\\equiv',
            0xBB: '\\approx', 0xAA: '\\sum', 0xD0: '\\prod',
            0xD3: '\\cup',
        }
        return symbol_map.get(code, f'[sym:{code}]')

    def _parse_font_def(self):
        """解析 FONT_DEF 记录 (0x11): encoding_index(1) + font_name(null-term)"""
        idx = self._read_byte()       # encoding index
        name = self._read_null_terminated_string()
        self.font_map[idx] = name

    def _parse_encoding_def(self):
        """解析 ENCODING_DEF 记录 (0x13) — 跳过"""
        name = self._read_null_terminated_string()

    def _parse_eqn_prefs(self):
        """解析 EQN_PREFS (0x12): options(1) + size_count(1) + nibble_array + style_count(1) + style_defs"""
        if self.pos >= len(self.data):
            return

        # Step 0: 跳过 options 字节
        options = self._read_byte()

        # Step 1: 消费 dimension array（nibble 编码的尺寸字符串）
        size_count = self._read_byte()
        self._consume_dimension_array(size_count)

        # Step 2: 消费 style definition array
        if self.pos >= len(self.data):
            return
        style_count = self._read_byte()
        self._consume_style_definitions(style_count)

    def _consume_dimension_array(self, count: int):
        """消费 EQN_PREFS 的 nibble 压缩尺寸数组，并自校准找到 style_count 边界"""
        start = self.pos
        # 硬性步长上限: count * 12 为自适应估值, 200 bytes 为绝对安全帽
        scan_limit = min(len(self.data), start + min(count * 12, 200))
        best_pos = start

        for scan_pos in range(start, min(scan_limit, len(self.data) - 1)):
            candidate_count = self.data[scan_pos]
            if not (1 <= candidate_count <= 20):
                continue
            # 验证 style definitions: 每个 index 应为 uint8 (0-10), char_style (0-255)
            test_pos = scan_pos + 1
            valid = True
            for _ in range(candidate_count):
                if test_pos >= len(self.data):
                    valid = False
                    break
                idx = self.data[test_pos]
                test_pos += 1
                if idx > 10:               # FONT_DEF index 超出合理范围
                    valid = False
                    break
                if idx != 0 and test_pos < len(self.data):
                    test_pos += 1  # char_style byte
            if valid and test_pos < len(self.data):
                # 验证后一字节是否为 MTEF 方程数据 (LINE/CHAR/TMPL tag)
                next_tag = self.data[test_pos]
                next_rt = next_tag & 0x0F if next_tag not in (0x10, 0x11, 0x12, 0x13) else -1
                if next_rt in (0x01, 0x02, 0x03, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E):
                    best_pos = scan_pos
                    break

        if best_pos == start:
            log.warning(
                f"EQN_PREFS: 自校准未找到 style_count 边界 (count={count}, "
                f"range=[{start}, {scan_limit}), data: {self.data[start:min(start+30, len(self.data))].hex()}"
            )
        self.pos = best_pos

    def _consume_style_definitions(self, count: int):
        """消费 EQN_PREFS 的 style definition 数组（uint8 index + optional char_style）"""
        for _ in range(count):
            if self.pos >= len(self.data):
                break
            index = self._read_byte()             # FONT_DEF index, 0 = unused
            if index != 0 and self.pos < len(self.data):
                self._read_byte()                 # char_style byte

    def _parse_color_def(self):
        """解析 COLOR_DEF 记录 (0x10): R(2) + G(2) + B(2) + spare(2)"""
        self._read_uint16_le()  # R
        self._read_uint16_le()  # G
        self._read_uint16_le()  # B
        self._read_uint16_le()  # spare

    # ========================================================================
    # Pre-AST: SIZE 前缀消费
    # ========================================================================

    def _consume_size_prefix(self):
        """在 AST 建树前仅消费 SIZE (0x09) 记录"""
        max_steps = 10
        for _ in range(max_steps):
            if self.pos >= len(self.data):
                return
            tag = self.data[self.pos]                    # 窥探
            if tag in (LINE, CHAR, TMPL, END, FULL, SUB, SUB2, SYM, SUBSYM):
                return                                    # AST 记录 — 停手
            tag = self._read_byte()                      # 消费
            if self.pos >= len(self.data):
                return
            options = self._read_byte()
            if tag == SIZE:
                self._parse_size_record()
            else:
                return

    # ========================================================================
    # 主解析循环
    # ========================================================================

    def parse(self) -> str:
        """入口：解析 MTEF 字节流，返回 LaTeX 字符串"""
        self.pos = 0
        # 跳过 MTEF v5 头部的元数据记录，直达方程数据
        self._skip_header()

        if self.pos >= len(self.data):
            return ''

        # Pre-AST: 消费 SIZE/FULL/SUB/SUB2/SYM/SUBSYM 前缀
        self._consume_size_prefix()

        if self.pos >= len(self.data):
            return ''

        parts = self._parse_line()
        return ''.join(parts)

    def _skip_header(self):
        """跳过 MTEF v5 头部的 FONT_DEF / ENCODING_DEF / EQN_PREFS / COLOR_DEF 等元数据，直到方程数据起始"""
        try:
            if self.data[0] != 5:
                return
        except IndexError:
            return

        # 读取 header: version(1) + platform(1) + product(1) + major(1) + minor(1)
        self.pos = 5

        # 跳过 app key (null-terminated)
        try:
            end = self.data.index(0, self.pos)
            self.pos = end + 1
        except ValueError:
            return

        # 跳过 equation options (1 byte) + option value if present
        if self.pos < len(self.data):
            eq_opt = self.data[self.pos]
            self.pos += 1
            if eq_opt & 0x80:
                if self.pos < len(self.data):
                    self.pos += self.data[self.pos] + 1

        # === 元数据跳过循环（三层防御体系） ===
        while self.pos < len(self.data):
            tag = self.data[self.pos]

            # ── 第一道防线：扩展记录绝对优先权 ──
            # MTEF v5 扩展记录 (0x10-0x13) 的低 4 位与标准记录冲突，
            # 必须在任何掩码操作之前用完整 tag 字节判定。
            if tag == FONT_DEF:           # 0x11
                self.pos += 1
                self._parse_font_def()
            elif tag == ENCODING_DEF:     # 0x13
                self.pos += 1
                self._parse_encoding_def()
            elif tag == EQN_PREFS:        # 0x12
                self.pos += 1
                self._parse_eqn_prefs()
            elif tag == COLOR_DEF:        # 0x10
                self.pos += 1
                self._parse_color_def()
            else:
                # ── 非扩展记录：tag 字节即 record_type ──
                # 仅嗅探，不吞噬 options — 交由后继处理器消费
                if tag in (LINE, CHAR, TMPL, PILE, MATRIX, EMBELL, RULER,
                           SIZE, FULL, SUB, SUB2, SYM, SUBSYM):
                    break
                elif tag == END:
                    self.pos += 1
                    break
                else:
                    self.pos += 1

        # 对齐修正：style 定义后残留 1 字节 typesize 标签
        if self.pos < len(self.data) and self.data[self.pos] in (FULL, SUB, SUB2, SYM, SUBSYM):
            self.pos += 1  # 未知记录，保守跳过 1 字节

    def _parse_line(self) -> list:
        """解析 LINE 记录 → LaTeX 片段列表"""
        self._depth += 1
        if self._depth > 32:
            self._depth -= 1
            return ['[NESTING_LIMIT]']
        parts = []
        while self.pos < len(self.data):
            tag = self._read_byte()
            if tag == END:
                self._depth -= 1
                return parts
            if self.pos >= len(self.data):
                break
            options = self._read_byte()

            if tag == LINE:
                if options & xfNULL:
                    continue
                if options & xfLMOVE:
                    self._read_nudge()
                if options & xfLSPACE:
                    self._read_uint16_le()
                subs = self._parse_line()
                parts.extend(subs)
            elif tag == CHAR:
                if options & xfLMOVE:
                    self._read_nudge()
                typeface = self._read_byte()
                ch = self._read_char_value(typeface, options)
                # 防污染断言: 控制字符/私用区字符 → 游标可能错位
                if ch and len(ch) == 1:
                    cp = ord(ch[0])
                    if cp <= 0x001F or 0xE000 <= cp <= 0xF8FF:
                        ctx_start = max(0, self.pos - 8)
                        ctx_end = min(len(self.data), self.pos + 4)
                        log.debug(
                            f"[AST_POLLUTION] CHAR→U+{cp:04X} (tf={typeface}) "
                            f"near pos={self.pos}: {self.data[ctx_start:ctx_end].hex()}"
                        )
                parts.append(ch)
                if options & xfEMBELL:
                    self._parse_embell_list()
            elif tag == TMPL:
                self.pos -= 1           # 回退: TMPL options 用 nibble-packed
                options = tag >> 4
                if options & xfLMOVE:
                    self._read_nudge()
                selector = self._read_byte()
                variation = self._read_byte()
                tmpl_opts = self._read_byte()

                # Parent Context Callback: subscript/superscript/big_ops/integrals
                # 需要从父 LINE 栈顶取最后一个 CHAR 作为 base
                if selector in (0x00, 0x0F, 0x15, 0x16, 0x17, 0x18, 0x19,
                                 0x1D, 0x1E, 0x21, 0x22, 0x23):
                    base = parts.pop() if parts else ''
                else:
                    base = ''

                latex = self._template_to_latex(selector, variation, tmpl_opts, base)
                parts.append(latex)
            elif tag == PILE:
                if options & xfLMOVE:
                    self._read_nudge()
                self._read_byte()  # halign
                self._read_byte()  # valign
                parts.extend(self._parse_line())
            elif tag == SIZE:
                self._parse_size_record()
            elif tag in (FULL, SUB, SUB2, SYM, SUBSYM):
                self.pos -= 1       # typesize 标签用 nibble-packed, 回退 options
                pass
            elif tag == FONT_DEF:
                self._parse_font_def()
            elif tag == FONT_STYLE_DEF:
                # skip for now
                self._read_byte()
                self._read_byte()
            elif tag == COLOR_DEF:                     # 0x10
                if options == 0: self.pos += 6          # index(2) + RGB(4) = 6 bytes
            elif tag == COLOR:                          # 0x0F
                pass                                     # 消费 tag+options, 无 payload
            else:
                # unknown → skip
                pass

        return parts

    def _template_to_latex(self, selector: int, variation: int, opts: int,
                            base: str = '') -> str:
        """TMPL → LaTeX。base 由父 LINE 通过 parts.pop() 传入（后置修饰语义）"""
        if selector == 0x00:
            # selector=0x00 → 私有下标变体 (angle 劫持), 跳过前导 END(0x00/0x10) 后消费
            if self.pos < len(self.data) and self.data[self.pos] == END:
                self.pos += 1
            subs = self._parse_line()
            return f'{base}_{{{subs[0]}}}' if subs else base

        subs = self._parse_line()

        if selector == 0x0C:  # square root
            arg = ''.join(subs)
            return f'\\sqrt{{{arg}}}'
        elif selector == 0x0B:  # radical (nth root)
            if len(subs) >= 2:
                deg = ''.join(subs[:-1])
                rad = subs[-1]
                return f'\\sqrt[{deg}]{{{rad}}}'
            return f'\\sqrt{{{subs[0] if subs else ""}}}'
        elif selector == 0x0D:  # fraction
            num = subs[0] if len(subs) > 0 else ''
            den = subs[1] if len(subs) > 1 else ''
            return f'\\frac{{{num}}}{{{den}}}'
        elif selector == 0x0F:  # sub/superscript — base from parent context
            if variation == 0:  # superscript only
                return f'{base}^{{{subs[0] if subs else ""}}}'
            elif variation == 1:  # subscript only
                return f'{base}_{{{subs[0] if subs else ""}}}'
            elif variation == 2:  # both
                sub = subs[0] if len(subs) > 0 else ''
                sup = subs[1] if len(subs) > 1 else ''
                return f'{base}_{{{sub}}}^{{{sup}}}'
            return base
        elif selector == 0x10:  # underbar
            arg = ''.join(subs)
            return f'\\underline{{{arg}}}'
        elif selector == 0x11:  # overbar
            arg = ''.join(subs)
            return f'\\overline{{{arg}}}'
        elif selector in (0x00, 0x01):  # fences (selector=0x00 已在上方处理, 此处仅 0x01)
            fences = {'paren': ('(', ')'), 'brack': ('[', ']'),
                      'brace': ('\\{', '\\}'), 'bar': ('|', '|'),
                      'angle': ('\\langle', '\\rangle')}
            fencetype = TMPL_SELECTORS.get(selector, ('paren', ''))[0]
            f = fences.get(fencetype, ('(', ')'))
            arg = ''.join(subs)
            return f'{f[0]}{arg}{f[1]}'
        elif selector in (0x15, 0x16, 0x17, 0x18, 0x19):  # big ops
            ops = {0x15: '\\sum', 0x16: '\\prod', 0x17: '\\coprod',
                   0x18: '\\bigcup', 0x19: '\\bigcap'}
            op = ops.get(selector, '\\sum')
            return f'{op}_{{{base}}}' if base else op
        elif selector in (0x1D, 0x1E):  # integral-style big op
            return f'\\int_{{{base}}}' if base else '\\int'
        elif selector in (0x21, 0x22, 0x23):  # integrals
            ints = {0x21: '\\int', 0x22: '\\iint', 0x23: '\\iiint'}
            op = ints.get(selector, '\\int')
            return f'{op}_{{{base}}}' if base else op
        else:
            # 未知模板 → 降级为 \boxed{} 包裹子节点
            arg = ''.join(subs) if subs else ''
            return f'\\boxed{{{arg}}}' if arg else ''

    def _parse_size_record(self):
        """解析 SIZE 记录"""
        sizes = ['', 'szFULL', 'szSUB', 'szSUB2', 'szSYM', 'szSUBSYM']
        b = self._read_byte()
        if b == 101:
            self._read_uint16_le()
        elif b == 100:
            self._read_byte()
            self._read_uint16_le()
        else:
            self._read_byte()

    def _parse_embell_list(self):
        """跳过修饰列表"""
        while self.pos < len(self.data):
            tag = self._read_byte()
            if self.pos >= len(self.data):
                return
            options = self._read_byte()
            if tag == END:
                return
            elif tag == EMBELL:
                if options & xfLMOVE:
                    self._read_nudge()
                self._read_byte()
            else:
                pass  # 未知记录，已消费 tag+options


# ============================================================================
# 顶层接口
# ============================================================================

def mtef_to_latex(mtef_bytes: bytes) -> str:
    """
    将 MTEF v5 二进制转换为 LaTeX。
    先尝试 Fast-Path (TeX 元数据提取)，失败时降级到递归解析。
    """
    from .mtef_fast import extract_tex_metadata

    # Track 1: Fast-Path
    tex = extract_tex_metadata(mtef_bytes)
    if tex:
        return tex

    # Track 2: Slow-Path
    try:
        parser = MTEFParser(mtef_bytes)
        return parser.parse()
    except Exception as e:
        log.warning(f"MTEF 解析失败: {e}")
        return "[公式: MTEF 解析失败]"
