"""
Track 1 (Fast-Path): 从 MTEF v5 字节流提取原始 TeX 元数据
- 微秒级提取，零递归开销
- MathType 7.x 将 TeX 输入语言以明文形式存储在 MTEF 头部
- 格式: "...TeX Input Language\0{original_tex}\0..."
"""
import struct


def extract_tex_metadata(mtef_bytes: bytes) -> str | None:
    """
    从 MTEF 字节流搜索并提取原始 TeX 输入。
    若公式通过 TeX 语法输入（而非 MathType UI 点击），MathType 会保存原始 TeX 明文。
    失败时返回 None，由上游降级到 Track 2。

    数据格式：
      b'TeX Input Language\x00' + TeX 源码 + \x00
    """
    marker = b'TeX Input Language\x00'
    idx = mtef_bytes.find(marker)
    if idx < 0:
        return None

    try:
        start = idx + len(marker)
        # 查找终止符
        end_pos = mtef_bytes.index(0, start)
        raw = mtef_bytes[start:end_pos]
        # 编码容错：历史手稿可能混入非 UTF-8 字符
        return raw.decode('utf-8', errors='replace').strip()
    except (ValueError, UnicodeDecodeError):
        return None


def extract_app_key(mtef_bytes: bytes) -> str | None:
    """提取 MTEF v5 应用标识键 (如 'DSMT7')"""
    try:
        if len(mtef_bytes) < 11:
            return None
        end = mtef_bytes.index(0, 5)
        return mtef_bytes[5:end].decode('ascii', errors='replace')
    except (ValueError, UnicodeDecodeError):
        return None


def get_mtef_version(mtef_bytes: bytes) -> int:
    """读取 MTEF 版本号（字节 0）"""
    return mtef_bytes[0] if len(mtef_bytes) > 0 else 0


def get_platform(mtef_bytes: bytes) -> str:
    """读取生成平台（字节 1）"""
    if len(mtef_bytes) < 2:
        return 'unknown'
    return 'Windows' if mtef_bytes[1] == 1 else 'Mac'


def read_nudge(buffer: bytes, pos: int = 0) -> tuple[int, int, int]:
    """
    读取 nudge 偏移量，返回 (dx, dy, 消耗字节数)
    小端序锁死：所有 16 位值使用 <H 显式声明
    """
    if len(buffer) - pos < 2:
        return 0, 0, 0
    dx = struct.unpack('<b', buffer[pos:pos+1])[0]
    dy = struct.unpack('<b', buffer[pos+1:pos+2])[0]
    if dx == 128 and dy == 128 and len(buffer) - pos >= 6:
        dx = struct.unpack('<h', buffer[pos+2:pos+4])[0]
        dy = struct.unpack('<h', buffer[pos+4:pos+6])[0]
        return dx, dy, 6
    return dx, dy, 2


def is_inline_equation(mtef_bytes: bytes) -> bool:
    """
    解析 MTEF v5 头部 Equation Options 字节的 Bit 0，
    判断公式是 Inline (行内) 还是 Display (块级)。

    MTEF v5 头部结构:
      [0]     version (0x05)
      [1]     platform
      [2-4]   product + version
      [5..n]  app_key, null-terminated (e.g., "DSMT7\0")
      [n+1]   equation options byte (bit 0 = inline)

    容错：非 v5、截断或异常时默认返回 True (inline)，保证出稿不断裂。
    """
    try:
        if len(mtef_bytes) < 2 or mtef_bytes[0] != 5:
            return True  # 非 MTEF v5，默认当作 inline 安全兜底

        # 动态寻找 app_key 的 null 终止符（起点索引 5）
        app_key_end = mtef_bytes.index(0, 5)
        eq_options_idx = app_key_end + 1

        if eq_options_idx >= len(mtef_bytes):
            return True

        eq_options = mtef_bytes[eq_options_idx]
        return (eq_options & 0x01) != 0

    except (ValueError, IndexError):
        return True
