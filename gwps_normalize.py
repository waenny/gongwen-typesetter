# -*- coding: utf-8 -*-
"""文本规范化层（重建版）

职责：在排版前把任意来源的文字规范化为公文用字——
  1) 去除多余空格（保留英文词间单个半角空格）
  2) 全半角统一：全角数字/字母→半角；中文语境标点半角→全角
  3) 引号智能配对：直引号 / 异形引号 → 中文上下引号 “ ” ‘ ’
  4) 括号：含中文的 () [] → 全角 （） ［］
  5) 省略号：连续的 . ． 。 → ……
  6) 数字间连字符 → 一字线 —；破折号/波浪号规范化
说明：本层只处理"文字/用字"，不碰版式几何（字体、页边距等由引擎负责）。
"""
import re

# 需要删除的各种"空白"字符（半角/全角/不间断/零宽/制表等）
SPACE_CHARS = set(
    "\t\n\r \u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005"
    "\u2006\u2007\u2008\u2009\u200a\u2028\u2029\u202f\u205f\u3000\ufeff\u200b\u200c\u200d"
)
SPACE_CHARS.remove("\n")  # 换行交由段落逻辑处理，这里不吞
SPACE_CHARS.remove("\r")

_DATE_HYPHEN_RE = re.compile(r"\d{4}-\d{1,2}-\d{1,2}")


def _is_ascii_alnum(ch):
    return ch.isascii() and ch.isalnum()


def _is_cjk(ch):
    return bool(ch) and "\u4e00" <= ch <= "\u9fff"


def _keep_space(a, b):
    """仅在“两侧都是英文数字”或“一侧中文一侧英文/数字”时保留一个半角空格。"""
    aa, bb = _is_ascii_alnum(a), _is_ascii_alnum(b)
    if aa and bb:
        return True
    if (_is_cjk(a) and bb) or (_is_cjk(b) and aa):
        return True
    return False


def remove_all_spaces(text):
    """删除空格类字符；仅保留英文单词/短语之间的单个半角空格。"""
    if not text:
        return text
    out = []
    n = len(text)
    i = 0
    while i < n:
        ch = text[i]
        if ch in SPACE_CHARS:
            # 判断是否"英文词间空格"：左右都是 ASCII 字母/数字
            if ch in (" ", "\u00a0"):
                prev = text[i - 1] if i > 0 else ""
                nxt = ""
                j = i + 1
                while j < n and text[j] in SPACE_CHARS:
                    j += 1
                nxt = text[j] if j < n else ""
                if _keep_space(prev, nxt):
                    out.append(" ")
                    i = j
                    continue
            i += 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _fullwidth_to_halfwidth(ch):
    code = ord(ch)
    if code == 0x3000:
        return " "
    if 0xFF01 <= code <= 0xFF5E:
        return chr(code - 0xFEE0)
    return ch


def _is_date_hyphen(text, i):
    """位置 i 的 '-' 是否属于日期（如 2026-08-31），是则保留不转一字线。"""
    for m in _DATE_HYPHEN_RE.finditer(text):
        if m.start() <= i <= m.end():
            return True
    return False


def _normalize_digits_letters(text):
    """仅将全角数字/字母转半角；不动全角标点（、（）：，等）。"""
    out = []
    for ch in text:
        code = ord(ch)
        if (0xFF10 <= code <= 0xFF19) or (0xFF21 <= code <= 0xFF3A) or (0xFF41 <= code <= 0xFF5A):
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def _normalize_ellipsis(text):
    """连续的 . ． 合并为省略号 ……（不误伤 '1.' 之类单点序号）。"""
    return re.sub(r"[.．。]{2,}", "……", text)


def _normalize_period(text):
    """句末英文句点在中文语境下转句号。保留小数点与 '1.' 序号点。"""
    out = []
    n = len(text)
    for i, ch in enumerate(text):
        if ch == ".":
            prev = text[i - 1] if i > 0 else ""
            nxt = text[i + 1] if i + 1 < n else ""
            # 小数点：两侧数字 → 保留
            if prev.isdigit() and nxt.isdigit():
                out.append(ch)
                continue
            # 序号点：前面是数字且后面非句读/空白后是汉字标题 → 保留半角（引擎 RE_H3 依赖）
            if prev.isdigit():
                out.append(ch)
                continue
            # 中文后 + 后面是中文/结尾 → 视作句号
            if prev and ord(prev) > 0x2E00:
                out.append("。")
                continue
            out.append(ch)
        else:
            out.append(ch)
    return "".join(out)


def _normalize_brackets(text):
    """含中文内容的 () [] 转全角 （） ［］。"""
    chars = list(text)
    stack_paren = []
    stack_square = []
    for idx, ch in enumerate(chars):
        if ch == "(":
            stack_paren.append(idx)
        elif ch == ")" and stack_paren:
            o = stack_paren.pop()
            if re.search(r"[\u4e00-\u9fff]|\d", "".join(chars[o + 1:idx])):
                chars[o] = "（"
                chars[idx] = "）"
        elif ch == "[":
            stack_square.append(idx)
        elif ch == "]" and stack_square:
            o = stack_square.pop()
            if re.search(r"[\u4e00-\u9fff]", "".join(chars[o + 1:idx])):
                chars[o] = "［"
                chars[idx] = "］"
    return "".join(chars)


# 引号规范化：双引号家族 / 单引号家族
_OPEN_BEFORE_D = set("（(，,。；;：:？?！!、…—–·「『《〈〔【")
_CLOSE_AFTER_D = set("，,。；;：:？?！!、…—–·」』》〉〕】）)")


def _normalize_quotes(text):
    """直引号/异形引号 → 中文上下引号，采用"上下文 + 配对"判断开闭方向。"""
    if not text:
        return text
    # 先归一异形双引号到 "，异形单引号到 '
    for c in ("„", "‟", "〝", "〞", "＂"):
        text = text.replace(c, '"')
    for c in ("‚", "‛", "＇"):
        text = text.replace(c, "'")
    for c in ("「", "『"):
        text = text.replace(c, '"')
    for c in ("」", "』"):
        text = text.replace(c, '"')

    out = []
    dbl_open = False   # 当前双引号是否处于"待闭合"（已开）状态
    # 逐字符扫描，按配对决定 “ ”
    n = len(text)
    for i, ch in enumerate(text):
        if ch == '"':
            prev = text[i - 1] if i > 0 else ""
            nxt = text[i + 1] if i + 1 < n else ""
            # 开：前无/前是开侧标点/前空白；且(后非空白非闭合类)
            is_open = (prev == "" or prev in _OPEN_BEFORE_D or prev in " \t") and nxt not in _CLOSE_AFTER_D and nxt != ""
            if not dbl_open:
                out.append("“")
                dbl_open = True
            else:
                out.append("”")
                dbl_open = False
            _ = is_open  # 配对优先于上下文，保留变量以便后续调优
        elif ch == "'":
            prev = text[i - 1] if i > 0 else ""
            # 英文缩写撇号（字母'字母）→ ’
            if prev.isascii() and prev.isalnum():
                out.append("’")
            else:
                # 单引号：按开合交替（外双内单由整体顺序近似处理）
                if getattr(_normalize_quotes, "_s_open", False):
                    out.append("’")
                    _normalize_quotes._s_open = False
                else:
                    out.append("‘")
                    _normalize_quotes._s_open = True
        else:
            out.append(ch)
    _normalize_quotes._s_open = False
    return "".join(out)


def normalize_punctuation(text):
    """统一标点为公文用字。"""
    if not text:
        return text
    text = _normalize_digits_letters(text)
    # 半角句读 → 全角
    table = {",": "，", ";": "；", "?": "？", "!": "！", ":": "："}
    out = []
    n = len(text)
    for i, ch in enumerate(text):
        if ch in table:
            if ch in ",;:" and i > 0 and i + 1 < n and text[i - 1].isdigit() and text[i + 1].isdigit():
                out.append(ch)  # 时间 12:30 / 千分位 1,000 / 比例 保留半角
            elif ch == ":" and i + 1 < n and text[i + 1] == "/":
                out.append(ch)  # URL/协议 http:// 等
            else:
                out.append(table[ch])
        elif ch == "~":
            out.append("～")
        else:
            out.append(ch)
    text = "".join(out)
    # 破折号：-- / —— 归一为 ——
    text = text.replace("--", "——").replace("――", "——")
    # 仅年份区间转一字线（19xx/20xx-19xx/20xx），不误伤电话/编号/型号
    text = re.sub(r"((?:19|20)\d{2})-((?:19|20)\d{2})", r"\1—\2", text)
    text = _normalize_ellipsis(text)
    text = _normalize_period(text)
    text = _normalize_brackets(text)
    text = _normalize_quotes(text)
    return text


def normalize_text(text, config=None):
    """按配置规范化：先标点（此时空格仍在，利于识别序号），再去空格。"""
    config = config or {}
    if config.get("normalize_punctuation", True):
        text = normalize_punctuation(text)
    if config.get("remove_spaces", True):
        text = remove_all_spaces(text)
    return text
