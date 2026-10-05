# -*- coding: utf-8 -*-
"""结构识别层（重建版）—— 修复 A1/A2/A3 的核心

- heading_level(text)：与原版一致的四級序数判定（RE_H1..H4）。
- cn_num / render_ordinal：数字→公文规范序数文字。
- expand_auto_numbering(doc)：★A1★ 把 Word 自动编号（含"编号绑定到样式"的情况）
  展开成字面序数前缀，并移除 w:numPr 防止双份编号；解析不到或项目符号则安全跳过。
- find_title(blocks, cfg)：★A3★ 主/副标题识别（居中/长度/句读保护，两分支一致）。
"""
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from gwps_config import (
    RE_H1_PATTERN, RE_H2_PATTERN, RE_H3_PATTERN, RE_H4_PATTERN,
)

_RE_H1 = re.compile(RE_H1_PATTERN)
_RE_H2 = re.compile(RE_H2_PATTERN)
_RE_H3 = re.compile(RE_H3_PATTERN)
_RE_H4 = re.compile(RE_H4_PATTERN)

_CN = "零一二三四五六七八九"


def cn_num(n):
    """1→一 … 10→十 11→十一 21→二十一 110→一百一十 105→一百零五。"""
    try:
        n = int(n)
    except (TypeError, ValueError):
        return str(n)
    if n < 1 or n > 9999:
        return str(n)
    if n < 10:
        return _CN[n]
    if n < 20:
        return "十" + (_CN[n % 10] if n % 10 else "")
    if n < 100:
        return _CN[n // 10] + "十" + (_CN[n % 10] if n % 10 else "")
    if n < 1000:
        bai, r = divmod(n, 100)
        s = _CN[bai] + "百"
        if r == 0:
            return s
        if r < 10:
            return s + "零" + _CN[r]
        if r < 20:
            return s + "一" + cn_num(r - 10)
        return s + cn_num(r)
    qian, r = divmod(n, 1000)
    s = _CN[qian] + "千"
    if r == 0:
        return s
    if r < 100:
        return s + "零" + cn_num(r)
    return s + cn_num(r)


def is_cn_numfmt(numfmt):
    return numfmt in (
        "chineseCounting", "chineseCountingThousand", "chineseLegalSimplified",
        "ideographTraditional", "ideographModern",
    )


def _render_num(numfmt, n):
    if is_cn_numfmt(numfmt):
        return cn_num(n)
    if numfmt == "lowerLetter":
        return chr(ord("a") + (n - 1) % 26)
    if numfmt == "upperLetter":
        return chr(ord("A") + (n - 1) % 26)
    return str(n)


def render_ordinal(n, ilvl, chinese):
    """按层级渲染规范序数：偶数层=一、/1.；奇数层=（一）/（1）。"""
    paren = (int(ilvl) % 2 == 1)
    tok = cn_num(n) if chinese else str(n)
    if paren:
        return "（%s）" % tok
    return ("%s、" % tok) if chinese else ("%s." % tok)


def heading_level(text):
    """返回 1..4 或 None。"""
    t = (text or "").strip()
    if not t:
        return None
    if _RE_H1.match(t):
        return 1
    if _RE_H2.match(t):
        return 2
    if _RE_H3.match(t):
        return 3
    if _RE_H4.match(t):
        return 4
    return None


_EXISTING_PREFIX = re.compile(
    r"^\s*(?:[一二三四五六七八九十百千零]+\s*、|[（(][一二三四五六七八九十]+[）)]"
    r"|\d+\s*[.．]|[（(]\d+[）)])"
)


def _child_val(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        return None
    return el.get(qn("w:val"))


def _numpr_from(pPr):
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    nid = _child_val(numPr, "w:numId")
    ilvl = _child_val(numPr, "w:ilvl")
    if nid is None:
        return None
    return (nid, ilvl if ilvl is not None else "0")


def _resolve_numpr(para):
    """段落直接 numPr；否则沿样式链（含 basedOn）继承的 numPr。"""
    ref = _numpr_from(para._p.pPr)
    if ref:
        return ref
    try:
        style = para.style
    except Exception:
        style = None
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            st_pPr = style.element.find(qn("w:pPr"))
        except Exception:
            st_pPr = None
        ref = _numpr_from(st_pPr)
        if ref:
            return ref
        style = getattr(style, "base_style", None)
    return None


def _load_numbering(doc):
    """numId -> {ilvl: {numFmt, lvlText, start}}。缺 numbering.xml 时返回空表。"""
    info = {}
    try:
        root = doc.part.numbering_part.element
    except Exception:
        return info
    abstract = {}
    for an in root.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        lvls = {}
        for lvl in an.findall(qn("w:lvl")):
            il = lvl.get(qn("w:ilvl")) or "0"
            lvls[il] = {
                "numFmt": _child_val(lvl, "w:numFmt") or "decimal",
                "lvlText": _child_val(lvl, "w:lvlText"),
                "start": _child_val(lvl, "w:start"),
            }
        abstract[aid] = lvls
    for num in root.findall(qn("w:num")):
        nid = num.get(qn("w:numId"))
        aref = _child_val(num, "w:abstractNumId")
        for il, spec in abstract.get(aref, {}).items():
            info[(nid, str(il))] = spec
    return info


def _render_marker(num_id, ilvl_i, counters, numbering, spec):
    lvltext = spec.get("lvlText")
    if lvltext and "%" in lvltext:
        def rep(m):
            k = int(m.group(1)) - 1
            sk = numbering.get((num_id, str(k))) or {}
            return _render_num(sk.get("numFmt", "decimal"), counters.get((num_id, k), 1))
        return re.sub(r"%(\d+)", rep, lvltext)
    chinese = is_cn_numfmt(spec.get("numFmt", "decimal"))
    return render_ordinal(counters.get((num_id, ilvl_i), 1), ilvl_i, chinese)


def _prepend_text(para, text):
    """在段落开头插入文字，不清空原 run（保留格式/图片/超链接）。"""
    new = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new.append(t)
    pPr = para._p.pPr
    if pPr is not None:
        pPr.addnext(new)
    else:
        para._p.insert(0, new)


def _strip_numpr(para):
    pPr = para._p.pPr
    if pPr is None:
        return
    for el in pPr.findall(qn("w:numPr")):
        pPr.remove(el)
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        for a in ("w:hanging", "w:hangingChars", "w:left", "w:leftChars"):
            ind.attrib.pop(qn(a), None)


def expand_auto_numbering(doc, log=None):
    """★A1★ 展开自动编号为字面序数并移除 numPr。解析不到/项目符号则跳过（不臆造）。"""
    def note(m):
        if log:
            try:
                log(m)
            except Exception:
                pass

    try:
        from docx.text.paragraph import Paragraph  # noqa
    except Exception as e:
        note("  > 无法导入 docx，跳过自动编号展开：%s" % e)
        return 0

    numbering = _load_numbering(doc)
    changed = 0
    skipped = 0
    counters = {}
    for para in doc.paragraphs:
        try:
            ref = _resolve_numpr(para)
        except Exception:
            continue
        if not ref:
            continue
        num_id, ilvl = ref
        try:
            ilvl_i = int(ilvl)
        except Exception:
            ilvl_i = 0
        spec = numbering.get((num_id, str(ilvl_i)))
        if not spec or spec.get("numFmt") in ("bullet", "none"):
            skipped += 1
            continue
        key = (num_id, ilvl_i)
        if key not in counters:
            st = spec.get("start")
            counters[key] = (int(st) - 1) if (st and str(st).isdigit()) else 0
        counters[key] += 1
        for k in list(counters):
            if k[0] == num_id and k[1] > ilvl_i:
                counters[k] = 0
        text = para.text or ""
        if _EXISTING_PREFIX.match(text.strip()):
            _strip_numpr(para)  # 已有规范序数，去掉 numPr 防双份
            continue
        try:
            marker = _render_marker(num_id, ilvl_i, counters, numbering, spec)
        except Exception:
            skipped += 1
            continue
        _prepend_text(para, marker)
        _strip_numpr(para)
        changed += 1
    note("  > 已展开 %d 处自动编号%s。" % (
        changed, ("；%d 处无法判定/项目符号，跳过" % skipped) if skipped else ""))
    return changed


def find_title(blocks, config=None):
    """★A3★ 识别主标题与副标题段落。居中与"首段兜底"两分支同等使用长度/句读保护。"""
    cfg = config or {}
    first_as_title = cfg.get("first_para_as_title", True)
    title_max = cfg.get("title_max_len", 60)
    subtitle_max = cfg.get("subtitle_max_len", 50)
    punct = re.compile(r"[。！？；]")

    def is_center(p):
        try:
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            return p.alignment == WD_ALIGN_PARAGRAPH.CENTER
        except Exception:
            return False

    title, subtitle = [], []
    n = len(blocks)
    start = 0
    while start < n and not (getattr(blocks[start], "text", "") or "").strip():
        start += 1
    if start >= n:
        return title, subtitle

    # 优先匹配“关于……的<文种>”作为主标题（跳过机关标志/发文字号等版头行）
    WZ = (r"(通知|通报|报告|请示|批复|意见|决定|决议|公告|通告|函|会议纪要"
          r"|规划|方案|办法|细则|规则|命令|议案)")
    tp = re.compile(r"关于.+?" + WZ + r"$")
    seen = 0
    for k in range(start, min(start + 8, n)):
        tt = (getattr(blocks[k], "text", "") or "").strip()
        if not tt:
            continue
        seen += 1
        if tp.search(re.sub(r"[。！？；，、：]+$", "", tt)):
            title.append(blocks[k])
            j = k + 1
            while j < n:
                t2 = (getattr(blocks[j], "text", "") or "").strip()
                if not t2:
                    j += 1
                    continue
                if not is_center(blocks[j]) or heading_level(t2):
                    break
                subtitle.append(blocks[j])
                j += 1
            return title, subtitle
        if seen >= 6:
            break

    first = blocks[start]
    ftext = (getattr(first, "text", "") or "").strip()
    if heading_level(ftext):
        return title, subtitle

    long_body = len(ftext) > title_max
    has_punct = bool(punct.search(ftext)) or ftext.endswith(("：", "，", "、"))
    centered = is_center(first)

    accept = False
    if centered:
        accept = not (long_body and has_punct)
    else:
        accept = first_as_title and not (long_body and has_punct)
    if accept:
        title.append(first)
    else:
        return title, subtitle

    i = start + len(title)
    while i < n:
        blk = blocks[i]
        t = (getattr(blk, "text", "") or "").strip()
        if not t:
            i += 1
            continue
        if heading_level(t) or not is_center(blk) or len(t) > subtitle_max:
            break
        subtitle.append(blk)
        i += 1
    return title, subtitle
