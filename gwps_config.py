# -*- coding: utf-8 -*-
"""公文一键排版助手（重建版）—— 常量与默认配置

行为对齐原版逆向结果，字体/正则/扩展名与原始 exe 一致；
默认值按 GB/T 9704—2012 校准。所有键沿用原版 DEFAULT_CONFIG 命名，
便于与 GUI 参数面板一一对应。
"""

APP_NAME = "公文一键排版助手"
APP_VERSION = "1.0.8"

# ---------- 版头样式 ----------
HEADER_STYLE_NONE = "无红色版头"
HEADER_STYLE_BIG = "正式发文（大红头）"
HEADER_STYLE_SMALL = "简化文头（小红头）"
HEADER_STYLE_OPTIONS = [HEADER_STYLE_NONE, HEADER_STYLE_BIG, HEADER_STYLE_SMALL]

# ---------- 字体常量（与原版一致）----------
BODY_FONT = "仿宋_GB2312"
TITLE_FONT = "方正小标宋简体"
H1_FONT = "黑体"
H2_FONT = "楷体_GB2312"
WESTERN_FONT = "Times New Roman"
PAGE_NUMBER_FONT = "宋体"

MAST_RED_RGB = (210, 31, 43)
MAST_RED_HEX = "D21F2B"

# ---------- 支持的文件扩展名 ----------
SUPPORTED_EXTENSIONS = (".docx", ".doc", ".wps", ".txt", ".md")

# ---------- 序数正则（与原版逐字一致）----------
CHINESE_NUM = r"[一二三四五六七八九十百千万零]+"
RE_HAS_CHINESE = r"[一-\u9fff]"
# 第一层：一、二、……        RE_H1
# 第二层：（一）（二）……     RE_H2
# 第三层：1. 2. ……（半/全角点）RE_H3
# 第四层：（1）（2）……       RE_H4
RE_H1_PATTERN = r"^" + CHINESE_NUM + r"\s*、"
RE_H2_PATTERN = r"^[（(]" + CHINESE_NUM + r"[）)]"
RE_H3_PATTERN = r"^\d+\s*[.．]"
RE_H4_PATTERN = r"^[（(]\d+[）)]"

# ---------- 默认配置（GB/T 9704 校准）----------
# 版面：天头37 订口28 版心156×225 → 下≈35 右≈26（mm）；页码距版心下缘7mm
# 字号：二号=22pt 三号=16pt 四号=14pt；行距固定≈28磅（22行/28字网格）
DEFAULT_CONFIG = {
    # 页面
    "margin_top_cm": 3.7,
    "margin_bottom_cm": 3.5,
    "margin_left_cm": 2.8,
    "margin_right_cm": 2.6,
    "footer_distance_cm": 2.8,
    "page_number_font": PAGE_NUMBER_FONT,
    "page_number_size": 14,
    # 标题
    "title_font": TITLE_FONT,
    "title_size": 22,
    "subtitle_font": TITLE_FONT,
    "subtitle_size": 16,
    "h1_font": H1_FONT,
    "h1_size": 16,
    "h2_font": H2_FONT,
    "h2_size": 16,
    # 正文
    "body_font": BODY_FONT,
    "body_size": 16,
    "western_font": WESTERN_FONT,
    "line_spacing": 28,
    "title_line_spacing": 33,
    "indent_chars": 2,
    "title_before_blank": 0,
    "title_after_blank": 0,
    # 落款
    "sign_blank": 1,
    "sign_right_chars": 4,
    "first_para_as_title": True,
    # 版头
    "header_style": HEADER_STYLE_NONE,
    "mast_font": TITLE_FONT,
    "mast_size_big": 26,
    "mast_size_small": 22,
    "mast_top_blank_lines": 2,
    "mast_top_gap_mm": 35,
    "grid_line_pitch": 580,
    "grid_char_space": 0,
    "org_name": "",
    "doc_number": "",
    "signer": "",
    "issuer": "",
    "date_text": "",
    # 智能开关
    "remove_spaces": True,
    "normalize_punctuation": True,
    "normalize_western": True,
    "remove_blank_lines": True,
    "set_outline": True,
    "enable_attachment_formatting": True,
    "enable_table_formatting": True,
    "table_body_center": True,
    "enable_image_formatting": True,
    "figcaption_size": 12,
    "enable_footer_notes": True,
    "h1_bold": False,
    "h2_bold": False,
    "h3_bold": False,
    "h4_bold": False,
    "cc": "",
    "print_org": "",
    "print_date": "",
    # 表格
    "table_header_font": H1_FONT,
    "table_font": BODY_FONT,
    "table_size": 14,
    "table_line_spacing": 22,
    "table_row_height_cm": 0.8,
    "table_width_percent": 100,
    "table_header_bold": True,
    "table_auto_col_width": True,
    "table_unified_borders": True,
    "table_smart_align": True,
    "table_border_size_pt": 0.5,
}
