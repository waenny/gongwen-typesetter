# -*- coding: utf-8 -*-
"""字体检测回归：4 种公文字体都要能正确识别"已安装/缺失"。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gwps_fonts as F

st = F.check_status()
assert isinstance(st, list) and len(st) == 4, st
print("FONTS_OK", [(a, int(b)) for a, b in st])
