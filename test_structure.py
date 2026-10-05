# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gwps_structure as S

def esc(s): return s.encode('unicode_escape').decode('ascii')

# cn_num
assert S.cn_num(1) == '一'
assert S.cn_num(2) == '二'
assert S.cn_num(10) == '十'
assert S.cn_num(11) == '十一', esc(S.cn_num(11))
assert S.cn_num(20) == '二十', esc(S.cn_num(20))
assert S.cn_num(21) == '二十一', esc(S.cn_num(21))
assert S.cn_num(100) == '一百', esc(S.cn_num(100))

# heading_level
assert S.heading_level('一、总体要求') == 1
assert S.heading_level('（一）加强领导') == 2
assert S.heading_level('1.第一项') == 3
assert S.heading_level('（1）细化') == 4
assert S.heading_level('这是一段正文') is None
assert S.heading_level('  （二）带前导空格 ') == 2

# render_ordinal
assert S.render_ordinal(1, 0, True) == '一、', esc(S.render_ordinal(1, 0, True))
assert S.render_ordinal(2, 1, True) == '（二）', esc(S.render_ordinal(2, 1, True))
assert S.render_ordinal(3, 2, False) == '3.'
assert S.render_ordinal(1, 3, False) == '（1）', esc(S.render_ordinal(1, 3, False))

# is_cn_numfmt
assert S.is_cn_numfmt('chineseCounting') is True
assert S.is_cn_numfmt('decimal') is False

print('STRUCTURE_OK')
