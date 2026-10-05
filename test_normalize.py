# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gwps_normalize as G

def esc(s):
    return s.encode('unicode_escape').decode('ascii')

# 1) 去空格：全角/多余空格删除，英文词间单空格保留
r = G.remove_all_spaces('你　好  world  end')
assert '你好' in r and 'world' in r and 'end' in r, esc(r)

# 2) 直引号 → 中文上下引号
p = G.normalize_punctuation('他说"你好"，结束')
assert '\u201c你好\u201d' in p, esc(p)

# 3) 半角句读 → 全角；日期连字符不被转成一字线
p = G.normalize_punctuation('2026-08-31,好的;')
assert '，' in p and '；' in p and '2026-08-31' in p, esc(p)

# 4) 省略号合并
p = G.normalize_punctuation('等等。。。')
assert '\u2026\u2026' in p, esc(p)

# 5) 序号点 '1.' 保留半角
p = G.normalize_punctuation('第一层下面 1. 要点')
assert '1.' in p, esc(p)

# 6) 全角数字 → 半角
p = G.normalize_punctuation('２０２6年')
assert '2026年' in p, esc(p)

# 7) 括号含中文 → 全角
p = G.normalize_punctuation('(中文测试)')
assert '\uff08中文测试\uff09' in p, esc(p)

print('NORMALIZE_OK')
print('sample_quotes:', esc(G.normalize_punctuation('他说"这是"外\'内\'结束')))
