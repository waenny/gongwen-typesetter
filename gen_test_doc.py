# -*- coding: utf-8 -*-
"""生成一份“故意排版混乱”的公文测试文档到桌面，用于检验助手自动排版能力。"""
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

DESK = os.path.join(os.environ.get("USERPROFILE", "."), "Desktop")
docx_out = os.path.join(DESK, "公文排版测试.docx")

doc = Document()

def para(text="", center=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    if center:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p

# 1) 大标题：前面塞全角空格+半角空格、结尾带句读（测标题识别/去空格/句读）
p = para("　 　关于开展2026年度安全生产大检查工作的通知。")

# 2) 主送机关：半角冒号、没顶格前有空格
para(" 各有关单位:")

# 3) 正文：半角逗号、双半角空格、不间断空格、零宽空格混入
p = doc.add_paragraph()
p.add_run("为 切实做好安全生产 工作，现将有关事项通\u00a0知如下，请认真\u200b贯彻落实,")

# 4) 一级标题（正常序数，但内部有多余空格）
para("一、总体要求")

# 5) 二级标题（独立，应整行楷体）
para("（一）提 高政治站位")

# 6) 二级引导句：（二）标题。正文… 应“（二）…标题。”楷体，其后仿宋
para("（二）压实工作责任。各单位要迅速行动，结合实际制定方案，确保各项任务落地见效、取得实效。")

# 7) 三级序数（仿宋，不加粗）
para("1.细化工作方案")
# 8) 四级序数
para("（1）明确责任分工")

# 9) 含半角引号 + 省略号 + 破折号 + 错别字 + 语句不通 + 时间/千分位/IP（半角应保留）
p = doc.add_paragraph()
p.add_run("通知指出\"高度重视\"，要坚决防止事故的发生，做到防微杜渐迫不急待，确保万无一失。会议定在12:30召开，预算1,000万元，服务器地址192.168.1.10:8080，等等...都要提前布置。这次大检查历时两个月--覆盖面很广。")

# 10) 软回车（段内换行）应作为同一段处理
p = doc.add_paragraph()
r = p.add_run("第二，强化督导检查")
r.add_break()
p.add_run("对重点区域开展不间断巡查")

# 11) 段落之间多空行（应被清理）
para("")
para("   ")
para("")

# 12) 附件说明（半角冒号、序号乱）
para("附件: 1.安全生产检查表 2.隐患整改清单")

# 13) 小表格（测表格套字体/表头加粗）
tb = doc.add_table(rows=2, cols=3)
tb.cell(0, 0).text = "序号"
tb.cell(0, 1).text = "检查项目"
tb.cell(0, 2).text = "结果"
tb.cell(1, 0).text = "1"
tb.cell(1, 1).text = "消防设施"
tb.cell(1, 2).text = "合格"

para("")
# 14) 落款：右空、成文日期带虚位（应去虚位）
para("某某市安全生产委员会办公室")
para("二〇二六年八月五日")   # 中文数字日期（测是否保留/或数字日期分支另测）
para("2026年08月05日")       # 阿拉伯数字带虚位（应→2026年8月5日）

doc.save(docx_out)
print("DOCX_SAVED", docx_out, os.path.getsize(docx_out), "bytes")

# 尝试转成真正的 .doc（Word/WPS COM）
doc_out = os.path.join(DESK, "公文排版测试.doc")
converted = False
try:
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    app = None
    for prog in ("Word.Application", "KWPS.Application"):
        try:
            app = win32com.client.DispatchEx(prog)
            break
        except Exception:
            app = None
    if app:
        try:
            app.Visible = False; app.DisplayAlerts = 0
        except Exception:
            pass
        d = app.Documents.Open(docx_out)
        d.SaveAs2(doc_out, FileFormat=0)  # 0 = Word 97-2003 .doc
        d.Close(SaveChanges=0)
        app.Quit()
        converted = True
except Exception as e:
    print("DOC_CONVERT_FAIL", repr(e))
if converted:
    print("DOC_SAVED", doc_out, os.path.getsize(doc_out), "bytes")
else:
    print("NO_DOC_USING_DOCX")
