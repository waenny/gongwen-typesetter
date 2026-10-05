# -*- coding: utf-8 -*-
"""图形界面（重建版）—— 复原原版政务红主题/横幅徽标/菜单/联系开发者/自绘勾选框/图标。

保留 C1 命名 / C2 自检 / 一键装字体 / AI(可选降级)；线程安全日志、防重入、唯一临时名。
"""
import os
import sys
import queue
import tempfile
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import gwps_engine as E
import gwps_fonts as F
import gwps_settings as ST
from gwps_config import (
    APP_NAME, APP_VERSION, HEADER_STYLE_OPTIONS, HEADER_STYLE_NONE, DEFAULT_CONFIG,
)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _ROOT = TkinterDnD.Tk
    _HAS_DND = True
except Exception:
    _ROOT = tk.Tk
    _HAS_DND = False

# 政务配色（取自原版：界面用深红 #7A0F22，非红头那个亮红 #D21F2B）
ACCENT = "#7A0F22"
ACCENT_DARK = "#5E0E16"
GOLD = "#D4AF37"
BG = "#FAF6EF"
PANEL = "#F0E6D0"
BORDER = "#E3D5B8"
TEXT = "#333333"
SUBTLE = "#666666"
HINT = "#8A8272"
CHECK_EMPTY = "#B8AFA0"
CHECK_GREEN = "#2E8B57"
FONT = "Microsoft YaHei UI"

_BOOL_KEYS = [
    ("remove_spaces", "去除多余空格"),
    ("normalize_punctuation", "统一标点(含引号)"),
    ("normalize_western", "数字英文用 Times New Roman"),
    ("remove_blank_lines", "删除段落空行"),
    ("set_outline", "写入标题大纲级别"),
    ("enable_attachment_formatting", "识别附件说明"),
    ("enable_table_formatting", "格式化表格"),
    ("table_body_center", "表格内容上下左右居中"),
    ("enable_image_formatting", "图片居中/缩回版心/不裁切"),
    ("first_para_as_title", "首段兜底当大标题"),
    ("enable_footer_notes", "插入版记(抄送/印发)"),
    ("h3_bold", "三级标题加粗"),
]
_DONE = object()

HELP_TEXT = """公文一键排版助手 v{ver}

【软件简介】
本软件严格遵循 GB/T 9704—2012《党政机关公文格式》，提供版头样式、智能增强等公文排版功能。

【操作流程】
1. 在“导入文件”页添加 .docx/.doc/.wps/.txt/.md 文件（支持拖拽，也可“添加文件夹”批量），
   或在“粘贴文字”页直接粘贴内容（第一行自动作为标题）。
2. 在右侧选择“版头样式”（无红色版头 / 正式发文大红头 / 简化文头小红头），
   填写发文机关、发文字号、签发人（可选）。
3. 可选填写落款署名、成文日期（落款将按公文格式右空排版）。
4. 勾选所需的“智能增强”项（默认全部开启）。
5. 点击“开始排版导出”，选择输出位置即可。

【版头样式说明】
· 正式发文（大红头）：发文机关标志红色大字，下空一行排发文字号，其下红色分隔线，
  分隔线下空一行排标题。
· 简化文头（小红头）：发文机关标志为红色较小字号。
· 无红色版头：不插入红头，仅排版正文。

【智能增强】
1. 去除空格：自动删除各种形式的空格（半角、全角、不间断、制表符、零宽字符等），
   仅保留英文单词之间的空格。
2. 标点规范化：统一全半角与中英文标点；上下引号智能配对为 “ ” ‘ ’；
   省略号转为……；数字间连字符转为一字线；全角数字字母转半角。
3. 数字与英文字母统一使用公文要求的 Times New Roman 字体。
4. 另有：删除段落空行、写入标题大纲级别、识别附件说明、格式化表格、插入版记等，
   可在右侧“选项”里逐项开关。

【排版规则】
· 页面：A4，上 3.7cm / 下 3.5cm / 左 2.8cm / 右 2.6cm。
· 标题：二号方正小标宋，居中。
· 一级标题（一、）：三号黑体；二级标题（（一））：三号楷体；
  三级（1.）/四级（（1））：三号仿宋。
· 正文：三号仿宋，首行缩进 2 字符，固定行距 28 磅，两端对齐。
· 页码：一字线“— 1 —”，单页靠右、双页靠左。
· 落款：正文下空一行起排；成文日期右空 4 字，发文署名与成文日期右端对齐。
· 附件说明：左对齐 + 悬挂缩进，序号自动逐项对齐（无需手工换行）。

【字体】
本程序内置 4 种公文字体（仿宋_GB2312、楷体_GB2312、方正小标宋简体、黑体）。
首次使用建议点右侧“一键安装字体”：用户级安装、免管理员权限；已装好的会自动跳过。

【安全说明】
· 本工具不会修改您的任何原始文件，所有操作均在临时副本上进行，排版结果另存为新文件。
· 本工具全程在本机离线运行，不连接任何网络、不向任何服务器发送数据。

【使用须知与免责声明】
· 本软件为免费工具，供个人及单位内部日常办公使用，可自由复制分发，
  但不得以任何形式出售、转售或用于商业牟利（包括捆绑销售、代装收费、付费提供下载等）。
· 本软件仅用于对非涉密、可公开的公文材料进行格式排版。
  严禁使用本软件处理涉及国家秘密、工作秘密、商业秘密或个人隐私的涉密文件。
  使用者应自行确认所处理材料不属于涉密信息；因违规使用产生的一切后果由使用者自行承担。
· 排版结果由使用者自行审核把关，确保内容与格式符合发文单位要求及国家有关法规。
  开发者不对排版结果的正确性、完整性、适用性作出任何承诺。
· 本软件按“现状”提供，不构成任何明示或默示的担保；
  因使用或无法使用本软件产生的任何直接或间接损失，开发者不承担法律责任。
· 使用本软件即表示您已阅读、理解并接受上述全部条款；如不同意，请立即停止使用并删除本软件。
"""


class Section(ttk.Frame):
    """可折叠分区：点标题栏或右侧 ▶/▼ 收起/展开，收起后只留标题栏。

    右侧三个分区（版头/要素、选项、字体）统一用这一种样式，保证布局一致。
    """

    def __init__(self, master, title, expanded=True):
        super().__init__(master, style="TFrame")
        self._expanded = bool(expanded)
        self.head = tk.Frame(self, bg=PANEL, highlightbackground=BORDER,
                             highlightthickness=1, cursor="hand2")
        self.head.pack(fill="x")
        self.title = tk.Label(self.head, text=title, bg=PANEL, fg=ACCENT_DARK,
                              font=(FONT, 10, "bold"), anchor="w", cursor="hand2")
        self.title.pack(side="left", fill="x", expand=True, padx=(9, 0), pady=5)
        self.arrow = tk.Label(self.head, text="", bg=PANEL, fg=ACCENT_DARK,
                              font=(FONT, 9, "bold"), width=3, cursor="hand2")
        self.arrow.pack(side="right", padx=(0, 2))
        self.body = tk.Frame(self, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        self.body.pack(fill="x", expand=False)
        for w in (self.head, self.title, self.arrow):
            w.bind("<Button-1>", lambda e: self.toggle())
        self._apply()

    def toggle(self):
        self._expanded = not self._expanded
        self._apply()

    def expand(self, yes=True):
        self._expanded = bool(yes)
        self._apply()

    def is_expanded(self):
        return self._expanded

    def _apply(self):
        self.arrow.configure(text="▼" if self._expanded else "▶")
        if self._expanded:
            self.body.pack(fill="x", expand=False)
        else:
            self.body.pack_forget()


def _res(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", name)


class CheckRow(ttk.Frame):
    """自绘勾选框：未勾选灰□，勾选绿√（避开 clam 把勾显示成×）。"""
    def __init__(self, master, text, var, command=None):
        super().__init__(master, style="TFrame")
        self.var = var
        self.command = command
        self.lbl = tk.Label(self, text="□ " + text, bg=BG, fg=TEXT,
                            font=(FONT, 9), anchor="w", cursor="hand2")
        self.lbl.pack(fill="x")
        self.lbl.bind("<Button-1>", self._toggle)
        self._refresh()

    def _toggle(self, _e=None):
        self.var.set(not self.var.get())
        self._refresh()
        if self.command:
            self.command()

    def _refresh(self):
        on = self.var.get()
        self.lbl.configure(text=("√ " if on else "□ ") + self._text(),
                           fg=(CHECK_GREEN if on else CHECK_EMPTY))

    def _text(self):
        return self.lbl.cget("text")[2:]


class App:
    def __init__(self, root):
        self.root = root
        self.files = []
        self.busy = False
        self._q = queue.Queue()
        self.vars = {}
        self.checks = {}
        self.var_header = tk.StringVar(value=HEADER_STYLE_NONE)
        self.var_outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self._init_style()
        self._build_menu()
        self._build()
        self.root.after(120, self._drain)

    # ---------- 主题 ----------
    def _init_style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            return
        st.configure(".", background=BG, foreground=TEXT, font=(FONT, 9))
        st.configure("TFrame", background=BG)
        st.configure("Banner.TFrame", background=ACCENT)
        st.configure("TLabel", background=BG, foreground=TEXT)
        st.configure("Title.TLabel", background=ACCENT, foreground="#FFFFFF", font=(FONT, 15, "bold"))
        st.configure("Sub.TLabel", background=ACCENT, foreground=GOLD, font=(FONT, 9))
        st.configure("Section.TLabel", background=BG, foreground=ACCENT_DARK, font=(FONT, 10, "bold"))
        st.configure("Hint.TLabel", background=BG, foreground=HINT, font=(FONT, 8))
        st.configure("TLabelframe", background=BG, bordercolor=BORDER, relief="flat")
        st.configure("TLabelframe.Label", background=BG, foreground=ACCENT_DARK, font=(FONT, 10, "bold"))
        st.configure("TButton", background=PANEL, foreground=TEXT, bordercolor=BORDER,
                     focuscolor=PANEL, padding=(8, 5), font=(FONT, 9))
        st.map("TButton", background=[("active", "#E8DCC0"), ("disabled", "#DDD6C8")])
        st.configure("Accent.TButton", background=ACCENT, foreground="#FFFFFF",
                     bordercolor=ACCENT_DARK, focuscolor=ACCENT, padding=(10, 8), font=(FONT, 11, "bold"))
        st.map("Accent.TButton", background=[("active", ACCENT_DARK)])
        st.configure("TNotebook", background=BG, bordercolor=BORDER)
        st.configure("TNotebook.Tab", background=PANEL, foreground=SUBTLE, font=(FONT, 10), padding=(12, 5))
        st.map("TNotebook.Tab", background=[("selected", BG)], foreground=[("selected", ACCENT_DARK)])
        st.configure("TEntry", fieldbackground="#FFFFFF", bordercolor=BORDER, insetcolor="#FFFFFF")
        st.map("TEntry", bordercolor=[("focus", ACCENT)])
        st.configure("TCombobox", fieldbackground="#FFFFFF", background=PANEL, arrowcolor=ACCENT_DARK)
        st.configure("TCheckbutton", background=BG)
        st.configure("Horizontal.TSeparator", background=BORDER)
        st.configure("TProgressbar", background=ACCENT, troughcolor=PANEL,
                     bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)

    def _build_menu(self):
        mbar = tk.Menu(self.root)
        helpm = tk.Menu(mbar, tearoff=0)
        helpm.add_command(label="使用说明（含免责声明）", command=self.show_help_window)
        helpm.add_separator()
        helpm.add_command(label="联系开发者", command=self.show_contact)
        helpm.add_command(label="关于", command=self.show_about)
        mbar.add_cascade(label="操作", menu=helpm)
        self.root.config(menu=mbar)

    # ---------- 横幅徽标 ----------
    def _make_seal(self, size=40):
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk
        except Exception:
            return None
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=8, fill=ACCENT_DARK, outline=GOLD, width=2)
        try:
            fnt = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", int(size * 0.62))
        except Exception:
            fnt = ImageFont.load_default()
        d.text((size // 2, size // 2), "公", font=fnt, fill="#FFE9A8", anchor="mm")
        photo = ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))
        self._seal_photo = photo
        return photo

    def _build(self):
        r = self.root
        r.title("%s v%s" % (APP_NAME, APP_VERSION))
        r.geometry("1180x880")
        r.minsize(920, 640)
        try:
            ico = _res("app.ico")
            if os.path.isfile(ico):
                r.iconbitmap(ico)
        except Exception:
            pass
        r.protocol("WM_DELETE_WINDOW", self._on_close)

        # 顶部横幅
        banner = ttk.Frame(r, style="Banner.TFrame")
        banner.pack(fill="x")
        seal = self._make_seal(44)
        if seal is not None:
            tk.Label(banner, image=seal, bg=ACCENT).pack(side="left", padx=(14, 8), pady=8)
        ttk.Label(banner, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Label(banner, text="　按 GB/T 9704-2012 党政机关公文格式", style="Sub.TLabel").pack(side="left")
        tk.Button(banner, text="联系开发者", command=self.show_contact, bg=GOLD,
                  fg=ACCENT_DARK, activebackground="#E6C55A", relief="flat", bd=0,
                  font=(FONT, 9, "bold"), cursor="hand2",
                  padx=12, pady=4).pack(side="right", padx=(6, 14))

        body = ttk.Frame(r)
        body.pack(fill="both", expand=True, padx=10, pady=8)

        # 左：输入 + 日志
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        nb = ttk.Notebook(left)
        nb.pack(fill="both", expand=True)
        tf = ttk.Frame(nb); nb.add(tf, text=" 导入文件 ")
        tt = ttk.Frame(nb); nb.add(tt, text=" 粘贴文字 ")
        fr = ttk.Frame(tf); fr.pack(fill="x", pady=4)
        ttk.Button(fr, text="添加文件…", command=self.add_files).pack(side="left")
        ttk.Button(fr, text="添加文件夹…", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(fr, text="清空列表", command=self.clear_files).pack(side="left")
        self.file_list = tk.Listbox(tf, height=7, exportselection=False, font=(FONT, 9),
                                    selectmode=tk.EXTENDED, activestyle="none",
                                    selectbackground=ACCENT, selectforeground="#FFFFFF")
        self.file_list.pack(fill="both", expand=True, padx=4, pady=2)
        ttk.Label(tf, text="可直接把 .docx/.doc/.wps/.txt/.md 文件（或装着这些文件的文件夹）拖到窗口任意位置",
                  style="Hint.TLabel").pack(fill="x", padx=6, pady=(0, 4))
        self.text_in = tk.Text(tt, wrap="word", font=(FONT, 10))
        self.text_in.pack(fill="both", expand=True, padx=4, pady=4)
        self.nb, self.tab_import, self.tab_paste = nb, tf, tt
        self._enable_dnd(self.file_list, nb, tf, tt, left, self.root)

        ofr = ttk.Frame(left); ofr.pack(fill="x", pady=4)
        ttk.Label(ofr, text="导出目录：").pack(side="left")
        ttk.Entry(ofr, textvariable=self.var_outdir).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(ofr, text="浏览…", command=self.pick_outdir).pack(side="left")

        logf = ttk.LabelFrame(left, text="处理日志 / 自检报告")
        logf.pack(fill="both", expand=False, pady=4)
        self.console = scrolledtext.ScrolledText(logf, height=10, font=(FONT, 9), state="disabled", wrap="word")
        self.console.pack(fill="both", expand=True, padx=4, pady=4)
        self._enable_dnd(self.console)

        # 右：参数 + 操作（三个分区统一样式、可折叠；“选项”默认收起）
        right = ttk.Frame(body, width=380)
        right.pack(side="right", fill="y", padx=(12, 0))
        right.pack_propagate(False)

        sec_head = Section(right, "版头 / 要素", expanded=True)
        sec_head.pack(fill="x", pady=(0, 6))
        pf = sec_head.body
        ttk.Label(pf, text="版头样式：").grid(row=0, column=0, sticky="w", padx=9, pady=3)
        self.combo_header = ttk.Combobox(pf, textvariable=self.var_header,
                                         values=HEADER_STYLE_OPTIONS, state="readonly")
        self.combo_header.grid(row=0, column=1, sticky="ew", padx=(0, 9))
        for i, (k, lab) in enumerate([("org_name", "发文机关"), ("doc_number", "发文字号"),
                                      ("signer", "签发人(上行文)"), ("issuer", "落款署名"),
                                      ("date_text", "成文日期"), ("cc", "抄送"),
                                      ("print_org", "印发机关"), ("print_date", "印发日期")]):
            ttk.Label(pf, text=lab + "：").grid(row=1 + i, column=0, sticky="w", padx=9, pady=2)
            v = tk.StringVar(value=DEFAULT_CONFIG.get(k, ""))
            self.vars[k] = v
            ttk.Entry(pf, textvariable=v).grid(row=1 + i, column=1, sticky="ew", padx=(0, 9))
        pf.columnconfigure(1, weight=1)

        # 选项默认收起：点标题栏或右侧 ▶ 展开
        sec_opt = Section(right, "选项（点这里展开/收起）", expanded=False)
        sec_opt.pack(fill="x", pady=6)
        sf = sec_opt.body
        for k, lab in _BOOL_KEYS:
            var = tk.BooleanVar(value=bool(DEFAULT_CONFIG.get(k, True)))
            self.checks[k] = var
            CheckRow(sf, lab, var).pack(fill="x", padx=9, pady=1)
        self.sec_opt = sec_opt

        sec_font = Section(right, "字体", expanded=True)
        sec_font.pack(fill="x", pady=6)
        bf = sec_font.body
        for text, cmd in (("检测常用字体", self.check_fonts),
                          ("一键安装字体", lambda: self._spawn(self.install_fonts))):
            ttk.Button(bf, text=text, command=cmd).pack(fill="x", padx=9, pady=3)

        bottom = ttk.Frame(right)
        bottom.pack(side="bottom", fill="x", pady=(12, 6))
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))
        self.btn_run = ttk.Button(bottom, text="开 始 排 版 导 出", style="Accent.TButton", command=self.run)
        self.btn_run.pack(fill="x")

    # ---------- 线程安全日志 ----------
    def log(self, msg):
        self._q.put(str(msg))

    def _drain(self):
        try:
            while True:
                item = self._q.get_nowait()
                if item is _DONE:
                    self.busy = False
                    self.btn_run.configure(state="normal")
                    continue
                self.console.configure(state="normal")
                self.console.insert("end", item + "\n")
                self.console.see("end")
                self.console.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(120, self._drain)

    def _spawn(self, fn, *a):
        if self.busy:
            return
        self.busy = True
        self.btn_run.configure(state="disabled")

        def w():
            try:
                fn(*a)
            except Exception as e:
                self.log("出错：%s" % e)
            finally:
                self._q.put(_DONE)
        threading.Thread(target=w, daemon=True).start()

    # ---------- 菜单对话框 ----------
    def show_help_window(self):
        win = tk.Toplevel(self.root)
        win.title("使用说明　·　%s v%s" % (APP_NAME, APP_VERSION))
        win.configure(bg=BG)
        win.transient(self.root)
        win.geometry("780x660")
        txt = scrolledtext.ScrolledText(win, wrap="word", font=(FONT, 10),
                                        bg="#FFFFFF", fg=TEXT, relief="flat",
                                        padx=14, pady=12, spacing1=1, spacing3=3)
        txt.pack(fill="both", expand=True, padx=10, pady=(10, 6))
        txt.insert("1.0", HELP_TEXT.format(ver=APP_VERSION))
        txt.configure(state="disabled")
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=(0, 10))
        win.update_idletasks()
        win.geometry("+%d+%d" % (self.root.winfo_rootx() + 120, self.root.winfo_rooty() + 60))

    def show_about(self):
        self._info("关于", "%s v%s\n遵循 GB/T 9704-2012《党政机关公文格式》。\n"
                             "排版为确定性引擎，AI 仅做结构识别与文字核对，不改动版式与敏感表述。" % (APP_NAME, APP_VERSION))

    def show_contact(self):
        win = tk.Toplevel(self.root)
        win.title("联系开发者")
        win.configure(bg=BG)
        ttk.Label(win, text="扫码加微信联系开发者", style="Section.TLabel").pack(pady=(10, 4))
        png = _res("contact.png")
        photo = None
        if os.path.isfile(png):
            try:
                from PIL import Image, ImageTk
                im = Image.open(png)
                ratio = min(1.0, 280.0 / max(im.size))
                im = im.resize((int(im.size[0] * ratio), int(im.size[1] * ratio)))
                photo = ImageTk.PhotoImage(im)
                tk.Label(win, image=photo, bg=BG).pack(padx=16, pady=(0, 12))
            except Exception:
                photo = None
        if photo is None:
            ttk.Label(win, text="(未能加载微信二维码图片)", style="Hint.TLabel").pack(pady=20)
        win.image_ref = photo  # 防 GC

    def _info(self, title, msg):
        messagebox.showinfo(title, msg, parent=self.root)

    # ---------- 文件/目录 ----------
    def add_files(self):
        for p in filedialog.askopenfilenames(
                title="选择公文文件",
                filetypes=[("支持的文件", "*.docx *.doc *.wps *.txt *.md"), ("所有文件", "*.*")]):
            if p not in self.files:
                self.files.append(p); self.file_list.insert("end", p)

    def add_folder(self):
        d = filedialog.askdirectory(title="选择文件夹（批处理其中所有支持文件）")
        if not d:
            return
        exts = (".docx", ".doc", ".wps", ".txt", ".md")
        for name in os.listdir(d):
            fp = os.path.join(d, name)
            if os.path.isfile(fp) and os.path.splitext(name)[1].lower() in exts and fp not in self.files:
                self.files.append(fp); self.file_list.insert("end", fp)

    def clear_files(self):
        self.files = []
        self.file_list.delete(0, "end")

    def _enable_dnd(self, *widgets):
        """把拖拽接收注册到多个控件：拖到窗口任意位置都能收文件/文件夹。"""
        if not _HAS_DND:
            return
        for w in widgets:
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

    def on_drop(self, event):
        exts = (".docx", ".doc", ".wps", ".txt", ".md")
        added = 0
        for p in self.root.tk.splitlist(event.data):
            if os.path.isdir(p):
                for name in sorted(os.listdir(p)):
                    fp = os.path.join(p, name)
                    if (os.path.isfile(fp) and os.path.splitext(name)[1].lower() in exts
                            and fp not in self.files):
                        self.files.append(fp)
                        self.file_list.insert("end", fp)
                        added += 1
            elif os.path.isfile(p) and p not in self.files:
                self.files.append(p)
                self.file_list.insert("end", p)
                added += 1
        if added:
            try:
                self.nb.select(self.tab_import)
            except Exception:
                pass
            self.log("已拖入 %d 个文件。" % added)
        return getattr(event, "action", None)

    def pick_outdir(self):
        d = filedialog.askdirectory(initialdir=self.var_outdir.get())
        if d:
            self.var_outdir.set(d)

    def build_config(self):
        cfg = {"header_style": self.var_header.get()}
        for k, v in self.vars.items():
            cfg[k] = v.get()
        for k, var in self.checks.items():
            cfg[k] = var.get()
        return cfg

    # ---------- 字体 ----------
    def check_fonts(self):
        self.log("— 字体检测 —")
        for name, ok in F.check_status():
            self.log("  %s ： %s" % (name, "已安装" if ok else "缺失"))

    def install_fonts(self):
        self.log("— 一键安装字体 —")
        F.install(force=False, log=self.log)

    # ---------- 主流程 ----------
    def run(self):
        if self.busy:
            return
        cfg = self.build_config()
        outdir = self.var_outdir.get()
        tasks = list(self.files)
        paste = self.text_in.get("1.0", "end").strip()
        if not tasks and not paste:
            messagebox.showinfo("提示", "请添加文件或粘贴文字。", parent=self.root)
            return
        self.log("\n==================== 开始排版 ====================")
        self._spawn(self._do_work, tasks, paste, cfg, outdir)

    def _do_work(self, tasks, paste, cfg, outdir):
        if paste:
            fd, src = tempfile.mkstemp(suffix=".txt", prefix="gwps_paste_")
            os.close(fd)
            with open(src, "w", encoding="utf-8") as f:
                f.write(paste)
            self._one(src, cfg, outdir)
            try:
                os.remove(src)
            except OSError:
                pass
        total = len(tasks)
        for idx, p in enumerate(tasks, 1):
            self.log("— (%d/%d) %s" % (idx, total, os.path.basename(p)))
            self._one(p, cfg, outdir)
        self.log("全部完成。")

    def _one(self, in_path, cfg, outdir):
        proc = E.WordProcessor(config=cfg, log=self.log)
        fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="gwps_out_")
        os.close(fd)
        try:
            report = proc.format_document(in_path, tmp)
            name = E.build_filename(proc.last_title)
            final = self._unique(os.path.join(outdir, name))
            os.replace(tmp, final)
            self.log("  ✓ 输出：%s" % os.path.basename(final))
            self.log("  — 自检 —")
            for rname, ok, detail in report:
                self.log("    [%s] %s %s" % ("通过" if ok else "需核", rname,
                                             ("| " + detail) if detail else ""))
        except Exception as e:
            self.log("  处理失败：%s" % e)
            try:
                if os.path.isfile(tmp):
                    os.remove(tmp)
            except OSError:
                pass

    @staticmethod
    def _unique(path):
        if not os.path.exists(path):
            return path
        base, ext = os.path.splitext(path)
        i = 1
        while os.path.exists("%s(%d)%s" % (base, i, ext)):
            i += 1
        return "%s(%d)%s" % (base, i, ext)

    def _on_close(self):
        try:
            self._save_default()
        except Exception:
            pass
        self.root.destroy()

    def _save_default(self):
        # 存注册表：exe 旁边不留任何文件（旧版会在同目录写 gongwen_default_config.json）
        try:
            ST.save(self.build_config())
        except Exception:
            pass

    def load_initial(self):
        try:
            ST.migrate_legacy(self.log)
            d = ST.load()
            if d.get("header_style"):
                self.var_header.set(d["header_style"])
            for k, v in self.vars.items():
                if k in d:
                    v.set(d[k])
            for k, var in self.checks.items():
                if k in d:
                    var.set(bool(d[k]))
        except Exception:
            pass


def main():
    root = _ROOT()
    app = App(root)
    app.load_initial()
    root.mainloop()


if __name__ == "__main__":
    main()
