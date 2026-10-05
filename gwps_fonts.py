# -*- coding: utf-8 -*-
"""公文常用字体检测与一键安装（重建版）

- 检测：按注册表字体值名"族名 (TrueType)"精确匹配（不用松散子串，避免只有系统"仿宋"
  却被判成"仿宋_GB2312 已装"）。
- 安装（用户级，免管理员）：字体文件复制到用户字体目录，并在
  HKCU\\...\\CurrentVersion\\Fonts 写入 **绝对路径**（Win7 无 %LOCALAPPDATA% 字体目录
  自动加载，但 Word/WPS 会读 HKCU\\Fonts，绝对路径值在 Win7/10/11 均可用），
  再调用 gdi32.AddFontResourceW 让当前会话立即生效并广播 WM_FONTCHANGE。
"""
import os
import sys

# 展示名 -> (注册表匹配字干, 随包候选文件名)
COMMON_FONTS = [
    ("仿宋_GB2312", "仿宋", ("仿宋_GB2312.ttf",)),
    ("楷体_GB2312", "楷体", ("楷体_GB2312.ttf",)),
    ("方正小标宋简体", "方正小标宋", ("方正小标宋简体.ttf", "方正简体小标宋体.TTF",
                                     "方正简体小标宋体.ttf")),
    ("黑体", "黑体", ("黑体.ttf", "simhei.ttf")),
]

# 系统自带字体的英文族名：注册表里常常只写 SimHei/SimSun，中文名匹配不上会误报“缺失”
_ALIASES = {
    "黑体": ("SimHei",),
    "仿宋": ("FangSong",),
    "楷体": ("KaiTi",),
    "宋体": ("SimSun",),
    "仿宋_GB2312": ("FangSong_GB2312",),
    "楷体_GB2312": ("KaiTi_GB2312",),
    "方正小标宋简体": ("FZXiaoBiaoSong-B05S",),
}


REG_FONTS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def _user_font_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
    return os.path.join(base, "Microsoft", "Windows", "Fonts")


def _bundled_fonts_dir():
    meipass = getattr(sys, "_MEIPASS", None)
    root = meipass or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(root, "fonts")


def _registry_font_names():
    names = set()
    try:
        import winreg
    except ImportError:
        return names
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, REG_FONTS_KEY)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    names.add(winreg.EnumValue(key, i)[0])
                    i += 1
                except OSError:
                    break
        finally:
            winreg.CloseKey(key)
    return names


def _family_installed(disp, reg_names):
    """按族名匹配注册表值名：'仿宋_GB2312 (TrueType)' 命中；'SimHei (TrueType)'
    这类系统自带字体用英文别名兜底；'A & B (TrueType)' 形式两边都算。"""
    heads = set()
    for r in reg_names:
        for part in r.split(" (")[0].split("&"):
            part = part.strip()
            if part:
                heads.add(part)
    if disp in heads:
        return True
    return any(a in heads for a in _ALIASES.get(disp, ()))


def check_status():
    """返回 [(展示名, 是否已安装)]。注册表族名精确匹配 或 系统/用户字体目录有同名文件。"""
    reg = _registry_font_names()
    sys_dir = os.path.join(os.environ.get("WINDIR", r"C:\\Windows"), "Fonts")
    user_dir = _user_font_dir()
    result = []
    for disp, _stem, fnames in COMMON_FONTS:
        ok = _family_installed(disp, reg)
        if not ok:
            for d in (sys_dir, user_dir):
                for cand in tuple(fnames) + (disp + ".ttf", disp + ".ttc"):
                    if os.path.isfile(os.path.join(d, cand)):
                        ok = True
                        break
                if ok:
                    break
        result.append((disp, ok))
    return result


def _add_font_resource(abs_path):
    try:
        import ctypes
        ctypes.windll.gdi32.AddFontResourceW(abs_path)
    except Exception:
        pass


def _broadcast_font_change():
    try:
        import ctypes
        HWND_BROADCAST = 0xFFFF
        WM_FONTCHANGE = 0x001D
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_FONTCHANGE, 0, 0, 0x0002, 1000, ctypes.byref(ctypes.c_void_p()))
    except Exception:
        pass


def install(force=False, log=None):
    """从随包 fonts/ 目录安装公文字体（用户级，免管理员）。返回 (已装, 跳过, 失败)。"""
    def note(m):
        if log:
            try:
                log(m)
            except Exception:
                pass

    src = _bundled_fonts_dir()
    if not os.path.isdir(src):
        note("  > 未找到随包 fonts 目录，无法一键安装（可手动放入字体）。")
        return [], [], [d for d, _, _ in COMMON_FONTS]
    try:
        import winreg
    except ImportError:
        note("  > 非 Windows 环境，跳过字体安装。")
        return [], [], []
    dst = _user_font_dir()
    try:
        os.makedirs(dst, exist_ok=True)
    except OSError as e:
        note("  > 创建用户字体目录失败：%s" % e)
        return [], [], [d for d, _, _ in COMMON_FONTS]
    existing = _registry_font_names()
    installed, skipped, failed = [], [], []
    for disp, _stem, fnames in COMMON_FONTS:
        if (not force) and _family_installed(disp, existing):
            skipped.append(disp)
            continue
        srcf = None
        for fn in fnames:
            cand = os.path.join(src, fn)
            if os.path.isfile(cand):
                srcf, out_name = cand, fn
                break
        if srcf is None:
            failed.append(disp + "（内嵌字体缺失）")
            continue
        try:
            shutil_copy(srcf, os.path.join(dst, out_name))
            tgt = os.path.abspath(os.path.join(dst, out_name))
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_FONTS_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, "%s (TrueType)" % disp, 0, winreg.REG_SZ, tgt)  # 绝对路径
            winreg.CloseKey(key)
            _add_font_resource(tgt)
            installed.append(disp)
        except Exception as e:
            failed.append("%s（%s）" % (disp, e))
    if installed:
        _broadcast_font_change()
    note("  > 字体：已装 %s；跳过 %s；失败 %s" % (installed, skipped, failed))
    return installed, skipped, failed


def shutil_copy(a, b):
    import shutil
    if os.path.isfile(b):
        try:
            os.remove(b)
        except OSError:
            pass
    shutil.copy2(a, b)
