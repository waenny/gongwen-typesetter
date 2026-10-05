# -*- coding: utf-8 -*-
"""设置持久化：写 HKCU 注册表，**不在 exe 旁边产生任何文件**（用户要求"只要一个 exe"）。

存成一条 JSON 字符串值，键名 HKCU\\SOFTWARE\\公文一键排版助手\\settings。
注册表不可用时退回进程内存（本次有效），保证功能不中断。
"""
import json
import os
import sys

_REG_PATH = r"SOFTWARE\公文一键排版助手"
_VALUE = "settings"

_mem = {}


def load():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_PATH) as k:
            raw, _ = winreg.QueryValueEx(k, _VALUE)
        d = json.loads(raw)
        return d if isinstance(d, dict) else {}
    except Exception:
        return dict(_mem)


def save(d):
    if isinstance(d, dict):
        _mem.update(d)
    try:
        import winreg
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _REG_PATH, 0, winreg.KEY_WRITE) as k:
            winreg.SetValueEx(k, _VALUE, 0, winreg.REG_SZ, json.dumps(d, ensure_ascii=False))
        return True
    except Exception:
        return False


def legacy_path():
    """旧版把配置写在 exe 同目录的 gongwen_default_config.json。"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "gongwen_default_config.json")


def migrate_legacy(log=None):
    """把旧版的 json 配置迁进注册表，然后删掉那个文件。返回是否迁移过。"""
    p = legacy_path()
    if not os.path.isfile(p):
        return False
    try:
        with open(p, "r", encoding="utf-8-sig") as f:
            d = json.load(f)
        if isinstance(d, dict) and d:
            save(d)
        os.remove(p)
        if log:
            log("  > 已把旧配置文件迁入注册表并删除：%s" % p)
        return True
    except Exception as e:
        if log:
            log("  > 旧配置迁移失败(%s)，保留原文件。" % e)
        return False
