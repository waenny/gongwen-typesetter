import sys, os, traceback
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gui_trace.txt")
try:
    import tkinter as tk
    import gwps_gui as G
    r = tk.Tk(); r.withdraw()
    app = G.App(r)
    r.update_idletasks(); r.update()
    assert app.btn_run.winfo_exists()
    cfg = app.build_config()
    assert "header_style" in cfg and "cc" in cfg
    r.destroy()
    open(out, "w", encoding="utf-8").write("GUI_OK dnd=%s keys=%d" % (G._HAS_DND, len(cfg)))
except Exception:
    open(out, "w", encoding="utf-8").write(traceback.format_exc())
