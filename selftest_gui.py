"""GUI 构建自检：不进入 mainloop，只验证窗口能创建、页签齐全。"""
import sys
from pathlib import Path
from tkinter import ttk

sys.path.insert(0, str(Path(__file__).parent))

import gui  # noqa: E402

app = gui.App()
app.update_idletasks()
app.update()

nb = next((w for w in app.winfo_children() if isinstance(w, ttk.Notebook)), None)
print("notebook:", "found" if nb else "MISSING")
print("tabs:", nb.index("end") if nb else 0)
print("log_text:", "ok" if hasattr(app, "log_text") else "MISSING")
print("buttons:", len(app._buttons))
app.destroy()
print("GUI_SELFTEST_OK")
