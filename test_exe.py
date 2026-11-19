"""验证打包出的 GUI exe 能真正启动（启动 6 秒仍存活即认为窗口已创建）。"""
import subprocess
import sys
import time

EXE = r"D:\freelance\2026-11-17_01-Excel批量工具箱\dist\ExcelToolGUI\ExcelToolGUI.exe"

p = subprocess.Popen([EXE], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(6)

if p.poll() is None:
    print("GUI_EXE_ALIVE_OK  (窗口已启动，PID=%d)" % p.pid)
    p.terminate()
    try:
        p.wait(timeout=5)
        print("已关闭")
    except Exception:
        p.kill()
    sys.exit(0)
else:
    print("GUI_EXE_EXITED  code=%s" % p.returncode)
    sys.exit(1)
