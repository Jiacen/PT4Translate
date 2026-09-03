from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

import uvicorn

ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
(ROOT / "sitrain_data").mkdir(exist_ok=True)
(ROOT / "sitrain_data" / "server.pid").write_text(str(os.getpid()), encoding="ascii")


def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:8080")


threading.Thread(target=open_browser, daemon=True).start()
uvicorn.run("sitrain_app.main:app", host="127.0.0.1", port=8080, reload=False)
