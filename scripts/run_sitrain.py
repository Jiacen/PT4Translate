from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

import uvicorn

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
PORT = int(os.getenv("SITRAIN_PORT", "9000"))
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
from sitrain_app.main import app
data_root = Path(os.getenv("SITRAIN_DATA_ROOT", ROOT / "sitrain_data"))
data_root.mkdir(parents=True, exist_ok=True)
(data_root / "server.pid").write_text(str(os.getpid()), encoding="ascii")


def open_browser():
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{PORT}")


threading.Thread(target=open_browser, daemon=True).start()
uvicorn.run(app, host="127.0.0.1", port=PORT, reload=False)
