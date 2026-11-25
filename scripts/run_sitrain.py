from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    # sys.executable is SITRAIN-Teacher.exe in a frozen build. Dispatch the
    # renderer BEFORE importing the web app or starting a browser/server.
    if sys.argv[1:2] == ["--render-ppt-preview"]:
        if len(sys.argv) != 5:
            return 2
        from sitrain_app.ppt_render import main as render_ppt
        return render_ppt(*sys.argv[2:])

    if sys.argv[1:2] == ["--index-word-images"]:
        if len(sys.argv) != 4:
            return 2
        from sitrain_app.word_image_index import main as index_word
        return index_word(*sys.argv[2:])

    os.chdir(ROOT)
    import uvicorn
    from sitrain_app.main import app

    port = int(os.getenv("SITRAIN_PORT", "9000"))
    data_root = Path(os.getenv("SITRAIN_DATA_ROOT", ROOT / "sitrain_data"))
    data_root.mkdir(parents=True, exist_ok=True)
    (data_root / "server.pid").write_text(str(os.getpid()), encoding="ascii")
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=port, reload=False)
    return 0


def open_browser(port):
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{port}")


if __name__ == "__main__":
    raise SystemExit(main())
