"""Render one PPTX slide through PowerPoint in an isolated Windows process.

PowerPoint COM can terminate its client process on RPC faults.  Keeping it in a
short-lived helper means a render failure never takes down the Web server.
"""
from __future__ import annotations

import sys
from pathlib import Path


def main(source_arg: str, target_arg: str, page_arg: str) -> int:
    source, target = Path(source_arg), Path(target_arg)
    app = presentation = None
    initialized = False
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        initialized = True
        app = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = app.Presentations.Open(str(source.resolve()), WithWindow=False)
        if page_arg == "all":
            target.mkdir(parents=True, exist_ok=True)
            for page in range(1, presentation.Slides.Count + 1):
                image = target / f"page-{page}-rendered.png"
                presentation.Slides.Item(page).Export(str(image.resolve()), "PNG", 1600, 900)
            return 0
        page = int(page_arg)
        if page < 1 or page > presentation.Slides.Count:
            return 2
        target.parent.mkdir(parents=True, exist_ok=True)
        presentation.Slides.Item(page).Export(str(target.resolve()), "PNG", 1600, 900)
        return 0 if target.exists() and target.stat().st_size else 1
    except Exception:
        return 1
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main(*sys.argv[1:]))
