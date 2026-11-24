"""Isolated, read-only PowerPoint rendering for Python and the frozen launcher."""
from __future__ import annotations

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
        app.AutomationSecurity = 3
        presentation = app.Presentations.Open(str(source.resolve()), ReadOnly=True, WithWindow=False)
        height = round(1600 * presentation.PageSetup.SlideHeight / presentation.PageSetup.SlideWidth)
        pages = list(range(1, presentation.Slides.Count + 1)) if page_arg == "all" else [int(p) for p in page_arg.split(",")]
        if not pages or any(page < 1 or page > presentation.Slides.Count for page in pages):
            return 2
        multiple = page_arg == "all" or len(pages) > 1
        for page in pages:
            image = target / f"page-{page}-rendered.png" if multiple else target
            image.parent.mkdir(parents=True, exist_ok=True)
            presentation.Slides.Item(page).Export(str(image.resolve()), "PNG", 1600, height)
            if not image.exists() or not image.stat().st_size:
                return 1
        return 0
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
