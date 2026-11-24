"""Render one PPTX slide through PowerPoint in an isolated Windows process.

PowerPoint COM can terminate its client process on RPC faults.  Keeping it in a
short-lived helper means a render failure never takes down the Web server.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sitrain_app.ppt_render import main


if __name__ == "__main__":
    raise SystemExit(main(*sys.argv[1:]))
