from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sitrain_app.word_image_index import main

if __name__ == "__main__":
    raise SystemExit(main(*sys.argv[1:]))
