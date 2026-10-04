"""WordMerge V0.1.1 console-capable launcher."""

from pathlib import Path
import runpy


if __name__ == "__main__":
    runpy.run_path(
        str(Path(__file__).with_name("WordMerge_V0.1.1.pyw")),
        run_name="__main__",
    )
