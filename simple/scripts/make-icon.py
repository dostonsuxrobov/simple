"""Installs the icon pack (source/simple-icon-pack) as the app's one icon.

The pack's ICO is used as it is: its 16-24 px frames are drawn separately for
small sizes, which resizing the large picture would lose. The same files go to
every workspace (window/file icon and the mark in its title bar). After a new
pack, update the icon-source.png hash in scripts/verify.cjs.
"""
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent.parent
SOURCE_ROOT = ROOT.parent
PACK = SOURCE_ROOT / "simple-icon-pack"
WORKSPACES = ["simple_pdf_source", "simple_calc_source", "simple_doc_source", "simple_image_source", "simple_video_source"]


def main() -> None:
    (ROOT / "build").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(PACK / "simple.ico", ROOT / "build" / "icon.ico")
    shutil.copyfile(PACK / "simple-512.png", ROOT / "build" / "icon.png")
    shutil.copyfile(PACK / "simple-512.png", ROOT / "assets" / "icon-source.png")
    for name in WORKSPACES:
        workspace = SOURCE_ROOT / name
        if (workspace / "build").is_dir():
            shutil.copyfile(PACK / "simple.ico", workspace / "build" / "icon.ico")
        (workspace / "public").mkdir(exist_ok=True)
        shutil.copyfile(PACK / "simple-64.png", workspace / "public" / "brand-icon.png")
    shutil.copyfile(PACK / "simple-512.png", SOURCE_ROOT / "simple_calc_source" / "icon.png")
    print(f"Installed the icon pack from {PACK}")


if __name__ == "__main__":
    main()
