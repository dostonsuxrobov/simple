from pathlib import Path
from PIL import Image


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "assets" / "icon-source.png"
OUTPUT_DIR = ROOT / "build"


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    source = Image.open(SOURCE).convert("RGBA")
    side = max(source.size)
    square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    square.alpha_composite(source, ((side - source.width) // 2, (side - source.height) // 2))
    square.save(OUTPUT_DIR / "icon.png", optimize=True)
    square.save(
        OUTPUT_DIR / "icon.ico",
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print(f"Created {OUTPUT_DIR / 'icon.ico'} from {SOURCE}")


if __name__ == "__main__":
    main()
