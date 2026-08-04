"""Extract the manuscript's embedded figures as separately uploadable PNGs."""

from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from PIL import Image


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "Manuscript.docx"
OUT = ROOT / "Submission_Figures"


def main():
    OUT.mkdir(exist_ok=True)
    with ZipFile(SOURCE) as archive:
        for number in range(1, 6):
            source = f"word/media/image{number}.png"
            destination = OUT / f"Figure_{number}.png"
            # Preserve the final manuscript's pixels and set explicit 300 dpi
            # metadata for separate-file submission. No interpolation is used.
            with Image.open(BytesIO(archive.read(source))) as image:
                image.save(destination, dpi=(300, 300))
    print(f"Wrote five figures to {OUT}")


if __name__ == "__main__":
    main()
