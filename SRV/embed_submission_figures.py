"""Replace selected embedded Word figure binaries with submission PNG files."""

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[2]
MANUSCRIPT = ROOT / "Manuscript.docx"
FIGURES = ROOT / "Submission_Figures"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-docx", type=Path, help="Copy the five image binaries from this Word file instead of Submission_Figures.")
    parser.add_argument(
        "--numbers",
        type=int,
        nargs="+",
        default=[1, 2, 3, 4, 5],
        help="Figure numbers to replace (default: all five).",
    )
    args = parser.parse_args()
    numbers = sorted(set(args.numbers))
    if not numbers or any(number not in range(1, 6) for number in numbers):
        parser.error("--numbers accepts values 1 through 5.")
    if args.source_docx:
        with ZipFile(args.source_docx) as source_images:
            replacements = {
                f"word/media/image{number}.png": source_images.read(f"word/media/image{number}.png")
                for number in numbers
            }
    else:
        replacements = {
            f"word/media/image{number}.png": (FIGURES / f"Figure_{number}.png").read_bytes()
            for number in numbers
        }
    temporary = MANUSCRIPT.with_suffix(".repacked.docx")
    with ZipFile(MANUSCRIPT) as source, ZipFile(temporary, "w", ZIP_DEFLATED) as destination:
        for item in source.infolist():
            data = replacements.get(item.filename, source.read(item.filename))
            destination.writestr(item, data)
    temporary.replace(MANUSCRIPT)
    print(f"Embedded figure binaries {numbers} in the manuscript.")


if __name__ == "__main__":
    main()
