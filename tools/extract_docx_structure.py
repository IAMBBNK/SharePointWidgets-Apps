from __future__ import annotations

from pathlib import Path
from docx import Document


def main() -> None:
    docx_path = Path(
        r"Schedule 15 - SoW template MSA IT Consulting (v5.1, 2017-04-26, NJD).docx"
    )
    doc = Document(docx_path)

    for idx, para in enumerate(doc.paragraphs, 1):
        text = (para.text or "").replace("\t", " ").strip()
        if not text:
            continue
        style = getattr(para.style, "name", "")
        print(f"{idx:04d}\t{style}\t{text}")


if __name__ == "__main__":
    main()

