# -*- coding: utf-8 -*-
"""Export a Word .docx to PDF via Microsoft Word (win32com)."""

import sys
from pathlib import Path


def export_pdf(docx_path: Path, pdf_path: Path | None = None) -> Path:
    docx_path = docx_path.resolve()
    if not docx_path.exists():
        raise FileNotFoundError(docx_path)
    pdf_path = (pdf_path or docx_path.with_suffix(".pdf")).resolve()

    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        # 17 = wdExportFormatPDF
        doc.ExportAsFixedFormat(str(pdf_path), ExportFormat=17)
        doc.Close(False)
    finally:
        word.Quit()
    return pdf_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python export_docx_to_pdf.py <docx> [pdf]")
        sys.exit(1)
    docx = Path(sys.argv[1])
    pdf = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    out = export_pdf(docx, pdf)
    print(f"OK {out} {out.stat().st_size}")


if __name__ == "__main__":
    main()
