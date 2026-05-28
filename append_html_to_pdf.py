#!/usr/bin/env python3
"""Converte gli HTML generati in PDF e li aggiunge a un PDF esistente."""

import argparse
from pathlib import Path


def merge_pdfs(input_pdf_path: str, new_pdfs: list, output_path: str) -> None:
    """Unisce i nuovi PDF al PDF di input."""
    try:
        from PyPDF2 import PdfWriter, PdfReader
    except ImportError:
        print("❌ PyPDF2 non installato. Installo...")
        import subprocess
        subprocess.check_call(["pip", "install", "PyPDF2"])
        from PyPDF2 import PdfWriter, PdfReader

    pdf_writer = PdfWriter()

    # Leggi il PDF originale
    with open(input_pdf_path, "rb") as f:
        pdf_reader = PdfReader(f)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

    # Aggiungi i nuovi PDF
    for pdf_path in new_pdfs:
        with open(pdf_path, "rb") as f:
            pdf_reader = PdfReader(f)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)

    # Scrivi il risultato
    with open(output_path, "wb") as f:
        pdf_writer.write(f)
    print(f"✓ Merged PDF saved to: {output_path}")


def html_to_pdf(html_path: str, pdf_path: str) -> None:
    """Converte HTML a PDF usando weasyprint."""
    try:
        from weasyprint import HTML
    except ImportError:
        print("❌ weasyprint non installato. Installo...")
        import subprocess
        subprocess.check_call(["pip", "install", "weasyprint"])
        from weasyprint import HTML

    HTML(html_path).write_pdf(pdf_path)
    print(f"✓ Converted {html_path} → {pdf_path}")


def main():
    parser = argparse.ArgumentParser(description="Unisce gli HTML RBS/WBS generati al PDF di riferimento.")
    parser.add_argument(
        "--html-dir",
        default="/Users/federico/magistrale/PM/liveshare",
        help="Directory contenente i file HTML ",
    )
    parser.add_argument(
        "--pdf",
        default="/Users/federico/Downloads/Online Gantt 20260528.pdf",
        help="File PDF di riferimento",
    )
    parser.add_argument(
        "--output",
        default="/Users/federico/Downloads/Online Gantt 20260528_merged.pdf",
        help="File PDF di output",
    )
    args = parser.parse_args()

    html_dir = Path(args.html_dir)
    rbs_html = html_dir / "rbs_eventonight.html"
    wbs_html = html_dir / "wbs_eventonight.html"

    if not rbs_html.exists() or not wbs_html.exists():
        print(f"❌ File HTML non trovati in {html_dir}")
        return

    # Converti HTML a PDF
    rbs_pdf = Path("/tmp/rbs_eventonight.pdf")
    wbs_pdf = Path("/tmp/wbs_eventonight.pdf")

    html_to_pdf(str(rbs_html), str(rbs_pdf))
    html_to_pdf(str(wbs_html), str(wbs_pdf))

    # Unisci i PDF
    merge_pdfs(args.pdf, [str(rbs_pdf), str(wbs_pdf)], args.output)
    print(f"\n✅ File finale: {args.output}")


if __name__ == "__main__":
    main()
