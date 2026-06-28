# -*- coding: utf-8 -*-
"""
Export Word DOCX formulas to PNG images for visual AI auditing.
Uses Word COM to export PDF, then Poppler to render pages to PNG.
"""
import argparse
import os
import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("⚠️  pywin32 not installed. Installing...")
    os.system(f"{sys.executable} -m pip install pywin32")
    import win32com.client as win32

try:
    from pdf2image import convert_from_path
except ImportError:
    print("⚠️  pdf2image not installed. Installing...")
    os.system(f"{sys.executable} -m pip install pdf2image")
    from pdf2image import convert_from_path


def docx_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF using Word COM."""
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # wdFormatPDF = 17
        doc.Close()
        print(f"✅ PDF saved: {pdf_path}")
        return True
    except Exception as e:
        print(f"❌ PDF export failed: {e}")
        return False
    finally:
        word.Quit()


def pdf_to_pngs(pdf_path, output_dir, dpi=200):
    """Convert PDF pages to PNG images."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    pages = convert_from_path(str(pdf_path), dpi=dpi)
    for i, page in enumerate(pages, 1):
        png_path = Path(output_dir) / f"page-{i:02d}.png"
        page.save(str(png_path), "PNG")
        print(f"✅ PNG saved: {png_path}")
    
    return pages


def main():
    parser = argparse.ArgumentParser(description="Export Word formulas to PNG for visual auditing")
    parser.add_argument("docx", help="Path to DOCX file")
    parser.add_argument("--out-dir", help="Output directory for PNGs", default=None)
    parser.add_argument("--dpi", type=int, default=200, help="Render DPI (default: 200)")
    
    args = parser.parse_args()
    
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"❌ File not found: {docx_path}")
        sys.exit(1)
    
    if args.out_dir:
        out_dir = Path(args.out_dir).resolve()
    else:
        out_dir = docx_path.parent / "visual-audit"
    
    pdf_path = out_dir / docx_path.with_suffix(".pdf").name
    
    print(f"📄 Exporting to PDF: {docx_path}")
    if not docx_to_pdf(docx_path, pdf_path):
        sys.exit(1)
    
    print(f"🖼️  Rendering PDF pages to PNG (DPI={args.dpi})...")
    pdf_to_pngs(pdf_path, out_dir, args.dpi)
    
    print(f"\n🎉 Visual audit materials ready in: {out_dir}")
    print(f"   Now you can upload these PNGs to a visual AI model for review.")


if __name__ == "__main__":
    main()
