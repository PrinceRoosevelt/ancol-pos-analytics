"""
Script Eksekutif: Kompilasi Dokumen Panduan Storytelling Executive Dashboard ke Format PDF
Menggunakan engine Microsoft Edge Headless bawaan Windows.
"""

import subprocess
import os
import tempfile
import sys

def build_pdf():
    edge_bin = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if not os.path.exists(edge_bin):
        edge_bin = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    
    html_file = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "docs", "PANDUAN_STORYTELLING_EXECUTIVE_DASHBOARD.html"))
    pdf_out = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "PANDUAN_STORYTELLING_EXECUTIVE_ANCOL_POS.pdf"))
    temp_profile = os.path.join(tempfile.gettempdir(), "edge_executive_pdf_gen")

    if not os.path.exists(html_file):
        print(f"Error: berkas HTML tidak ditemukan di {html_file}")
        sys.exit(1)

    cmd = [
        edge_bin,
        "--headless",
        "--disable-gpu",
        f"--user-data-dir={temp_profile}",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_out}",
        html_file
    ]

    print(f"Mengekspor PDF dari: {html_file}")
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    
    if os.path.exists(pdf_out) and os.path.getsize(pdf_out) > 0:
        size_kb = os.path.getsize(pdf_out) / 1024
        print(f"Berhasil membuat PDF: {pdf_out} ({size_kb:.1f} KB)")
    else:
        print(f"Gagal membuat PDF: {res.stderr}")
        sys.exit(1)

if __name__ == "__main__":
    build_pdf()
