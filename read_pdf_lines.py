#!/usr/bin/env python3
"""
PDF Line-by-Line Reader
Reads a PDF file and prints its content line by line for analysis.
"""

import pdfplumber
import sys

def read_pdf_lines(pdf_path):
    """
    Read a PDF file and print its content line by line.
    
    Args:
        pdf_path: Path to the PDF file
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Total pages: {len(pdf.pages)}")
            print("=" * 80)
            
            for page_num, page in enumerate(pdf.pages, start=1):
                print(f"\n--- PAGE {page_num} ---")
                
                # Extract text from the page
                text = page.extract_text()
                
                if text:
                    # Split into lines and print each one
                    lines = text.split('\n')
                    for line_num, line in enumerate(lines, start=1):
                        print(f"Line {line_num}: {line}")
                else:
                    print("(No text found on this page)")
                
                print("-" * 80)
                
    except FileNotFoundError:
        print(f"Error: File '{pdf_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading PDF: {e}")
        sys.exit(1)

def main():
    if len(sys.argv) != 2:
        print("Usage: python read_pdf_lines.py <path_to_pdf>")
        print("\nExample:")
        print("  python read_pdf_lines.py document.pdf")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    read_pdf_lines(pdf_path)

if __name__ == "__main__":
    main()
