#!/usr/bin/env python3
"""
PDF Batch Processor
Processes multiple PDF files in a folder to:
1. Count three-letter codes from the last line of each page across all PDFs
2. Combine all PDFs (excluding those with "multi-page" in filename) into one,
   sorted alphabetically by the first line of each page
"""

import pdfplumber
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import sys
import os
import re
from collections import Counter
from pathlib import Path

def extract_three_letter_code(line):
    """
    Extract the three-letter code from a line.
    Expected format: "xxx MM/DD/YY cccc"
    """
    pattern = r'^([a-zA-Z0-9]{3})\s+\d{1,2}/\d{1,2}/\d{2,4}\s+[a-zA-Z]{4}\s*$'
    match = re.match(pattern, line.strip())
    if match:
        return match.group(1)
    return None

def get_first_line(page):
    """
    Extract the first non-empty line from a page.
    
    Args:
        page: pdfplumber page object
        
    Returns:
        First non-empty line as string, or empty string if none found
    """
    text = page.extract_text()
    if text:
        lines = text.split('\n')
        for line in lines:
            if line.strip():
                return line.strip()
    return ""

def count_codes_in_folder(folder_path, verbose=False):
    """
    Count three-letter codes from all PDFs in a folder.
    
    Args:
        folder_path: Path to the folder containing PDFs
        verbose: If True, print details about each file
        
    Returns:
        Counter object with code counts
    """
    code_counter = Counter()
    pdf_files = list(Path(folder_path).glob('*.pdf'))
    
    if not pdf_files:
        print(f"No PDF files found in {folder_path}")
        return code_counter
    
    print(f"Found {len(pdf_files)} PDF file(s)")
    print("=" * 80)
    
    for pdf_file in sorted(pdf_files):
        print(f"\nProcessing: {pdf_file.name}")
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text()
                    
                    if text:
                        lines = text.split('\n')
                        
                        # Get last non-empty line
                        last_line = None
                        for line in reversed(lines):
                            if line.strip():
                                last_line = line.strip()
                                break
                        
                        if last_line:
                            code = extract_three_letter_code(last_line)
                            
                            if code:
                                code_counter[code] += 1
                                if verbose:
                                    print(f"  Page {page_num}: Found code '{code}'")
                            elif verbose:
                                print(f"  Page {page_num}: No code found in: {last_line}")
        
        except Exception as e:
            print(f"  Error processing {pdf_file.name}: {e}")
    
    print("=" * 80)
    return code_counter

def combine_pdfs_alphabetically(folder_path, output_path, code_counter, verbose=False):
    """
    Combine PDFs from a folder, sorted alphabetically by 3-character code.
    Pages with code count of 1 appear first, then pages with higher counts.
    Both groups are sorted alphabetically by code.
    Excludes files with "multi-page" in the filename.
    
    Args:
        folder_path: Path to the folder containing PDFs
        output_path: Path for the output combined PDF
        code_counter: Counter object with code counts
        verbose: If True, print details about each page
        
    Returns:
        Number of pages in combined PDF
    """
    pdf_files = list(Path(folder_path).glob('*.pdf'))
    
    # Filter out files with "multi-page" in the name
    pdf_files = [f for f in pdf_files if "multi-page" not in f.name.lower()]
    
    if not pdf_files:
        print("No PDF files to combine (after filtering)")
        return 0
    
    print(f"\nCombining {len(pdf_files)} PDF file(s) (excluding 'multi-page' files)")
    print("=" * 80)
    
    # Store pages with their code for sorting
    pages_with_keys = []
    
    for pdf_file in sorted(pdf_files):
        if verbose:
            print(f"Reading: {pdf_file.name}")
        
        try:
            # Use pdfplumber to extract text for sorting
            with pdfplumber.open(pdf_file) as plumber_pdf:
                # Use pypdf to get the actual pages for combining
                pypdf_reader = PdfReader(pdf_file)
                
                for page_num, plumber_page in enumerate(plumber_pdf.pages):
                    # Get the last line to extract the code
                    text = plumber_page.extract_text()
                    code = None
                    
                    if text:
                        lines = text.split('\n')
                        # Get last non-empty line
                        for line in reversed(lines):
                            if line.strip():
                                code = extract_three_letter_code(line.strip())
                                break
                    
                    pypdf_page = pypdf_reader.pages[page_num]
                    
                    # Get the count for this code
                    code_count = code_counter.get(code, 0) if code else 0
                    
                    pages_with_keys.append({
                        'code': code if code else '',
                        'code_count': code_count,
                        'page': pypdf_page,
                        'source': pdf_file.name,
                        'page_num': page_num + 1
                    })
                    
                    if verbose:
                        print(f"  Page {page_num + 1}: Code '{code}' (count: {code_count})")
        
        except Exception as e:
            print(f"Error reading {pdf_file.name}: {e}")
    
    # Sort pages: first by whether count is 1 (single occurrences first),
    # then alphabetically by code (case-insensitive)
    pages_with_keys.sort(key=lambda x: (0 if x['code_count'] == 1 else 1, x['code'].lower()))
    
    single_count = sum(1 for x in pages_with_keys if x['code_count'] == 1)
    multiple_count = len(pages_with_keys) - single_count
    
    print(f"\nSorted {len(pages_with_keys)} pages:")
    print(f"  - Single occurrence codes (sorted A-Z): {single_count}")
    print(f"  - Multiple occurrence codes (sorted A-Z): {multiple_count}")
    
    # Create the combined PDF
    writer = PdfWriter()
    
    for item in pages_with_keys:
        writer.add_page(item['page'])
    
    # Write the output
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)
    
    print(f"Combined PDF saved to: {output_path}")
    print("=" * 80)
    
    return len(pages_with_keys)

def create_excel_spreadsheet(folder_path, code_counter, verbose=False):
    """
    Create a new Excel spreadsheet with code counts.
    Codes in column A, counts in column D.
    
    Args:
        folder_path: Path to the folder where Excel file will be saved
        code_counter: Counter object with code counts
        verbose: If True, print detailed information
        
    Returns:
        Path to created Excel file
    """
    if not code_counter:
        print("No codes to write to Excel file")
        return None
    
    excel_path = os.path.join(folder_path, "add to service charges.xlsx")
    
    print(f"\nCreating Excel spreadsheet: {excel_path}")
    print("=" * 80)
    
    try:
        # Create a new workbook
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Code Counts"
        
        # Add header row
        sheet['A1'] = 'Code'
        sheet['D1'] = 'Count'
        
        # Style header
        header_font = Font(bold=True)
        sheet['A1'].font = header_font
        sheet['D1'].font = header_font
        sheet['A1'].alignment = Alignment(horizontal='center')
        sheet['D1'].alignment = Alignment(horizontal='center')
        
        # Add data rows (sorted alphabetically by code)
        row = 2
        for code in sorted(code_counter.keys()):
            count = code_counter[code]
            sheet[f'A{row}'] = code
            sheet[f'D{row}'] = count
            
            if verbose:
                print(f"  Row {row}: {code} = {count}")
            
            row += 1
        
        # Adjust column widths
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['D'].width = 12
        
        # Save the workbook
        wb.save(excel_path)
        
        print(f"\nSpreadsheet created successfully:")
        print(f"  - Total codes: {len(code_counter)}")
        print(f"  - Total occurrences: {sum(code_counter.values())}")
        print(f"  - File: {excel_path}")
        print("=" * 80)
        
        return excel_path
        
    except Exception as e:
        print(f"Error creating Excel file: {e}")
        return None

def print_code_results(code_counter):
    """Print the code counting results in a formatted way."""
    if not code_counter:
        print("\nNo codes found in any PDF.")
        return
    
    print("\n=== CODE COUNT RESULTS ===")
    print(f"\nTotal unique codes: {len(code_counter)}")
    print(f"Total occurrences: {sum(code_counter.values())}")
    print("\nCode counts (sorted by frequency):")
    print("-" * 40)
    
    for code, count in code_counter.most_common():
        print(f"  {code}: {count}")
    
    print("\nAlphabetical listing:")
    print("-" * 40)
    for code in sorted(code_counter.keys()):
        print(f"  {code}: {code_counter[code]}")

def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Usage: python batch_process.py <folder_path> [--verbose]")
        print("\nThis script will:")
        print("  1. Count three-letter codes from all PDFs in the folder")
        print("  2. Create Excel spreadsheet 'add to service charges.xlsx' with counts")
        print("  3. Combine PDFs (excluding 'multi-page' files) sorted by code")
        print("     - Pages with single occurrence codes appear first (A-Z)")
        print("     - Pages with multiple occurrence codes appear second (A-Z)")
        print("\nOptions:")
        print("  --verbose  Show detailed processing information")
        print("\nExample:")
        print("  python batch_process.py /path/to/pdf/folder")
        print("  python batch_process.py ./pdfs --verbose")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    verbose = len(sys.argv) == 3 and sys.argv[2] == "--verbose"
    
    # Verify folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: '{folder_path}' is not a valid directory")
        sys.exit(1)
    
    # Count codes from all PDFs
    print("\n" + "=" * 80)
    print("STEP 1: Counting three-letter codes from all PDFs")
    print("=" * 80)
    code_counter = count_codes_in_folder(folder_path, verbose)
    print_code_results(code_counter)
    
    # Create Excel spreadsheet
    print("\n" + "=" * 80)
    print("STEP 2: Creating Excel spreadsheet")
    print("=" * 80)
    excel_path = create_excel_spreadsheet(folder_path, code_counter, verbose)
    
    # Combine PDFs
    print("\n" + "=" * 80)
    print("STEP 3: Combining PDFs alphabetically by code")
    print("=" * 80)
    output_path = os.path.join(folder_path, "combined_single_page(print).pdf")
    num_pages = combine_pdfs_alphabetically(folder_path, output_path, code_counter, verbose)
    
    print(f"\n{'=' * 80}")
    print("COMPLETE!")
    print(f"{'=' * 80}")
    print(f"Total codes counted: {sum(code_counter.values())}")
    if excel_path:
        print(f"Excel file created: {excel_path}")
    print(f"Combined PDF pages: {num_pages}")
    print(f"Combined PDF output: {output_path}")

if __name__ == "__main__":
    main()
