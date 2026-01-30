#!/usr/bin/env python3
"""
PDF Three-Letter Code Counter
Reads a PDF file and counts occurrences of three-letter codes from the last line of each page.
Expected format: "xxx MM/DD/YY cccc" where xxx is the 3-letter code to count.
"""

import pdfplumber
import sys
import re
from collections import Counter

def extract_three_letter_code(line):
    """
    Extract the three-letter code from a line.
    Expected format: "xxx MM/DD/YY cccc"
    
    Args:
        line: String to extract the code from
        
    Returns:
        The three-letter code if found, None otherwise
    """
    # Pattern: 3 characters (letters/numbers), space, date, space, 4 characters
    # This pattern is flexible to handle variations
    pattern = r'^([a-zA-Z0-9]{3})\s+\d{1,2}/\d{1,2}/\d{2,4}\s+[a-zA-Z]{4}\s*$'
    
    match = re.match(pattern, line.strip())
    if match:
        return match.group(1)
    return None

def count_codes_in_pdf(pdf_path, verbose=False):
    """
    Read a PDF file and count three-letter codes from the last line of each page.
    
    Args:
        pdf_path: Path to the PDF file
        verbose: If True, print details about each page
        
    Returns:
        Counter object with code counts
    """
    code_counter = Counter()
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Processing {total_pages} pages...")
            print("=" * 80)
            
            for page_num, page in enumerate(pdf.pages, start=1):
                # Extract text from the page
                text = page.extract_text()
                
                if text:
                    # Get all lines
                    lines = text.split('\n')
                    
                    # Get the last non-empty line
                    last_line = None
                    for line in reversed(lines):
                        if line.strip():
                            last_line = line.strip()
                            break
                    
                    if last_line:
                        # Try to extract the three-letter code
                        code = extract_three_letter_code(last_line)
                        
                        if code:
                            code_counter[code] += 1
                            if verbose:
                                print(f"Page {page_num}: Found code '{code}' in line: {last_line}")
                        else:
                            if verbose:
                                print(f"Page {page_num}: No code found. Last line: {last_line}")
                    else:
                        if verbose:
                            print(f"Page {page_num}: No text found")
                else:
                    if verbose:
                        print(f"Page {page_num}: Empty page")
            
            print("=" * 80)
            
    except FileNotFoundError:
        print(f"Error: File '{pdf_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading PDF: {e}")
        sys.exit(1)
    
    return code_counter

def print_results(code_counter):
    """Print the results in a formatted way."""
    if not code_counter:
        print("\nNo codes found in the PDF.")
        return
    
    print("\n=== RESULTS ===")
    print(f"\nTotal unique codes: {len(code_counter)}")
    print(f"Total occurrences: {sum(code_counter.values())}")
    print("\nCode counts (sorted by frequency):")
    print("-" * 40)
    
    # Sort by count (descending), then by code name (alphabetically)
    for code, count in code_counter.most_common():
        print(f"  {code}: {count}")
    
    print("\nAlphabetical listing:")
    print("-" * 40)
    for code in sorted(code_counter.keys()):
        print(f"  {code}: {code_counter[code]}")

def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Usage: python count_codes.py <path_to_pdf> [--verbose]")
        print("\nOptions:")
        print("  --verbose  Show details for each page")
        print("\nExample:")
        print("  python count_codes.py document.pdf")
        print("  python count_codes.py document.pdf --verbose")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    verbose = len(sys.argv) == 3 and sys.argv[2] == "--verbose"
    
    code_counter = count_codes_in_pdf(pdf_path, verbose)
    print_results(code_counter)

if __name__ == "__main__":
    main()
