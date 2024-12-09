import pdfplumber
import csv
import tempfile
import os
from docx import Document
from docx.shared import Pt

def extract_rows_with_optional_filters(pdf_path, main_term_index, rank_column_index, college_column_index, min_rank, max_rank, target_colleges,main_term):
    temp_csv_path = tempfile.mktemp(suffix='.csv')
    
    with open(temp_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for page_num, page in enumerate(pdf.pages):
                print(f"Processing page {page_num + 1} of {total_pages}")
                tables = page.extract_tables()
                
                for table_num, table in enumerate(tables):
                    if not table or len(table) < 2:
                        print(f"Skipping empty or malformed table on page {page_num + 1}, table {table_num + 1}.")
                        continue

                    try:
                        for row in table[1:]:  # Start directly from the first data row
                            if main_term_index >= len(row):
                                print(f"column index out of bounds for table on page {page_num + 1}, table {table_num + 1}.")
                                break

                            
                            if main_term.upper() in str(row[main_term_index]).upper():
                                # Initialize flags for rank and college checks
                                rank_within_range = True
                                college_matches = True
                                
                                # Rank filtering if rank_column_index is specified
                                if rank_column_index is not None and rank_column_index < len(row):
                                    try:
                                        rank = int(row[rank_column_index])  # Convert rank to integer
                                        rank_within_range = min_rank <= rank <= max_rank
                                    except ValueError:
                                        print(f"Skipping row due to invalid rank value on page {page_num + 1}, table {table_num + 1}.")
                                        continue
                                elif rank_column_index is not None:
                                    print(f"Rank column index {rank_column_index} out of bounds for table on page {page_num + 1}, table {table_num + 1}.")
                                    continue

                                # College filtering if college_column_index and target_colleges are specified
                                if college_column_index is not None and target_colleges and college_column_index < len(row):
                                    college_name = ' '.join(str(row[college_column_index]).replace('\n', ' ').split()).lower()  # Convert to lowercase for comparison
                                    college_matches = any(target_college.lower() in college_name for target_college in target_colleges)
                                elif college_column_index is not None and target_colleges:
                                    print(f"College column index {college_column_index} out of bounds for table on page {page_num + 1}, table {table_num + 1}.")
                                    continue

                                # Write row if it meets searchterm, rank, and college criteria
                                if rank_within_range and college_matches:
                                    csv_writer.writerow(row)
                    except Exception as e:
                        print(f"Error processing table on page {page_num + 1}, table {table_num + 1}: {e}")

                page.flush_cache()

    return temp_csv_path

def save_to_docx_table(csv_path, output_path):
    doc = Document()
    doc.add_heading('Extracted Rows', 0)

    table = None
    row_count = 0

    with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
        csv_reader = csv.reader(csvfile)
        
        for row in csv_reader:
            if table is None:
                table = doc.add_table(rows=1, cols=len(row))
                table.autofit = True
                hdr_cells = table.rows[0].cells
                for i, column_name in enumerate(row):
                    hdr_cells[i].text = str(column_name)
                    run = hdr_cells[i].paragraphs[0].runs[0]
                    run.font.size = Pt(10)
            else:
                row_cells = table.add_row().cells
                for i, cell_value in enumerate(row):
                    row_cells[i].text = str(cell_value) if cell_value is not None else ''
                    run = row_cells[i].paragraphs[0].runs[0]
                    run.font.size = Pt(10)
            
            row_count += 1
            if row_count % 100 == 0:
                print(f"Processed {row_count} rows")

    if table is None:
        print("No data to save.")
    else:
        doc.save(output_path)
        print(f"Extracted rows saved in table format to {output_path}")

    # Remove temporary CSV file
    os.remove(csv_path)

# Example usage
pdf_path = 'sample.pdf'        # Replace with the path to your PDF
min_rank = 200                 # Min rank for filtering
max_rank = 1000                # Max rank for filtering
main_term = 'All India'        # Use the term for filtering rows(ex here using All India)
main_term_index = 2            # In which number coulmn the main_term located(start column with 0)

# Call with optional parameters
temp_csv_path = extract_rows_with_optional_filters(
    pdf_path, 
    main_term_index,
    rank_column_index=1,        # Omit rank filtering by passing None
    college_column_index=None,  # Omit college filtering by passing None
    min_rank=min_rank, 
    max_rank=max_rank, 
    target_colleges=['M.D. (PAEDIATRICS)'],  # Omit college filtering by passing None
    main_term=main_term
)

# Save the output to a Word document with auto-adjusted column widths
output_docx_path = 'output.docx'     # Specify the desired output path
save_to_docx_table(temp_csv_path, output_docx_path)


