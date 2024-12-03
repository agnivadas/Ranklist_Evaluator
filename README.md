
# College Counseling Rank Evaluator
[![GPL License](https://img.shields.io/badge/license-GPL-violet.svg)](http://www.gnu.org/licenses/gpl-3.0) [![Image_Categorizer](https://img.shields.io/badge/source-GitHub-303030.svg?style=flat-square)](https://github.com/agnivadas/Ranklist_Evaluator) ![Maintenance](https://img.shields.io/maintenance/yes/2024) ![Static Badge](https://img.shields.io/badge/contributions-welcome-blue)

College Counseling Rank Evaluator is a Python-based tool designed to analyze PDF file of rank lists for the college counseling process. It allows users to filter ranks, evaluate eligibility for specific colleges, and extract relevant data into user-friendly formats Word document. This tool could be used for filtering pdf for other purpose also as per user.




## Features

- **PDF Rank List Extraction:** Extracts tabular data from PDF files.
- **Search and Filter:** Filters ranks based on specific criteria such as search terms, rank ranges, and college preferences.
- **Customizable Output:** Saves filtered data as CSV and generates Word tables for easy review.
- **Automated Workflow:** Simplifies rank evaluation for efficient college counseling.


## Requirements

Ensure you have the following dependencies installed:

- Python 3.7 or later
- pdfplumber
- python-docx
Install dependencies using::
```bash
pip install pdfplumber python-docx

```
## Usage
1.**Input Requirements:**

Provide a rank list in PDF format.
Specify relevant parameters like column indices, rank ranges, and search terms.

2.**Run the Script:**

```python
python ranklist_evaluator.py
```
3.**Output:**

A filtered rank list saved as a DOCX file.
A Word document containing the extracted data in a tabular format.

## Example
```python
# Input parameters
pdf_path = 'sample.pdf' 
min_rank = 200    
max_rank = 1000  
main_term = 'All India' 
main_term_index = 2  

# Extract data
temp_csv_path = extract_rows_with_optional_filters(
    pdf_path,
    sc_column_index,
    rank_column_index=1,
    college_column_index=None,
    min_rank=min_rank,
    max_rank=max_rank,
    target_colleges=['M.D. (PAEDIATRICS)'],
    search_term=search_term
)

# Save as Word document
output_docx_path = 'output.docx'
save_to_docx_table(temp_csv_path, output_docx_path)
```


## Demo

Suppose the tables of ranklist pdf look like this below table structure.

<img src="/assets/screenshot5.jpg" width="600px">

- To filter all rows with alloted quota  'All India'
     ` main_term = 'All India'`
- Put main_term column number in main_term_index(column number starts with 0)
    ` main_term_index = 2`
- For rank filtering min and max rank and the rank column index
- For target college or subject use college_column_index and target_colleges. 
  ```python
     # For example 
      college_column_index= 4
      target_colleges=['M.D. (PAEDIATRICS)']
  ```
  Multiple target_colleges could be used .
   ```python
     # For example 
      college_column_index= 3
      target_colleges=['Rajasthan','West Bengal']
  ```
## Contributing

Contributions are welcome! Feel free to submit a pull request or open an issue for suggestions and bug reports.



## Acknowledgements

Special thanks to the open-source community for tools like pdfplumber and python-docx, which make this project possible.

