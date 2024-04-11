# Pillar II TreeViewer

The Pillar II TreeViewer is a Python libary that converts an Excel file to an Excel output from which a treeview PivotTable can be created.

## Installation

This package was created using Python 3.12.2

1. Clone from github
2. Create a virtual environment in the cloned directory:
    ```
    python3 -m venv venv
    ```
3. Install the libraries:
    ```
    pip install -r requirements.txt
    ```

## Usage
The packes presumes that there is an Excel file with three columns:
1. One column with all the shareholders
2. One column with all the subsidiaries (which can also appear in the shareholders column).
3. One column with the ownership percentages

The package can create a .json or an .xlsx file.

### Usage as command line tool

When using this package as a command line tool you have to provide the following command line arguments:

options:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  The location of the file to process.
  -s SHEETNAME, --sheetname SHEETNAME
                        The Excel sheet that contains the data
  -u ULTIMATE_COMPANY, --ultimate_company ULTIMATE_COMPANY
                        The name of ultimate company in the group.
  -o OUTPUT_FILE, --output_file OUTPUT_FILE
                        The name of the outputfile
  -x SKIPROWS, --skiprows SKIPROWS
                        Rows to skip in the Excel file
  -shc SHAREHOLDER_COLUMN, --shareholder_column SHAREHOLDER_COLUMN
                        The column with shareholder data
  -sbc SUBSIDIARY_COLUMN, --subsidiary_column SUBSIDIARY_COLUMN
                        The column with subsidiary data
  -ow OWNERSHIP_COLUMN, --ownership_column OWNERSHIP_COLUMN
                        The column with ownership data

Example with the example.xlsx file:

Output .xlsx:
```
python create_org.py -f "example.xlsx" -s "orgchart" -u "UltimateCompany_0" -o "output.xlsx" -shc shareholder -sbc subsidiary -ow ownership
```

Output .json:
```
python create_org.py -f "example.xlsx" -s "orgchart" -u "UltimateCompany_0" -o "output.json" -shc shareholder -sbc subsidiary -ow ownership
```

### Usage as a Python function

The example.py file holds two examples of the create_orgchart function to convert the data from example.xlsx to a json and xlsx output.


### Conversion to pivot table
The conversion to a pivot table has to take place in Excel. I am aiming to have this implemented in this package in the future. For now you can convert it to a pivot tree view as follows:
1. Select all the data in the 'orgchart' sheet.
2. On the 'insert tab' select 'Pivot table' and click 'Ok'.
3. In the PivotTable fields tab you can add all fields to the rows area.


