from argparse import ArgumentParser

from orgchart.orgchart import Orgchart

import pandas as pd
pd.options.display.max_columns = 1000


parser = ArgumentParser()
parser.add_argument('-f', '--file', type=str, help='The location of the file to process.', required=True)
parser.add_argument('-s', '--sheetname', type=str, help='The Excel sheet that contains the data', required=True)
parser.add_argument('-u', '--ultimate_company', type=str, help='The name of ultimate company in the group.', required=True)
parser.add_argument('-o', '--output_file', type=str, help='The name of the outputfile', required=True)
parser.add_argument('-x', '--skiprows', type=int, help='Rows to skip in the Excel file', required=False)

if __name__ == '__main__':

    args = parser.parse_args()

    org = Orgchart(
        file=args.file, 
        sheet_name=args.sheetname, 
        skip_rows=args.skiprows,
        ultimate_shareholder=args.ultimate_company,
        shareholder_column='shareholder',
        subsidiary_column='subsidiary',
        percentage_column='ownership'
    )

    org.create_orgchart(output_file=args.output_file)