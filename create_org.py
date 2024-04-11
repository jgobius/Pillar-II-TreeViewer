from argparse import ArgumentParser
from typing import Union

from orgchart.orgchart import Orgchart

parser = ArgumentParser()
parser.add_argument(
    "-f", "--file", type=str, help="The location of the file to process.", required=True
)
parser.add_argument(
    "-s",
    "--sheetname",
    type=str,
    help="The Excel sheet that contains the data",
    required=True,
)
parser.add_argument(
    "-u",
    "--ultimate_company",
    type=str,
    help="The name of ultimate company in the group.",
    required=True,
)
parser.add_argument(
    "-o", "--output_file", type=str, help="The name of the outputfile", required=True
)
parser.add_argument(
    "-x", "--skiprows", type=int, help="Rows to skip in the Excel file", required=False
)

parser.add_argument(
    "-shc", "--shareholder_column", type=str, help="The column with shareholder data"
)

parser.add_argument(
    "-sbc", "--subsidiary_column", type=str, help="The column with subsidiary data"
)

parser.add_argument(
    "-ow", "--ownership_column", type=str, help="The column with ownership data"
)


def create_orgchart(
    file: str,
    sheet_name: str,
    skip_rows: Union[int, None],
    ultimate_shareholder: str,
    shareholder_column: str,
    subsidiary_column: str,
    percentage_column: str,
    output_file:str
):

    org = Orgchart(
        file=file,
        sheet_name=sheet_name,
        skip_rows=skip_rows,
        ultimate_shareholder=ultimate_shareholder,
        shareholder_column=shareholder_column,
        subsidiary_column=subsidiary_column,
        percentage_column=percentage_column,
    )

    org.create_orgchart(output_file=output_file)


if __name__ == "__main__":

    args = parser.parse_args()

    create_orgchart(
        file=args.file,
        sheet_name=args.sheetname,
        skip_rows=args.skiprows,
        ultimate_shareholder=args.ultimate_company,
        shareholder_column=args.shareholder_column,
        subsidiary_column=args.subsidiary_column,
        percentage_column=args.ownership_column,
        output_file=args.output_file
    )

