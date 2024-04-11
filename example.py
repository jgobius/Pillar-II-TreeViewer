from create_org import create_orgchart

if __name__ == '__main__':

    create_orgchart(
        file='dummy.xlsx',
        sheet_name='orgchart',
        skip_rows=None,
        ultimate_shareholder='UltimateCompany_0',
        shareholder_column='shareholder',
        subsidiary_column='subsidiary',
        percentage_column='ownership',
        output_file='output.json'
    )

    create_orgchart(
        file='dummy.xlsx',
        sheet_name='orgchart',
        skip_rows=None,
        ultimate_shareholder='UltimateCompany_0',
        shareholder_column='shareholder',
        subsidiary_column='subsidiary',
        percentage_column='ownership',
        output_file='output.xlsx'
    )