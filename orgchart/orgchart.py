import pandas as pd
from typing import Union
from pprint import pprint
from tqdm import tqdm
import json

class Orgchart:

    def __init__(
            self, 
            file:str,
            sheet_name:str,
            skip_rows:Union[int, None],
            ultimate_shareholder:str,
            shareholder_column:str,
            subsidiary_column:str,
            percentage_column:str    
        ) -> None:

        self.df = pd.read_excel(file, sheet_name=sheet_name, skiprows=skip_rows)

        self.ultimate_shareholder = ultimate_shareholder
        self.shareholder_column = shareholder_column
        self.subsidiary_column = subsidiary_column
        self.percentage_column = percentage_column

    def create_orgchart(self, output_file:str):
        
        data = self._find_subsidiaries(shareholder=self.ultimate_shareholder)
        
        if output_file.endswith('.xlsx'):
            df = self._nested_dict_to_df(nested_dict=data)
            df.to_excel(output_file, sheet_name='orgchart', index=False)

        else:
            with open(output_file, 'w') as f:
                json.dump(data, f)
        
    def _find_subsidiaries(self, shareholder, subsidiary_dict=None):

        if subsidiary_dict is None:
            subsidiary_dict = {}
        
        subsidiaries_data = self.df[self.df[self.shareholder_column] == shareholder][[self.subsidiary_column, self.percentage_column]]
        
        for _, row in subsidiaries_data.iterrows():
            subsidiary = row[self.subsidiary_column]
            ownership = row[self.percentage_column]

            if subsidiary == shareholder:
                continue
            
            subsidiary_entry = {
                'ownership': ownership,
                'subsidiaries': self._find_subsidiaries(subsidiary)
            }
            
            subsidiary_dict[subsidiary] = subsidiary_entry

        return subsidiary_dict


    def _nested_dict_to_df(self, nested_dict):

        def flatten_dict(d, parent_keys=[]):
            items = []
            for k, v in d.items():
                company_with_ownership = "{} ({}%)".format(k, v['ownership']) if 'ownership' in v else k  # Append ownership

                keys = parent_keys + [company_with_ownership]

                if isinstance(v, dict) and 'subsidiaries' in v and v['subsidiaries']:
                    items.extend(flatten_dict(v['subsidiaries'], keys))
                else:
                    items.append(keys)
            return items

        flat_list = flatten_dict(nested_dict)

        max_depth = max(len(item) for item in flat_list)
        df_columns = {'Level {}'.format(i+2): [] for i in range(max_depth)}
        
        for path in flat_list:
            for i, level in enumerate(path):
                df_columns['Level {}'.format(i+2)].append(level)

            for j in range(len(path), max_depth):
                df_columns['Level {}'.format(j+1)].append(None)

        df = pd.DataFrame(df_columns)
        df['Level 1'] = self.ultimate_shareholder
        reordered_columns = ['Level 1']
        reordered_columns.extend(df.columns.to_list()[:-2])
        df = df[reordered_columns]

        return df