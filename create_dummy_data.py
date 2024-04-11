import pandas as pd
import random
from argparse import ArgumentParser

import pandas as pd
import random

def generate_company_data(max_subs_per_level=3, num_levels=3):
    def create_companies(prefix, num_companies):
        return [f"{prefix}{i}" for i in random.sample(range(100, 999), num_companies)]

    def create_relationships(shareholder, level, max_subs):
        relationships = []
        if level < num_levels:
            # Determine the random number of subsidiaries for this shareholder
            num_subsidiaries = random.randint(1, max_subs)
            # Generate unique subsidiaries names
            subsidiaries = create_companies(f"Company{level}_", num_subsidiaries)
            for subsidiary in subsidiaries:
                ownership = random.randint(51, 100)  # Minimum 51% to ensure control
                relationships.append((shareholder, subsidiary, ownership))
                # Proceed to the next level and generate subsidiaries for the current subsidiary
                relationships.extend(create_relationships(subsidiary, level + 1, max_subs))
        return relationships

    # Create a single top-level ultimate shareholder
    ultimate_shareholder = "UltimateCompany_0"
    
    # Create and return the relationships DataFrame along with the ultimate shareholder
    relationships = create_relationships(ultimate_shareholder, 1, max_subs_per_level)  # Start from level 1 as level 0 is the ultimate shareholder
    df = pd.DataFrame(relationships, columns=['shareholder', 'subsidiary', 'ownership'])

    return df, ultimate_shareholder

# Example usage
df_fake_companies, ultimate_shareholder = generate_company_data()
print("Ultimate Shareholder:", ultimate_shareholder)


parser = ArgumentParser()
parser.add_argument('-f', '--output_file', type=str, help='The output file to write the dataframe to. Must be an Excel file')


if __name__ == '__main__':

    args = parser.parse_args()

    df, ultimate_shareholder = generate_company_data(max_subs_per_level=5, num_levels=5)

    df_shareholder = pd.DataFrame({"Shareholder": ultimate_shareholder}, index=[0])
    print(f'Ultimate shareholder: {df_shareholder}')

    with pd.ExcelWriter(args.output_file) as writer:
        df.to_excel(writer, sheet_name='orgchart', index=False)
        df_shareholder.to_excel(writer, sheet_name='shareholder', index=False)