import pandas as pd
from pathlib import Path
import os

def main():
    import_path = Path("C:/Users/jong_chenmark/projects/MGH/analytics_automator/reports/for_import")

    def find_latest_estimate_summary(directory: Path):
        estimate_files = list(directory.glob("*Estimate_Summary_without_OGF*.csv"))
        if not estimate_files:
            print("No EstimateSummary file found.")
            return None
        return max(estimate_files, key=os.path.getmtime)
    am_win_rate = pd.read_csv(find_latest_estimate_summary(import_path), usecols = [
            'CustomerNumber', 'CustomerSize', 'ResidentialOrCommercial', 'ProgramCode', 'TotalPrice',
            'EstimateRequestedDate', 'EstimateGivenDate', 'RejectDate',
            'SoldDate', 'CancelDate', 'EstBy', 'EmpName', 'SoldBy1'
        ], low_memory=False)

    estimate_summary_df = pd.read_csv(find_latest_estimate_summary(import_path), usecols=[
            'CustomerNumber', 'CustomerSize', 'ResidentialOrCommercial',
            'BranchNumberOfCustomer', 'ProgramCode', 'TotalPrice',
            'EstimateRequestedDate', 'EstimateGivenDate', 'RejectDate',
            'SoldDate', 'CancelDate'
        ], low_memory = False)

    am_win_rate = am_win_rate[am_win_rate['TotalPrice'] != 0]
    estimate_summary_df = estimate_summary_df[estimate_summary_df['TotalPrice'] != 0]

    residential_df = estimate_summary_df[estimate_summary_df['ResidentialOrCommercial'].astype(str) == 'R']

    pbt_2023 = pd.read_csv(f'{import_path}/pbt_2023.csv', low_memory=False)
    # pbt_2023 = pbt_2023[pbt_2023['GrossSalesAmount'] != 0]
    pbt_2023.drop_duplicates(subset='CustomerNumber')
    # pbt_2023.to_csv(f'{import_path}/pbt_2023_cleaned.csv', index=False)
    pbt_2023_customers = set(pbt_2023['CustomerNumber'])

    non_renewal_customers = residential_df[~residential_df['CustomerNumber'].isin(pbt_2023_customers)]

    # non_renewal_customers.to_csv(f'{import_path}/non_renewal_customers.csv', index=False)

    sold_customers = non_renewal_customers[
        (non_renewal_customers['SoldDate'].notna()) & 
        (non_renewal_customers['SoldDate'] != '') & 
        (non_renewal_customers['CancelDate'].isna() | (non_renewal_customers['CancelDate'] == ''))
    ]


    def send_to_csv():
        df_list = [
            ('am_win_rate', am_win_rate),
            ('estimate_summary', estimate_summary_df),
            # ('pbt_2023', pbt_2023),
            # ('residential', residential_df),
            # ('non_renewal_customers', non_renewal_customers),
            # ('sold_customers', sold_customers)
        ]
        for name, df in df_list:
            file_path = Path(import_path/'exports') / f'{name}_export.csv'
            df.to_csv(file_path, index=False)
            print(f"Successfully exported {name} to {file_path}")

    send_to_csv()
    # print(non_renewal_customers.head())

if __name__ == '__main__':
    main()
