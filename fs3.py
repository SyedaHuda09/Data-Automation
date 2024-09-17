import pandas as pd
import re

def process_excel_pik(input_file, sheet_name, output_file):
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    percentage_pattern = re.compile(r'(\d+(\.\d+)?)')

    def format_floor(floor_value):
        try:
            floor_value = str(floor_value)
            match = percentage_pattern.search(floor_value)
            if match:
                return f"({match.group(1)}% Floor)"
        except Exception:
            return floor_value
        return floor_value  

    def merge_rate_floor(rate, floor):
        try:
            formatted_floor = format_floor(floor)
            if pd.isna(rate): 
                return formatted_floor if pd.notna(formatted_floor) else None
            if pd.isna(formatted_floor):  
                return rate
            return f"{rate}, {formatted_floor}"
        except Exception:
            return f"{rate}, {floor}" if rate and floor else rate or floor

    df['Interest rate'] = df.apply(lambda row: merge_rate_floor(row['Rate(b)'], row['Floor(b)']), axis=1)

    df.to_excel(output_file, index=False)

input_file = 'BDC-Normalization/FS_norm/FS KKR Capital Corp.^FSK^20240806^June 30, 2024^10Q^Extracted.xlsx'
sheet_name = 'Investments'
output_file = 'outputH_Interest_Rate111.xlsx'

process_excel_pik(input_file, sheet_name, output_file)
