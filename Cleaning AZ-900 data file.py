import pandas as pd

excel_file = "AZ-900 data.xlsx"
df = pd.read_excel(excel_file)

column_name = "Question" 

remove_duplicates_df = df.drop_duplicates(subset=[column_name])
remove_duplicates_and_empty_df = remove_duplicates_df.dropna(subset=[column_name])

remove_duplicates_and_empty_df.to_excel(excel_file, index=False)

print(f"Data edited and saved to {excel_file}")
