import pandas as pd
import openpyxl
import re
import os
import sys

EXCEL = r"Survey.xlsx"

def rename_columns(excel=EXCEL):

    df = pd.read_excel(excel)


    Old_Columns = list(df.columns)

    New_Columns = []

    for old_column in Old_Columns:
        new_column = old_column.split("-")[-1]
        new_column = new_column.strip()
        df = df.rename(columns={
            old_column: new_column
        })
    

    return df

renamed_columns_df = rename_columns()

