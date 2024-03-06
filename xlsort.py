import pandas as pd
import re

def clean_filename(name):
        return re.sub(r'[<>:"/\\|?*]', '_', str(name))


def xlsort():
    df = pd.read_excel("all.xlsx")

    grouped = df.groupby("Acronym line")

    for group_name, group_df in grouped:
        filename = f"{clean_filename(group_name)}.xlsx" if group_name else "blank.xlsx"
        
        group_df.to_excel(filename, index=False)


if __name__ == "__main__":
    xlsort()