import pandas as pd
import re
import os
from dotenv import load_dotenv

def clean_filename(name):
        return re.sub(r'[<>:"/\\|?*]', '_', str(name))


def xlsort(full_aniobio_datafile: str):
    df = pd.read_excel(full_aniobio_datafile)

    grouped = df.groupby(os.getenv("GROUP_BY_KEYWORD", "Acronym line"))

    for group_name, group_df in grouped:
        filename = f"{clean_filename(group_name)}.xlsx" if group_name else "blank.xlsx"
        
        group_df.to_excel(filename, index=False)


if __name__ == "__main__":
    load_dotenv()
    xlsort(os.getenv("FULL_FILE"))