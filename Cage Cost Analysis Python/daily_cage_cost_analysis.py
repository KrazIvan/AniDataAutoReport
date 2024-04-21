import pandas as pd
import json
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta

def load_cage_costs(file_path: str) -> dict:
    with open(file_path, "r") as file:
        return json.load(file)

def main():
    df = pd.read_excel(os.getenv("SOURCE_FILE"))

    df["Creation D."] = pd.to_datetime(df["Creation D."])
    df["Elimination D."] = pd.to_datetime(df["Elimination D."])

    cage_costs = load_cage_costs("cage_costs.json")

    currency = os.getenv("CURRENCY", "SEK")

    protocols = set(df["Protocol"])

    daily_usage = []

    for index, row in df.iterrows():
        creation_date = row["Creation D."]
        elimination_date = row["Elimination D."]
        cage_type = row["Cage type"]
        protocol = row["Protocol"]

        normalized_cage_type = cage_type.split()[0]

        
        if normalized_cage_type not in cage_costs:
            continue
            
        num_days = (elimination_date - creation_date).days + 1
        dates_used = [creation_date + timedelta(days=i) for i in range(num_days)]

        for date in dates_used:
            if date not in [item["Date"] for item in daily_usage]:
                daily_usage_entry = {"Date": date, f"Total Cost ({currency})": 0}
                for protocol_name in protocols:
                    daily_usage_entry[protocol_name + f" (Cost {currency})"] = 0
                daily_usage.append(daily_usage_entry)

            daily_usage_entry = next(item for item in daily_usage if item["Date"] == date)
            daily_usage_entry[normalized_cage_type + " (Usage)"] = daily_usage_entry.get(normalized_cage_type + " (Usage)", 0) + 1
            daily_usage_entry[normalized_cage_type + f" (Cost {currency})"] = daily_usage_entry.get(normalized_cage_type + f" (Cost {currency})", 0) + cage_costs[normalized_cage_type]
            daily_usage_entry[f"Total Cost ({currency})"] += cage_costs[normalized_cage_type]
            daily_usage_entry[protocol + f" (Cost {currency})"] = daily_usage_entry.get(protocol + f" (Cost {currency})", 0) + cage_costs[normalized_cage_type]

    daily_usage_df = pd.DataFrame(daily_usage)

    daily_usage_df.sort_values(by="Date", inplace=True)

    daily_usage_df.to_excel(os.getenv("TARGET_FILE", "DailyCageCost.xlsx"), index=False)

if __name__ == "__main__":
    load_dotenv()
    main()