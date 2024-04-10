import pandas as pd
from datetime import datetime, timedelta

def main():
    df = pd.read_excel('CagesFrom2023.xlsx')

    df['Creation D.'] = pd.to_datetime(df['Creation D.'])
    df['Elimination D.'] = pd.to_datetime(df['Elimination D.'])

    cage_costs = {'GM500B': 9.75, 'GM500': 7.25, 'GM500S': 9.75}

    protocols = set(df['Protocol'])

    daily_usage = []

    for index, row in df.iterrows():
        creation_date = row['Creation D.']
        elimination_date = row['Elimination D.']
        cage_type = row['Cage type']
        protocol = row['Protocol']

        normalized_cage_type = cage_type.split()[0]

        
        if normalized_cage_type not in cage_costs:
            continue
            
        num_days = (elimination_date - creation_date).days + 1
        dates_used = [creation_date + timedelta(days=i) for i in range(num_days)]

        for date in dates_used:
            if date not in [item['Date'] for item in daily_usage]:
                daily_usage_entry = {'Date': date, 'Total Cost (SEK)': 0}
                for protocol_name in protocols:
                    daily_usage_entry[protocol_name + ' (Cost SEK)'] = 0
                daily_usage.append(daily_usage_entry)

            daily_usage_entry = next(item for item in daily_usage if item['Date'] == date)
            daily_usage_entry[normalized_cage_type + ' (Usage)'] = daily_usage_entry.get(normalized_cage_type + ' (Usage)', 0) + 1
            daily_usage_entry[normalized_cage_type + ' (Cost SEK)'] = daily_usage_entry.get(normalized_cage_type + ' (Cost SEK)', 0) + cage_costs[normalized_cage_type]
            daily_usage_entry['Total Cost (SEK)'] += cage_costs[normalized_cage_type]
            daily_usage_entry[protocol + ' (Cost SEK)'] = daily_usage_entry.get(protocol + ' (Cost SEK)', 0) + cage_costs[normalized_cage_type]

    daily_usage_df = pd.DataFrame(daily_usage)

    daily_usage_df.sort_values(by='Date', inplace=True)

    daily_usage_df.to_excel('DailyCageUsage.xlsx', index=False)

if __name__ == "__main__":
    main()