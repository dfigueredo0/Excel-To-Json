import pandas as pd
import json
from datetime import datetime

def excel_to_json(excel_file):
    df = pd.read_excel(excel_file, header=1)

    actives = []
    advisors = []
    skip_values = ['Townsmen', "Housing Corp Members", "Last Name", "New Members"]
    advisors_position = ['President', "Asst. Chapter Advisor", "Chapter Advisor", "Resident Advisor"]

    for index, row in df.iterrows():
        if row['Last Name'] in skip_values or pd.isna(row['Last Name']):
            continue

        positions = [row['Current Office']] if pd.notna(row["Current Office"]) else []
        brother = {
            "lastname": row['Last Name'],
            "name": row['First Name'],
            "positions": positions,
            "classOf": f"AE {row['Pledge Class']}",
            "visible": "true",
            "hasImg": "false"
        }

        # Index would have to change depending on the number of members in the house
        if index < 25:
            actives.append(brother)
        else:
            if row['Current Office'] in advisors_position:
                advisors.append(brother)
        

    date_str = datetime.now().strftime("%Y-%m-%d")
    file_name = f"brothers_{date_str}.json"

    data = {
        "actives": actives,
        "advisors": advisors
    }

    with open(file_name, 'w') as f:
        json.dump(data, f, indent=6)

def main():
    excel_to_json(excel_file='rosters/AEPKS Roster Fall 2024.xlsx')

if __name__ == "__main__":
    main()
