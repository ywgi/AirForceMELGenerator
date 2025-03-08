import pandas as pd
from accounting_date_check import accounting_date_check
from board_filter import board_filter
from initial_mel_pdf_generator import generate_roster_pdf
from datetime import datetime, timedelta


#UIF 1,2,3 - ONLY 2,3 MAKE YOU INELIGIBLE

eligible_service_members = []
ineligible_service_members = []


alpha_roster_path = rf'C:\Users\Trent\Documents\Alpha Roster.xlsx'
test_path = rf'C:\Users\Trent\Documents\7 Oct 2024 - Sanitized Alpha Roster.xlsx'
required_columns = ['FULL_NAME', 'GRADE', 'ASSIGNED_PAS_CLEARTEXT', 'DAFSC', 'DOR', 'DATE_ARRIVED_STATION', 'TAFMSD','REENL_ELIG_STATUS', 'ASSIGNED_PAS']
optional_columns = ['GRADE_PERM_PROJ', 'UIF_CODE', 'UIF_DISPOSITION_DATE']
pdf_columns = ['FULL_NAME','GRADE', 'DATE_ARRIVED_STATION','DAFSC', 'ASSIGNED_PAS_CLEARTEXT', 'DOR', 'TAFMSD', 'ASSIGNED_PAS']
pascodes = []
pascodeMap = {}

boards = ['E5', 'E6', 'E7', 'E8', 'E9']
grade_map = {
    "SRA": "E4",
    "SSG": "E5",
    "TSG": "E6",
    "MSG": "E7",
    "SMS": "E8"
}
alpha_roster = pd.read_excel(alpha_roster_path)
filtered_alpha_roster = alpha_roster[required_columns + optional_columns]
pdf_roster = filtered_alpha_roster[pdf_columns]
valid_upload = True

cycle = 'SMS'
year = 2025

for index, row in filtered_alpha_roster.iterrows():
    for column, value in row.items():
        if pd.isna(value) and column in required_columns:
            valid_upload = False
            print(rf"error at {index}, {column}")
            break
    valid_member = accounting_date_check(row['DATE_ARRIVED_STATION'], cycle, year)
    if not valid_member:
        print(f"{index} is not a valid member.")
        continue
    if row['ASSIGNED_PAS'] not in pascodes:
        pascodes.append(row['ASSIGNED_PAS'])
    if row['GRADE_PERM_PROJ'] == cycle:
        print(f"{row['FULL_NAME']} IS NOT ELIGIBLE.")
        ineligible_service_members.append(index)
        continue
    if row['GRADE'] == cycle:
        member_status = board_filter(row['GRADE'], year, row['DOR'], row['UIF_CODE'], row['UIF_DISPOSITION_DATE'], row['TAFMSD'], row['REENL_ELIG_STATUS'])
        if member_status:
            print(f"{row['FULL_NAME']} IS ELIGIBLE.")
            eligible_service_members.append(index)
        else:
            print(f"{row['FULL_NAME']} IS NOT ELIGIBLE.")
            ineligible_service_members.append(index)

pascodes = sorted(pascodes)
for pascode in pascodes:
    name = input(f'Enter the name for {pascode}: ')
    rank = input(f'Enter rank for {name}: ')
    title = input(f'Enter the title for {name}: ')
    srid = '0R173'
    pascodeMap[pascode] = (name, rank, title, srid)

eligible_df = pdf_roster.loc[eligible_service_members]
ineligible_df = pdf_roster.loc[ineligible_service_members]
generate_roster_pdf(eligible_df, ineligible_df, cycle, year, pascodeMap)
