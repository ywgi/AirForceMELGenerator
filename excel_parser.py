import pandas as pd
from accounting_date_check import accounting_date_check
from board_filter import board_filter
from initial_mel_pdf_generator import generate_roster_pdf
from datetime import datetime, timedelta

eligible_service_members = []
ineligible_service_members = []

alpha_roster_path = rf'C:\Users\Trent\Documents\Alpha Roster.xlsx'
test_path = rf'C:\Users\Trent\Documents\7 Oct 2024 - Sanitized Alpha Roster.xlsx'
required_columns = ['FULL_NAME', 'GRADE', 'ASSIGNED_PAS_CLEARTEXT', 'DAFSC', 'DOR', 'DATE_ARRIVED_STATION', 'TAFMSD','REENL_ELIG_STATUS', 'ASSIGNED_PAS', 'CAFSC']
optional_columns = ['GRADE_PERM_PROJ', 'UIF_CODE', 'UIF_DISPOSITION_DATE', '2AFSC', '3AFSC', '4AFSC']
pdf_columns = ['FULL_NAME','GRADE', 'DATE_ARRIVED_STATION','DAFSC', 'ASSIGNED_PAS_CLEARTEXT', 'DOR', 'TAFMSD', 'ASSIGNED_PAS']
pascodes = []
pascodeMap = {}
reason_for_ineligible_map = {}
boards = ['E5', 'E6', 'E7', 'E8', 'E9']
grade_map = {
    "SRA": "E4",
    "SSG": "E5",
    "TSG": "E6",
    "MSG": "E7",
    "SMS": "E8"
}

alpha_roster = pd.read_excel(alpha_roster_path, parse_dates=True)
filtered_alpha_roster = alpha_roster[required_columns + optional_columns]
pdf_roster = filtered_alpha_roster[pdf_columns]
valid_upload = True

# cycle = input('Enter Cycle: ')
# year = input('Enter Year: ')
cycle = 'SRA'
year = 2025

for index, row in filtered_alpha_roster.iterrows():
    for column, value in row.items():
        if pd.isna(value) and column in required_columns:
            valid_upload = False
            print(rf"error at {index}, {column}")
            break
        if isinstance(value, pd.Timestamp):
            row[column] = value.strftime('%d-%b-%Y').upper()
            continue
    valid_member = accounting_date_check(row['DATE_ARRIVED_STATION'], cycle, year)
    if not valid_member:
        continue
    if row['ASSIGNED_PAS'] not in pascodes:
        pascodes.append(row['ASSIGNED_PAS'])
    if row['GRADE_PERM_PROJ'] == cycle:
        ineligible_service_members.append(index)
        reason_for_ineligible_map[index] = f'Projected for {cycle}.'
        continue
    if row['GRADE'] == cycle or (row['GRADE'] == 'A1C' and cycle == 'SRA'):
        member_status = board_filter(row['GRADE'], year, row['DOR'], row['UIF_CODE'], row['UIF_DISPOSITION_DATE'], row['TAFMSD'], row['REENL_ELIG_STATUS'], row['CAFSC'], row['2AFSC'], row['3AFSC'], row['4AFSC'])
        if member_status == None:
            continue
        elif member_status == True:
            eligible_service_members.append(index)
        elif member_status[0] == False:
            ineligible_service_members.append(index)
            reason_for_ineligible_map[index] = member_status[1]

pascodes = sorted(pascodes)
for pascode in pascodes:
    name = input(f'Enter the name for {pascode}: ')
    rank = input(f'Enter rank for {name}: ')
    title = input(f'Enter the title for {name}: ')
    srid = '0R173'
    pascodeMap[pascode] = (name, rank, title, srid)

eligible_df = pdf_roster.loc[eligible_service_members]
for column in eligible_df.columns:
    if pd.api.types.is_datetime64_any_dtype(eligible_df[column].dtype):
        eligible_df[column] = eligible_df[column].dt.strftime('%d-%b-%Y').str.upper()

ineligible_df = pdf_roster.loc[ineligible_service_members].copy()
ineligible_df['REASON'] = ineligible_df.index.map(reason_for_ineligible_map)
for column in ineligible_df.columns:
    if pd.api.types.is_datetime64_any_dtype(ineligible_df[column].dtype):
        ineligible_df[column] = ineligible_df[column].dt.strftime('%d-%b-%Y').str.upper()

print(ineligible_df)

generate_roster_pdf(eligible_df, ineligible_df, cycle, year, pascodeMap)
