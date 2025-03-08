import pandas as pd
from accounting_date_check import accounting_date_check
from board_filter import board_filter
from pdf_generator_A1C_addition import generate_roster_pdf
from final_mel_pdf_generator import generate_final_roster_pdf
from datetime import datetime, timedelta

eligible_service_members = []
ineligible_service_members = []


alpha_roster_path = rf'C:\Users\Gamer\Downloads\Book4.xlsx'
required_columns = ['FULL_NAME', 'GRADE', 'ASSIGNED_PAS_CLEARTEXT', 'DAFSC', 'DOR', 'DATE_ARRIVED_STATION', 'TAFMSD','REENL_ELIG_STATUS', 'ASSIGNED_PAS']
optional_columns = ['GRADE_PERM_PROJ', 'UIF_CODE', 'UIF_DISPOSITION_DATE','CAFSC']
reject_columns = ['SSAN', 'DATE_OF_BIRTH', 'HOME_ADDRESS','HOME_CITY','HOME_STATE','HOME_ZIP_CODE','MARITAL_STATUS', 'SPOUSE_SSAN', 'HOME_PHONE_NUMBER',
                  'SEC_CLR','TYPE_SEC_INV','DT_SCTY_INVES_COMPL','SEC_ELIG_DT','TECH_ID','ACDU_STATUS','ANG_ROLL_INDICATOR','AFR_SECTION_ID','CIVILIAN_ART_ID','ATTACHED_PAS']
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

cycle = 'SSG'
year = 2025

for index, row in filtered_alpha_roster.iterrows():
    for column, value in row.items():
        if pd.isna(value) and column in required_columns:
            valid_upload = False
            break
    valid_member = accounting_date_check(row['DATE_ARRIVED_STATION'], cycle, year)
    if not valid_member:
        continue
    if row['ASSIGNED_PAS'] not in pascodes:
        pascodes.append(row['ASSIGNED_PAS'])
    if row['GRADE_PERM_PROJ'] == cycle:
        continue
    if row['GRADE'] == cycle:
        member_status = board_filter(row['GRADE'], year, row['DOR'], row['UIF_CODE'], row['UIF_DISPOSITION_DATE'], row['TAFMSD'], row['REENL_ELIG_STATUS'], row['CAFSC'])
        if member_status is True:
            eligible_service_members.append(index)
        else:
            ineligible_service_members.append((index, member_status))


pascodes = sorted(pascodes)
for pascode in pascodes:
    name = input(f'Enter the name for {pascode}: ')
    rank = input(f'Enter rank for {name}: ')
    title = input(f'Enter the title for {name}: ')
    srid = input(f'Enter the SRID for {pascode} ')
    pascodeMap[pascode] = (name, rank, title, srid)

# Create DataFrames
eligible_df = pdf_roster.loc[eligible_service_members]

# For ineligible, we need to handle the reason
ineligible_df = pd.DataFrame()
for idx, reason in ineligible_service_members:
    row = pdf_roster.loc[idx].copy()
    row['REASON'] = reason
    # Use concat instead of append
    ineligible_df = pd.concat([ineligible_df, pd.DataFrame([row])], ignore_index=True)

print("\nGenerating PDFs...")
generate_roster_pdf(eligible_df, ineligible_df, cycle, year, pascodeMap)
generate_final_roster_pdf(eligible_df, ineligible_df, cycle, year, pascodeMap)
print("PDF generation complete!")
