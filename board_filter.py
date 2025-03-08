
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Static closeout date = annual report due date
SCODs = {'SRA': f'31-MAR',
         'SSG': f'31-JAN',
         'TSG': f'30-NOV',
         'MSG': f'30-SEP',
         'SMS': f'31-JUL'
         }

# Time in grade = how long you've been in a rank
TIG = {'SRA': f'01-AUG',
       'SSG': f'01-JUL',
       'TSG': f'01-MAY',
       'MSG': f'01-MAR',
       'SMS': f'01-DEC'
       }

tig_months_required = {
    'SRA': 6,
    'SSG': 23,
    'TSG': 24,
    'MSG': 20,
    'SMS': 21
}

# Total active federal military service date = time in military (years)
TAFMSD = {
    'SRA': 3,
    'SSG': 5,
    'TSG': 8,
    'MSG': 11,
    'SMS': 14
}

# manditory date of separation = the day you have to exit the military
mdos = {
    'SRA': f'01-SEP',
    'SSG': f'01-AUG',
    'TSG': f'01-AUG',
    'MSG': f'01-APR',
    'SMS': f'01-JAN'
}

main_higher_tenure = {
    'SRA': 10,
    'SSG': 20,
    'TSG': 22,
    'MSG': 24,
    'SMS': 26
}

exception_higher_tenure = {
    'SRA': 12,
    'SSG': 22,
    'TSG': 24,
    'MSG': 26,
    'SMS': 28
}

# reenlistment codes
re_codes = {
    "2A": "HQ AFPC denied reenlistment for quality reasons.",
    "2B": "Discharged under general conditions.",
    "2C": "Involuntary separation with honorable discharge.",
    "2F": "Undergoing rehabilitation in a DOD facility.",
    "2G": "Failed Substance Abuse Treatment for drugs.",
    "2H": "Failed Substance Abuse Treatment for alcohol.",
    "2J": "Under investigation, may result in discharge.",
    "2K": "Notified of involuntary separation.",
    "2M": "Serving or separated while under sentence.",
    "2P": "AWOL; deserter.",
    "2W": "Retired and recalled to active duty.",
    "2X": "Not selected for reenlistment.",
    "4H": "Ineligible due to Article 15.",
    "4I": "Ineligible due to Control Roster.",
    "4J": "Ineligible due to AF Weight Management Program.",
    "4K": "Medically disqualified or pending evaluation.",
    "4L": "Separated from a commissioning program.",
    "4M": "Breach of enlistment agreement.",
    "4N": "Convicted by civil authority."
}

exception_hyt_start_date = datetime(2023, 12, 8)
exception_hyt_end_date = datetime(2025, 9, 30)


def board_filter(grade, year, date_of_rank, uif_code, uif_disposition_date, tafmsd, re_status, cafsc):
    formatted_date_of_rank = datetime.strptime(date_of_rank, "%d-%b-%Y")
    formatted_tafmsd = datetime.strptime(tafmsd, "%d-%b-%Y")
    scod = f'{SCODs.get(grade)}-{year}'
    formatted_scod = datetime.strptime(scod, "%d-%b-%Y")
    tig_selection_month = f'{TIG.get(grade)}-{year}'
    formatted_tig_selection_month = datetime.strptime(tig_selection_month, "%d-%b-%Y")
    tig_eligibility_month = formatted_tig_selection_month - relativedelta(months=tig_months_required.get(grade))
    tafmsd_required_date = formatted_tig_selection_month - relativedelta(years=TAFMSD.get(grade) - 1)
    hyt_date = datetime.strptime(tafmsd, "%d-%b-%Y") + relativedelta(years=main_higher_tenure.get(grade))
    mdos_date = formatted_tig_selection_month + relativedelta(months=1)

    # Add safe datetime conversion for UIF disposition date
    try:
        if uif_disposition_date:
            formatted_uif_disposition_date = datetime.strptime(uif_disposition_date, "%d-%b-%Y")
        else:
            formatted_uif_disposition_date = None
    except (ValueError, TypeError):
        formatted_uif_disposition_date = None

    # Check eligibility criteria and return specific reason if not eligible
    if formatted_date_of_rank > tig_eligibility_month:
        return f"TIG: <{tig_months_required.get(grade)} months"

    if formatted_tafmsd > tafmsd_required_date:
        return f"TIS: <{TAFMSD.get(grade)} years"

    if exception_hyt_start_date < hyt_date < exception_hyt_end_date:
        hyt_date += relativedelta(years=2)

    if hyt_date < mdos_date:
        return "HYT: Mandatory DOS"

    if uif_code and uif_code > 1 and formatted_uif_disposition_date and formatted_uif_disposition_date < formatted_scod:
        return f"UIF: Code {uif_code}"

    if re_status and re_status in re_codes:
        # Get the description and truncate it if needed
        description = re_codes[re_status]
        short_desc = description[:30] + "..." if len(description) > 30 else description
        return f"RE {re_status}: {short_desc}"

    # Check CAFSC skill level requirements
    if cafsc and isinstance(cafsc, str) and len(cafsc) >= 5:
        # Skip 8X000 career fields (exempt from the skill level requirements)
        if not cafsc.startswith('8'):
            # Get the skill level (5th position)
            skill_level = cafsc[4] if len(cafsc) > 4 else '0'

            # A1C requires 5-level for SrA
            if grade == 'A1C' and skill_level < '5':
                return "AFSC: Requires 5-skill level"

            # SSG requires 7-level for TSG
            if grade == 'SSG' and skill_level < '7':
                return "AFSC: Requires 7-skill level"

            # MSG requires 9-level for SMS
            if grade == 'MSG' and skill_level < '9':
                return "AFSC: Requires 9-skill level"

    # If all checks pass, return True for eligible
    return True