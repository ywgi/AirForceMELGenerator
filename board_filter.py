from datetime import datetime
from dateutil.relativedelta import relativedelta

#Static closeout date = annual report due date
SCODs = {'SRA': f'31-MAR',
         'SSG': f'31-JAN',
         'TSG': f'30-NOV',
         'MSG': f'30-SEP',
         'SMS': f'31-JUL'
}

#Time in grade = how long you've been in a rank
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

#Total active federal military service date = time in military (years)
TAFMSD = {
    'SRA': 3,
    'SSG': 5,
    'TSG': 8,
    'MSG': 11,
    'SMS': 14
}

#manditory date of separation = the day you have to exit the military
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

#reenlistment codes
re_codes = ["2B", "4A", "4B", "4C", "4D", "4E", "4F", "4G", "4H", "4I", "4J", "4K", "4L", "4M", "4N"]

exception_hyt_start_date = datetime(2023, 12, 8)
exception_hyt_end_date = datetime(2025, 9,30)


def board_filter(grade, year, date_of_rank, uif_code, uif_disposition_date, tafmsd, re_status):
    formatted_date_of_rank = datetime.strptime(date_of_rank,"%d-%b-%Y")
    formatted_tafmsd = datetime.strptime(tafmsd,"%d-%b-%Y")
    scod = f'{SCODs.get(grade)}-{year}'
    tig_selection_month = f'{TIG.get(grade)}-{year}'
    formatted_tig_selection_month = datetime.strptime(tig_selection_month, "%d-%b-%Y")
    tig_eligibility_month = formatted_tig_selection_month - relativedelta(months=tig_months_required.get(grade))
    tafmsd_required_date = formatted_tig_selection_month - relativedelta(years=TAFMSD.get(grade)-1)
    hyt_date = datetime.strptime(tafmsd, "%d-%b-%Y") + relativedelta(years=main_higher_tenure.get(grade))
    mdos = formatted_tig_selection_month + relativedelta(months=1)

    if formatted_date_of_rank > tig_eligibility_month:
        return False
    if formatted_tafmsd > tafmsd_required_date:
        return False
    if exception_hyt_start_date < hyt_date < exception_hyt_end_date:
        hyt_date += relativedelta(years=2)
    if hyt_date < mdos:
        return False
    if uif_code > 1 and uif_disposition_date < scod:
        return False
    if re_status in re_codes:
        return False
    return True
