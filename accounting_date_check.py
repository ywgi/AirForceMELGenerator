from datetime import datetime
from dateutil.relativedelta import relativedelta

#Static closeout date = annual report due date
SCODs = {'SRA': f'31-MAR',
         'SSG': f'31-JAN',
         'TSG': f'30-NOV',
         'MSG': f'30-SEP',
         'SMS': f'31-JUL'
}

def accounting_date_check(date_arrived_station, grade, year):
    arrival_date = datetime.strptime(date_arrived_station, "%d-%b-%Y")
    scod = f'{SCODs.get(grade)}-{year}'
    formatted_scod_date = datetime.strptime(scod, "%d-%b-%Y")
    accounting_date = formatted_scod_date - relativedelta(days=120-1)
    adjusted_accounting_date = accounting_date.replace(day=3).replace(hour=23, minute=59, second=59)
    if arrival_date > adjusted_accounting_date:
        return False
    return True


