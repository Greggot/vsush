
from datetime import datetime
from openpyxl import load_workbook

"""
8 пар в день, по 2 строчки на числитель-знаменатель + 
пропуск строки перед следующим днем
"""
ROWS_IN_DAY = 17
OFFSETS = [5]

def calc_offsets() -> int:
    for i in range(5):
        OFFSETS.append(OFFSETS[-1] + ROWS_IN_DAY)

def date_to_schedule_column_offset(date: datetime) -> int:
    return OFFSETS[date.weekday()]

def time_to_schedule_column(date: datetime) -> int:
    hours = [9, 11, 13, 15, 16, 18, 20, 21]
    minutes = [35, 20, 5, 00, 45, 30, 00, 30]
    for i in range(len(hours)):
        newdate = datetime(date.year, date.month, date.day, hours[i], minutes[i])
        if date < newdate:
            return i
    return 0 

def date_to_offset(date: datetime) -> int:
    return date_to_schedule_column_offset(date) + 2 * time_to_schedule_column(date)

if __name__ == "__main__":
    wb = load_workbook('schedule.xlsx')
    ws = wb.active
    calc_offsets()


    date = datetime.strptime('04.09.2025 9:10', '%d.%m.%Y %H:%M')
    print(date_to_offset(date))

    date = datetime.strptime('04.09.2025 9:35', '%d.%m.%Y %H:%M')
    print(date_to_offset(date))

    date = datetime.strptime('04.09.2025 9:45', '%d.%m.%Y %H:%M')
    print(date_to_offset(date))

    date = datetime.strptime('04.09.2025 11:20', '%d.%m.%Y %H:%M')
    print(date_to_offset(date))

    # print(date_to_schedule_column_offset(date))
    # print(f"{date.strftime('%A')}")