from re import search

"""
Номер аудитории состоит из числа и, возможно, одной буквы - пристройка
"""
def auditory(value) -> int:
    if value is not None:
        match = search(r'\d+[а-яА-Яa-zA-Z]?', value)
        if match:
            return match.group()
    return 0

"""
Обычно единственное значащее число в ячейке окружено текстом -
курс, группа, номер аудитории вне пристройки
"""
def first_number(value):
    if value is None:
        return 0
    match = search(r'\d+', value)
    if match:
        return match.group()
    return value
