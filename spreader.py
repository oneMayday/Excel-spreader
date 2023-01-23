from openpyxl import *
from random import randint
from math import ceil


def read_file_to_list(file_name: str) -> list:
    # Читаем файл

    work_table = input('Введите рабочий диапазон таблицы: ')
    book = load_workbook(filename=file_name)
    sheet = book.worksheets[0][work_table.split(':')[0] : work_table.split(':')[1]]
    return sheet


def create_dict_from_table(sheet: list) -> dict:
    # Формируем словарь из ячеек со значениями отличными от None

    values_dict = {}
    for row in sheet:
        for cell in row:
            if cell.value is not None:
                values_dict.setdefault(cell.coordinate, cell.value)
    return values_dict


def find_max_and_values(values_dict: dict, limit_value=300) -> tuple[int, int, int, int, list, int]:
    # Находим из словаря максимальное количество продаж (и формируем из них список),
    # общее число продаж, количество оптовых продаж, сумму оптовых продаж

    sum_sales = 0
    sum_sales_count = 0
    opt_sum_sales = 0
    opt_sales_count = 0
    opt_sales_list = []

    for key, value in values_dict.items():
        if value >= limit_value:
            opt_sales_list.append((key, value))
            opt_sum_sales += value
            values_dict[key] = randint(0, 2)
            opt_sales_count += 1
        sum_sales += value
        sum_sales_count += 1

    clear_sales = sum_sales - opt_sum_sales

    return sum_sales, sum_sales_count, opt_sum_sales, opt_sales_count, opt_sales_list, clear_sales


def smudge_values(values_dict: dict, clear_sales: int, sum_sales: int, opt_sum_sales: int, limit_value=300)\
        -> tuple[dict, int, int]:
    # Приравниваем максимумы к (1, 10), размазываем значения по таблице, в соответствии с их процентным соотношением

    new_values_dict = values_dict.copy()
    sum_values_dict = sum([value for value in new_values_dict.values()])

    for key, value in new_values_dict.items():
        if value != 1:
            new_value = ceil(new_values_dict[key] + (new_values_dict[key] / clear_sales) * opt_sum_sales)
            difference = new_value - value

            if new_value < limit_value:
                if sum_values_dict + difference < sum_sales:
                    new_values_dict[key] = new_value
                    sum_values_dict += difference
                else:
                    new_diff = sum_sales - sum_values_dict
                    new_values_dict[key] = value + new_diff
                    sum_values_dict += new_diff
                    break
            else:
                corr = randint(1, 10)
                new_value = limit_value - corr
                difference = new_value - value
                new_values_dict[key] = new_value
                sum_values_dict += difference

    # Сумма новых значений в словаре
    new_values_dict_sum = sum([value for value in new_values_dict.values()])

    # Остаток после разбивки
    remain = sum_sales - sum_values_dict

    return new_values_dict, remain, new_values_dict_sum


def smudge_remain(new_values_dict: dict, remain: int, limit_value=300) -> tuple[dict, int]:
    # Создаем новый файл и наполняем его значениями из словаря

    remain_dec = remain % 10
    remain_all = (remain // 10) * 10

    for key, value in new_values_dict.items():
        if remain_dec:
            if 31 <= value and (value + remain_dec) < limit_value:
                new_values_dict[key] += remain_dec
                remain_dec = 0
        elif remain_all:
            if remain_all > 25:
                if 51 <= value and (value + 25) < limit_value:
                    new_values_dict[key] += 25
                    remain_all -= 25
            else:
                if 51 <= value and (value + remain_all) < limit_value:
                    new_values_dict[key] += remain_all
                    remain_all -= remain_all
                    break
                else:
                    continue
        else:
            continue

    remain = remain_all

    return new_values_dict, remain


def create_new_file(file_name: str, new_values_dict: dict) -> None:
    # Соаздем новый файл и копируем в него данные из получившегося словаря

    wb = load_workbook(filename=file_name)
    ws = wb.active

    for key, value in new_values_dict.items():
        ws[key] = value

    wb.save('new_' + file_name)


# Рабочий вариант!!!!
#
# for key, value in dict2.items():
#     if value != 1:
#         new_value = ceil(dict2[key] + (dict2[key] / clear_sales) * sum_max_sales)
#         difference = new_value - value
#
#         if sum_dict2 + difference < sum_sales:
#             dict2[key] = new_value
#             sum_dict2 += difference
#         else:
#             new_diff = sum_sales - sum_dict2
#             dict2[key] = value + new_diff
#             sum_dict2 += new_diff
#             break
