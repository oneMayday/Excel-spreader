from spreader import *
from openpyxl.utils.cell import coordinate_to_tuple


if __name__ == '__main__':

    # Считываем файл и адрес первой ячейки с товарами
    file_name = input('Введите имя файла: ')
    start_cell = input('Введите координаты первой ячейки с наименованием товара (например A2): ')
    book = load_workbook(filename=file_name)
    sheet = book.active

    # Находим координаты первой ячейки и количество столбцов в таблице
    start_cell_x, start_cell_y = coordinate_to_tuple(start_cell)
    start_column_values = start_cell_y + 1
    count_max_column = sheet.max_column

    # Объявляем результирующий список словарей значений и промежуточный словарь
    result_values_list = []
    values_dict = {}

    # Главный цикл
    for row in sheet.iter_rows(min_row=start_cell_x, min_col=start_cell_y, max_col=count_max_column):
        counter = 1

        if row[0].value not in (0, None):
            counter += 1

        if counter == 2:
            counter = 1

            if values_dict:

                sum_sales, sum_sales_count, opt_sum_sales, opt_sales_count, opt_sales_list, clear_sales \
                    = find_max_and_values(values_dict)

                # print(f'Продажи тотал, шт.: {sum_sales:,}')
                # print(f'Продажи розница, шт.: {clear_sales:,}')
                # print(f'Продажи ОПТ, шт.: {opt_sum_sales:,}')
                # print(f'Кол-во оптовых продаж: {opt_sales_count:,}')

                new_values_dict, remain, new_values_dict_sum = smudge_values(values_dict, clear_sales, sum_sales,
                                                                             opt_sum_sales)

                # print(f'Сумма продаж после разбивки, шт.: {new_values_dict_sum:,}')
                # print(f'Остаток после прогона №1: {remain:,}')
                # print(f'Проверка: сумма + остаток: {(new_values_dict_sum + remain):,}')

                iteration_count = 2

                while remain:
                    new_values_dict, remain = smudge_remain(new_values_dict, remain)
                    print(f'Остаток после прогона №{iteration_count}: {remain:,}')
                    iteration_count += 1

                result_sum = sum([value for value in new_values_dict.values()])
                print(f'Итоговая сумма значений после разбивки: {result_sum:,}')

                result_values_list.append(new_values_dict)

            values_dict = {}

        values_dict = create_dict_from_table(values_dict, row, start_column_values)

    if values_dict:
        sum_sales, sum_sales_count, opt_sum_sales, opt_sales_count, opt_sales_list, clear_sales \
            = find_max_and_values(values_dict)
        new_values_dict, remain, new_values_dict_sum = smudge_values(values_dict, clear_sales, sum_sales,
                                                                     opt_sum_sales)
        iteration_count = 2

        while remain:
            new_values_dict, remain = smudge_remain(new_values_dict, remain)
            print(f'Остаток после прогона №{iteration_count}: {remain:,}')
            iteration_count += 1

        result_sum = sum([value for value in new_values_dict.values()])
        print(f'Итоговая сумма значений после разбивки: {result_sum:,}')

        result_values_list.append(new_values_dict)

    create_new_file(file_name, result_values_list)

    print('Готово!')
    input()