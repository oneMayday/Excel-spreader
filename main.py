from spreader import *


if __name__ == '__main__':

    file_name = input('Введите имя файла: ')

    sheet = read_file_to_list(file_name)

    values_dict = create_dict_from_table(sheet)

    sum_sales, sum_sales_count, opt_sum_sales, opt_sales_count, opt_sales_list, clear_sales\
        = find_max_and_values(values_dict)

    print(f'Продажи тотал, шт.: {sum_sales:,}')
    print(f'Продажи розница, шт.: {clear_sales:,}')
    print(f'Продажи ОПТ, шт.: {opt_sum_sales:,}')
    print(f'Кол-во оптовых продаж: {opt_sales_count:,}')

    new_values_dict, remain, new_values_dict_sum = smudge_values(values_dict, clear_sales, sum_sales, opt_sum_sales)

    print(f'Сумма продаж после разбивки, шт.: {new_values_dict_sum:,}')
    print(f'Остаток после прогона №1: {remain:,}')
    print(f'Проверка: сумма + остаток: {(new_values_dict_sum + remain):,}')

    iteration_count = 2

    while remain:
        new_values_dict, remain = smudge_remain(new_values_dict, remain)
        print(f'Остаток после прогона №{iteration_count}: {remain:,}')
        iteration_count += 1

    result_sum = sum([value for value in new_values_dict.values()])
    print(f'Итоговая сумма значений после разбивки: {result_sum:,}')

    create_new_file(file_name, new_values_dict)

    print('Готово!')
    input()
