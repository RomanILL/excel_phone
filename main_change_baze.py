import openpyxl  # открывать xlsx файлы
import re
import os
from glob import glob
import func_xlsx_base as ffp


if __name__ == "__main__":

    # блок констант
    COUNTRY_SELECT = ["Россия"]
    exit_phone_base_file_name = "exit_phone_base.xlsx"
    dir_input_files = "input_xlsx"
    dir_output_files = "output_xlsx"
    dir_supporting_files = "support_tables"
    support_cities_table_name = "cities.xlsx"
    support_regions_table_name = "regions.xlsx"

    # проверяем папки назначения (откуда брать, куда класть, вспомогательная)
    ffp.check_destination_folders((dir_input_files, dir_output_files, dir_supporting_files))


    # собираем список файлов для обработки
    phone_base_files_list = glob(f'{dir_input_files}\\*.xlsx')
    # печатаем список файлов для обработки - FIXME можно убрать после отладки
    ffp.print_any_list(phone_base_files_list)
    print("-" * 30 + "\n")

    # формируем список заголовков таблиц из первого файла - перенесено в главный цикл программы

    # индексы заголовков, где что лежит
    # FIXME переделать в удобный блок или словарь
    parent_city_head_id = 2
    address_registration_head_id = 6
    actual_address_head_id = 7
    phones_list_head_id = 9
    vehicle_number_head_id = 18
    driver_phone_head_id = 23

    # список - добавка в шапку таблиц
    # FIXME после отладки блока принятия решений оставить только город / регион, страну и телефон по формату
    extend_head_list = ["Страна", "Город Юр.лица", "Город Факт.", "Регион", "Регион по номеру", "Телефон по формату"]
    extend_id_head_list = [27, 28, 29, 30, 31, 32]

    # формируем вспомогательные данные (словари городов и номеров регионов)
    # формируем словарь городов
    ffp.make_cities_dict(dir_supporting_files + "\\" + support_cities_table_name, COUNTRY_SELECT)
    # формируем словарь регионов
    ffp.make_regions_dict(dir_supporting_files + "\\" + support_regions_table_name)

    # открываем файл для записи
    if os.path.isfile(dir_output_files + "\\" + exit_phone_base_file_name):
        # delete current file
        print(f'Удаляется ранее созданный файл "{exit_phone_base_file_name}"')
        os.remove(dir_output_files + "\\" + exit_phone_base_file_name)
    exit_phone_xlsx = ffp.create_write_file(dir_output_files + "\\" + exit_phone_base_file_name)

    # перебираем исходные файлы (признак первого открытого файла - пустая переменная heads)
    heads = None
    for current_file_name in phone_base_files_list:
        print(f"Открываем файл для чтения: {current_file_name}")
        current_xlsx_obj = ffp.open_file_xlsx(current_file_name)
        if heads is None:
            # если это первый файл из списка, то собираем заголовки
            print(f"Читаем заголовки таблицы {current_file_name}")
            heads = ffp.get_heads(current_xlsx_obj)
            # печатаем список заголовков
            # FIXME можно убрать после отладки
            ffp.print_any_list(heads)
            print("-" * 30 + "\n")

        # блок обработки файла
        # берем первый по индексу лист
        current_xlsx_obj.active = 0
        active_sheet = current_xlsx_obj.active
        numbers_rows = active_sheet.max_row
        numbers_cols = len(heads)


        # блок закрытия файла после использования
        current_xlsx_obj.close()

    exit_phone_xlsx.save(dir_output_files + "\\" + exit_phone_base_file_name)
    exit_phone_xlsx.close()
    print("Программа успешно завершена.")
    input('Для заверения нажмите Enter...')
