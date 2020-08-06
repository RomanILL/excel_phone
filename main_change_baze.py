import os
from glob import glob
import func_xlsx_base as ffp
from make_phone_func import make_good_phone_list

if __name__ == "__main__":

    # блок констант
    COUNTRY_SELECT = ["Россия"]
    exit_phone_base_file_name = "exit_phone_base.xlsx"
    dir_input_files = "input_xlsx"
    dir_output_files = "output_xlsx"
    dir_supporting_files = "support_tables"
    support_cities_table_name = "cities.xlsx"
    support_regions_table_name = "regions.xlsx"
    ignore_cols_list = (1, 2, 11, 12, 13, 14, 18, 21, 22, 23, 25, 27)

    # индексы заголовков, где что лежит
    # FIXME переделать в удобный блок или словарь

    """
    parent_city_head_id = 0 родитель
    address_registration_head_id = 4 юр.адрес
    actual_address_head_id = 5 факт. адрес
    phones_list_head_id = 7 список телефонов
    vehicle_number_head_id = 11 номер ТС
    driver_phone_head_id = 13 номер водителя
    """

    parent_city_head_id = 0
    address_registration_head_id = 4
    actual_address_head_id = 5
    phones_list_head_id = 7
    vehicle_number_head_id = 11
    driver_phone_head_id = 13

    # проверяем папки назначения (откуда брать, куда класть, вспомогательная)
    ffp.check_destination_folders((dir_input_files, dir_output_files, dir_supporting_files))

    # собираем список файлов для обработки
    phone_base_files_list = glob(f'{dir_input_files}\\*.xlsx')
    # печатаем список файлов для обработки - FIXME можно убрать после отладки
    ffp.print_any_list(phone_base_files_list)
    print("-" * 30 + "\n")

    # формируем список заголовков таблиц из первого файла - перенесено в главный цикл программы

    # список - добавка в шапку таблиц
    # FIXME после отладки блока принятия решений оставить только город / регион, страну и телефон по формату
    extend_head_list = ["Страна", "Город Юр.лица", "Город Факт.", "Регион", "Регион по номеру", "Телефон компании"]

    # формируем вспомогательные данные (словари городов и номеров регионов)
    # формируем словарь городов
    cities_dict = ffp.make_cities_dict(dir_supporting_files + "\\" + support_cities_table_name, COUNTRY_SELECT)
    # формируем словарь регионов
    region_dict = ffp.make_regions_dict(dir_supporting_files + "\\" + support_regions_table_name)

    # перебираем исходные файлы (признак первого открытого файла - пустая переменная heads)
    heads = None
    for current_file_name in phone_base_files_list:
        if "~" not in current_file_name:
            temp_file_name = current_file_name.split("\\")[-1]
            print(f'Открываем файл для чтения: {temp_file_name}')
            current_xlsx_obj = ffp.open_file_xlsx(current_file_name)
            if heads is None:
                # если это первый файл из списка, то собираем заголовки

                print(f"Читаем заголовки таблицы {temp_file_name}")
                heads = ffp.get_heads(current_xlsx_obj, ignore_id_list=ignore_cols_list)
                # печатаем список заголовков
                # FIXME можно убрать после отладки
                ffp.print_any_list(heads)
                print("-" * 30 + "\n")

                # открываем файл для записи
                if os.path.isfile(dir_output_files + "\\" + exit_phone_base_file_name):
                    # delete current file
                    print(f'Удаляется ранее созданный файл "{exit_phone_base_file_name}"')
                    os.remove(dir_output_files + "\\" + exit_phone_base_file_name)
                exit_phone_xlsx = ffp.create_write_file(dir_output_files + "\\" + exit_phone_base_file_name)
                # расширяем заголовки
                heads.extend(extend_head_list)
                # прописываем заголовки в новый файл
                exit_phone_xlsx.active.append(heads)

            # блок обработки файла
            # берем первый по индексу лист
            current_xlsx_obj.active = 0

            numbers_rows = current_xlsx_obj.active.max_row
            numbers_cols = current_xlsx_obj.active.max_column

            for i_row in range(2, numbers_rows + 1):
                row_for_write = []
                extend_row = [None, None, None, None, None, None]
                """ "Страна", "Город Юр.лица", "Город Факт.", "Регион", "Регион по номеру", "Телефон по формату" """
                """ [27, 28, 29, 30, 31, 32] """
                for j_col in range(1, numbers_cols + 1):
                    # индексы заголовков, где что лежит
                    """
                    parent_city_head_id = 0 родитель
                    address_registration_head_id = 4 юр.адрес
                    actual_address_head_id = 5 факт. адрес
                    phones_list_head_id = 7 список телефонов
                    vehicle_number_head_id = 11 номер ТС
                    driver_phone_head_id = 13 номер водителя
                    """

                    # собираем новую строку, исключая ненужные столбцы
                    if j_col not in ignore_cols_list:
                        current_cell_value = str(current_xlsx_obj.active.cell(row=i_row, column=j_col).value)
                        row_for_write.append(current_cell_value)
                # в этом моменте у нас собралась строка для записи из оригинального файла

                # собираем строку extend_row из оригинала
                # ищем города в адресах
                for city_name in cities_dict:
                    if city_name in row_for_write[address_registration_head_id]:
                        extend_row[1] = city_name
                    if city_name in row_for_write[actual_address_head_id]:
                        extend_row[2] = city_name
                # определяем регион и страну по городу (приоритет фактическому адресу)
                if extend_row[2] is not None:

                    extend_row[0] = cities_dict[extend_row[2]][0]
                    extend_row[3] = cities_dict[extend_row[2]][1]
                elif extend_row[1] is not None:
                    extend_row[0] = cities_dict[extend_row[1]][0]
                    extend_row[3] = cities_dict[extend_row[1]][1]
                # определяем регион по номеру машины
                for region_id in region_dict:
                    if region_id in row_for_write[vehicle_number_head_id]:
                        extend_row[4] = region_dict[region_id]

                # блок работы с телефонами
                # находим сотовые телефоны и приводим их к стандартному виду
                phone_numbers_string = row_for_write[phones_list_head_id]
                # получаем список мобильных телефонов компании
                company_phone_list = make_good_phone_list(phone_numbers_string)
                # получаем список мобильных телефонов водителя
                phone_numbers_string = row_for_write[driver_phone_head_id]
                driver_phone_list = make_good_phone_list(phone_numbers_string)
                all_phone_list = company_phone_list + driver_phone_list
                row_for_write.extend(extend_row)
                for write_phone in all_phone_list:
                    row_for_write[20] = write_phone
                    exit_phone_xlsx.active.append(row_for_write)

            # блок закрытия файла после использования
            current_xlsx_obj.close()
        else:
            pass
    exit_phone_xlsx.save(dir_output_files + "\\" + exit_phone_base_file_name)
    print(f"В файле {exit_phone_base_file_name} сохранены изменения")
    exit_phone_xlsx.close()
    print("Программа успешно завершена.")
    input('Для заверения нажмите Enter...')
