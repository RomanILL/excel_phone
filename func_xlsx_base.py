import openpyxl  # открывать xlsx файлы
import os


def create_write_file(w_filename, name_first_sheet="Лист 1"):
    """ функция создания файла записи xlsx
    возвращает объект нового файла"""
    print(f'Создаем файл "{w_filename}"')
    # создаем новый excel-файл
    exit_file = openpyxl.Workbook()
    # добавляем новый лист
    exit_file.active.title = name_first_sheet
    return exit_file


def check_destination_folders(folders_list):
    for folder in folders_list:
        print(check_and_create_dir(folder))
    print("-" * 30 + "\n")


def make_cities_dict(file_name, country_select_list):
    # формируем словарь городов
    cities_xlsx = open_file_xlsx(file_name)
    cities_dict = dict()
    for i in range(2, cities_xlsx.active.max_row + 1):
        country_name = cities_xlsx.active.cell(row=i, column=2).value
        if country_name in country_select_list:
            city_name = cities_xlsx.active.cell(row=i, column=1).value
            region_name = cities_xlsx.active.cell(row=i, column=3).value
            cities_dict[city_name] = (country_name, region_name)
    cities_xlsx.close()
    del cities_xlsx


def make_regions_dict(file_name):
    # формируем словарь регионов
    regions_xlsx = open_file_xlsx(file_name)
    regions_dict = dict()
    for i in range(2, regions_xlsx.active.max_row + 1):
        for one_num in range(len(str(regions_xlsx.active.cell(row=i, column=1).value).split(", "))):
            region_name = regions_xlsx.active.cell(row=i, column=2).value
            regions_dict[one_num] = region_name
    regions_xlsx.close()
    del regions_xlsx


def print_any_list(any_list):
    """функция просто печатает списки в столбик"""
    count_any_list = len(any_list)
    for i in range(count_any_list):
        print(any_list[i])
    print("Количество элементов списка:", count_any_list)


def check_and_create_dir(dir_name):
    """функция проверяет существует ли папка, и если её нет, то создает новую"""
    if not os.path.isdir(dir_name):
        os.mkdir(dir_name)
        return f"Папка '{dir_name}' не существовала. Папка создана"
    return f"Папка '{dir_name}' существует"


def open_file_xlsx(origin_xlsx_name):
    """функция для открытия любых xlsx файлов
    возвращает кортеж: файл,
    при ошибке открытия возвращает None кортеж"""
    try:
        new_xlsx_file_object = openpyxl.load_workbook(origin_xlsx_name)
    except:
        print(f'ВНИМАНИЕ: Ошибка открытия файла "{origin_xlsx_name}"')
        input("Чтобы продолжить нажмите Enter...")
        return None
    return new_xlsx_file_object

    """ раньше это была часть функции, но надо упрощать"""

    """
    print(f'Все листы файла "{origin_xlsx_name}":', new_xlsx_file_object.sheetnames)
    new_xlsx_file_object.active = 0
    sheet = new_xlsx_file_object.active
    print("Выбран лист:", sheet)
    max_rows = sheet.max_row
    max_cols = sheet.max_column
    print("строк всего:", max_rows, "столбцов всего:", max_cols)
    return new_xlsx_file_object, sheet, max_rows, max_cols
    """


def get_heads(xlsx_object, sheet_number=0, start_row=1):
    """ функция формирования списка заголовков таблицы
    на вход принимает объект открытого файла xlsx, лист, с которого считывать, номер строки, где лежат заголовки
    нумерация листов идет с 0, нумрация строк с 1"""
    heads = list()
    if len(xlsx_object.sheetnames) < sheet_number:
        print(f"Ошибка: Количество листов в xlsx книге меньше, чем {sheet_number}")
        print(f"Будет использован лист, который был активен")
    else:
        xlsx_object.active = sheet_number
    print("Для создания заголовков выбран лист:", xlsx_object.active)
    for i in range(1, xlsx_object.active.max_column + 1):
        head = str(xlsx_object.active.cell(row=start_row, column=i).value)
        heads.append(head)
    return heads




class ConnectingTables:
    pass


class PhoneNumbersTable:
    def __init__(self, heads_table):
        self.heads = heads_table


    def get_head_xlsx(self):
        """ метод возвращает  
        :return: 
        """
        pass