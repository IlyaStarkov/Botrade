import os
from work_on_OS import search_extension


def get_name_of_file():
    return os.listdir()[search_extension(os.listdir())[1]]

def validation_of_excel_file():
    pass