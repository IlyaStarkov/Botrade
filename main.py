from work_on_OS import search_for_an_excel_file
from work_on_OS import pause
from work_on_excel import validation_of_excel_file
from work_on_excel import get_name_of_file



if search_for_an_excel_file() == 0:
    print("В данной директории нет excel-файла")
    pause()
elif search_for_an_excel_file() > 1:
    print("В данной директории находится более одного excel-файла")
    pause()
else:
    print(get_name_of_file())
    pause()
