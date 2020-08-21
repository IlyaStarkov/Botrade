from work_on_OS import search_for_an_excel_file
from work_on_OS import pause
from work_on_excel import validation_of_excel_file
from work_on_excel import main_function
from mail import send_mail, message



if search_for_an_excel_file() == 0:
    print("В данной директории нет excel-файла")
    pause()

elif search_for_an_excel_file() > 1:
    print("В данной директории находится более одного excel-файла")
    pause()
else:

    if validation_of_excel_file() == 1:
        print("Файл не соответствует шаблону")
        pause()
    else:
        if send_mail("logicApp@yandex.ru", "hczignejcbzufxxr", message.encode("utf-8")) == -1:
            pause()
        else:
            print("Начинаем проверку...\n")
            main_function()
            pause()

