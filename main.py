from work_on_OS import pause, \
     make_file, make_folder, open_folder
from work_on_excel import validation_of_excel_file
from work_on_excel import main_function
from mail import send_mail, message


if validation_of_excel_file():
    print("Файл не соответствует шаблону")
    pause()
else:
    # if send_mail("logicApp@yandex.ru", "hczignejcbzufxxr", message.encode("utf-8")) == -1:
    #     pause()
    # else:
        make_folder()
        print("Начинаем проверку...\n")
        make_file(main_function())
        open_folder()
