import os
import datetime


def make_folder():
    try:
        os.mkdir("Reports")
    except OSError:
        pass


def pause():
    print("Для завершения нажмите Enter...")
    return input()


def get_user():
    return os.getlogin()


def get_date():
    return datetime.date.today()


def make_file(number_of_mechanics):
    print("Cоздание файла продаж...")
    f = open("Reports/" + number_of_mechanics[5] + "_Информация о количестве продаж.txt", 'w')
    f.write("Европейский:\n")
    f.write("SOB 4 уп. капсул в подарок - " + str(number_of_mechanics[0]) + "\n")
    f.write("SOB Core девайс в подарок - " + str(number_of_mechanics[1]) + "\n\n")
    f.write("FR 2+1 - " + str(number_of_mechanics[2]) + "\n")
    f.write("FR 4+2 - " + str(number_of_mechanics[3]) + "\n")
    f.write("FR 6+3 - " + str(number_of_mechanics[4]) + "\n\n")
    f.write("Чехол в подарок - " + str(number_of_mechanics[7]) + "\n\n")
    f.write("Ошибок - " + str(number_of_mechanics[6]) + "\n")
    f.close()


def date_picker():
    print("Если данный отчёт по сегодняшнему дню, нажмитие Enter.\n"
          "В ином случае введите дату:\n")
    name = input()
    if name == '':
        return str(get_date())
    else:
        return str(name)


def open_folder():
    os.system(r"explorer.exe Reports")