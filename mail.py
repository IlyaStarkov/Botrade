import  smtplib
from work_on_excel import file_mame
from work_on_OS import get_user

message = "Запуск программы:\n\nИмя файла: " + file_mame\
          + '\n' + "Пользователь: " + get_user()

def send_mail(email, password, message):
    try:
        server = smtplib.SMTP("smtp.yandex.ru", 587)
        server.starttls()
        server.login(email, password)
        server.sendmail(email, email, message)
        server.quit()
    except smtplib.socket.gaierror:
        print("Ошибка\nСкорее всего отсутствует интернет-соединение")
        return -1

