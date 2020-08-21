import  smtplib
from work_on_excel import get_name_of_file
from work_on_OS import get_user, pause

message = "Запуск программы:\n\nИмя файла: " + get_name_of_file()\
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

