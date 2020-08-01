import os

extension = "xlsx"

def search_for_an_excel_file():
    count_of_file = search_extension(os.listdir())
    #return search_extension(os.listdir())
    return count_of_file[0]


def search_extension(list):
    acc = 0
    key = 0
    for i in range(len(list)):
        if list[i].find(extension) != -1:
            acc+=1
            key = i
    return acc, key

def pause():
    print("Для завершения нажмите Enter...")
    return input()