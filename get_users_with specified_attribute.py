import simple_functions
from openpyxl import Workbook, load_workbook
from ADinfo import username_list, surname_list, getinfo_by_username, \
                   get_all_workers, getinfo_by_surname, get_status_by_username, \
                   get_status_by_surname, get_users_from_group, get_users_from_specified_containers, \
                   get_users_with_specified_attribute, save_data_to_excel
import datetime
now = datetime.datetime.now()
if now.month < 10:
        month_string = '0' + str(now.month)
else:
        month_string = str(now.month)

if now.day < 10:
        day_string = '0' + str(now.day)
else:
        day_string = str(now.day)
date_string = day_string + month_string

result_username =[]
result_surname = []
status = []

name = input('Введите название атрибута: ')
value = input('Введите значение атрибута:')
usernames, surnames = get_users_with_specified_attribute(name, value)
for i in usernames:
        status.append(get_status_by_username(i))
save_data_to_excel(usernames, surnames,status, colona2 = 'usernames', colonna2='Surnames', STATUS = 'STATUS')
#test = save_data_to_excel(usernames, surnames,status, filename = '.\\results\\testing.xlsx', colona2 = 'usernames', colonna2='Surnames', STATUS = 'STATUS')
