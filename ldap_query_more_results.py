from typing import List

from ldap3 import Server, Connection, ALL, NTLM, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES, ObjectDef, AttrDef, Reader, \
    Writer, Entry, Attribute, OperationalAttribute, SUBTREE
import os
import csv
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Color, colors, PatternFill, Alignment
from openpyxl.styles.borders import Border, Side

import re
import datetime

pswd = ''
sAMAccountName = []

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

server = Server('ldap://svodc01.svo.air.loc', port=389, use_ssl=True, get_info=ALL)
conn = Connection(server, user="SVO.AIR.LOC\\av.kuzmin", password=pswd, authentication=NTLM)
conn.bind()
conn.start_tls()
search_filter ='(&(' + str('company') + '=' + str('ОАО "Международный аэропорт Шереметьево"') + ')' + '(objectclass=user))'
result = conn.extend.standard.paged_search('DC=SVO,DC=AIR,DC=LOC', search_filter, search_scope=SUBTREE,attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES], generator=False) #attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES])
conn.extend.standard.paged_search('DC=SVO,DC=AIR,DC=LOC', search_filter, search_scope=SUBTREE,attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES], generator=False)
i = 0
entries = []
for item in result:
    #print(item)
    entries.append(item)
    i += 1
print(i)  
print(entries[1]['attributes']['uSNCreated'])
