import sys
import os
import getpass

import xlrd, xlwt

from netmiko import ConnectHandler
from login import logincisco
from findcrash import findcrash


USERNAME = logincisco()
print('Your username: ', USERNAME)
PASSWORD = getpass.getpass()
Command = 'show flash:'
count = 1
countfailed = 1
countcrash = 1
bookresult = xlwt.Workbook()
sheetcrash = bookresult.add_sheet('Crashinfo')
sheetfail = bookresult.add_sheet('No connection')


sheetcrash.write(0, 0, 'hostname')
sheetcrash.write(0, 1, 'IP address')

sheetfail.write(0, 0, 'hostname')
sheetfail.write(0, 1, 'IP address')


rb=xlrd.open_workbook("C:\\...\\loopbacks.xlsx")

sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]



for IP in vals:

  try:
        DEVICE_PARAMS = {'device_type': 'cisco_ios',
                         'ip': IP[1],
                         'username': USERNAME,
                         'password': PASSWORD}

        print('\nConnecting to device {}...'.format(IP[1]))
        with ConnectHandler(**DEVICE_PARAMS) as ssh:

            ssh.enable()
            flash = ssh.send_command(Command)


  except Exception:
      print('Unable to connect')
      print('Devices checked: ', count)
      sheetfail.write(countfailed, 0, IP[0])
      sheetfail.write(countfailed, 1, IP[1])
      countfailed = countfailed + 1
      count = count + 1
      continue

  match = findcrash(flash)

  if match == None:
     print('Not Found')
  else:
     print('File CRASHINFO found!')
     sheetcrash.write(countcrash, 0, IP[0])
     sheetcrash.write(countcrash, 1, IP[1])
     countcrash = countcrash + 1
  print('Devices checked: ', count)
  count = count + 1

bookresult.save('C:\\...\\crashes.xls')

