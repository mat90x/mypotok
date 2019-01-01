#!/usr/bin/python3
import pandas
import argparse
import os.path
import sys

parser = argparse.ArgumentParser(
    description = 'analytical tool for potok investor'
    )
parser.add_argument(
    'file'
    ,help = 'path to the potok file'
    ,default = 'potok.xlsx'
    )
parser.add_argument(
    '-s'
    ,'--sort'
    ,choices = ['start', 'end', 'loan']
    ,help = 'sort the result by field'
    ,default = 'start'
    )
parser.add_argument(
    '--header'
    ,type = int
    ,help = 'print table header after each HEADER record'
    ,default = 35
    )
parser.add_argument(
    '-r'
    ,'--reverse'
    ,action = 'store_true'
    ,help = 'reverse sorting order'
    ,default = False
    )
args = parser.parse_args()

if not os.path.isfile(args.file):
    print("file '" + args.file + "' doesn't exist!")
    sys.exit(1)

potok = pandas.read_excel(args.file)

del potok['Номер счета инвестора']
del potok['Номер счета заемщика']
potok = potok.replace(
    regex = True
    ,to_replace = [
         'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ'
        ,'Общество с ограниченной ответственностью'
        ,'Индивидуальный предприниматель'
        ,'Акционерное общество'
        ,'ЧАСТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ']
    ,value = [
         'ООО'
        ,'ООО'
        ,'ИП'
        ,'АО'
        ,'ЧАО']
    )

loans = {}
fin_tbl = []

for row in potok.itertuples():
    date = row[1]
    paid = row[2]
    typ  = row[3]
    name = row[4]
    num  = row[5]

    if typ == 'инвестирование':
        if num not in loans:
            loans[num] = len(fin_tbl)
            fin_tbl.append([num, name, date, date, paid, 0, 0])
            continue
        else:
            print("found twice")
            sys.exit(1)

    i = loans[num]

    if typ == 'выплата ОД':
        fin_tbl[i][5] = fin_tbl[i][5] + paid 
    elif typ == 'пени' or typ == 'проценты':
        fin_tbl[i][6] = fin_tbl[i][6] + paid

    if date > fin_tbl[i][3]:
        fin_tbl[i][3] = date

if args.sort == 'start':
    fin_tbl.sort(key=lambda x: x[2], reverse=args.reverse)
elif args.sort == 'end':
    fin_tbl.sort(key=lambda x: x[3], reverse=args.reverse)
elif args.sort == 'loan':
    fin_tbl.sort(key=lambda x: x[4], reverse=args.reverse)

# total mess of mysql-like table output

headers = [
'      Contract #      ',
'               Debtor name                ',
'  Start date  ',
'   End date   ',
'   Loan   ',
' Debt paid ',
' Interest ',
' Days ',
' Daily int. ',
' Int. rate '
]

header = ''
border = ''
for h in headers:
    header = header + '|' + h
    border = border + '+' + '-'*len(h)
header = header + '|'
border = border + '+'

for i, r in enumerate(fin_tbl):
    if i % args.header == 0:
        print('\n'.join([border, header, border]))

    days = (r[3] - r[2]).days
    
    daily_interest = 0
    if days:
        daily_interest = r[6] / days
    
    interest_rate = 0
    if r[4]:
        interest_rate = 100 * daily_interest * 365 / r[4]

    s = ''
    s = s + '|' + ('%-' + str(len(headers[0])) + 's'  ) % r[0]
    s = s + '|' + ('%-' + str(len(headers[1])) + 's'  ) % r[1][:len(headers[1])-1]
    s = s + '|  ' + str(r[2].date()) + '  '
    s = s + '|  ' + str(r[3].date()) + '  '
    s = s + '|' + ('%'  + str(len(headers[4])) + '.2f') % r[4]
    s = s + '|' + ('%'  + str(len(headers[5])) + '.2f') % r[5]
    s = s + '|' + ('%'  + str(len(headers[6])) + '.2f') % r[6]
    s = s + '|' + ('%'  + str(len(headers[7])) + 'd'  ) % days
    s = s + '|' + ('%'  + str(len(headers[8])) + '.2f') % daily_interest
    s = s + '|' + ('%'  + str(len(headers[9])) + '.2f') % interest_rate
    s = s + '|'

    print(s)

print(border)