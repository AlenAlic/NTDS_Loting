import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import openpyxl

document_name = 'NTDS 2018 Inschrijflijst (Responses)'
end_date = datetime.date(datetime.now())

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(credentials)
sheet = client.open(document_name).sheet1

late_signup = sheet.get_all_records()
new_contestants = list()
title = sheet.row_values(1)
title.append('Nr.')
title.append('Teamcaptain')
del title[0]
order = [22, 1, 2, 3, 0, 5, 6, 7, 8, 9, 10, 11, 12, 23, 15, 16, 17, 18, 19, 20, 21, 14, 13, 4]
title = [title[i] for i in order]


for index, contestant in enumerate(late_signup):
    timestamp = datetime.date(datetime.strptime(contestant['Timestamp'], '%d/%m/%Y %H:%M:%S'))
    signup_date = timestamp - end_date
    if signup_date.days <= 0:
        new_entry = list()
        new_entry.append(index + 1)
        new_entry.append(contestant['Voornaam'])
        new_entry.append(contestant['Tussenvoegsel'])
        new_entry.append(contestant['Achternaam'])
        new_entry.append(contestant['Email address'])
        new_entry.append(contestant['Ballroom niveau'])
        new_entry.append(contestant['Latin niveau'])
        new_entry.append(contestant['Hoe heet je Ballroom partner?'])
        new_entry.append(contestant['Hoe heet je Latin partner?'])
        new_entry.append(contestant['Ballroom rol'])
        new_entry.append(contestant['Latin rol'])
        new_entry.append(contestant['Ballroom verplicht blind daten'])
        new_entry.append(contestant['Latin verplicht blind daten'])
        new_entry.append('')
        new_entry.append(contestant['EHBO'])
        new_entry.append(contestant['BHV'])
        new_entry.append(contestant['Jury Ballroom'])
        new_entry.append(contestant['Jury Latin'])
        new_entry.append(contestant['Student'])
        new_entry.append(contestant['Slaapplek'])
        new_entry.append(contestant['Allergiën / Dieet'])
        new_entry.append(contestant['Wil je vrijwilliger zijn voor dit NTDS?'])
        new_entry.append(contestant['Ben je op een eerder ETDS of NTDS vrijwilliger geweest?'])
        new_entry.append(contestant['Team'])
        new_contestants.append(new_entry)

for index, contestant in enumerate(new_contestants):
    for i, item in enumerate(contestant):
        if item == 'Nee':
            new_contestants[index][i] = ''

ballroom_partners = list()
for contestant in new_contestants:
    ballroom_partners.append(contestant[7])

latin_partners = list()
for contestant in new_contestants:
    latin_partners.append(contestant[8])

contestant_names = list()
for contestant in new_contestants:
    if contestant[2] != '':
        seq = (contestant[1], contestant[2], contestant[3])
    else:
        seq = (contestant[1], contestant[3])
    contestant_names.append(' '.join(seq))

for index, contestant in enumerate(contestant_names):
    if contestant in ballroom_partners:
        ballroom_partners[ballroom_partners.index(contestant)] = index+1
    if contestant in latin_partners:
        latin_partners[latin_partners.index(contestant)] = index+1

for index, partner in enumerate(ballroom_partners):
    new_contestants[index][7] = partner

for index, partner in enumerate(latin_partners):
    new_contestants[index][8] = partner


workbook = openpyxl.Workbook()
worksheet = workbook.worksheets[0]
# worksheet.title = 'Backup'
for row in range(len(new_contestants)):
    for column in range(len(new_contestants[0])):
        cell = worksheet.cell(row=row + 1, column=column + 1)
        cell.value = new_contestants[row][column]
workbook.save('NTDS_Backup.xlsx')

print('Done!')
