from openpyxl import Workbook, load_workbook
from datetime import date
import os

webinar = False
conference = False

wb = load_workbook('/Users/cameronhakenson/triage_transfer.xlsx')
ws = wb.active

today = date.today()
full_date = f"{today.month}/{today.day}/23"

ws['A1'] = f'CAMERON {full_date}'
ws['A12'] = 'Overview from the organization:'
ws['A23'] = '* * *' ; ws['A25'] = '* * *'
ws['A26'] = 'T41-Cameron'

date = ""
title_and_date = ""

event_check = input('Webinar or Conference (web or con): ')
if event_check == 'web':
    webinar = True
elif event_check == 'con':
    conference = True

url_id = input('Enter URL ID: ')
ws['A5'] = url_id

full_url = input('Enter Full URL: ')
ws['A6'] = full_url

month = input('Enter the month:  ')
day = input('Enter the day: ')
year = input('Enter the year: ')
if webinar:
    time = input('Enter the time: ')
    ws['A9'] = f'Date of event: {month}. {day}, {year}, {time}'
    date = f' Webinar Scheduled for {month}. {day}, {year}'
elif conference:
    ws['A9'] = f'Date of event: {month}. {day}, {year}'
    date = f' Conference Scheduled for {month}. {day}, {year}'

title = input('Enter the title: ')
title_and_date = title.title() + date
ws['A8'] = f'Title: {title.title()}'
ws['A7'] = title_and_date

if webinar:
    ws['A10'] = ''
    ws['A11'] = ''
elif conference:
    location = input('Enter location: ')
    ws['A10'] = 'Location of event:'
    ws['A11'] = location

if webinar:
    os.system("open /Users/cameronhakenson/triage_description.txt")
    blank_read = input('Press Enter to continue...')
    with open('triage_description.txt') as file_description:
        ws['A14'] = file_description.read()
    with open('triage_description.txt', 'w') as file_description:
        file_description.write('')
    ws['A15'] = ""
elif conference:
    os.system("open /Users/cameronhakenson/triage_agenda.txt")
    blank_read = input('Press Enter to continue...')
    with open('triage_agenda.txt') as file_agenda:
        ws['A15'] = 'Agenda: ' + file_agenda.read()
    with open('triage_agenda.txt', 'w') as file_agenda:
        file_agenda.write('')
    ws['a14'] = ''

if webinar:
    ws['A19'] = ''
elif conference:
    link_agenda = input('Enter agenda link: ')
    ws['A19'] = f'Link to full agenda: {link_agenda}'

speakers = input('Enter speakers: ')
if speakers == "" or speakers == " ":
    ws['A20'] = ""
else:
    ws['A20'] = f'Speakers: {speakers}'

sponsors = input('Enter sponsors: ')
if sponsors == "" or sponsors == " ":
    ws['A21'] = ""
    ws['A22'] = ""
else:
    ws['A21'] = '* * *'
    ws['A22'] = f'Sponsors: {sponsors}'

register = input('Enter registration link: ')
ws['A24'] = f'Registration: {register}'

wb.save('/Users/cameronhakenson/triage_transfer.xlsx')
os.system("open /Users/cameronhakenson/triage_transfer.xlsx")
