from openpyxl import Workbook, load_workbook
from datetime import date
import os
import event


def main():
    x = 0
    events = True
    while events:
        if x == 0:
            event.url()
        event.restart_sheet()
        event.date_made()
        web = event.event_check()
        month = input('Enter the month:  ')
        day = input('Enter the day: ')
        year = '2023' # *** HARD CODED, CHANGE AT END OF YEAR ***
        if web:
            event.title_section(event.webinar_date(month, day, year))
            event.web_description()
        if not web:
            event.title_section(event.conference_date(month, day, year))
            event.con_location()
            event.con_agenda()
        event.speakers()
        event.sponsors()
        event.register()
        event.wb.save('/Users/cameronhakenson/triage_transfer.xlsx') # Location of excel file
        os.system("open /Users/cameronhakenson/triage_transfer.xlsx")
        stay = input('Are you staying on the same website? (yes / no): ')
        if stay == 'no':
            break
        else:
            x += 1
main()
