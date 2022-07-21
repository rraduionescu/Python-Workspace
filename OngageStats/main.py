import pygsheets
import pandas as pd
import datetime
import glob
import os
from html.parser import HTMLParser


class OngageStatParser(HTMLParser):
    def handle_data(self, data):
        global all_data
        all_data.append(data)


def validate_date(date_text):
    try:
        return datetime.datetime.strptime(date_text, '%b %d, %Y')
    except ValueError:
        pass


all_data = []

list_names = []
campaign_dates = []
subjects = []
campaign_names = []
sent = []
opens = []
open_rates = []
clicks = []
click_rates = []

html_data = ''
parser = OngageStatParser()

path = '/Users/IonescuRadu/Downloads'
for filename in glob.glob(os.path.join(path, '*.html')):
    with open(os.path.join(os.getcwd(), filename), 'r') as file:
        html_data += file.read()
parser.feed(html_data)

campaign_date_indices = [(i + 2) for i, x in enumerate(all_data) if 'Schedule' in x]
subject_indices = [(i + 1) for i, x in enumerate(all_data) if 'Subject:' in x]
list_name_indices = [(i + 2) for i, x in enumerate(all_data) if 'List Name' in x]
campaign_name_indices = [(i + 1) for i, x in enumerate(all_data) if 'Name:' in x]
sent_indices = [(i + 2) for i, x in enumerate(all_data) if 'Sent -' in x]
open_indices = [(i + 2) for i, x in enumerate(all_data) if 'Unique Opens -' in x]
click_indices = [(i + 2) for i, x in enumerate(all_data) if 'Unique Clicks -' in x]

for i in campaign_date_indices:
    first = [idx for idx in range(len(all_data[i])) if all_data[i][idx].isupper()]
    last = [idx for idx in range(len(all_data[i])) if all_data[i][idx] == '2']
    if first and last:
        c_date = all_data[i][first[0]: last[-1] + 1]
        if validate_date(c_date):
            campaign_dates.append(validate_date(c_date))
for i in subject_indices:
    subjects.append(all_data[i])
for i in list_name_indices:
    list_names.append(''.join([c for c in all_data[i] if c.isupper() or c == ' ' or c == '-']).strip().title())
for i in campaign_name_indices:
    campaign_names.append(all_data[i].strip())
for i in sent_indices:
    sent.append(int(str(all_data[i]).replace(',', '', 1)))
for i in open_indices:
    opens.append(int(str(all_data[i]).replace(',', '', 1)))
    open_rates.append(all_data[i + 2])
for i in click_indices:
    clicks.append(int(str(all_data[i]).replace(',', '', 1)))
    click_rates.append(all_data[i + 2])

campaign_dates_string = []
for date in campaign_dates:
    campaign_dates_string.append(date.strftime('%d.%m.%Y'))

for i in range(4):
    if list_names[i] == 'All Lists-Go':
        list_names[i] = 'Get Openers'
    if list_names[i] == 'Networks':
        list_names[i] = 'Network'
    if list_names[i] == 'Networks Guests':
        list_names[i] = 'Network Guest'

for i in range(4):
    print('Date: {} Subject: {} Open rate: {} Click rate: {} Sent: {} Open: {} Clicks: {} List: {}, Campaign: {}'.format(
        campaign_dates_string[i],
        subjects[i],
        open_rates[i],
        click_rates[i],
        sent[i],
        opens[i],
        clicks[i],
        list_names[i],
        campaign_names[i],
    ))

column_1 = 'Date'
column_2 = 'List'
column_3 = 'Subject'
column_4 = 'Mailer count'
column_5 = 'Open Rate'
column_6 = 'Template'
column_7 = 'Partner|Network|Channel'
column_8 = 'CTR'
column_9 = 'Sent number'
column_10 = 'Open Number'
column_11 = 'Clicks Number'

# ============ INTERNAL FORMAT ============
data_frame_internal = pd.DataFrame(
    {
        column_1: campaign_dates_string,
        'Empty1': ['', '', '', ''],
        column_2: list_names,
        column_3: subjects,
        'Empty2': ['', '', '', ''],
        column_5: open_rates,
        'Empty3': ['', '', '', ''],
        column_7: campaign_names,
        column_8: click_rates,
        column_9: sent,
        column_10: opens,
        column_11: clicks
    })

# ============ EXTERNAL FORMAT ============
data_frame_external = pd.DataFrame(
    {
        column_1: campaign_dates_string,
        'Empty1': ['', '', '', ''],
        column_2: list_names,
        column_3: subjects,
        'Empty2': ['', '', '', ''],
        column_5: open_rates,
        column_6: ['', '', '', ''],
        column_7: campaign_names,
        column_8: click_rates,
        'Empty3': ['', '', '', ''],
        'Empty4': ['', '', '', ''],
        'Empty5': ['', '', '', ''],
        'Empty6': ['', '', '', ''],
        column_9: sent,
        'Empty7': ['', '', '', ''],
        'Empty8': ['', '', '', ''],
        'Empty9': ['', '', '', ''],
        'Empty10': ['', '', '', ''],
        column_10: opens,
        column_11: clicks
    })


# data_frame.to_excel('/Users/IonescuRadu/Downloads/5.xlsx', sheet_name='sheet1', index=False)

google_client = pygsheets.authorize(service_file='/Users/IonescuRadu/Downloads/ongagestats-6436213ca729.json')
book = google_client.open('MAILER_REPORT')
work_sheet = book[1]

work_sheet.set_dataframe(data_frame_external, (4, 2), copy_head=False)
# TODO - order lists !!!
