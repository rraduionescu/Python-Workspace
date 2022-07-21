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
    f = [idx for idx in range(len(all_data[i])) if all_data[i][idx].isupper()]
    l = [idx for idx in range(len(all_data[i])) if all_data[i][idx] == '2']
    if f and l:
        c_date = all_data[i][f[0]: l[-1] + 1]
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
    open_rates.append(float(str(all_data[i + 2]).replace('%', '', 1)) / 100)
for i in click_indices:
    clicks.append(int(str(all_data[i]).replace(',', '', 1)))
    click_rates.append(float(str(all_data[i + 2]).replace('%', '', 1)) / 100)

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
column_a = 'A'
column_2 = 'List'
column_3 = 'Subject'
column_4 = 'Mailer count'
column_5 = 'Open Rate'
column_b = 'B'
column_c = 'C'
column_6 = 'CTR'
column_7 = 'Sent number'
column_8 = 'Open Number'
column_9 = 'Clicks Number'

data_frame = pd.DataFrame(
    {
        column_1: campaign_dates_string,
        column_a: ['', '', '', ''],
        column_2: list_names,
        column_3: subjects,
        column_4: ['', '', '', ''],
        column_5: open_rates,
        column_b: ['', '', '', ''],
        column_c: ['', '', '', ''],
        column_6: click_rates,
        column_7: sent,
        column_8: opens,
        column_9: clicks
    })
data_frame.to_excel('/Users/IonescuRadu/Downloads/5.xlsx', sheet_name='sheet1', index=False)

google_client = pygsheets.authorize(service_file='/Users/IonescuRadu/Downloads/ongagestats-6436213ca729.json')
book = google_client.open('MAILER_REPORT')
work_sheet = book[0]

work_sheet.set_dataframe(data_frame, (494, 2), copy_head=False)
