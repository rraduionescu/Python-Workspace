import datetime
import glob
import os
from html.parser import HTMLParser


class OngageStatParser(HTMLParser):
    def handle_data(self, data):
        global all_data
        all_data.append(data)


def validate(date_text):
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
file = open("Campaigns Overview.html", "r")

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
        if validate(c_date):
            campaign_dates.append(validate(c_date))
for i in subject_indices:
    subjects.append(all_data[i])
for i in list_name_indices:
    list_names.append(''.join([c for c in all_data[i] if c.isupper() or c == ' ' or c == '-']).strip())
for i in campaign_name_indices:
    campaign_names.append(all_data[i].strip())
for i in sent_indices:
    sent.append(all_data[i])
for i in open_indices:
    opens.append(all_data[i])
    open_rates.append(all_data[i + 2])
for i in click_indices:
    clicks.append(all_data[i])
    click_rates.append(all_data[i + 2])

for i in range(4):
    print('Date: {} Subject: {} Open rate: {} Click rate: {} Sent: {} Open: {} Clicks: {} List: {}'.format(
        campaign_dates[i].strftime('%d.%m.%Y'),
        subjects[i],
        open_rates[i],
        click_rates[i],
        sent[i],
        opens[i],
        clicks[i],
        list_names[i],
    ))
