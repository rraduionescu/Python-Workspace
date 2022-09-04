from __future__ import print_function
from html.parser import HTMLParser
from PIL import Image
from Google import Create_Service
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pygsheets
import imgkit
import pandas as pd
import datetime
import glob
import os


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

options = {'width': 500, 'disable-smart-width': '', 'height': 5000}
imgkit.from_file('/Users/IonescuRadu/Downloads/1_files/blank.html', 'raw_template.jpg', options=options)
print('\t Email template > JPG')

raw_image = Image.open('raw_template.jpg')
raw_pixels = raw_image.load()
bottom = 0
for y in range(raw_image.height):
    if raw_pixels[1, y] == (215, 215, 215):
        bottom = y
        break
template = raw_image.crop((0, 0, raw_image.width, bottom))
template.save('template.jpg')

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
    if ',' in str(campaign_names[i]):
        campaign_names[i] = str(campaign_names[i]).replace(',', '\n').title()


list_names_ordered = list_names.copy()
campaign_dates_string_ordered = campaign_dates_string.copy()
campaign_names_ordered = campaign_names.copy()
subjects_ordered = subjects.copy()
sent_ordered = sent.copy()
open_rates_ordered = open_rates.copy()
click_rates_ordered = click_rates.copy()
opens_ordered = opens.copy()
clicks_ordered = clicks.copy()
for i in range(4):
    if list_names[i] == 'Never Joined':
        list_names_ordered[0] = list_names[i]
        campaign_dates_string_ordered[0] = campaign_dates_string[i]
        campaign_names_ordered[0] = campaign_names[i]
        subjects_ordered[0] = subjects[i]
        sent_ordered[0] = sent[i]
        open_rates_ordered[0] = open_rates[i]
        click_rates_ordered[0] = click_rates[i]
        opens_ordered[0] = opens[i]
        clicks_ordered[0] = clicks[i]
    if list_names[i] == 'Network':
        list_names_ordered[1] = list_names[i]
        campaign_dates_string_ordered[1] = campaign_dates_string[i]
        campaign_names_ordered[1] = campaign_names[i]
        subjects_ordered[1] = subjects[i]
        sent_ordered[1] = sent[i]
        open_rates_ordered[1] = open_rates[i]
        click_rates_ordered[1] = click_rates[i]
        opens_ordered[1] = opens[i]
        clicks_ordered[1] = clicks[i]
    if list_names[i] == 'Network Guest':
        list_names_ordered[2] = list_names[i]
        campaign_dates_string_ordered[2] = campaign_dates_string[i]
        campaign_names_ordered[2] = campaign_names[i]
        subjects_ordered[2] = subjects[i]
        sent_ordered[2] = sent[i]
        open_rates_ordered[2] = open_rates[i]
        click_rates_ordered[2] = click_rates[i]
        opens_ordered[2] = opens[i]
        clicks_ordered[2] = clicks[i]
    if list_names[i] == 'Get Openers':
        list_names_ordered[3] = list_names[i]
        campaign_dates_string_ordered[3] = campaign_dates_string[i]
        campaign_names_ordered[3] = campaign_names[i]
        subjects_ordered[3] = subjects[i]
        sent_ordered[3] = sent[i]
        open_rates_ordered[3] = open_rates[i]
        click_rates_ordered[3] = click_rates[i]
        opens_ordered[3] = opens[i]
        clicks_ordered[3] = clicks[i]

print('\nCampaign Stats ---------------------------------------------------------------------------------------------')
for i in range(4):
    print(
        'Date: {} Campaign: {} List: {} \nSubject: {} \n\tSent: {} \n\tOpen rate: {} \n\tClick rate: {} \n\tOpen: {} \n\tClicks: {}'.format(
            campaign_dates_string_ordered[i],
            str(campaign_names_ordered[i]).replace('\n', ','),
            list_names_ordered[i],
            subjects_ordered[i],
            sent_ordered[i],
            open_rates_ordered[i],
            click_rates_ordered[i],
            opens_ordered[i],
            clicks_ordered[i],
        ))
print('------------------------------------------------------------------------------------------------------------')

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

google_client = pygsheets.authorize(service_file='ongagestats.json')
book = google_client.open('MAILER REPORT TEST')
work_sheet = book[1]

CL = '/Users/IonescuRadu/Downloads/secret.json'
AS = 'sheets'
AV = 'v4'
SC = ['https://www.googleapis.com/auth/spreadsheets']
ID = '1kKda6OBFREefrlOBBSeIuEBGT_DttdTbtmfk5qjrrWE'

service = Create_Service(CL, AS, AV, SC)
response = service.spreadsheets().values().get(
    spreadsheetId=ID,
    range='OUTGOING | EXTERNAL!F4:F',
).execute()
last_row = len(response['values']) + 4
counter = int(response['values'][-1][0]) + 1
campaign_names_ordered.append(campaign_names_ordered[3])

auth = GoogleAuth()
drive = GoogleDrive(auth)
gfile = drive.CreateFile({'parents': [{'id': '1zzkbvlaaUZBQbJeD6-UktFT368Ui94Jl'}]})
gfile.SetContentFile('template.jpg')
gfile.Upload()
file_id = gfile.get('id')
print('testestest')
print(gfile)
print(gfile.get('embedLink'))

# ============ INTERNAL FORMAT ============
# data_frame_internal = pd.DataFrame(
#     {
#         column_1: campaign_dates_string_ordered,
#         'Empty1': ['', '', '', ''],
#         column_2: list_names_ordered,
#         column_3: subjects_ordered,
#         'Empty2': ['', '', '', ''],
#         column_5: open_rates_ordered,
#         'Empty3': ['', '', '', ''],
#         column_7: campaign_names_ordered[],
#         column_8: click_rates_ordered,
#         column_9: sent_ordered,
#         column_10: opens_ordered,
#         column_11: clicks_ordered,
#     })

# ============ EXTERNAL FORMAT ============
data_frame_external1 = pd.DataFrame(
    {
        column_1: campaign_dates_string_ordered,
        'Empty1': ['', '', '', ''],
        column_2: list_names_ordered,
        column_3: subjects_ordered,
    })
data_frame_external2 = pd.DataFrame(
    {
        'Empty2': ['', '', '', '', str(counter)],
    })
data_frame_external3 = pd.DataFrame(
    {
        column_5: open_rates_ordered,
        column_6: ['', '', '', ''],
    })
data_frame_external4 = pd.DataFrame(
    {
        column_7: campaign_names_ordered,
    })
data_frame_external5 = pd.DataFrame(
    {
        column_8: click_rates_ordered,
        'Empty3': ['', '', '', ''],
        'Empty4': ['', '', '', ''],
        'Empty5': ['', '', '', ''],
        'Empty6': ['', '', '', ''],
        column_9: sent_ordered,
        'Empty7': ['', '', '', ''],
        'Empty8': ['', '', '', ''],
        'Empty9': ['', '', '', ''],
        'Empty10': ['', '', '', ''],
        column_10: opens_ordered,
        column_11: clicks_ordered,
    })
data_frame_external6 = pd.DataFrame(
    {
        #column_6: ['=image(\"https://drive.google.com/uc?export=view&id={}\")'.format(file_id)],
        column_6: ['=image("' + gfile.get('embedLink') + '")'],
    })

work_sheet.set_dataframe(data_frame_external1, (last_row, 2), copy_head=False)
work_sheet.set_dataframe(data_frame_external2, (last_row, 6), copy_head=False)
work_sheet.set_dataframe(data_frame_external3, (last_row, 7), copy_head=False)
work_sheet.set_dataframe(data_frame_external4, (last_row, 9), copy_head=False)
work_sheet.set_dataframe(data_frame_external5, (last_row, 10), copy_head=False)
work_sheet.set_dataframe(data_frame_external6, (last_row, 8), copy_head=False)

print('\nDone\n\tStat data > Google Sheet')

# TODO 1. Order lists !!!               DONE
#      2. Email template height         DONE
#      3. Insert after last line !!!    DONE
#      4. Insert template !!!           ALMOST
#      5. Mailer sheet connection !!!
