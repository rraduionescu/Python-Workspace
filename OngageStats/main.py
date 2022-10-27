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

options = {'width': 500, 'disable-smart-width': '', 'height': 9000}
imgkit.from_file('/Users/IonescuRadu/Downloads/1_files/blank.html', 'raw_template.jpg', options=options)
print('\t Email template > JPG')

raw_image = Image.open('raw_template.jpg')
raw_pixels = raw_image.load()
bottom = 0
for y in range(raw_image.height):
    if raw_pixels[1, raw_image.height - y - 1] == (255, 255, 255):
        bottom = raw_image.height - y - 1
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
    last = [idx for idx in range(len(all_data[i])) if all_data[i][idx] == ',']
    if first and last:
        c_date = all_data[i][first[0]: last[-1]]
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
    campaign_dates_string.append(date.strftime('%d/%m/%Y'))

for i in range(len(list_names)):
    if list_names[i] == 'All Lists-Go':
        list_names[i] = 'GO'
    if list_names[i] == 'Networks':
        list_names[i] = 'NW'
    if list_names[i] == 'Never Joined':
        list_names[i] = 'NJ'
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
for i in range(len(list_names)):
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
    if list_names[i] == 'Get Openers':
        list_names_ordered[2] = list_names[i]
        campaign_dates_string_ordered[2] = campaign_dates_string[i]
        campaign_names_ordered[2] = campaign_names[i]
        subjects_ordered[2] = subjects[i]
        sent_ordered[2] = sent[i]
        open_rates_ordered[2] = open_rates[i]
        click_rates_ordered[2] = click_rates[i]
        opens_ordered[2] = opens[i]
        clicks_ordered[2] = clicks[i]

print('\nCampaign Stats ---------------------------------------------------------------------------------------------')
for i in range(len(list_names_ordered)):
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

column_A = 'Mail #'
column_B = 'Date'
column_C = 'Partner'
column_D = 'Subject'
column_E = 'Template'
column_F = 'List'
column_G = 'Sent #'
column_H = 'Open %'
column_I = 'Open #'
column_J = 'Click %'
column_K = 'Click #'
column_L = 'Unsub #'
column_M = 'Compl #'
column_N = 'Swipe'

google_client = pygsheets.authorize(service_file='ongagestats.json')
book = google_client.open('Mailer Report')
work_sheet = book[0]
work_sheet_raw = book[1]

CL = 'client_secrets.json'
AS = 'sheets'
AV = 'v4'
SC = ['https://www.googleapis.com/auth/spreadsheets']
ID = '1Em4gIwCA73nPi_dUJY5hTvSdH6b57JD9Oz5X2qnZ1AY'

service = Create_Service(CL, AS, AV, SC)
response = service.spreadsheets().values().get(
    spreadsheetId=ID,
    range='OUTGOING!A2:A',
).execute()
last_row = len(response['values']) + 4
counter = int(response['values'][-1][0].split('\n')[0]) + 1
response = service.spreadsheets().values().get(
    spreadsheetId=ID,
    range='DATA!A1:A',
).execute()
last_raw_row = len(response['values']) + 1

auth = GoogleAuth()
drive = GoogleDrive(auth)
googleFile = drive.CreateFile({'parents': [{'id': '1ECMflv2bw6ecDtt7q0TsG5Q4EIaDnsW5'}],
                               'title': datetime.datetime.strptime(campaign_dates_string_ordered[2], '%d/%m/%Y').strftime('%b %d, %Y') + '.jpg'})
googleFile.SetContentFile('template.jpg')
googleFile.Upload()
file_id = googleFile.get('id')

mail_no = (str(counter) + '\n' + 'INT') if campaign_names_ordered[0] == 'LDI' else (str(counter) + '\n' + 'EXT')
data_frame_1 = pd.DataFrame(
    {
        column_A: [mail_no],
        column_B: [campaign_dates_string_ordered[0]],
        column_C: [campaign_names_ordered[0]],
        column_D: [subjects_ordered[0]],
        column_E: ['=image("' + googleFile.get('embedLink') + '")'],
    })
data_frame_2 = pd.DataFrame(
    {
        column_F: list_names_ordered,
        column_G: sent_ordered,
        column_H: open_rates_ordered,
        column_I: opens_ordered,
        column_J: click_rates_ordered,
        column_K: clicks_ordered,
    })
data_raw = pd.DataFrame(
    {
        column_A: campaign_dates_string_ordered,
        column_B: list_names_ordered,
        column_C: campaign_names_ordered,
        column_D: subjects_ordered,
        column_E: sent_ordered,
        column_F: open_rates_ordered,
        column_G: opens_ordered,
        column_H: click_rates_ordered,
        column_I: clicks_ordered,
        column_J: [0, 0, 0],
        column_K: [0, 0, 0],
        column_L: [0, 0, 0],
        column_M: [0, 0, 0],
        column_N: ['=HYPERLINK("https://drive.google.com/uc?id=' + file_id + '", IMAGE("https://drive.google.com/uc?id=' + file_id + '"))',
                   '=HYPERLINK("https://drive.google.com/uc?id=' + file_id + '", IMAGE("https://drive.google.com/uc?id=' + file_id + '"))',
                   '=HYPERLINK("https://drive.google.com/uc?id=' + file_id + '", IMAGE("https://drive.google.com/uc?id=' + file_id + '"))']
    })


# work_sheet.set_dataframe(data_frame_1, (last_row, 1), copy_head=False)
# work_sheet.set_dataframe(data_frame_2, (last_row, 6), copy_head=False)
# work_sheet_raw.set_dataframe(data_raw, (last_raw_row, 1), copy_head=False)

print('\nDone\n\tStat data > Google Sheet')

# TODO insert unsub and complaint numbers
