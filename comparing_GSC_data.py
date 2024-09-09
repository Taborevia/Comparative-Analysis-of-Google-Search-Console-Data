import os
import sys
import csv
import datetime
import argparse
from asyncio import exceptions
import calendar

from fpdf import FPDF
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.auth.exceptions import RefreshError

# Zakresy, które pozwalają na odczytanie danych z GSC
SCOPES = ['https://www.googleapis.com/auth/webmasters.readonly']
SCOPES_SPREADSHEET = ['https://www.googleapis.com/auth/spreadsheets']


# authorization for GSC
def authorize(credentials_file):
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES_SPREADSHEET)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


# authorization for spreadsheet
def authorize_spreadsheet(credentials_file):
    creds = None
    if os.path.exists('token_spreadsheets.json'):
        creds = Credentials.from_authorized_user_file('token_spreadsheets.json', SCOPES_SPREADSHEET)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES_SPREADSHEET)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES_SPREADSHEET)
            creds = flow.run_local_server(port=0)
        with open('token_spreadsheets.json', 'w') as token:
            token.write(creds.to_json())

    # Połączenie z Google Sheets
    return creds


# data insertion to spreadsheet
def insert_data(data, service, spreadsheet_id):
    body = {'values': data}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="A1:D5",
        valueInputOption="RAW",
        body=body
    ).execute()


def create_spreadsheet(token, title):
    creds = authorize_spreadsheet(token)
    service = build("sheets", "v4", credentials=creds)
    spreadsheet_body = {
        "properties": {
            "title": title
        },
        'sheets': [
            {'properties': {'title': 'Sheet1'}}
        ]
    }
    spreadsheet = service.spreadsheets().create(body=spreadsheet_body).execute()
    return service, spreadsheet  # .get("spreadsheetId"), spreadsheet['sheets'][0]['properties']['sheetId']


# Sending request to GSC for data from `site_url` between `start_date` and `end_date`

def get_search_console_data(service, site_url, start_date, end_date):
    request = {
        'startDate': start_date,  # Początkowa data
        'endDate': end_date,  # Końcowa data
    }

    # Wysłanie zapytania do API i pobranie odpowiedzi
    response = service.searchanalytics().query(siteUrl=site_url, body=request).execute()
    return response['rows'][0]


# Comparison GSC data for two time periods
def comparison(oldResults, newResults):
    results = [["Metrics", "2023", "2024", "Change(%)"]]
    for key in oldResults.keys():
        if newResults[key] > oldResults[key]:
            # print(f'{key} increased by {(newResults[key] / oldResults[key]) * 100 - 100:.2f}%')
            results.append([key, float(f'{oldResults[key]:.2f}'), float(f'{newResults[key]:.2f}'),
                            float(f'{(newResults[key] / oldResults[key]) * 100 - 100:.2f}')])
        else:
            # print(f'{key} decreased by {100 - (newResults[key] / oldResults[key]) * 100:.2f}%')
            results.append([key, float(f'{oldResults[key]:.2f}'), float(f'{newResults[key]:.2f}'),
                            float(f'-{100 - (newResults[key] / oldResults[key]) * 100:.2f}')])
    return results


def generate_csv_report(data, filename):
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for row in data:
            writer.writerow(row)


def generate_pdf_report(data, filename):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Nagłówki
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(200, 10, txt="Report SEO", ln=True, align='C')

    # Dodanie tabeli danych
    pdf.ln(10)
    pdf.set_font('Arial', '', 10)
    # Zapis danych
    for row in data:
        for item in row:
            pdf.cell(40, 10, str(item), border=1)
        pdf.ln()

    # Zapis pliku PDF
    pdf.output(filename)


def generate_spreadsheets_report(token, data, title):
    service, spreadsheet = create_spreadsheet(token, title)
    spreadsheet_ID = spreadsheet.get("spreadsheetId")
    insert_data(data, service, spreadsheet_ID)
    sheet_ID = spreadsheet['sheets'][0]['properties']['sheetId']
    current_date = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    year = current_date.year
    month = current_date.strftime("%B")
    # Utworzenie wykresu słupkowego
    request_body = {
        "requests": [
            {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": f"Changes of metrics {year - 1} {month} - {year} {month}",
                            "basicChart": {
                                "chartType": "COLUMN",
                                "legendPosition": "BOTTOM_LEGEND",
                                "axis": [
                                    {
                                        "position": "BOTTOM_AXIS",
                                        "title": "Metrics"
                                    },
                                    {
                                        "position": "LEFT_AXIS",
                                        "title": "Results"
                                    }
                                ],
                                "domains": [
                                    {
                                        "domain": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": sheet_ID,
                                                        "startRowIndex": 0,
                                                        "endRowIndex": 5,
                                                        "startColumnIndex": 0,
                                                        "endColumnIndex": 1
                                                    }
                                                ]
                                            }
                                        }
                                    }
                                ],
                                "series": [
                                    {
                                        "series": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": sheet_ID,
                                                        "startRowIndex": 0,
                                                        "endRowIndex": 5,
                                                        "startColumnIndex": 3,
                                                        "endColumnIndex": 4
                                                    }
                                                ]
                                            }
                                        },
                                        "targetAxis": "LEFT_AXIS"
                                    }
                                ],
                                "headerCount": 1
                            }
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": sheet_ID,
                                    "rowIndex": 1,
                                    "columnIndex": 5
                                },
                                "offsetXPixels": 0,
                                "offsetYPixels": 0
                            }
                        }
                    }
                }
            }
        ]
    }

    # Wykonanie zapytania o dodanie wykresu
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_ID,
        body=request_body
    ).execute()


def generate_report(data, file_format, filename):
    if file_format == "PDF":
        generate_pdf_report(data, filename)
    elif file_format == "CSV":
        generate_csv_report(data, filename)
    else:
        raise ValueError("wrong file name")


def datesInput():
    dates = []
    dates.append(datetime.datetime.strptime((input("Podaj datę (YYYY-MM): ") + '-01'), '%Y-%m-%d').date())
    if dates[0].month == 12:
        dates.append(datetime.date(dates[0].year + 1, 1, 1) - datetime.timedelta(days=1))
    else:
        dates.append(datetime.date(dates[0].year, dates[0].month + 1, 1) - datetime.timedelta(days=1))
    dates.append(datetime.date(dates[0].year - 1, dates[0].month, dates[0].day))
    dates.append(datetime.date(dates[1].year - 1, dates[1].month, dates[1].day))
    return dates


def main():
    print("---------Aplikacja porównawcza danych z GSC---------")
    dates = datesInput()
    credentials = authorize(CREDENTIALS_FILE)
    service = build('searchconsole', 'v1', credentials=credentials)
    newResults = get_search_console_data(service, SITE_URL, str(dates[0]), str(dates[1]))
    oldResults = get_search_console_data(service, SITE_URL, str(dates[2]), str(dates[3]))
    results = comparison(oldResults, newResults)
    while True:
        answer = input("Czy chcesz wygenerowac raport? (Tak/Nie)")
        if answer == "Tak":
            while True:
                answer2 = input("W jakim formacie chcesz zapisac plik? (CSV/PDF)")
                if answer2 == 'PDF' or answer2 == 'CSV':
                    break
            while True:
                answer3 = input("Jak chcesz nazwac plik z raportem?")
                if not os.path.exists(answer3):
                    break
            generate_report(results, answer2, answer3)
            break
            generate_report(results, input("Jaka nazwa pliku?"))
        elif answer == "Nie":
            break


def forSchedule(token, url):
    dates = [datetime.date.today()]
    if dates[0].month == 1:
        dates.insert(0, datetime.date(dates[0].year - 1, 12, 1) - datetime.timedelta(days=1))
    else:
        dates.insert(0, datetime.date(dates[0].year, dates[0].month - 1, 1))
    dates.append(datetime.date(dates[0].year - 1, dates[0].month, dates[0].day))
    dates.append(datetime.date(dates[1].year - 1, dates[1].month, dates[1].day))
    credentials = authorize(token)
    service = build('searchconsole', 'v1', credentials=credentials)
    newResults = get_search_console_data(service, url, str(dates[0]), str(dates[1]))
    oldResults = get_search_console_data(service, url, str(dates[2]), str(dates[3]))
    results = comparison(oldResults, newResults)
    current_date = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    year = current_date.year
    month = current_date.strftime("%B")
    generate_spreadsheets_report(token, results, f"SEO Report - {month} {year} ({url})")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument('--url', type=str, help="URL of the site from which data will be retrieved")
    parser.add_argument('--token', type=str, help="path to JSON file with credentials", default="client_secret.json")
    parser.add_argument('--schedule', action='store_true', help="schedule mode")

    args = parser.parse_args()

    if args.schedule:
        forSchedule(args.token, args.url)
    else:
        main()
