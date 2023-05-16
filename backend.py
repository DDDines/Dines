#backend

import os.path
import requests
from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from threading import Thread


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1nSQih5Dbb4IRKUca1hnGeC_tD5m5_yhCFQFK80tXMuw'
SAMPLE_RANGE_NAME = 'Data!A:F'


def loadDataFromSheets():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentialsOAuth.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Salva as credenciais para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result['values']

        return values
    except HttpError as error:
        print(f"Ocorreu um erro: {error}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


def writeData(values, row_num, column=4):
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentialsOAuth.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Salva as credenciais para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Atualiza os valores na planilha
        data_range = f'Data!{chr(65 + column).upper()}{row_num + 1}:{chr(65 + column).upper()}{row_num + 1}'
        body = {
            'values': [values]
        }
        result = sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=data_range,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()

        print(
            f"Valores atualizados na planilha: {result.get('updatedCells')} células")

        return result

    except HttpError as error:
        print(f"Ocorreu um erro: {error}")
        return None
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None


def generate_title_and_description(query):
    search_url = f"https://www.bing.com/search?q={'+'.join(query)}"
    try:
        response = requests.get(search_url)
        soup = BeautifulSoup(response.text, "html.parser")
        search_results = soup.find_all('li', {"class": "b_algo"})

        if search_results:
            first_result = search_results[0]
            link = first_result.find('a')['href']
            title = first_result.find('a').text
            description = first_result.find(
                'p').text if first_result.find('p') else ''

            return title, description
        else:
            return '', ''
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return '', ''


def searchAndGenerate(query, item_info):
    query_with_extra_info = " ".join(item_info)
    query = query.strip().split(" ") + item_info

    title, description = "", ""
    search_url = f"https://www.bing.com/search?q={'+'.join(query)}"

    try:
        response = requests.get(search_url)
        soup = BeautifulSoup(response.text, "html.parser")
        search_results = soup.find_all('li', {"class": "b_algo"})

        if search_results:
            first_result = search_results[0]
            link = first_result.find('a')['href']
            title = first_result.find('a').text
            description = first_result.find(
                'p').text if first_result.find('p') else ''
    except Exception as e:
        print(f"Ocorreu um erro na pesquisa: {e}")

    return title, description


def processSearchData():
    values = loadDataFromSheets()
    if values is not None:
        for row_num, row_data in enumerate(values):
            item_info = row_data[:3]
            query = row_data[4]

            if all(item_info) and query:
                title, description = searchAndGenerate(query, item_info)

                if title and description:
                    values[row_num][3] = title
                    values[row_num][4] = description
                    writeData([" ".join([title, description])], row_num)

        print("Pesquisa e registro concluídos.")


def runSearchData():
    search_thread = Thread(target=searchData)
    search_thread.start()
