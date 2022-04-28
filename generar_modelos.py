from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive.appdata','https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID of a sample document.
#DOCUMENT_ID = '1aD7EO_FKZA7m1ZCvl3i7g828CzkL9h3k1iCtnHkX5g8'
MODELOS_FOLDER = '1rOu-f9lXmloeMCd_OJuDaR5r5L3nF7Iy';
ALUMNOS_FOLDER = '1-x0KBIR9a_EMZq1IzdwppJLGhwCxy8-v';
GRADOS_FOLDER = '1EcJ-vTiuhXidEP35Ry9y0TkzOh6sWs18';
PERIODO = "Marzo - Abril"

def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        docs = build('docs', 'v1', credentials=creds)
        drive = build('drive', 'v3', credentials=creds)
        sheet = build('sheets', 'v4', credentials=creds)

        # Retrieve the documents contents from the Docs service.
        #document = docs.documents().get(documentId=DOCUMENT_ID).execute()
        #print('The title of the document is: {}'.format(document.get('title')))

        # TRAER MODELOS DE DOCUMENTO
        results = drive.files().list(
            q="'"+MODELOS_FOLDER+"' in parents and mimeType contains 'document' and trashed = False",
            pageSize=100, fields="nextPageToken, files(id, name, parents)").execute()
        modelos = results.get('files', [])
        print(modelos)

        # TRAER PLANILLAS DE ALUMNOS
        results = drive.files().list(
            q="'"+ALUMNOS_FOLDER+"' in parents and mimeType contains 'spreadsheet' and trashed = False",
            pageSize=100, fields="nextPageToken, files(id, name)").execute()
        sheets = results.get('files', [])
        print(sheets)

        # PARA CADA PLANILLA
        for s in sheets:
            print(s['name'])
            grado = s['name']
            # LEER CELDAS
            result = sheet.spreadsheets().values().get(spreadsheetId=s['id'], range="A1:A100").execute()
            rows = result.get('values', [])
            # CREAR CARPETA DEL GRADO SI NO EXISTE
            results = drive.files().list(
                q="name = '"+grado+"' and '" + GRADOS_FOLDER + "' in parents and mimeType contains 'folder' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name)").execute()
            folders = results.get('files', [])
            if len(folders) == 0:
                file_metadata = {
                    'name': grado,
                    'parents': [GRADOS_FOLDER],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = drive.files().create(body=file_metadata,fields='id').execute()
                gradoFolderId = folder.get('id');
                print("Carpeta creada: "+grado )
            else:
                gradoFolderId = folders[0].get('id')

            # CREAR CARPETA PERIODO SI NO EXISTE
            results = drive.files().list(
                q="name = '"+PERIODO+"' and '" + gradoFolderId + "' in parents and mimeType contains 'folder' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name)").execute()
            folders = results.get('files', [])
            if len(folders) == 0:
                file_metadata = {
                    'name': PERIODO,
                    'parents': [gradoFolderId],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = drive.files().create(body=file_metadata,fields='id').execute()
                periodoFolderId = folder.get('id');
                print("Carpeta creada: "+PERIODO )
            else:
                periodoFolderId = folders[0].get('id')

            # PARA CADA ALUMNO
            for a in rows:
                alumno = a[0]
                print(alumno)
                for m in modelos:
                    if m['name'] == grado:
                        # CREAR LA COPIA
                        body = {
                            "name": alumno,
                            "parents": [periodoFolderId]
                        }
                        copia = drive.files().copy(fileId=m['id'], body=body, fields='id').execute();
                        copiaId = copia.get('id')
                        # ESCRIBIR EL ALUMNO DENTRO DEL ARCHIVO
                        requests = [
                            {
                                'replaceAllText': {
                                    'containsText': {
                                        'text': '{{alumno}}',
                                        'matchCase': 'true'
                                    },
                                    'replaceText': alumno
                                }
                            }
                        ]
                        docs.documents().batchUpdate(documentId=copiaId, body={'requests': requests}).execute()
            # FIN ALUMNO
        # FIN PLANILLA GRADO

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()