from __future__ import print_function

import sys

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import credentials

MODELOS_FOLDER = '1rOu-f9lXmloeMCd_OJuDaR5r5L3nF7Iy';
ALUMNOS_FOLDER = '1-x0KBIR9a_EMZq1IzdwppJLGhwCxy8-v';
GRADOS_FOLDER = '1EcJ-vTiuhXidEP35Ry9y0TkzOh6sWs18';
PERIODO = "Marzo - Abril"

def main():

    creds = credentials.credentials()



    try:
        docs = build('docs', 'v1', credentials=creds)
        drive = build('drive', 'v3', credentials=creds)
        sheet = build('sheets', 'v4', credentials=creds)

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
            grado = s['name']
            if len(sys.argv) > 1:
                if not (grado in sys.argv):
                    continue;
            print(s['name'])

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

            # TRAER LAS COPIAS DE LA CARPETA
            results = drive.files().list(
                q="'" + periodoFolderId + "' in parents and mimeType contains 'document' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name, parents)").execute()
            cs = results.get('files', [])
            copias = []
            for c in cs:
                copias.append(c['name'])

            # PARA CADA ALUMNO
            for a in rows:
                alumno = a[0]
                print(alumno)
                # SI YA EXISTE SALTEAR
                if(alumno in copias):
                    continue;
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