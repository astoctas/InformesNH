from __future__ import print_function

import sys

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaInMemoryUpload
import credentials

INFORMES_FOLDER = '1UyBR1FxWw6_6yoZQ-s5f656hVFHI4Dkh';
ALUMNOS_FOLDER = '1-x0KBIR9a_EMZq1IzdwppJLGhwCxy8-v';
GRADOS_FOLDER = '1EcJ-vTiuhXidEP35Ry9y0TkzOh6sWs18';
PERIODO = "Marzo - Abril"

def main():

    creds = credentials.credentials()



    try:
        docs = build('docs', 'v1', credentials=creds)
        drive = build('drive', 'v3', credentials=creds)

        # TRAER GRADOS
        results = drive.files().list(
            q="'" + GRADOS_FOLDER + "' in parents and mimeType contains 'folder' and trashed = False",
            pageSize=100, fields="nextPageToken, files(id, name)").execute()
        grados = results.get('files', [])

        # PARA CADA GRADO
        for s in grados:
            grado = s['name']
            if len(sys.argv) > 1:
                if not (grado in sys.argv):
                    continue;
            print(grado)

            gradoFolderId = s.get('id')

            # CARPETA PERIODO
            results = drive.files().list(
                q="name = '"+PERIODO+"' and '" + gradoFolderId + "' in parents and mimeType contains 'folder' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name)").execute()
            folders = results.get('files', [])
            periodoFolderId = folders[0].get('id')

            # TRAER LOS INFORMES DE LA CARPETA
            results = drive.files().list(
                q="'" + periodoFolderId + "' in parents and mimeType contains 'document' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name, parents)").execute()
            informes = results.get('files', [])

            # CREAR CARPETA GRADO SI NO EXISTE
            # CREAR CARPETA DEL GRADO SI NO EXISTE
            results = drive.files().list(
                q="name = '"+grado+"' and '" + INFORMES_FOLDER + "' in parents and mimeType contains 'folder' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name)").execute()
            folders = results.get('files', [])
            if len(folders) == 0:
                file_metadata = {
                    'name': grado,
                    'parents': [INFORMES_FOLDER],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = drive.files().create(body=file_metadata,fields='id').execute()
                gradoInformesFolderId = folder.get('id');
                print("Carpeta creada: "+grado )
            else:
                gradoInformesFolderId = folders[0].get('id')

            # CREAR CARPETA PERIODO SI NO EXISTE
            results = drive.files().list(
                q="name = '"+PERIODO+"' and '" + gradoInformesFolderId + "' in parents and mimeType contains 'folder' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name)").execute()
            folders = results.get('files', [])
            if len(folders) == 0:
                file_metadata = {
                    'name': PERIODO,
                    'parents': [gradoInformesFolderId],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = drive.files().create(body=file_metadata,fields='id').execute()
                periodoInformesFolderId = folder.get('id');
                print("Carpeta creada: "+PERIODO )
            else:
                periodoInformesFolderId = folders[0].get('id')

            # TRAER LOS PDF DE LA CARPETA
            results = drive.files().list(
                q="'" + periodoInformesFolderId + "' in parents and mimeType contains 'pdf' and trashed = False",
                pageSize=100, fields="nextPageToken, files(id, name, parents)").execute()
            cs = results.get('files', [])
            pdfs = []
            for c in cs:
                pdfs.append(c['name'])

            # PARA CADA INFORME
            for i in informes:
                alumno = i.get('name');
                print(alumno)
                # CREAR PDF
                if(alumno in pdfs):
                    continue;
                filebody = drive.files().export(fileId=i.get('id'),mimeType="application/pdf").execute();
                media_body = MediaInMemoryUpload(filebody, "application/pdf")
                body = {
                    'name': alumno,
                    'parents': [periodoInformesFolderId],
                    'mimeType': "application/pdf"
                }
                file = drive.files().create(body=body,media_body=media_body).execute()

                # FIN ALUMNO

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()