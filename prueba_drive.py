import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/drive.file']

def conectar_drive():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('drive', 'v3', credentials=creds)

        metadatos_carpeta = {
            'name': 'ERP_Licitaciones_Maestra',
            'mimeType': 'application/vnd.google-apps.folder'
        }
        
        print("Conectando con Google Drive...")
        carpeta = service.files().create(body=metadatos_carpeta, fields='id').execute()
        print(f'¡Éxito absoluto! La carpeta maestra se creó en tu Drive con el ID: {carpeta.get("id")}')

    except Exception as error:
        print(f'Ha ocurrido un error en la Matrix: {error}')

if __name__ == '__main__':
    conectar_drive()