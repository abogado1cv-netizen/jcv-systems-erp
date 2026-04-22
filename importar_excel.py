import os
import django

# 1. Encendemos el motor de Django en modo "invisible"
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'erp_core.settings')
django.setup()

from licitaciones.models import CatalogoMedicamento
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import io

# Pedimos permiso total de lectura/escritura para Drive
SCOPES = ['https://www.googleapis.com/auth/drive']

def importar_catalogo():
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

    print("Conectando con el Ordenador Maestro (Drive)...")
    service = build('drive', 'v3', credentials=creds)

    print("Buscando tu archivo 'catalogo_maestro.xlsx'...")
    resultados = service.files().list(
        q="name='catalogo_maestro.xlsx'",
        spaces='drive',
        fields='files(id, name)'
    ).execute()
    
    archivos = resultados.get('files', [])
    
    if not archivos:
        print("❌ No se encontró el archivo. Revisa el nombre en tu Drive.")
        return

    archivo_id = archivos[0]['id']
    print("✅ Archivo encontrado. Descargando datos de forma segura...")

    # Descargando el archivo a la memoria temporal del programa
    request = service.files().get_media(fileId=archivo_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    fh.seek(0)
    print("✅ Descarga completada. Leyendo datos con Pandas...")
    
    # Leemos el Excel
    df = pd.read_excel(fh)

    print("Limpiando catálogo anterior para evitar duplicados...")
    CatalogoMedicamento.objects.all().delete()

    print("Guardando medicamentos en la base de datos...")
    contador = 0
    for index, row in df.iterrows():
        try:
            CatalogoMedicamento.objects.create(
                clave_sector=str(row['Clave_Sector']),
                descripcion=str(row['Descripcion']),
                denominacion_generica=str(row['Denominacion_Generica']),
                denominacion_distintiva=str(row['Denominacion_Distintiva']),
                fabricante=str(row['Fabricante']),
                rfc_fabricante=str(row['RFC_Fabricante']),
                pais_fabricacion=str(row['Pais_Fabricacion']),
                num_registro_sanitario=str(row['Registro_Sanitario']),
                num_prorroga=str(row.get('Prorroga', 'NO APLICA')),
                codigo_barras=str(row['Codigo_Barras']),
                fecha_expedicion=pd.to_datetime(row['Fecha_Expedicion']).date(),
                fecha_vigencia=pd.to_datetime(row['Fecha_Vigencia']).date()
            )
            contador += 1
        except Exception as e:
            print(f"Error en la fila {index + 2} de Excel: {e}")

    print(f"🚀 ¡Éxito absoluto! Se importaron {contador} medicamentos a tu ERP.")

if __name__ == '__main__':
    importar_catalogo()
