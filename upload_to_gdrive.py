import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Scopes requis pour lire/écrire sur Drive et gérer les feuilles de calcul
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

SOURCE_DIR = "/Users/koffiefabienmichael/Downloads/FAAB APPLI/FAAB PROJET/KIT CHEF PROJET"

def authenticate():
    """Initialise l'authentification OAuth2 avec le compte Google."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if not os.path.exists('credentials.json'):
            print("❌ ERREUR: Le fichier 'credentials.json' est introuvable.")
            print("Veuillez générer ce fichier depuis la Google Cloud Console et le placer ici.")
            exit(1)
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        print("\n⏳ Le serveur local est lancé, veuillez valider sur la page web qui s'ouvre...")
        creds = flow.run_local_server(port=8080)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def get_or_create_folder(drive_service, folder_name, parent_id=None):
    """Vérifie si un dossier existe, sinon le crée."""
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
        
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if not items:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_id:
            file_metadata['parents'] = [parent_id]
            
        file = drive_service.files().create(body=file_metadata, fields='id').execute()
        return file.get('id')
    
    return items[0]['id']

def upload_excel_to_sheets(drive_service, file_path, file_name, folder_id):
    """Téléverse un fichier Excel et le convertit automatiquement en Google Sheets."""
    file_metadata = {
        'name': file_name.replace('.xlsm', '').replace('.xlsb', '').replace('.xlsx', ''),
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [folder_id]
    }
    
    media = MediaFileUpload(file_path, mimetype='application/vnd.ms-excel.sheet.macroEnabled.12', resumable=True)
    
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"✅ Converti : {file_name} -> ID: {file.get('id')}")
    return file.get('id')

def main():
    print("🔑 Authentification au compte Google Workspace en cours...")
    drive_service = authenticate()
    
    print("\n📂 Création du dossier maître 'KIT CHEF PROJET_SHEETS'")
    root_folder_id = get_or_create_folder(drive_service, 'KIT CHEF PROJET_SHEETS')
    print(f"ID Dossier Racine: {root_folder_id}")

    folders_cache = {}

    for root, dirs, files in os.walk(SOURCE_DIR):
        relative_path = os.path.relpath(root, SOURCE_DIR)
        
        current_parent_id = root_folder_id
        if relative_path != '.':
            path_parts = relative_path.split(os.sep)
            for part in path_parts:
                cache_key = f"{current_parent_id}_{part}"
                if cache_key not in folders_cache:
                    new_id = get_or_create_folder(drive_service, part, current_parent_id)
                    folders_cache[cache_key] = new_id
                current_parent_id = folders_cache[cache_key]

        for file in files:
            if file.endswith(('.xlsm', '.xlsb', '.xlsx')) and not file.startswith('~$'):
                print(f"Téléversement en cours: {file}...")
                file_path = os.path.join(root, file)
                try:
                    upload_excel_to_sheets(drive_service, file_path, file, current_parent_id)
                except Exception as e:
                    print(f"❌ Erreur sur {file}: {e}")

    print("\n🎉 Migration d'infrastructure terminée avec succès !")

if __name__ == '__main__':
    main()
