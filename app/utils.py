# app/utils.py
from pathlib import Path
from typing import Dict, Any, Optional
import json
import re
import shutil
import sys
import os


def get_base_path():
    """Get the base path for the application (supports .exe and script mode)"""
    if getattr(sys, 'frozen', False):
        # Running as .exe - use the temp folder where PyInstaller extracts files
        return sys._MEIPASS
    else:
        # Running as script - use the current directory
        return os.path.dirname(os.path.abspath(__file__))


# Get base path
BASE_PATH = get_base_path()

# Data directory - try multiple paths for .exe support
DATA_DIR = Path("data")
if not DATA_DIR.exists():
    # Try with base path for .exe
    DATA_DIR = Path(BASE_PATH) / "data"
    if not DATA_DIR.exists():
        # Try one level up (if running from different location)
        DATA_DIR = Path(os.path.dirname(BASE_PATH)) / "data"

CLIENTS_DIR = DATA_DIR / "clients"
CLIENTS_DIR.mkdir(parents=True, exist_ok=True)


def get_client_folder_name(client_name: str) -> str:
    """
    Generate a folder name from client name with number if duplicate exists.
    Example: "Jan Jansen" -> "Jan_Jansen", if exists -> "1_Jan_Jansen"
    """
    clean_name = re.sub(r'[^\w\s-]', '', client_name)
    clean_name = re.sub(r'[-\s]+', '_', clean_name)
    clean_name = clean_name.strip('_')

    base_name = clean_name
    counter = 1

    while True:
        folder_path = CLIENTS_DIR / base_name
        if not folder_path.exists():
            return base_name
        json_file = folder_path / "client_data.json"
        if json_file.exists():
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get('name') == client_name:
                        return base_name
            except:
                pass
        base_name = f"{counter}_{clean_name}"
        counter += 1


def get_all_clients() -> list:
    """Get all clients from the clients directory"""
    clients = []
    for client_folder in CLIENTS_DIR.iterdir():
        if client_folder.is_dir():
            json_file = client_folder / "client_data.json"
            if json_file.exists():
                try:
                    with open(json_file, 'r', encoding='utf-8') as f:
                        client_data = json.load(f)
                        clients.append(client_data)
                except Exception as e:
                    print(f"Error loading client {json_file}: {e}")
    clients.sort(key=lambda x: x.get('name', '').lower())
    return clients


def get_client(client_id: str) -> Optional[Dict[str, Any]]:
    """Get a specific client by ID"""
    for client_folder in CLIENTS_DIR.iterdir():
        if client_folder.is_dir():
            json_file = client_folder / "client_data.json"
            if json_file.exists():
                try:
                    with open(json_file, 'r', encoding='utf-8') as f:
                        client_data = json.load(f)
                        if client_data.get('id') == client_id:
                            return client_data
                except Exception:
                    continue
    return None


def save_client(client_data: Dict[str, Any]) -> bool:
    """Save client data to a JSON file in the client's folder"""
    try:
        client_id = client_data.get('id')
        if not client_id:
            return False

        folder_name = client_data.get('folder_name')
        if not folder_name:
            folder_name = get_client_folder_name(client_data.get('name', 'unknown'))
            client_data['folder_name'] = folder_name

        folder_path = CLIENTS_DIR / folder_name
        folder_path.mkdir(parents=True, exist_ok=True)

        json_file = folder_path / "client_data.json"

        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(client_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error saving client: {e}")
        return False


def get_client_excel_path(client_id: str) -> Path:
    """Get the path to the client's Excel file"""
    client = get_client(client_id)
    if client:
        folder_name = client.get('folder_name', client_id)
        return CLIENTS_DIR / folder_name / "loopbaan_onderzoek.xlsx"
    return CLIENTS_DIR / client_id / "loopbaan_onderzoek.xlsx"


def get_client_word_path(client_id: str) -> Path:
    """Get the path to the client's Word document"""
    client = get_client(client_id)
    if client:
        folder_name = client.get('folder_name', client_id)
        return CLIENTS_DIR / folder_name / "loopbaan_rapport.docx"
    return CLIENTS_DIR / client_id / "loopbaan_rapport.docx"


def delete_client(client_id: str) -> bool:
    """Delete a client and all their data"""
    client = get_client(client_id)
    if client:
        folder_name = client.get('folder_name', client_id)
        client_folder = CLIENTS_DIR / folder_name
        if client_folder.exists():
            try:
                shutil.rmtree(client_folder)
                return True
            except Exception as e:
                print(f"Error deleting client: {e}")
                return False
    return False