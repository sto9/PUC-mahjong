"""設定ファイル"""
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent

SPREADSHEET_ID = "1whqVie4Ql2R2NqtNZKGOB1i1rJVF9jnDXhzzYlQJqAs"

CREDENTIAL_FILE = BASE_DIR / "credential" / "google_service_account.json"
MEMBERS_FILE = BASE_DIR / "credential" / "members.json"
FANS_FILE = BASE_DIR / "constants" / "fans.json"

PAIFU_DIR = BASE_DIR / "paifu"

DORA_FANS = [31, 32, 33, 34]
RARE_FANS = [-1, 3, 4, 5, 6, 18, 19, 20, 24, 28]
ORIGIN_POINT_4 = 25000
ORIGIN_POINT_3 = 35000

def load_fans():
    """役の名前をロード"""
    with open(FANS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_members():
    """メンバー情報をロード"""
    with open(MEMBERS_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)
        return data['members']