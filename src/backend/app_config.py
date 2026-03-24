"""Configuration et chemins de l'application ISENAPP.

Centralise toutes les constantes de configuration : port du serveur,
chemins des fichiers de données, répertoires du graphe, scopes OAuth
Google, et bootstrap des fichiers par défaut.

Dépendances internes :
    (aucune — module racine de configuration)

Dépendances externes :
    (aucune)
"""

import os
import shutil
from pathlib import Path

PORT = 8080
DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = str(Path(DIR).resolve().parents[1])
BUNDLED_DATA_DIR = os.path.join(PROJECT_ROOT, "data")
RENDERER_INDEX = os.path.join(PROJECT_ROOT, "src", "renderer", "index.html")


def get_app_data_dir():
    """Return a writable data directory for runtime files."""
    env_override = os.environ.get("ISENAPP_DATA_DIR", "").strip()
    if env_override:
        return env_override

    xdg_data_home = os.environ.get("XDG_DATA_HOME", "").strip()
    if xdg_data_home:
        return os.path.join(xdg_data_home, "isenapp")

    return os.path.join(str(Path.home()), ".local", "share", "isenapp")


APP_DATA_DIR = get_app_data_dir()
os.makedirs(APP_DATA_DIR, exist_ok=True)


def bootstrap_file(filename):
    """Copy bundled defaults to writable app data dir when missing."""
    src = os.path.join(BUNDLED_DATA_DIR, filename)
    if not os.path.isfile(src):
        src = os.path.join(DIR, filename)
    dst = os.path.join(APP_DATA_DIR, filename)
    if os.path.isfile(src) and not os.path.exists(dst):
        shutil.copy2(src, dst)
    return dst if os.path.exists(dst) else src


DATA = bootstrap_file("data.json")
CONTACTS_CSV = bootstrap_file("contacts_complets_v2.csv")
LOG_FILE = os.path.join(APP_DATA_DIR, "api_errors.log")
DOWNLOADS = str(Path.home() / "Téléchargements")

MAILS_DIR = str(Path.home() / "mails")
SEEN_UIDS_FILE = os.path.join(APP_DATA_DIR, "seen_uids.json")
ACCOUNTS_FILE = os.path.join(APP_DATA_DIR, "accounts.json")
INBOX_INDEX_FILE = os.path.join(APP_DATA_DIR, "inbox_index.json")

ISENAPP_DATA = str(Path.home() / "Documents" / "isenapp_mails")
GRAPH_MD_DIR = os.path.join(ISENAPP_DATA, "mails")
GRAPH_ATT_DIR = os.path.join(ISENAPP_DATA, "attachements")
GRAPH_VAULT = ISENAPP_DATA

GOOGLE_CALENDAR_SCOPE = "https://www.googleapis.com/auth/calendar"
GOOGLE_MAIL_SCOPE = "https://mail.google.com/"

os.makedirs(MAILS_DIR, exist_ok=True)
os.makedirs(GRAPH_MD_DIR, exist_ok=True)
os.makedirs(GRAPH_ATT_DIR, exist_ok=True)

if not os.path.isdir(DOWNLOADS):
    DOWNLOADS = str(Path.home() / "Downloads")
if not os.path.isdir(DOWNLOADS):
    DOWNLOADS = str(Path.home())
