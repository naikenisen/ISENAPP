"""Gestion des comptes email (CRUD + normalisation).

Permet de charger, sauvegarder et rechercher les comptes email
configurés, avec normalisation des champs d'authentification
pour assurer la compatibilité ascendante.

Dépendances internes :
    - app_config  : chemin du fichier accounts.json
    - json_store  : lecture/écriture atomique JSON

Dépendances externes :
    (aucune)
"""

from app_config import ACCOUNTS_FILE
from json_store import atomic_write_json, read_json_with_backup


def normalize_auth_fields(account):
    """Normalize auth/provider fields for backward compatibility."""
    provider = (account.get("provider", "") or "").lower()
    auth_type = (account.get("auth_type", "") or "").lower()
    if provider == "gmail_oauth" and not auth_type:
        auth_type = "oauth2"
    if auth_type:
        account["auth_type"] = auth_type
    return account


def load_accounts():
    accounts = read_json_with_backup(ACCOUNTS_FILE, [])
    if not isinstance(accounts, list):
        return []
    for acc in accounts:
        normalize_auth_fields(acc)
    return accounts


def save_accounts(accounts):
    atomic_write_json(ACCOUNTS_FILE, accounts)


def find_account_index_by_email(accounts, email_addr):
    target = (email_addr or "").strip().lower()
    for idx, acc in enumerate(accounts):
        if (acc.get("email", "") or "").strip().lower() == target:
            return idx
    return -1


def find_account_by_email(email_addr):
    """Find account config matching a sender email."""
    target = (email_addr or "").lower()
    for acc in load_accounts():
        if (acc.get("email", "") or "").lower() == target:
            return acc
    return None
