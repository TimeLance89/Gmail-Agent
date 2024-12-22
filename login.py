import imaplib
import json
import os

CONFIG_FILE = "credentials.json"

def save_credentials(email_addr, password):
    """Speichert Anmeldedaten."""
    data = {"email": email_addr, "password": password}
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"Fehler beim Speichern von {CONFIG_FILE}: {e}")

def load_credentials():
    """Lädt Anmeldedaten."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
                return data["email"], data["password"]
            except (json.JSONDecodeError, KeyError) as e:
                print(f"Fehler beim Laden von {CONFIG_FILE}: {e}")
                return None, None
    return None, None

def test_imap_login(email_addr, password):
    """Testet IMAP-Login."""
    try:
        conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        conn.login(email_addr, password)
        conn.logout()
        return True
    except imaplib.IMAP4.error as e:
        print(f"IMAP-Login-Fehler: {e}")
        return False
    except Exception as e:
        print(f"Allgemeiner IMAP-Fehler: {e}")
        return False
