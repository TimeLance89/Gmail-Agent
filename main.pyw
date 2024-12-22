import imaplib
import email
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime, timedelta
import uuid
from threading import Thread
import sys
from login import test_imap_login, save_credentials, load_credentials

from tkcalendar import Calendar

import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import openai
import pyttsx3  # Für Text-to-Speech

import pystray
from PIL import Image, ImageDraw

from translator import Translator
import re

CONFIG_FILE = "credentials.json"
ANSWERED_FILE = "answered.json"
TERMINE_FILE = "termine.json"
BLOCKED_TIMES_FILE = "blocked_times.json"
SETTINGS_FILE = "settings.json"

GPT_MODEL_NAME = "gpt-4"


# ---------------------- Signatur für KI-AGENT ----------------------
def add_signature(message, user_name, translator):
    footer = translator.gettext("automatic_reply_footer").format(user_name=user_name)
    return f"{message}\n\n---\n{footer}"


# ---------------------- Text-to-Speech ----------------------
def speak(text):
    try:
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Text-to-Speech Fehler: {e}")


# ---------------------- JSON-Helferfunktionen ----------------------
def load_answered_ids():
    if not os.path.exists(ANSWERED_FILE):
        return set()
    with open(ANSWERED_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            return set(data)
        except json.JSONDecodeError as e:
            print(f"Fehler beim Laden von {ANSWERED_FILE}: {e}")
            return set()

def save_answered_ids(answered_ids_set):
    try:
        with open(ANSWERED_FILE, "w", encoding="utf-8") as f:
            json.dump(list(answered_ids_set), f, indent=2)
    except Exception as e:
        print(f"Fehler beim Speichern von {ANSWERED_FILE}: {e}")

def load_termine():
    if not os.path.exists(TERMINE_FILE):
        return []
    with open(TERMINE_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            return data  # Liste von Dicts
        except json.JSONDecodeError as e:
            print(f"Fehler beim Laden von {TERMINE_FILE}: {e}")
            return []

def save_termine(termine_list):
    try:
        with open(TERMINE_FILE, "w", encoding="utf-8") as f:
            json.dump(termine_list, f, indent=2)
    except Exception as e:
        print(f"Fehler beim Speichern von {TERMINE_FILE}: {e}")

def load_blocked_times():
    if not os.path.exists(BLOCKED_TIMES_FILE):
        return []
    with open(BLOCKED_TIMES_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            return data  # Liste von Dicts
        except json.JSONDecodeError as e:
            print(f"Fehler beim Laden von {BLOCKED_TIMES_FILE}: {e}")
            return []

def save_blocked_times(blocked_times_list):
    try:
        with open(BLOCKED_TIMES_FILE, "w", encoding="utf-8") as f:
            json.dump(blocked_times_list, f, indent=2)
    except Exception as e:
        print(f"Fehler beim Speichern von {BLOCKED_TIMES_FILE}: {e}")

def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {"user_name": "Benutzer", "auftraggeber_email": ""}
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            if "user_name" not in data:
                data["user_name"] = "Benutzer"
            if "auftraggeber_email" not in data:
                data["auftraggeber_email"] = ""
            return data
        except json.JSONDecodeError as e:
            print(f"Fehler beim Laden von {SETTINGS_FILE}: {e}")
            return {"user_name": "Benutzer", "auftraggeber_email": ""}

def save_settings(settings_dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings_dict, f, indent=2)
    except Exception as e:
        print(f"Fehler beim Speichern von {SETTINGS_FILE}: {e}")


# ---------------------- Hilfsfunktionen ----------------------
def parse_email_address(from_field):
    if "<" in from_field and ">" in from_field:
        start = from_field.find("<") + 1
        end = from_field.find(">")
        return from_field[start:end].strip()
    else:
        parts = from_field.split()
        for p in parts:
            if "@" in p:
                return p.strip("<>")
        return from_field

def parse_email_date(date_str):
    try:
        parsed_date = email.utils.parsedate_to_datetime(date_str)
        return parsed_date.strftime("%Y-%m-%d %H:%M")
    except Exception as e:
        print(f"Fehler beim Parsen des E-Mail-Datums: {e}")
        return date_str

def construct_conversation_history(thread_emails):
    conversation = ""
    for email_item in sorted(thread_emails, key=lambda x: parse_email_date(x["date"])):
        sender = parse_email_address(email_item["from"])
        date = parse_email_date(email_item["date"])
        subject = email_item["subject"]
        body = email_item["body"]
        conversation += f"Von: {sender}\nDatum: {date}\nBetreff: {subject}\nNachricht:\n{body}\n\n"
    return conversation


# ---------------------- KI-Funktionen ----------------------
def generate_ai_reply(conversation_history, context, openai_api_key, translator, system_message=None):
    if not openai_api_key:
        return translator.gettext("error_no_api_key")

    try:
        openai.api_key = openai_api_key
        if system_message is None:
            system_message = (
                "Du bist ein vielseitiger Assistent, der Aufgaben effizient und präzise bearbeitet. "
                "Deine Aufgabe ist es, die Anweisungen des Nutzers vollständig zu verstehen und eine Antwort zu generieren, "
                "die genau auf die spezifische Anfrage eingeht. Formuliere professionell, höflich und eindeutig."
            )

        response = openai.ChatCompletion.create(
            model=GPT_MODEL_NAME,
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": conversation_history}
            ],
            max_tokens=500,
            temperature=0.3
        )
        return response.choices[0].message["content"].strip()
    except Exception as e:
        print(f"Fehler beim Generieren der KI-Antwort: {e}")
        return f"({translator.gettext('error_ai_reply')}: {str(e)})"


def generate_conflict_reply(subject, date, time, user_name, openai_api_key, translator):
    settings = load_settings()
    user_name = settings.get("user_name", "Benutzer")
    if not openai_api_key:
        return translator.gettext("error_no_api_key")
    try:
        openai.api_key = openai_api_key
        system_msg = (
            f"Du bist ein KI-Assistent für die Terminplanung im Namen von {user_name}. "
            "Deine Aufgabe ist es, höflich zu kommunizieren, dass ein gewünschter Termin nicht verfügbar ist, "
            "und mögliche Alternativen anzubieten. Stelle klar, dass du ein automatisierter Assistent bist."
        )
        user_msg = (
            f"Betreff: {subject}\n"
            f"Nachricht: Der gewünschte Termin am {date} um {time} ist leider nicht verfügbar. "
            "Bitte schlagen Sie einen alternativen Termin vor."
        )

        response = openai.ChatCompletion.create(
            model=GPT_MODEL_NAME,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg}
            ],
            max_tokens=150,
            temperature=0.7
        )
        ai_reply = response.choices[0].message["content"].strip()
        ai_reply = add_signature(ai_reply, user_name, translator)
        return ai_reply
    except Exception as e:
        print(f"Fehler beim Generieren der Konflikt-Antwort: {e}")
        return f"({translator.gettext('error_ai_reply')}: {str(e)})"


# -- NEUE VERSION: generate_confirmation_reply nimmt 2 Parameter (recipient_name, butler_name) --
def generate_confirmation_reply(subject, date, time, summary, recipient_name, butler_name, translator):
    """
    Generiert eine Bestätigungsantwort für einen erfolgreich eingetragenen Termin,
    wobei 'recipient_name' die Person ist, die uns gemailt hat,
    und 'butler_name' der Name unseres KI-Assistenten.
    """
    try:
        confirmation_text = (
            f"Sehr geehrte/r {recipient_name},\n\n"
            f"vielen Dank für Ihre Terminanfrage. Ihr Termin am {date} um {time} wurde erfolgreich in unserem System eingetragen.\n"
            f"Worum es in diesem Termin geht: {summary}\n\n"
            f"Mit freundlichen Grüßen,\n"
            f"{butler_name}"
        )
        confirmation_text = add_signature(confirmation_text, butler_name, translator)
        return confirmation_text
    except Exception as e:
        print(f"Fehler beim Generieren der Bestätigungsantwort: {e}")
        return f"({translator.gettext('error_ai_reply')}: {str(e)})"


def parse_appointment_request(from_addr, subject, body, existing_appointments, openai_api_key, translator):
    settings = load_settings()
    user_name = settings.get("user_name", "Benutzer")
    try:
        system_msg = (
            f"Du bist ein automatisierter KI-Kalender-Assistent, der im Auftrag von {user_name} arbeitet. "
            "Deine Aufgabe ist es, Termine aus E-Mails präzise zu extrahieren und bestehende Termine zu verwalten. "
            "Du darfst keine Annahmen machen, die über die Angaben in der E-Mail hinausgehen. "
            "Falls die E-Mail einen Termin enthält oder ändert, extrahiere: Datum, Uhrzeit, Betreff, Zusammenfassung. "
            "Nutze 'identifier' aus E-Mail-Adresse und Betreff.\n"
            "Gib NUR ein JSON im Format:\n"
            '{"date": "YYYY-MM-DD", "time": "HH:MM", "betreff": "Betreff", "identifier": "sender@example.com|Betreff", "summary": "Zusammenfassung"}\n'
            "Falls kein Termin, gib {\"date\": null} zurück.\n"
        )
        user_msg = f"Absender: {from_addr}\nBetreff: {subject}\nText:\n{body}\n\nNur JSON ausgeben!"

        if not openai_api_key:
            raise ValueError(translator.gettext("error_no_api_key"))
        openai.api_key = openai_api_key

        response = openai.ChatCompletion.create(
            model=GPT_MODEL_NAME,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg}
            ],
            max_tokens=250,
            temperature=0.0
        )
        raw_json = response.choices[0].message["content"].strip()

        print(f"ChatGPT Response for Appointment Parsing: {raw_json}")
        data = json.loads(raw_json)
        if "date" not in data:
            data["date"] = None
        if data["date"] is not None:
            for appt in existing_appointments:
                if appt["from_addr"] == from_addr and appt["betreff"].lower() == data["betreff"].lower():
                    data["identifier"] = appt["identifier"]
                    break
            else:
                data["identifier"] = f"{from_addr}|{subject}"

        if "time" not in data:
            data["time"] = None
        elif data["time"] == "":
            data["time"] = None

        return data
    except json.JSONDecodeError as e:
        print(f"JSON-Parsing-Fehler: {e}")
        return {"date": None}
    except Exception as e:
        print(f"Fehler beim Parsen des Termins: {e}")
        return {"date": None}


# ---------------------- Thread-Fetch ----------------------
def fetch_email_thread(conn, message_id):
    thread_emails = []
    try:
        search_criteria = f'(OR (HEADER In-Reply-To "{message_id}") (HEADER References "{message_id}"))'
        status, data = conn.search(None, search_criteria)
        if status != "OK":
            print(f"IMAP-Suche für Thread fehlgeschlagen: {status}")
            return thread_emails
        
        mail_ids = data[0].split()
        for mid in mail_ids:
            status, msg_data = conn.fetch(mid, "(RFC822)")
            if status != "OK":
                print(f"Fehler beim Abrufen der E-Mail-ID {mid}: {status}")
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            from_addr = msg.get("From", "")
            date_str = msg.get("Date", "")
            body_text = extract_email_body(msg)
            thread_message_id = msg.get("Message-ID", "")

            thread_emails.append({
                "message_id": thread_message_id,
                "subject": subject,
                "from": from_addr,
                "date": date_str,
                "body": body_text
            })
    except Exception as ex:
        print(f"IMAP-Fehler beim Thread-Fetch: {ex}")
    
    return thread_emails


# ---------------------- Termin-Konfliktprüfung ----------------------
def is_time_conflict(new_date, new_start, new_end, termine_list, blocked_times_list):
    try:
        new_start_dt = datetime.strptime(f"{new_date} {new_start}", "%Y-%m-%d %H:%M")
        new_end_dt = datetime.strptime(f"{new_date} {new_end}", "%Y-%m-%d %H:%M")
    except ValueError as e:
        print(f"Fehler beim Parsen der neuen Terminzeiten: {e}")
        return True

    for termin in termine_list:
        if termin["date"] != new_date:
            continue
        try:
            existing_start = datetime.strptime(f"{termin['date']} {termin['time']}", "%Y-%m-%d %H:%M")
            existing_end = existing_start + timedelta(hours=1)
            if (new_start_dt < existing_end) and (new_end_dt > existing_start):
                print(f"Konflikt mit bestehendem Termin: {termin}")
                return True
        except ValueError as e:
            print(f"Fehler beim Parsen Terminzeiten: {e}")
            continue

    for block in blocked_times_list:
        if block["date"] != new_date:
            continue
        try:
            block_start = datetime.strptime(f"{block['date']} {block['start_time']}", "%Y-%m-%d %H:%M")
            block_end = datetime.strptime(f"{block['date']} {block['end_time']}", "%Y-%m-%d %H:%M")
            if (new_start_dt < block_end) and (new_end_dt > block_start):
                print(f"Konflikt mit blockierter Zeit: {block}")
                return True
        except ValueError as e:
            print(f"Fehler beim Parsen blockierter Zeiten: {e}")
            continue

    return False


# ---------------------- IMAP-Funktionen ----------------------
def list_emails_imap(email_addr, password, max_results=10):
    emails = []
    try:
        conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        conn.login(email_addr, password)
        conn.select("INBOX", readonly=True)

        status, data = conn.search(None, 'ALL')
        if status != "OK":
            print(f"IMAP-Suche fehlgeschlagen: {status}")
            return emails

        mail_ids = data[0].split()
        mail_ids = mail_ids[-max_results:]

        for mid in reversed(mail_ids):
            status, msg_data = conn.fetch(mid, "(RFC822)")
            if status != "OK":
                print(f"Fehler beim Abrufen E-Mail-ID {mid}: {status}")
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            from_addr = msg.get("From", "")
            to_addr = msg.get("To", "")
            date_str = msg.get("Date", "")
            body_text = extract_email_body(msg)
            message_id = msg.get("Message-ID", "")
            in_reply_to = msg.get("In-Reply-To", "")
            references = msg.get("References", "")

            thread_emails = []
            if in_reply_to:
                thread_emails = fetch_email_thread(conn, in_reply_to)
            elif references:
                first_ref = references.split()[0]
                thread_emails = fetch_email_thread(conn, first_ref)

            full_thread = thread_emails + [{
                "message_id": message_id,
                "subject": subject,
                "from": from_addr,
                "date": date_str,
                "body": body_text
            }]

            emails.append({
                "message_id": message_id,
                "subject": subject,
                "from": from_addr,
                "to": to_addr,
                "date": date_str,
                "body": body_text,
                "thread": full_thread
            })

        conn.close()
        conn.logout()
    except Exception as ex:
        print(f"IMAP-Fehler: {ex}")
    return emails


def extract_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get("Content-Disposition"))
            if ctype == "text/plain" and "attachment" not in cdisp:
                try:
                    return part.get_payload(decode=True).decode('utf-8', errors='replace')
                except Exception as e:
                    print(f"Fehler beim Dekodieren: {e}")
                    return "(Fehler beim Dekodieren des Textes)"
        return "(Kein Plain-Text gefunden.)"
    else:
        try:
            return msg.get_payload(decode=True).decode('utf-8', errors='replace')
        except Exception as e:
            print(f"Fehler beim Dekodieren: {e}")
            return "(Fehler beim Dekodieren des Textes)"


# ---------------------- SMTP-Funktionen ----------------------
def send_email_smtp(sender_email, password, to_email, subject, body_text, translator):
    port = 465
    context = ssl.create_default_context()

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body_text, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, to_email, msg.as_string())
        print(f"E-Mail erfolgreich gesendet an {to_email}: {subject}")
        return True, None
    except Exception as e:
        print(f"Fehler beim Senden der E-Mail an {to_email}: {e}")
        return False, str(e)


# --------------------------- GUI-Klassen ---------------------------
class LoginWindow(tk.Toplevel):
    def __init__(self, master=None, translator=None):
        super().__init__(master)
        self.translator = translator or Translator()
        self.title(self.translator.gettext("login_title"))
        self.geometry("400x300")
        self.bind("<Escape>", self.exit_fullscreen)
        self.transient(master)

        # Versuche, das Fenster direkt nach Erstellen in den Vordergrund zu holen
        self.lift()
        self.focus_force()

        frm = ttk.Frame(self, padding="10 10 10 10")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=self.translator.gettext("email_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(50, 5))
        self.entry_email = ttk.Entry(frm, width=50, font=("Arial", 12))
        self.entry_email.pack(pady=(0,15))

        ttk.Label(frm, text=self.translator.gettext("password_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(5,5))
        self.entry_password = ttk.Entry(frm, width=50, show="*", font=("Arial", 12))
        self.entry_password.pack(pady=(0,25))

        btn_login = ttk.Button(frm, text=self.translator.gettext("login_button"), command=self.do_login, width=20)
        btn_login.pack(pady=20)

    def exit_fullscreen(self, event=None):
        self.attributes("-fullscreen", False)

    def do_login(self):
        email_addr = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        print(f"Versuche, mit E-Mail: {email_addr} und Passwort: {'*' * len(password)} einzuloggen.")

        if not email_addr or not password:
            messagebox.showwarning(self.translator.gettext("warning_title"), self.translator.gettext("error_invalid_input"))
            return

        success = test_imap_login(email_addr, password)
        print(f"Login erfolgreich: {success}")
        if success:
            save_credentials(email_addr, password)
            messagebox.showinfo(self.translator.gettext("success_title"), self.translator.gettext("login_success"))
            self.destroy()
            self.master.open_main_window(email_addr, password)
        else:
            messagebox.showerror(
                self.translator.gettext("error_title"),
                self.translator.gettext("login_failure")
            )


class ManageBlockedTimesWindow(tk.Toplevel):
    def __init__(self, master=None, translator=None):
        super().__init__(master)
        self.translator = translator or Translator()
        self.title(self.translator.gettext("blocked_times_manage_title"))
        self.geometry("500x400")
        self.resizable(False, False)
        self.bind("<Escape>", self.close_window)
        self.transient(master)

        self.blocked_times = load_blocked_times()

        frm = ttk.Frame(self, padding="10 10 10 10")
        frm.pack(fill=tk.BOTH, expand=True)

        add_frame = ttk.LabelFrame(frm, text=self.translator.gettext("add_blocked_time_label"), padding="10 10 10 10")
        add_frame.pack(fill=tk.X, pady=(0,10))

        ttk.Label(add_frame, text=self.translator.gettext("appointment_date"), font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=5)
        self.entry_date = ttk.Entry(add_frame, width=20, font=("Arial", 12))
        self.entry_date.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(add_frame, text=self.translator.gettext("appointment_start_time"), font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=5)
        self.entry_start_time = ttk.Entry(add_frame, width=20, font=("Arial", 12))
        self.entry_start_time.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(add_frame, text=self.translator.gettext("appointment_end_time"), font=("Arial", 12)).grid(row=2, column=0, sticky="w", pady=5)
        self.entry_end_time = ttk.Entry(add_frame, width=20, font=("Arial", 12))
        self.entry_end_time.grid(row=2, column=1, padx=5, pady=5)

        btn_add = ttk.Button(add_frame, text=self.translator.gettext("add_button"), command=self.add_blocked_time, width=20)
        btn_add.grid(row=3, column=0, columnspan=2, pady=20)

        display_frame = ttk.LabelFrame(frm, text=self.translator.gettext("blocked_times_label"), padding="10 10 10 10")
        display_frame.pack(fill=tk.BOTH, expand=True)

        self.blocked_tree = ttk.Treeview(display_frame, columns=("date", "start_time", "end_time"), show="headings", height=10)
        self.blocked_tree.heading("date", text=self.translator.gettext("appointment_date"))
        self.blocked_tree.heading("start_time", text=self.translator.gettext("appointment_start_time"))
        self.blocked_tree.heading("end_time", text=self.translator.gettext("appointment_end_time"))
        self.blocked_tree.pack(fill=tk.BOTH, expand=True, pady=(0,10))

        btn_remove = ttk.Button(display_frame, text=self.translator.gettext("remove_button"), command=self.remove_blocked_time, width=30)
        btn_remove.pack(pady=5)

        self.update_blocked_tree()

    def close_window(self, event=None):
        self.destroy()

    def add_blocked_time(self):
        date = self.entry_date.get().strip()
        start_time = self.entry_start_time.get().strip()
        end_time = self.entry_end_time.get().strip()

        try:
            datetime.strptime(date, "%Y-%m-%d")
            datetime.strptime(start_time, "%H:%M")
            datetime.strptime(end_time, "%H:%M")
        except ValueError:
            messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_invalid_input"))
            return

        if start_time >= end_time:
            messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_time_conflict"))
            return

        new_block = {"date": date, "start_time": start_time, "end_time": end_time}
        self.blocked_times.append(new_block)
        save_blocked_times(self.blocked_times)
        self.update_blocked_tree()

        self.entry_date.delete(0, tk.END)
        self.entry_start_time.delete(0, tk.END)
        self.entry_end_time.delete(0, tk.END)
        print(f"Blockierte Zeit hinzugefügt: {new_block}")

    def update_blocked_tree(self):
        for item in self.blocked_tree.get_children():
            self.blocked_tree.delete(item)
        for block in self.blocked_times:
            self.blocked_tree.insert("", tk.END, values=(block["date"], block["start_time"], block["end_time"]))

    def remove_blocked_time(self):
        selected_item = self.blocked_tree.selection()
        if not selected_item:
            messagebox.showwarning(self.translator.gettext("warning_title"), self.translator.gettext("error_no_selection"))
            return
        values = self.blocked_tree.item(selected_item)["values"]
        date, start_time, end_time = values

        self.blocked_times = [block for block in self.blocked_times if not (block["date"] == date and block["start_time"] == start_time and block["end_time"] == end_time)]
        save_blocked_times(self.blocked_times)
        self.update_blocked_tree()
        print(f"Blockierte Zeit entfernt: {values}")


class SettingsWindow(tk.Toplevel):
    def __init__(self, master=None, translator=None):
        super().__init__(master)
        self.translator = translator or Translator()
        self.title(self.translator.gettext("settings_title"))
        self.geometry("400x500")
        self.resizable(False, False)
        self.bind("<Escape>", self.close_window)
        self.transient(master)

        self.settings = load_settings()

        frm = ttk.Frame(self, padding="10 10 10 10")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=self.translator.gettext("settings_name_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 5))
        self.entry_name = ttk.Entry(frm, width=50, font=("Arial", 12))
        self.entry_name.pack(pady=(0,15))

        ttk.Label(frm, text=self.translator.gettext("settings_api_key_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 5))
        self.entry_api_key = ttk.Entry(frm, width=50, show="*", font=("Arial", 12))
        self.entry_api_key.pack(pady=(0,15))

        ttk.Label(frm, text=self.translator.gettext("settings_auftraggeber_email_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 5))
        self.entry_auftraggeber_email = ttk.Entry(frm, width=50, font=("Arial", 12))
        self.entry_auftraggeber_email.pack(pady=(0,15))
        current_auftraggeber_email = self.settings.get("auftraggeber_email", "")
        self.entry_auftraggeber_email.insert(0, current_auftraggeber_email)

        ttk.Label(frm, text=self.translator.gettext("settings_language_label"), font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 5))
        self.language_var = tk.StringVar()
        self.language_combo = ttk.Combobox(frm, textvariable=self.language_var, state="readonly", width=47, font=("Arial", 12))
        self.language_combo['values'] = (self.translator.gettext("language_german"), self.translator.gettext("language_english"))
        current_language = self.settings.get("language", "de")
        if current_language == "de":
            self.language_combo.current(0)
        else:
            self.language_combo.current(1)
        self.language_combo.pack(pady=(0,15))

        current_name = self.settings.get("user_name", "")
        self.entry_name.insert(0, current_name)

        current_api_key = self.settings.get("openai_api_key", "")
        self.entry_api_key.insert(0, current_api_key)

        btn_save = ttk.Button(frm, text=self.translator.gettext("settings_save_button"), command=self.save_settings, width=20)
        btn_save.pack(pady=10)

    def close_window(self, event=None):
        self.destroy()

    def save_settings(self):
        name = self.entry_name.get().strip()
        api_key = self.entry_api_key.get().strip()
        auftraggeber_email = self.entry_auftraggeber_email.get().strip()
        language = self.language_var.get()

        if language == self.translator.gettext("language_german"):
            language_code = "de"
        elif language == self.translator.gettext("language_english"):
            language_code = "en"
        else:
            language_code = "de"

        if not name:
            messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_invalid_input"))
            return

        if not api_key:
            messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_invalid_input"))
            return

        if auftraggeber_email and not self.is_valid_email(auftraggeber_email):
            messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_invalid_auftraggeber_email"))
            return

        self.settings["user_name"] = name
        self.settings["openai_api_key"] = api_key
        self.settings["auftraggeber_email"] = auftraggeber_email
        self.settings["language"] = language_code
        save_settings(self.settings)
        messagebox.showinfo(self.translator.gettext("success_title"), self.translator.gettext("settings_save_success"))
        print(f"Einstellungen gespeichert: Name={name}, API-Key gesetzt={bool(api_key)}, Sprache={language_code}, Auftraggeber-E-Mail={auftraggeber_email}")
        self.master.translator.set_language(language_code)
        self.master.update_ui_language()
        self.destroy()

    def is_valid_email(self, email):
        regex = r'^\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        return re.match(regex, email) is not None


class AppointmentDetailsWindow(tk.Toplevel):
    def __init__(self, master=None, appointment=None, translator=None):
        super().__init__(master)
        self.translator = translator or Translator()
        self.title(self.translator.gettext("appointment_details_title"))
        self.geometry("800x600")
        self.resizable(True, True)
        self.transient(master)
        self.appointment = appointment

        frm = ttk.Frame(self, padding="15")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=self.translator.gettext("appointment_details_title"), font=("Arial", 16, "bold")).pack(pady=(0, 15))

        if appointment:
            details_frame = ttk.Frame(frm, padding="10")
            details_frame.pack(fill=tk.BOTH, expand=True)

            ttk.Label(details_frame, text=self.translator.gettext("email_subject"), font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="nw", pady=5)
            ttk.Label(details_frame, text=appointment.get("betreff", ""), font=("Arial", 12), wraplength=400, anchor="w", justify="left").grid(row=0, column=1, sticky="w", pady=5)

            ttk.Label(details_frame, text=self.translator.gettext("email_date"), font=("Arial", 12, "bold")).grid(row=1, column=0, sticky="nw", pady=5)
            ttk.Label(details_frame, text=appointment.get("date", ""), font=("Arial", 12)).grid(row=1, column=1, sticky="w", pady=5)

            ttk.Label(details_frame, text=self.translator.gettext("appointment_time"), font=("Arial", 12, "bold")).grid(row=2, column=0, sticky="nw", pady=5)
            ttk.Label(details_frame, text=appointment.get("time", ""), font=("Arial", 12)).grid(row=2, column=1, sticky="w", pady=5)

            ttk.Label(details_frame, text=self.translator.gettext("appointment_summary"), font=("Arial", 12, "bold")).grid(row=3, column=0, sticky="nw", pady=5)
            summary_text = ScrolledText(details_frame, wrap=tk.WORD, height=8, font=("Arial", 12))
            summary_text.grid(row=3, column=1, sticky="w", pady=5)
            summary_text.insert(tk.END, appointment.get("summary", ""))
            summary_text.config(state=tk.DISABLED)

            ttk.Label(details_frame, text=self.translator.gettext("appointment_identifier"), font=("Arial", 12, "bold")).grid(row=4, column=0, sticky="nw", pady=5)
            ttk.Label(details_frame, text=appointment.get("identifier", ""), font=("Arial", 12), wraplength=400, anchor="w", justify="left").grid(row=4, column=1, sticky="w", pady=5)

            btn_delete = ttk.Button(frm, text=self.translator.gettext("delete_button"), command=self.delete_appointment)
            btn_delete.pack(pady=10)
        else:
            ttk.Label(frm, text=self.translator.gettext("no_details_available"), font=("Arial", 14)).pack(pady=100)
            print("Keine Termindetails verfügbar.")

    def delete_appointment(self):
        confirm = messagebox.askyesno(
            self.translator.gettext("confirm_delete_title"),
            self.translator.gettext("confirm_delete_message")
        )
        if confirm:
            self.master.delete_appointment(self.appointment)
            self.destroy()


class MainWindow(tk.Toplevel):
    def __init__(self, master=None, email_addr="", password="", translator=None):
        super().__init__(master)
        self.translator = translator or Translator()
        self.title(f"GMail {self.translator.gettext('inbox_title')}: {email_addr}")
        self.geometry("1900x1300")
        self.bind("<Escape>", self.exit_fullscreen)

        self.email_addr = email_addr
        self.password = password

        self.answered_ids = load_answered_ids()
        self.termine_list = load_termine()
        self.blocked_times_list = load_blocked_times()
        self.settings = load_settings()
        self.user_name = self.settings.get("user_name", "Ihre KI")
        self.openai_api_key = self.settings.get("openai_api_key", "")
        self.language = self.settings.get("language", "de")
        self.translator.set_language(self.language)

        container = ttk.Frame(self, padding="10 10 10 10")
        container.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        right_frame = ttk.Frame(container)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        lbl_title = ttk.Label(left_frame, text=self.translator.gettext("inbox_title"), font=("Arial", 18, "bold"))
        lbl_title.pack(pady=(0,20))

        self.email_list = tk.Listbox(left_frame, width=100, height=25, font=("Arial", 12))
        self.email_list.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0,20))

        details_frame = ttk.Frame(left_frame)
        details_frame.pack(fill=tk.BOTH, expand=True)

        self.details_header = ttk.Label(details_frame, text=self.translator.gettext("email_details_header"), anchor="w", justify="left", font=("Arial", 14, "bold"))
        self.details_header.pack(anchor="nw", pady=(0,10))

        self.details_text = ScrolledText(details_frame, wrap=tk.WORD, height=15, font=("Arial", 12))
        self.details_text.pack(fill=tk.BOTH, expand=True)

        # Kalender-Bereich
        kalender_label = ttk.Label(right_frame, text=self.translator.gettext("appointment_title"), font=("Arial", 18, "bold"))
        kalender_label.pack(pady=(0,20))

        self.calendar = Calendar(right_frame, selectmode='day', date_pattern='yyyy-mm-dd', font=("Arial", 12))
        self.calendar.pack(pady=(0,20))

        btn_show_appointments = ttk.Button(right_frame, text=self.translator.gettext("appointment_show_button"), command=self.show_appointments_for_selected_date, width=25)
        btn_show_appointments.pack(pady=10)

        lbl_all_appointments = ttk.Label(right_frame, text=self.translator.gettext("appointment_all_label"), font=("Arial", 14, "bold"))
        lbl_all_appointments.pack(pady=(20,10))

        self.all_appointments_tree = ttk.Treeview(right_frame, columns=("date", "time", "betreff"), show="headings", height=15, selectmode="browse")
        self.all_appointments_tree.heading("date", text=self.translator.gettext("appointment_date"))
        self.all_appointments_tree.heading("time", text=self.translator.gettext("appointment_time"))
        self.all_appointments_tree.heading("betreff", text=self.translator.gettext("appointment_subject"))
        self.all_appointments_tree.column("date", width=100, anchor="center")
        self.all_appointments_tree.column("time", width=100, anchor="center")
        self.all_appointments_tree.column("betreff", width=200, anchor="w")
        self.all_appointments_tree.pack(fill=tk.BOTH, expand=True, pady=(0,20))
        self.all_appointments_tree.bind("<Double-1>", self.on_double_click_termin_all)

        lbl_day_appointments = ttk.Label(right_frame, text=self.translator.gettext("appointment_day_label"), font=("Arial", 14, "bold"))
        lbl_day_appointments.pack(pady=(10,10))

        self.day_appointments_tree = ttk.Treeview(right_frame, columns=("time", "betreff"), show="headings", height=10, selectmode="browse")
        self.day_appointments_tree.heading("time", text=self.translator.gettext("appointment_time"))
        self.day_appointments_tree.heading("betreff", text=self.translator.gettext("appointment_subject"))
        self.day_appointments_tree.column("time", width=100, anchor="center")
        self.day_appointments_tree.column("betreff", width=200, anchor="w")
        self.day_appointments_tree.pack(fill=tk.BOTH, expand=True, pady=(0,20))
        self.day_appointments_tree.bind("<Double-1>", self.on_double_click_termin_day)

        btn_manage_blocked = ttk.Button(right_frame, text=self.translator.gettext("blocked_times_manage_button"), command=self.open_manage_blocked_times, width=30)
        btn_manage_blocked.pack(pady=10)

        btn_settings = ttk.Button(right_frame, text=self.translator.gettext("settings_button"), command=self.open_settings, width=20)
        btn_settings.pack(pady=10)

        btn_tray = ttk.Button(right_frame, text=self.translator.gettext("tray_minimize_button"), command=self.minimize_to_tray, width=25)
        btn_tray.pack(pady=10)


        self.email_data = list_emails_imap(email_addr, password, max_results=50)
        self.fill_listbox()

        for em in self.email_data:
            message_id = em.get("message_id", "")
            if message_id:
                self.answered_ids.add(message_id)
        save_answered_ids(self.answered_ids)
        print(self.translator.gettext("emails_marked_as_answered"))

        self.email_list.bind("<<ListboxSelect>>", self.show_email_details)

        self.mark_appointments_on_calendar()
        self.fill_all_appointments_tree()

        self.check_interval = 60_000
        self.poll_emails()

        self.icon = None
        self.create_tray_icon()

        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

    def exit_fullscreen(self, event=None):
        self.attributes("-fullscreen", False)

    def create_tray_icon(self):
        try:
            image = Image.new('RGB', (64, 64), (0, 0, 255))
            draw = ImageDraw.Draw(image)
            draw.rectangle([0, 0, 63, 63], fill=(0, 0, 255))
            draw.text((10, 20), "G", fill=(255, 255, 255))

            menu = (
                pystray.MenuItem(self.translator.gettext("tray_restore"), self.restore_window),
                pystray.MenuItem(self.translator.gettext("tray_exit"), self.exit_application)
            )

            self.icon = pystray.Icon("GMail Agent", image, "GMail Agent", menu)
            tray_thread = Thread(target=self.icon.run, daemon=True)
            tray_thread.start()
            print(self.translator.gettext("tray_icon_created"))
        except Exception as e:
            print(f"Fehler beim Erstellen des Tray-Icons: {e}")

    def restore_window(self, icon, item):
        self.deiconify()
        self.attributes("-fullscreen", True)
        print(self.translator.gettext("app_restored_tray"))

    def exit_application(self, icon, item):
        icon.stop()
        self.master.quit()
        print(self.translator.gettext("app_exit"))

    def minimize_to_tray(self):
        self.withdraw()
        speak(self.translator.gettext("speak_minimized"))
        print(self.translator.gettext("speak_minimized"))

    def fill_listbox(self):
        self.email_list.delete(0, tk.END)
        for i, em in enumerate(self.email_data):
            display_text = f"{self.translator.gettext('email_from')}{em['from']} | {self.translator.gettext('email_subject')}{em['subject']}"
            self.email_list.insert(i, display_text)

    def fill_all_appointments_tree(self):
        self.all_appointments_tree.delete(*self.all_appointments_tree.get_children())
        for t in self.termine_list:
            date_str = t.get("date", "")
            time_str = t.get("time", "")
            betreff = t.get("betreff", "")
            self.all_appointments_tree.insert("", tk.END, values=(date_str, time_str, betreff))

    def update_kalender_tree_day(self):
        self.day_appointments_tree.delete(*self.day_appointments_tree.get_children())
        selected_date = self.calendar.get_date()
        for t in self.termine_list:
            if t.get("date", "") == selected_date:
                time_str = t.get("time", "")
                betreff = t.get("betreff", "")
                self.day_appointments_tree.insert("", tk.END, values=(time_str, betreff))

    def show_appointments_for_selected_date(self):
        self.update_kalender_tree_day()
        print(f"{self.translator.gettext('appointments_for_date')} {self.calendar.get_date()} {self.translator.gettext('displayed')}.")

    def show_email_details(self, event):
        selection = self.email_list.curselection()
        if not selection:
            return
        idx = selection[0]
        if idx < 0 or idx >= len(self.email_data):
            return

        email_info = self.email_data[idx]
        header_txt = (
            f"{self.translator.gettext('email_from')}{email_info['from']}\n"
            f"{self.translator.gettext('email_subject')}{email_info['subject']}\n"
            f"{self.translator.gettext('email_date')}{email_info['date']}"
        )
        self.details_header.config(text=header_txt)

        self.details_text.delete("1.0", tk.END)
        self.details_text.insert(tk.END, email_info['body'])

    def open_manage_blocked_times(self):
        blocked_win = ManageBlockedTimesWindow(self, translator=self.translator)
        blocked_win.grab_set()
        self.blocked_times_list = load_blocked_times()
        print(self.translator.gettext("blocked_times_reloaded"))

    def open_settings(self):
        settings_win = SettingsWindow(self, translator=self.translator)
        settings_win.grab_set()
        self.settings = load_settings()
        self.user_name = self.settings.get("user_name", "Ihre KI")
        self.openai_api_key = self.settings.get("openai_api_key", "")
        self.language = self.settings.get("language", "de")
        self.translator.set_language(self.language)
        print(self.translator.gettext("settings_reloaded"))
        self.update_ui_language()

    def poll_emails(self):
        print(self.translator.gettext("checking_new_emails"))
        new_data = list_emails_imap(self.email_addr, self.password, max_results=50)
        known_ids = {d["message_id"] for d in self.email_data}
        new_emails_detected = False

        for em in new_data:
            mid = em.get("message_id", "")
            if mid and mid not in known_ids:
                self.email_data.append(em)
                new_emails_detected = True
                print(f"{self.translator.gettext('new_email_added')}: {em['subject']} {self.translator.gettext('from')} {em['from']}")

        if new_emails_detected:
            speak(self.translator.gettext("speak_new_email"))
            self.fill_listbox()

        appointments_changed = self.auto_process_new_emails()
        if appointments_changed:
            self.mark_appointments_on_calendar()
            self.fill_all_appointments_tree()
            print(self.translator.gettext("appointments_updated_marked"))

        self.after(self.check_interval, self.poll_emails)

    def auto_process_new_emails(self):
        count_replied = 0
        appointments_changed = False

        auftraggeber_email = self.settings.get("auftraggeber_email", "").lower()

        for em in self.email_data:
            message_id = em.get("message_id", "")
            if not message_id:
                continue
            if message_id in self.answered_ids:
                continue

            from_field = em['from']
            from_email = parse_email_address(from_field).lower()
            subject = em['subject']
            body = em['body']
            to_field = em.get('to', "")
            to_email = parse_email_address(to_field).lower()
            thread = em.get('thread', [])

            low_from = from_email.lower()
            if any(x in low_from for x in ["no-reply@", "noreply@", "accounts.google.com", "postmaster", "mailer-daemon"]):
                print(f"E-Mail '{subject}' von '{from_email}' -> HARDCODED FILTER: {self.translator.gettext('no_reply_no_response')}")
                self.answered_ids.add(message_id)
                continue

            # FALL 1: Auftraggeber
            if from_email == auftraggeber_email:
                print(f"E-Mail von Auftraggeber erkannt: {subject}")
                appointment_data = parse_appointment_request(from_email, subject, body, self.termine_list, self.openai_api_key, self.translator)
                if appointment_data.get("date") is not None:
                    if appointment_data.get("time") is None:
                        print(f"Termin '{subject}' am {appointment_data['date']} hat keine gültige Zeit. Termin wird übersprungen.")
                        conversation_history = construct_conversation_history(thread + [{
                            "from": from_field,
                            "subject": subject,
                            "body": body,
                            "date": em.get("date", "")
                        }])
                        fallback_ai_reply = generate_ai_reply(
                            conversation_history, "", self.openai_api_key, self.translator
                        )
                        to_email = parse_email_address(from_field)
                        reply_subject = f"Re: {subject}" if subject else "Re:"
                        success, err = send_email_smtp(
                            sender_email=self.email_addr,
                            password=self.password,
                            to_email=to_email,
                            subject=reply_subject,
                            body_text=fallback_ai_reply,
                            translator=self.translator
                        )
                        if success:
                            print(f"Fallback-Antwort ohne Zeit an den Auftraggeber geschickt: {subject}")
                        else:
                            print(f"Fehler beim Senden der Fallback-Mail: {err}")
                        self.answered_ids.add(message_id)
                        continue

                    appointment_data["from_addr"] = from_email
                    summary = appointment_data.get("summary", "")
                    try:
                        new_end_time = (
                            datetime.strptime(appointment_data["time"], "%H:%M") 
                            + timedelta(hours=1)
                        ).strftime("%H:%M")
                    except ValueError as e:
                        print(f"Fehler beim Parsen der Terminzeit: {e}")
                        self.answered_ids.add(message_id)
                        continue

                    if is_time_conflict(
                        appointment_data["date"],
                        appointment_data["time"],
                        new_end_time,
                        self.termine_list,
                        self.blocked_times_list
                    ):
                        print(f"{self.translator.gettext('appointment_conflict_detected')} {appointment_data['betreff']} am {appointment_data['date']} um {appointment_data['time']}")
                        conflict_reply = generate_conflict_reply(
                            subject=subject,
                            date=appointment_data["date"],
                            time=appointment_data["time"],
                            user_name=self.user_name,
                            openai_api_key=self.openai_api_key,
                            translator=self.translator
                        )
                        print(f"Konflikt: {conflict_reply}")
                        self.answered_ids.add(message_id)
                        continue

                    self.update_termin_in_list(appointment_data)
                    appointments_changed = True

                    # HIER: Bestätigungs-E-Mail JETZT DOCH an den Auftraggeber senden!
                    # => Wir extrahieren hier den "sender_name" aus from_email
                    sender_name = from_email.split("@")[0]  # z.B. "jeff" oder "steffnruh"
                    confirmation_reply = generate_confirmation_reply(
                        subject=subject,
                        date=appointment_data["date"],
                        time=appointment_data["time"],
                        summary=summary,
                        recipient_name=sender_name,  # <-- an den Absender
                        butler_name=self.user_name,  # <-- wir unterschreiben mit dem KI-Namen
                        translator=self.translator
                    )
                    to_email = parse_email_address(from_field)
                    reply_subject = f"Re: {subject}" if subject else "Re:"
                    success, err = send_email_smtp(
                        sender_email=self.email_addr,
                        password=self.password,
                        to_email=to_email,
                        subject=reply_subject,
                        body_text=confirmation_reply,
                        translator=self.translator
                    )
                    if success:
                        print(f"Bestätigungs-E-Mail an Auftraggeber gesendet: {subject} an {to_email}")
                        speak(self.translator.gettext("speak_confirmation_reply_generated"))
                    else:
                        print(f"Fehler beim Senden der Bestätigungs-Mail: {err}")

                    self.answered_ids.add(message_id)

                else:
                    # Kein Termin -> TROTZDEM KI-ANTWORT
                    conversation_history = construct_conversation_history(thread + [{
                        "from": from_field,
                        "subject": subject,
                        "body": body,
                        "date": em.get("date", "")
                    }])
                    fallback_ai_reply = generate_ai_reply(
                        conversation_history, "", self.openai_api_key, self.translator
                    )
                    to_email = parse_email_address(from_field)
                    reply_subject = f"Re: {subject}" if subject else "Re:"
                    success, err = send_email_smtp(
                        sender_email=self.email_addr,
                        password=self.password,
                        to_email=to_email,
                        subject=reply_subject,
                        body_text=fallback_ai_reply,
                        translator=self.translator
                    )
                    if success:
                        print(f"Fallback-Antwort an Auftraggeber geschickt (kein Termin): {subject}")
                    else:
                        print(f"Fehler beim Senden der Fallback-Mail: {err}")
                    self.answered_ids.add(message_id)
                continue

            # FALL 2: Nicht-Auftraggeber
            appointment_data = parse_appointment_request(from_email, subject, body, self.termine_list, self.openai_api_key, self.translator)
            if appointment_data.get("date") is not None:
                if appointment_data.get("time") is None:
                    print(f"Termin '{subject}' am {appointment_data['date']} hat keine gültige Zeit. Termin wird übersprungen.")
                    conversation_history = construct_conversation_history(thread + [{
                        "from": from_field,
                        "subject": subject,
                        "body": body,
                        "date": em.get("date", "")
                    }])
                    fallback_ai_reply = generate_ai_reply(
                        conversation_history, "", self.openai_api_key, self.translator
                    )
                    to_email = parse_email_address(from_field)
                    reply_subject = f"Re: {subject}" if subject else "Re:"
                    success, err = send_email_smtp(
                        sender_email=self.email_addr,
                        password=self.password,
                        to_email=to_email,
                        subject=reply_subject,
                        body_text=fallback_ai_reply,
                        translator=self.translator
                    )
                    if success:
                        count_replied += 1
                        print(f"Fallback-Mail ohne Zeit gesendet an: {to_email}")
                    else:
                        print(f"Fehler beim Senden Fallback: {err}")
                    self.answered_ids.add(message_id)
                    continue

                appointment_data["from_addr"] = from_email
                summary = appointment_data.get("summary", "")
                try:
                    new_end_time = (
                        datetime.strptime(appointment_data["time"], "%H:%M") 
                        + timedelta(hours=1)
                    ).strftime("%H:%M")
                except ValueError as e:
                    print(f"Fehler beim Parsen Zeit: {e}")
                    self.answered_ids.add(message_id)
                    continue

                if is_time_conflict(
                    appointment_data["date"],
                    appointment_data["time"],
                    new_end_time,
                    self.termine_list,
                    self.blocked_times_list
                ):
                    print(f"{self.translator.gettext('appointment_conflict_detected')} {appointment_data['betreff']} am {appointment_data['date']} um {appointment_data['time']}")
                    conflict_reply = generate_conflict_reply(
                        subject=subject,
                        date=appointment_data["date"],
                        time=appointment_data["time"],
                        user_name=self.user_name,
                        openai_api_key=self.openai_api_key,
                        translator=self.translator
                    )
                    to_email = parse_email_address(from_field)
                    reply_subject = f"Re: {subject}" if subject else "Re:"
                    success, err = send_email_smtp(
                        sender_email=self.email_addr,
                        password=self.password,
                        to_email=to_email,
                        subject=reply_subject,
                        body_text=conflict_reply,
                        translator=self.translator
                    )
                    if success:
                        count_replied += 1
                        print(f"{self.translator.gettext('conflict_reply_sent')}: {subject} {self.translator.gettext('to')} {to_email}")
                        speak(self.translator.gettext("speak_conflict_reply_generated"))
                    else:
                        print(f"{self.translator.gettext('error_sending_conflict_reply')} {to_email}: {err}")
                    self.answered_ids.add(message_id)
                    continue

                # ---------- HIER NEUER TEIL, WIR NEHMEN den "sender_name" aus from_email ----------
                sender_name = from_email.split("@")[0]  # z.B. "someone"
                self.update_termin_in_list(appointment_data)
                appointments_changed = True

                confirmation_reply = generate_confirmation_reply(
                    subject=subject,
                    date=appointment_data["date"],
                    time=appointment_data["time"],
                    summary=summary,
                    recipient_name=sender_name,  # <-- an die Person, die uns geschrieben hat
                    butler_name=self.user_name,  # <-- wir unterschreiben als KI
                    translator=self.translator
                )

                to_email = parse_email_address(from_field)
                reply_subject = f"Re: {subject}" if subject else "Re:"
                success, err = send_email_smtp(
                    sender_email=self.email_addr,
                    password=self.password,
                    to_email=to_email,
                    subject=reply_subject,
                    body_text=confirmation_reply,
                    translator=self.translator
                )
                if success:
                    count_replied += 1
                    print(f"{self.translator.gettext('confirmation_reply_sent')}: {subject} {self.translator.gettext('to')} {to_email}")
                    speak(self.translator.gettext("speak_confirmation_reply_generated"))
                else:
                    print(f"{self.translator.gettext('error_sending_confirmation_reply')} {to_email}: {err}")

            # Und hier: generische KI-Antwort
            conversation_history = construct_conversation_history(thread + [{
                "from": from_field,
                "subject": subject,
                "body": body,
                "date": em.get("date", "")
            }])
            ai_reply = generate_ai_reply(
                conversation_history, "", self.openai_api_key, self.translator
            )
            to_email = parse_email_address(from_field)
            reply_subject = f"Re: {subject}" if subject else "Re:"
            success, err = send_email_smtp(
                sender_email=self.email_addr,
                password=self.password,
                to_email=to_email,
                subject=reply_subject,
                body_text=ai_reply,
                translator=self.translator
            )
            if success:
                count_replied += 1
                print(f"{self.translator.gettext('auto_replied')}: {subject} {self.translator.gettext('to')} {to_email}")
                speak(self.translator.gettext("speak_reply_sent"))
            else:
                print(f"{self.translator.gettext('error_sending_reply')} {to_email}: {err}")

            self.answered_ids.add(message_id)

        if appointments_changed:
            print(self.translator.gettext("appointments_changed"))
            return True

        if count_replied > 0:
            print(f"{count_replied} {self.translator.gettext('new_emails_replied')}.")
            save_answered_ids(self.answered_ids)

        return appointments_changed

    def update_termin_in_list(self, appt):
        identifier = appt["identifier"]
        found = False
        for t in self.termine_list:
            if t.get("identifier") == identifier:
                t["date"] = appt["date"]
                t["time"] = appt.get("time", "")
                t["betreff"] = appt.get("betreff", t["betreff"])
                t["summary"] = appt.get("summary", t.get("summary", ""))
                t["from_addr"] = appt.get("from_addr", self.email_addr)
                found = True
                print(f"Termin aktualisiert: {identifier}")
                break

        if not found and appt["date"] is not None:
            new_entry = {
                "identifier": identifier,
                "from_addr": appt.get("from_addr", self.email_addr),
                "date": appt["date"],
                "time": appt.get("time", ""),
                "betreff": appt.get("betreff", self.translator.gettext("appointment_new_subject")),
                "summary": appt.get("summary", "")
            }
            self.termine_list.append(new_entry)
            print(f"Neuer Termin angelegt: {identifier}")

        print(f"Aktuelle Termine: {self.termine_list}")
        save_termine(self.termine_list)
        self.fill_all_appointments_tree()

    def on_double_click_termin_all(self, event):
        selected_item = self.all_appointments_tree.selection()
        if not selected_item:
            return
        item = self.all_appointments_tree.item(selected_item)
        values = item['values']
        if not values or len(values) < 3:
            messagebox.showwarning(self.translator.gettext("warning_title"), self.translator.gettext("error_no_details"))
            return

        date, time, betreff = values
        for appt in self.termine_list:
            if appt["date"] == date and appt["time"] == time and appt["betreff"] == betreff:
                AppointmentDetailsWindow(self, appointment=appt, translator=self.translator)
                print(f"Termindetails geöffnet für: {appt}")
                return

        messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_appointment_not_found"))
        print(f"Termin nicht gefunden: {values}")

    def on_double_click_termin_day(self, event):
        selected_item = self.day_appointments_tree.selection()
        if not selected_item:
            return
        item = self.day_appointments_tree.item(selected_item)
        values = item['values']
        if not values or len(values) < 2:
            messagebox.showwarning(self.translator.gettext("warning_title"), self.translator.gettext("error_no_details"))
            return

        time, betreff = values
        selected_date = self.calendar.get_date()
        for appt in self.termine_list:
            if appt["date"] == selected_date and appt["time"] == time and appt["betreff"] == betreff:
                AppointmentDetailsWindow(self, appointment=appt, translator=self.translator)
                print(f"Termindetails geöffnet für: {appt}")
                return

        messagebox.showerror(self.translator.gettext("error_title"), self.translator.gettext("error_appointment_not_found"))
        print(f"Termin nicht gefunden: Datum={selected_date}, Uhrzeit={time}, Betreff={betreff}")

    def mark_appointments_on_calendar(self):
        self.calendar.calevent_remove('all')
        for appt in self.termine_list:
            date_str = appt.get("date", "")
            betreff = appt.get("betreff", "")
            if date_str:
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                    self.calendar.calevent_create(date_obj, betreff, 'appointment')
                except ValueError:
                    print(f"{self.translator.gettext('error_invalid_date_format')} {date_str}. {self.translator.gettext('expected_format')}")
        self.calendar.tag_config('appointment', background='blue', foreground='white', font=("Arial", 10, "bold"))
        print(self.translator.gettext("appointments_marked_on_calendar"))

    def fill_all_appointments_tree(self):
        self.all_appointments_tree.delete(*self.all_appointments_tree.get_children())
        for t in self.termine_list:
            date_str = t.get("date", "")
            time_str = t.get("time", "")
            betreff = t.get("betreff", "")
            self.all_appointments_tree.insert("", tk.END, values=(date_str, time_str, betreff))

    def update_ui_language(self):
        self.title(f"GMail {self.translator.gettext('inbox_title')}: {self.email_addr}")
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Label):
                        text = child.cget("text")
                        mapping = {
                            "Posteingang": "inbox_title",
                            "Terminkalender": "appointment_title",
                            "Alle Termine": "appointment_all_label",
                            "Termine für ausgewählten Tag": "appointment_day_label",
                            "Termindetails": "appointment_details_title",
                            "Keine Details verfügbar.": "no_details_available",
                            "Blockierte Zeiten verwalten": "blocked_times_manage_button",
                            "Einstellungen": "settings_button",
                            "Minimieren in Tray": "tray_minimize_button",
                            "Neue blockierte Zeit hinzufügen": "add_blocked_time_label",
                            "Blockierte Zeiten": "blocked_times_label",
                            "Speichern": "settings_save_button",
                            "Name der KI:": "settings_name_label",
                            "OpenAI API Key:": "settings_api_key_label",
                            "Sprache:": "settings_language_label",
                            "Hinzufügen": "add_button",
                            "Ausgewählte blockierte Zeit entfernen": "remove_button",
                            "Termindetails": "appointment_details_title",
                            "Keine Details verfügbar.": "no_details_available",
                            "Refresh": "refresh_button",
                            "Aufgabe hinzufügen": "add_task_button",
                        }
                        key = mapping.get(text, None)
                        if key:
                            child.config(text=self.translator.gettext(key))
                    elif isinstance(child, ttk.Button):
                        btn_text = child.cget("text")
                        mapping = {
                            "Hinzufügen": "add_button",
                            "Ausgewählte blockierte Zeit entfernen": "remove_button",
                            "Blockierte Zeiten verwalten": "blocked_times_manage_button",
                            "Einstellungen": "settings_button",
                            "Minimieren in Tray": "tray_minimize_button",
                            "Speichern": "settings_save_button",
                            "Termine anzeigen": "appointment_show_button",
                            "Termindetails anzeigen": "appointment_details_title",
                            "Login": "login_button",
                            "Refresh": "refresh_button",
                            "Aufgabe hinzufügen": "add_task_button",
                            "Feedback anzeigen": "view_feedback_button"
                        }
                        key = mapping.get(btn_text, None)
                        if key:
                            child.config(text=self.translator.gettext(key))

        self.all_appointments_tree.heading("date", text=self.translator.gettext("appointment_date"))
        self.all_appointments_tree.heading("time", text=self.translator.gettext("appointment_time"))
        self.all_appointments_tree.heading("betreff", text=self.translator.gettext("appointment_subject"))

        self.day_appointments_tree.heading("time", text=self.translator.gettext("appointment_time"))
        self.day_appointments_tree.heading("betreff", text=self.translator.gettext("appointment_subject"))

        self.details_header.config(text=self.translator.gettext("email_details_header"))
        print(self.translator.gettext("ui_language_updated"))

    def delete_appointment(self, appointment):
        self.termine_list = [t for t in self.termine_list if t != appointment]
        save_termine(self.termine_list)
        self.fill_all_appointments_tree()
        self.update_kalender_tree_day()
        self.mark_appointments_on_calendar()
        messagebox.showinfo(
            self.translator.gettext("delete_success_title"),
            self.translator.gettext("delete_success_message")
        )
        print(f"Termin gelöscht: {appointment}")



class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        style = ttk.Style(self)
        style.theme_use("clam")

        #self.withdraw()

        stored_email, stored_password = load_credentials()
        self.settings = load_settings()
        language = self.settings.get("language", "de")
        self.translator = Translator(language)

        if stored_email and stored_password:
            if test_imap_login(stored_email, stored_password):
                print(self.translator.gettext("login_success"))
                self.open_main_window(stored_email, stored_password)
            else:
                print(self.translator.gettext("login_failure"))
                self.open_login_window()
        else:
            print(self.translator.gettext("login_no_credentials"))
            self.open_login_window()

    def open_login_window(self):
        login_win = LoginWindow(self, translator=self.translator)
        login_win.grab_set()
        login_win.lift()
        login_win.focus_force()

        # Warte, bis das Fenster tatsächlich sichtbar wird
        login_win.wait_visibility(login_win)
        
        self.update_idletasks()
        print(self.translator.gettext("login_window_opened"))



    def open_main_window(self, email_addr, password):
        main_win = MainWindow(self, email_addr=email_addr, password=password, translator=self.translator)
        main_win.grab_set()
        print(self.translator.gettext("main_window_opened"))

    def update_ui_language(self):
        pass


def main():
    app = MainApp()
    app.mainloop()

if __name__ == "__main__":
    main()