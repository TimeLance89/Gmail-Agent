import json
import os

class Translator:
    def __init__(self, language='de'):
        self.language = language
        self.translations = {}
        self.load_language()

    def load_language(self):
        lang_file = os.path.join('locales', f'{self.language}.json')
        if os.path.exists(lang_file):
            with open(lang_file, 'r', encoding='utf-8') as f:
                self.translations = json.load(f)
        else:
            print(f"Sprachdatei nicht gefunden: {lang_file}")

    def gettext(self, key):
        return self.translations.get(key, key)

    def set_language(self, language):
        self.language = language
        self.load_language()
