import os
import json

class SettingsManager:
    @staticmethod
    def _get_settings_file():
        appdata = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'BKL_Office_Tools')
        if not os.path.exists(appdata):
            os.makedirs(appdata)
        return os.path.join(appdata, 'settings.json')

    @classmethod
    def get(cls, key, default=None):
        file_path = cls._get_settings_file()
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    return data.get(key, default)
            except Exception:
                pass
        return default

    @classmethod
    def set(cls, key, value):
        file_path = cls._get_settings_file()
        data = {}
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
            except Exception:
                pass
        
        data[key] = value
        try:
            with open(file_path, 'w') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
