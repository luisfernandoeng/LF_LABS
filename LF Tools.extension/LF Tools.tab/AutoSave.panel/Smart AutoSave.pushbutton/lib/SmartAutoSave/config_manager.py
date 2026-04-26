# -*- coding: utf-8 -*-
import os
import io
import json

class ConfigManager:
    def __init__(self):
        self.appdata = os.getenv('APPDATA')
        self.config_dir = os.path.join(self.appdata, 'pyRevit', 'SmartAutoSave')
        self.config_path = os.path.join(self.config_dir, 'config.json')
        
        # Default settings
        self.settings = {
            "enabled":            True,
            "interval_minutes":   10,
            "countdown_seconds":  5,
            "show_toast":         True,
            "log_history":        True,
            "wait_safe_state":    True,
        }
        
        self.load()

    def _ensure_dir(self):
        if not os.path.exists(self.config_dir):
            try:
                os.makedirs(self.config_dir)
            except:
                pass

    def load(self):
        """Loads settings from JSON file"""
        if os.path.exists(self.config_path):
            try:
                with io.open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    
                    # Update defaults with loaded settings (handles missing keys gracefully)
                    for k, v in loaded_settings.items():
                        if k in self.settings:
                            self.settings[k] = v
            except:
                pass # Fallback to defaults if file is corrupted

    def save(self):
        """Saves current settings to JSON file"""
        self._ensure_dir()
        try:
            with io.open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            pass # Fails silently if permission issue, etc

    def get(self, key, default=None):
        """Get a setting by key"""
        return self.settings.get(key, default)
        
    def set(self, key, value):
        """Set a setting and save"""
        if key in self.settings:
            self.settings[key] = value
            self.save()

# Global config instance for the session
config = ConfigManager()
