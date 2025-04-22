"""
settings_manager.py
Handles loading and saving the application configuration from config.ini
"""

import configparser
import os

CONFIG_FILE = "config.ini"

def load_config():
    """
    Load the configuration from config.ini. Returns a configparser.ConfigParser object.
    If the file doesn't exist, create default sections in memory (you can decide to save them).
    """
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
    else:
        # Initialize default config if file is missing
        config["tableau_online"] = {
            "server": "https://us-west-2b.online.tableau.com",
            "api_version": "3.25",
            "personal_access_token_name": "",
            "personal_access_token_secret": "",
            "site_name": "tableaulibertyenergycom",
            "site_url": "tableaulibertyenergycom"
        }
        config["logging"] = {"level": "INFO"}
        config["appearance"] = {"theme": "darkly"}
        config["files"] = {
            "ref_data_path": r"C:\path\to\RefData.xlsx",
            "master_file_path": r"C:\path\to\MasterFile.xlsx"
        }
    return config

def save_config(config: configparser.ConfigParser):
    """Saves the given config object back to config.ini."""
    with open(CONFIG_FILE, "w") as f:
        config.write(f)
