from botcity.maestro import *
import os
from typing import Optional, List, Tuple
import requests
import json
from requests_ntlm import HttpNtlmAuth
from dotenv import load_dotenv

class BotcityMaestroConnection:

    _instance = None
    _maestro = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def get_credentials(self, label) -> str:

        load_dotenv()
        bot_identifier = os.getenv("bot_identifier")
        bot_environment = os.getenv("bot_environment")
        api_get_credential_url = os.getenv("api_get_credential_url").replace("%%bot_environment%%", bot_environment).replace("%%bot_identifier%%",bot_identifier).replace("%%cretential_label%%",label)

        payload = {}
        headers = {}

        response = requests.request("GET", api_get_credential_url, headers=headers, data=payload)
        data = json.loads(response.text)
        return {item["key"]: item["value"] for item in data}

    def log_sap_dialog_dismissed(self, tipo_dialogo: str, texto_mensagem: str) -> None:
        try:
            t = (texto_mensagem or "").strip()[:2000]
            print(f"[SAP_DialogWatcher] {tipo_dialogo}: {t}")
        except Exception:
            pass
