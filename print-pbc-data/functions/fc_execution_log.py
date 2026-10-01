import requests
import json
from typing import Optional
from dotenv import load_dotenv
import os
from functions.fc_utils  import now
from threading import Lock
from enum import Enum
import warnings
from functions.fc_botcity_maestro import BotcityMaestroConnection

class ExecutionLogStatus(Enum):
    SUCCESS = 1
    PARTIALLY_COMPLETED = 2
    ERROR = 3

class ExecutionLogConnection:

    _instance = None
    _log_identifier: Optional[str] = None
    _bot_identifier: Optional[str] = None
    _bot_enviroment: Optional[str] = None
    _api_insert_url: Optional[str] = None
    _api_update_url: Optional[str] = None
    _proxy: Optional[str] = None
    
    # Contadores
    _success_count: int = 0
    _error_count: int = 0
    _lock = Lock()
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)

        return cls._instance

    def __init__(self):
        # Evita reconfigurar caso já tenha sido inicializado
        if not self._bot_identifier:
            load_dotenv()
            credentials_apilog = BotcityMaestroConnection().get_credentials("API_Log_1-0")
            self._proxy = BotcityMaestroConnection().get_credentials("PROXY")["http_proxy"]

            self._bot_identifier = os.getenv("bot_identifier")
            self._bot_enviroment = os.getenv("bot_environment")
            self._api_insert_url = credentials_apilog["INSERT_URL"]
            self._api_update_url = credentials_apilog["UPDATE_URL"]

    def start_task(self):
        try:
            if self._bot_identifier.upper() == "DEFAULT" or self._bot_identifier == "":
                warnings.warn("To log the execution of a task, it is necessary to define a bot_identifier in the config/.env file.This method was ignored and no log is being recorded.", UserWarning)
                return 
                
            payload = {
                "environment": self._bot_enviroment,
                "bot_identifier": self._bot_identifier
            }
            headers = {
            'Content-Type': 'application/json'
            }

            os.environ['http_proxy'] = self._proxy
            os.environ['https_proxy'] = self._proxy            

            response = requests.post(f"{self._api_insert_url}", json=payload, timeout=60)
            response.raise_for_status()

            data = response.json()

            self._log_identifier = data.get("identifier")

            if not self._log_identifier:
                raise ValueError("The API did not return a valid log Identifier.")

            self._success_count = 0
            self._error_count = 0

            return self._log_identifier
        except Exception as ex:
            print(ex)

    def finish_task(self, status: ExecutionLogStatus = None, message: str = ""):
        try:
            if self._bot_identifier.upper() == "DEFAULT" or self._bot_identifier == "": 
                return
            
            if not self._log_identifier:
                raise RuntimeError("No tasks are currently running; you need to run the start_task() first")

            if status == None:
                if self._error_count == 0:
                    status = ExecutionLogStatus.SUCCESS
                elif self._success_count >  0: 
                    status = ExecutionLogStatus.PARTIALLY_COMPLETED
                else:
                    status = ExecutionLogStatus.ERROR
                
            payload = {
                "environment": self._bot_enviroment ,
                "identifier": self._log_identifier,
                "status":  status.value,
                "message": message,
                "processed_items": self._success_count,
                "failed_items": self._error_count
            }

            headers = {
            'Content-Type': 'application/json'
            }

            os.environ['http_proxy'] = self._proxy
            os.environ['https_proxy'] = self._proxy            

            response = requests.post(f"{self._api_update_url}", json=payload,timeout=60)
            response.raise_for_status()
            print(response)
        except Exception as ex:
            print(ex)

    def get_log_id(self) -> Optional[int]:
        return self._log_identifier

    def add_success(self, count: int = 1):
        with self._lock:
            self._success_count += count

    def add_error(self, count: int = 1):
        with self._lock:
            self._error_count += count

    def get_counts(self) -> dict:
        return {
            "success_count": self._success_count,
            "error_count": self._error_count
        }

    def reset_counts(self):
        with self._lock:
            self._success_count = 0
            self._error_count = 0
