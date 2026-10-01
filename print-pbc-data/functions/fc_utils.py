import os
from pathlib import Path
import glob
import shutil
import pandas as pd
from datetime import datetime, timezone
from enum import Enum
from typing import Tuple, List
import math
import pytz
from dotenv import load_dotenv
import locale

def getAbsolutePath(relativePath : str) -> str:
    script_path = os.path.abspath(__file__)
    absolute_path = script_path.replace(r"\functions\utils.py", relativePath)
    return absolute_path

def deleteFilesOnFolder(folder:str, criteria:str="*"):
    for file in glob.glob(os.path.join(folder, criteria)):
        os.remove(file)

def getResourcesFolderAbsolutePath() -> str:
    rootPath = Path(__file__).resolve().parent.parent
    resourcesPath = os.path.join(str(rootPath), "resources")
    if not os.path.exists(resourcesPath):
        os.makedirs(resourcesPath) 
    return resourcesPath

def getTempFolderAbsolutePath(deleteFiles:bool=False) -> str:
    rootPath = Path(__file__).resolve().parent.parent
    resourcesPath = os.path.join(str(rootPath), "temp")
    if not os.path.exists(resourcesPath):
        os.makedirs(resourcesPath) 
    elif deleteFiles:
        deleteFilesTempFolder()    
    
    return resourcesPath
def deleteFilesTempFolder(criteria:str="*"):
    caminhos = glob.glob(os.path.join(getTempFolderAbsolutePath(), criteria))
    for caminho in caminhos:
        try:
            if os.path.isfile(caminho):
                os.remove(caminho)
                print(f"Arquivo removido: {caminho}")
            elif os.path.isdir(caminho):
                shutil.rmtree(caminho)
                print(f"Pasta removida: {caminho}")
        except Exception as e:
            print(f"Erro ao remover {caminho}: {e}")      
       
def now():
    load_dotenv()
    tz_name = os.getenv("TIMEZONE", "UTC")
    tz = pytz.timezone(tz_name)
    return datetime.now(tz)

def formated_now():
    load_dotenv()
    language = os.getenv("LANGUAGE", "PT").upper()

    locales = {
        "PT": "Portuguese_Brazil.1252",
        "ES": "Spanish_Spain.1252",
        "EN": "English_United States.1252",
    }

    try:
        locale.setlocale(locale.LC_TIME, locales.get(language, ""))
    except locale.Error:
        locale.setlocale(locale.LC_TIME, "")

    return now().strftime("%a, %#d de %b de %Y %H:%M")

def template_version() -> str:
    return "2.1.0"    