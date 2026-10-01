
import os
import win32com.client as win32
import base64
import requests
import smtplib
import json
from typing import Optional, List, Tuple
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from functions import fc_utils
from functions.fc_utils import now
from functions import fc_botcity_maestro
from functions.fc_botcity_maestro import BotcityMaestroConnection
from dotenv import load_dotenv
from enum import Enum
from requests_ntlm import HttpNtlmAuth
import urllib.parse
import requests
from msal import ConfidentialClientApplication
from botcity.maestro import *
from time import time as _wall_time

__all__ = ["auth"]

class auth_enumType(Enum):
    Sharepoint = 1
    Email = 2

class auth:
    def _get_o365_token(type : auth_enumType) -> str:
        try:
            credentials = BotcityMaestroConnection().get_credentials("O365")
            client_id = credentials["client_id"]
            client_secret = credentials["client_secret"]    
            tenant_id = credentials["tenant_id"]
            tenant_name = credentials["tenant_name"]
            api_email_get_token = credentials["get_token_url"]

            if type == auth_enumType.Sharepoint:
                scope = f"https://{tenant_name}/.default"
            else:
                scope = f"{client_id}/.default"

            resp = requests.post(
                    api_email_get_token,
                    data={
                        "grant_type": "client_credentials",
                        "client_id": f"{client_id}",
                        "client_secret": f"{client_secret}",
                        "scope": scope
                    },
                    headers= {"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=120
                )
            resp.raise_for_status()
            resp_json = resp.json()

            token = f"{resp_json['access_token']}"
            return token
        except Exception as ex:
            print(ex)
class fc_email:    
    def sendMail(to:str, subject:str, message:str, cc:str = None, cco:str = None, lstFile: List[str] = None) -> bool:
        try:
            load_dotenv()
            bot_identifier = os.getenv("bot_identifier")
       
            if  bot_identifier == "" or bot_identifier.upper() == "DEFAULT":
                raise RuntimeError("To use the email sending functionality, you need to obtain a Bot ID from the APSe Digital team.")

            api_email_send_message = BotcityMaestroConnection().get_credentials("API_Email_2-0")["URL"]
            proxy = BotcityMaestroConnection().get_credentials("PROXY")["http_proxy"]
            os.environ['http_proxy'] = proxy
            os.environ['https_proxy'] = proxy                            

            token =  auth._get_o365_token(type=auth_enumType.Email)

            payload = {
                "bot_identifier": bot_identifier,
                "to": to,
                "cc": cc,
                "cco": cco,
                "subject": subject,
                "message": message
            }

            # Se tiver anexos
            if lstFile:
                attachments = []
                for file_path in lstFile:
                    with open(file_path, "rb") as f:
                        file_bytes = f.read()
                        encoded = base64.b64encode(file_bytes).decode("utf-8")
                        attachments.append({
                            "file_name": file_path.split("\\")[-1].split("/")[-1],  # nome do arquivo
                            "file_content": encoded
                        })
                payload["attachments"] = attachments        

            response = requests.post(
                        url=api_email_send_message , 
                        headers={"Authorization": "Bearer {}".format(token), "Content-Type": "application/json"}, 
                        json=payload, 
                        timeout=120
                    )

            print(response)
            return True
        except Exception as ex:
            return False
class fc_sharepoint:
    """Cliente para operações no SharePoint via Microsoft Graph."""

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    _UPLOAD_SIMPLE_LIMIT = 4 * 1024 * 1024  # 4 MB
    _UPLOAD_CHUNK_SIZE = 10 * 1024 * 1024  # 10 MB

    # -------------------------------------------------------------------------
    # Inicialização
    # -------------------------------------------------------------------------

    def __init__(self):
        """
        Inicializa o cliente SharePoint. 
        """
        
        self._timeout = 600  # 10 minutos (requests usa segundos)
        self._proxies = self._normalize_proxies(
            BotcityMaestroConnection().get_credentials("PROXY")["http_proxy"]
        )
        credentials = BotcityMaestroConnection().get_credentials("MSAL")
        self.client_id = credentials["client_id"]
        self.client_secret = credentials["client_secret"]
        self.tenant_id = credentials["tenant_id"]
        self.scope = ["https://graph.microsoft.com/.default"]

        #Altere essas informações abaixo informando a URL do Sharepoint e a pasta. Ao chamar os métodos de Download, Upload e etc você pode opcionalmente informar valores diferentes
        #Change the information below by specifying the SharePoint URL and folder. When calling the Download, Upload, etc. methods, you can optionally specify different values.
        #Modifique la información a continuación especificando la URL y la carpeta de SharePoint. Al llamar a los métodos Descargar, Cargar, etc., puede especificar valores diferentes opcionalmente.
        self.sharepoint_url = "https://bayergroup.sharepoint.com/sites/028933"
        self.sharepoint_folder = "RPA/Teste Template"

        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise ValueError("client_id, client_secret e tenant_id são obrigatórios")

        _cca_kwargs: dict = {
            "client_id": self.client_id,
            "client_credential": self.client_secret,
            "authority": f"https://login.microsoftonline.com/{self.tenant_id}",
        }
        if self._proxies:
            _cca_kwargs["proxies"] = self._proxies
        self._app = ConfidentialClientApplication(**_cca_kwargs)
        self._access_token: str | None = None
        self._token_expires_on: float | None = None
        self._site_id_cache: dict[str, str] = {}

    # -------------------------------------------------------------------------
    # Métodos públicos
    # -------------------------------------------------------------------------

    def download(
        self,
        output_dir: str,
        file_name: str | None = None,
        sharepoint_url: str | None = None,
        sharepoint_folder: str | None = None,
        recursive: bool = True,
    ) -> list[str]:
        """
        Faz download de arquivos de uma pasta.

        Args:
            output_dir: Diretório local de destino.
            sharepoint_url: ID do site ou URL. Usa self.sharepoint_url se None.
            sharepoint_folder: Caminho da pasta. Usa self.sharepoint_folder se None.
            file_name: Se informado, baixa apenas este arquivo.
            recursive: Se True, inclui subpastas.

        Returns:
            Lista de caminhos dos arquivos baixados.
        """
        site_id = self._ensure_site_id(sharepoint_url)
        folder_path = self._ensure_folder(sharepoint_folder)
        folder_path = self._normalize_drive_path(folder_path, site_id)
        os.makedirs(output_dir, exist_ok=True)

        if file_name:
            file_path = f"{folder_path}/{file_name}".strip("/")
            output_path = os.path.join(output_dir, file_name)
            self._download_file_by_path(site_id, file_path, output_path)
            return [output_path]

        return self._download_folder_recursive(
            site_id, folder_path, output_dir, recursive
        )

    def list_download_files(
        self,
        output_dir: str,
        sharepoint_url: str | None = None,
        sharepoint_folder: str | None = None,
    ) -> list[dict]:
        """
        Faz download dos arquivos (recursivo) e retorna lista com nome e criador.

        Args:
            output_dir: Diretório local de destino.
            sharepoint_url: ID do site ou URL. Usa self.sharepoint_url se None.
            sharepoint_folder: Caminho da pasta. Usa self.sharepoint_folder se None.

        Returns:
            Lista de dicts: nomeArquivo, caminhoSharepoint, pastaUsuario, salvoEm, criadoPor.
            Em erro: nome_arquivo, erro_download, pasta_usuario.
        """
        site_id = self._ensure_site_id(sharepoint_url)
        folder_path = self._ensure_folder(sharepoint_folder)
        folder_path = self._normalize_drive_path(folder_path, site_id)
        os.makedirs(output_dir, exist_ok=True)
        return self._list_download_files_recursive(
            site_id, folder_path, output_dir, None, set()
        )

    def upload(
        self,
        file_paths: list[str],
        sharepoint_url: str | None = None,
        sharepoint_folder: str | None = None,
    ) -> list[dict]:
        """
        Faz upload de uma lista de arquivos.

        Args:
            file_paths: Lista de caminhos locais dos arquivos.
            sharepoint_url: ID do site ou URL. Usa self.sharepoint_url se None.
            sharepoint_folder: Pasta de destino. Usa self.sharepoint_folder se None.

        Returns:
            Lista de driveItem para cada arquivo enviado.
        """
        site_id = self._ensure_site_id(sharepoint_url)
        folder_path = self._ensure_folder(sharepoint_folder)
        folder_path = folder_path.replace("%20", " ").strip("/")
        results: list[dict] = []

        for local_path in file_paths:
            if not os.path.isfile(local_path):
                raise FileNotFoundError(f"Arquivo não encontrado: {local_path}")
            filename = os.path.basename(local_path)
            size = os.path.getsize(local_path)

            if size <= self._UPLOAD_SIMPLE_LIMIT:
                result = self._upload_file_simple(site_id, folder_path, local_path, filename)
            else:
                result = self._upload_file_session(site_id, folder_path, local_path, filename)
            results.append(result)
        return results

    def move_file(
        self,
        source_folder: str,
        destination_folder: str,
        file_name: str,
        sharepoint_url: str | None = None,
        dest_name: str | None = None,
    ) -> dict:
        """
        Move um arquivo de uma pasta para outra.

        Args:
            source_folder: Pasta de origem.
            destination_folder: Pasta de destino (criada automaticamente se não existir).
            file_name: Nome do arquivo a mover.
            sharepoint_url: ID do site ou URL. Usa self.sharepoint_url se None.
            dest_name: Opcional. Novo nome do arquivo no destino.

        Returns:
            driveItem do arquivo movido.
        """
        site_id = self._ensure_site_id(sharepoint_url)
        source_folder = self._normalize_drive_path(source_folder, site_id)
        destination_folder = self._normalize_drive_path(destination_folder, site_id)

        dest_item = self._ensure_drive_folder_path_exists(site_id, destination_folder)
        dest_id = dest_item["id"]
        site_param = self._site_param(site_id)

        file_path = f"{source_folder}/{file_name}".strip("/")
        item = self._get_item_by_path(site_id, file_path)
        url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/items/{item['id']}"

        patch_body: dict = {"parentReference": {"id": dest_id}}
        if dest_name is not None:
            patch_body["name"] = dest_name

        resp = requests.patch(url, headers=self._get_headers(), json=patch_body, **self._request_kwargs())
        resp.raise_for_status()
        return resp.json()

    def delete_file(self, file_path: str, sharepoint_url: str | None = None) -> None:
        """Deleta um arquivo (envia para a Lixeira)."""
        try:
            site_id = self._ensure_site_id(sharepoint_url)
            file_path = file_path.replace("%20", " ").strip("/")
            item = self._get_item_by_path(site_id, file_path)
            site_param = self._site_param(site_id)
            url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/items/{item['id']}"
            response = requests.delete(url, headers=self._get_headers(), **self._request_kwargs())
            response.raise_for_status()
        except Exception as e:
            print(f"Erro ao deletar arquivo: {e}")

    # -------------------------------------------------------------------------
    # Métodos privados - Autenticação
    # -------------------------------------------------------------------------

    def _get_access_token(self) -> str:
        """Obtém o token de acesso para as requisições à API.
        Renova automaticamente quando expirado (token vale ~1 hora)."""
        buffer_seconds = 300  # Renova 5 min antes de expirar
        now = _wall_time()
        if (
            self._access_token
            and self._token_expires_on
            and now < self._token_expires_on - buffer_seconds
        ):
            return self._access_token

        result = self._app.acquire_token_for_client(scopes=self.scope)
        if "access_token" not in result:
            raise Exception(
                f"Falha na autenticação: {result.get('error_description', result)}"
            )
        self._access_token = result["access_token"]
        if "expires_on" in result:
            self._token_expires_on = float(result["expires_on"])
        else:
            self._token_expires_on = now + result.get("expires_in", 3600)
        return self._access_token

    def _get_headers(self) -> dict:
        """Retorna os headers para requisições à API."""
        return {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json",
        }

    @staticmethod
    def _normalize_proxies(raw) -> dict | None:
        """requests exige proxies como dict; credenciais costumam vir como URL única."""
        if not raw:
            return None
        if isinstance(raw, dict):
            return raw
        if isinstance(raw, str):
            url = raw.strip()
            if not url:
                return None
            return {"http": url, "https": url}
        return None

    def _request_kwargs(self, timeout: int | None = None) -> dict:
        """Retorna timeout e proxies para requisições."""
        kwargs: dict = {"timeout": timeout or self._timeout}
        if self._proxies:
            kwargs["proxies"] = self._proxies
        return kwargs

    def _ensure_site_id(self, sharepoint_url: str | None) -> str:
        """Retorna sharepoint_url ou self.sharepoint_url. Levanta se ambos forem None."""
        if sharepoint_url is not None:
            return sharepoint_url
        if self.sharepoint_url is not None:
            return self.sharepoint_url
        raise ValueError("sharepoint_url é obrigatório (informe no método ou defina self.sharepoint_url)")

    def _ensure_folder(self, sharepoint_folder: str | None) -> str:
        """Retorna sharepoint_folder ou self.sharepoint_folder. Levanta se ambos forem None."""
        if sharepoint_folder is not None:
            return sharepoint_folder
        if self.sharepoint_folder is not None:
            return self.sharepoint_folder
        raise ValueError("sharepoint_folder é obrigatório (informe no método ou defina self.sharepoint_folder)")

    # -------------------------------------------------------------------------
    # Métodos privados - Utilitários
    # -------------------------------------------------------------------------

    def _normalize_site_id(self, site_id_or_url: str) -> str:
        """Converte URL para o formato path da API Graph."""
        value = site_id_or_url.strip()
        if "," in value and not value.lower().startswith("http"):
            return value
        if ":/" in value and not value.lower().startswith("http"):
            return value
        if value.lower().startswith("http"):
            parsed = urllib.parse.urlparse(value)
            path = parsed.path.rstrip("/") or "/"
            return f"{parsed.netloc}:{path}"
        return value

    def _resolve_site_id(self, site_id_or_url: str | None) -> str:
        """Resolve o site para obter o ID real (hostname,guid,guid)."""
        site_id_or_url = self._ensure_site_id(site_id_or_url)
        value = site_id_or_url.strip()
        if "," in value:
            return value
        cache_key = self._normalize_site_id(value)
        if cache_key in self._site_id_cache:
            return self._site_id_cache[cache_key]
        encoded = urllib.parse.quote(cache_key, safe="")
        url = f"{self.GRAPH_BASE_URL}/sites/{encoded}"
        response = requests.get(url, headers=self._get_headers(), **self._request_kwargs())
        response.raise_for_status()
        resolved_id = response.json()["id"]
        self._site_id_cache[cache_key] = resolved_id
        return resolved_id

    def _normalize_drive_path(self, path: str, site_id: str | None = None) -> str:
        """Remove prefixo Shared Documents do path. Se site_id for informado, remove
        também o path do site do início (ex: sites/039317/) para compatibilidade."""
        path = path.replace("%20", " ").strip("/")

        # Remove path do site_id do início do folder_path (compatibilidade com formato antigo)
        if site_id:
            site_path = self._extract_site_path_from_id(site_id)
            if site_path and (
                path == site_path or path.startswith(site_path + "/")
            ):
                path = path[len(site_path) :].lstrip("/")

        for prefix in ("Shared Documents/", "Documents/", "Documentos Compartilhados/"):
            if path.lower().startswith(prefix.lower()):
                return path[len(prefix) :].strip("/")
        return path

    def _extract_site_path_from_id(self, site_id_or_url: str) -> str:
        """Extrai o path do site (ex: sites/039317) a partir do site_id ou URL."""
        value = site_id_or_url.strip()
        if not value:
            return ""
        if value.lower().startswith("http"):
            parsed = urllib.parse.urlparse(value)
            return (parsed.path or "").strip("/")
        if ":" in value:
            # Formato hostname:path (ex: bayergroup.sharepoint.com:/sites/039317)
            _, rest = value.split(":", 1)
            return (rest or "").strip("/")
        return ""

    def _site_param(self, site_id: str | None) -> str:
        """Retorna o site_id codificado para uso em URLs."""
        site_id = self._ensure_site_id(site_id)
        resolved = self._resolve_site_id(site_id)
        return (
            urllib.parse.quote(resolved, safe=",")
            if "," in resolved
            else urllib.parse.quote(resolved, safe="")
        )

    # -------------------------------------------------------------------------
    # Métodos privados - API Graph (drive)
    # -------------------------------------------------------------------------

    def _get_item_by_path(self, site_id: str | None, item_path: str) -> dict:
        """Obtém metadados de um item pelo caminho no SharePoint."""
        site_id = self._ensure_site_id(site_id)
        item_path = self._normalize_drive_path(item_path, site_id)
        encoded = urllib.parse.quote(item_path, safe="/") if item_path else ""
        site_param = self._site_param(site_id)
        url = (
            f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{encoded}"
            if encoded
            else f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root"
        )
        response = requests.get(url, headers=self._get_headers(), **self._request_kwargs())
        response.raise_for_status()
        return response.json()

    def _create_drive_child_folder(
        self, site_id: str, parent_path: str, segment_name: str
    ) -> dict:
        """Cria uma subpasta sob parent_path (vazio = raiz da biblioteca)."""
        site_id = self._ensure_site_id(site_id)
        site_param = self._site_param(site_id)
        if parent_path:
            parent_norm = self._normalize_drive_path(parent_path, site_id)
            enc = urllib.parse.quote(parent_norm, safe="/")
            url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{enc}:/children"
        else:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root/children"
        body = {
            "name": segment_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }
        resp = requests.post(
            url, headers=self._get_headers(), json=body, **self._request_kwargs()
        )
        if resp.status_code == 409:
            child_path = (
                f"{parent_path}/{segment_name}".strip("/") if parent_path else segment_name
            )
            child_path = self._normalize_drive_path(child_path, site_id)
            return self._get_item_by_path(site_id, child_path)
        resp.raise_for_status()
        return resp.json()

    def _ensure_drive_folder_path_exists(self, site_id: str, folder_path: str) -> dict:
        """Garante cada segmento do caminho no drive; retorna o driveItem da pasta final."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._normalize_drive_path(folder_path, site_id).strip("/")
        if not folder_path:
            return self._get_item_by_path(site_id, "")

        parts = [p for p in folder_path.split("/") if p]
        parent_path = ""
        for segment in parts:
            current_path = f"{parent_path}/{segment}".strip("/") if parent_path else segment
            try:
                item = self._get_item_by_path(site_id, current_path)
                if "folder" not in item:
                    raise ValueError(
                        f"Já existe um item que não é pasta em: {current_path}"
                    )
            except requests.HTTPError as e:
                if e.response is None or e.response.status_code != 404:
                    raise
                self._create_drive_child_folder(site_id, parent_path, segment)
            parent_path = current_path

        return self._get_item_by_path(site_id, folder_path)

    def _list_folder_items(
        self,
        site_id: str | None,
        folder_path: str | None,
        with_metadata: bool = False,
    ) -> list[dict]:
        """Lista itens de um diretório. Se with_metadata=True, inclui createdBy."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._ensure_folder(folder_path)
        folder_path = self._normalize_drive_path(folder_path, site_id)
        encoded = urllib.parse.quote(folder_path, safe="/") if folder_path else ""
        site_param = self._site_param(site_id)
        select = "id,name,file,folder,createdBy,parentReference,webUrl" if with_metadata else ""
        query = f"?$select={select}" if select else ""

        if encoded:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{encoded}:/children{query}"
        else:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root/children{query}"

        items = []
        while url:
            response = requests.get(url, headers=self._get_headers(), **self._request_kwargs())
            response.raise_for_status()
            data = response.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def _download_from_url(self, api_url: str, output_path: str) -> str:
        """Executa o download a partir da URL da API Graph."""
        response = requests.get(
            api_url,
            headers=self._get_headers(),
            allow_redirects=True,
            **self._request_kwargs(timeout=120),
        )
        response.raise_for_status()
        dir_path = os.path.dirname(output_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)
        with open(output_path, "wb") as f:
            f.write(response.content)
        return output_path

    def _download_file_by_path(
        self, site_id: str | None, file_path: str, output_path: str
    ) -> str:
        """Download interno de um arquivo pelo caminho."""
        site_id = self._ensure_site_id(site_id)
        file_path = self._normalize_drive_path(file_path, site_id)
        encoded = urllib.parse.quote(file_path, safe="/")
        site_param = self._site_param(site_id)
        url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{encoded}:/content"
        return self._download_from_url(url, output_path)

    def _extract_created_by_name(self, item: dict) -> str:
        """Extrai o nome do criador do driveItem."""
        created_by = item.get("createdBy") or {}
        user = created_by.get("user") or {}
        return user.get("displayName")  or "Desconhecido"

    def _extract_created_by_email(self, item: dict) -> str:
        """Extrai o nome do criador do driveItem."""
        created_by = item.get("createdBy") or {}
        user = created_by.get("user") or {}
        return user.get("email") or "Desconhecido"

    # -------------------------------------------------------------------------
    # Métodos privados - Download recursivo
    # -------------------------------------------------------------------------

    def _download_folder_recursive(
        self,
        site_id: str | None,
        folder_path: str | None,
        output_dir: str,
        recursive: bool,
    ) -> list[str]:
        """Baixa recursivamente todos os arquivos de uma pasta."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._ensure_folder(folder_path)
        downloaded: list[str] = []
        os.makedirs(output_dir, exist_ok=True)
        try:
            items = self._list_folder_items(site_id, folder_path)
        except requests.HTTPError as e:
            if e.response.status_code == 404:
                raise FileNotFoundError(f"Pasta não encontrada: {folder_path}") from e
            raise

        for item in items:
            name = item.get("name", "")
            if "folder" in item and recursive:
                sub_path = f"{folder_path}/{name}".strip("/")
                sub_dir = os.path.join(output_dir, name)
                downloaded.extend(
                    self._download_folder_recursive(site_id, sub_path, sub_dir, True)
                )
            elif "file" in item:
                file_path = f"{folder_path}/{name}".strip("/")
                output_path = os.path.join(output_dir, name)
                try:
                    self._download_file_by_path(site_id, file_path, output_path)
                    downloaded.append(output_path)
                except requests.HTTPError as e:
                    if e.response.status_code == 404:
                        continue
                    raise
        return downloaded

    def _list_download_files_recursive(
        self,
        site_id: str | None,
        folder_path: str | None,
        output_dir: str,
        nome_top_folder: str | None,
        arquivos_processados: set[str],
    ) -> list[dict]:
        """Percorre pasta recursivamente, baixa e coleta metadados."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._ensure_folder(folder_path)
        resultados: list[dict] = []
        try:
            items = self._list_folder_items(site_id, folder_path, with_metadata=True)
        except requests.HTTPError as e:
            if e.response.status_code == 404:
                raise FileNotFoundError(f"Pasta não encontrada: {folder_path}") from e
            raise

        nome_pasta = os.path.basename(folder_path) if folder_path else "root"
        if nome_top_folder is None:
            nome_top_folder = nome_pasta
        caminho_atual = os.path.join(output_dir, nome_pasta)
        os.makedirs(caminho_atual, exist_ok=True)

        for item in items:
            name = item.get("name", "")
            if "folder" in item:
                sub_path = f"{folder_path}/{name}".strip("/")
                resultados.extend(
                    self._list_download_files_recursive(
                        site_id, sub_path, caminho_atual, nome_top_folder, arquivos_processados
                    )
                )
            elif "file" in item:
                web_url = item.get("webUrl", "")
                if web_url in arquivos_processados:
                    continue
                arquivos_processados.add(web_url)

                file_path = f"{folder_path}/{name}".strip("/")
                download_path = os.path.join(caminho_atual, name)
                criado_por_nome = self._extract_created_by_name(item)
                criado_por_email = self._extract_created_by_email(item)

                try:
                    self._download_file_by_path(site_id, file_path, download_path)
                    resultados.append({
                        "nomeArquivo": name,
                        "caminhoSharepoint": folder_path,
                        "pastaUsuario": nome_pasta,
                        "salvoEm": caminho_atual,
                        "criadoPorNome": criado_por_nome,
                        "criadoPorEmail": criado_por_email
                    })
                except Exception as e:
                    resultados.append({
                        "nome_arquivo": name,
                        "erro_download": str(e),
                        "pasta_usuario": nome_top_folder,
                    })
        return resultados

    # -------------------------------------------------------------------------
    # Métodos privados - Upload
    # -------------------------------------------------------------------------

    def _upload_file_simple(
        self,
        site_id: str | None,
        folder_path: str | None,
        local_path: str,
        filename: str,
    ) -> dict:
        """Upload via PUT direto (arquivos < 4 MB)."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._ensure_folder(folder_path)
        folder_path = self._normalize_drive_path(folder_path, site_id)
        target = f"{folder_path}/{filename}".strip("/") if folder_path else filename
        encoded = urllib.parse.quote(target, safe="/")
        site_param = self._site_param(site_id)
        url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{encoded}:/content"

        with open(local_path, "rb") as f:
            content = f.read()
        headers = {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/octet-stream",
        }
        response = requests.put(url, headers=headers, data=content, **self._request_kwargs(timeout=120))
        response.raise_for_status()
        return response.json()

    def _upload_file_session(
        self,
        site_id: str | None,
        folder_path: str | None,
        local_path: str,
        filename: str,
    ) -> dict:
        """Upload via session (arquivos > 4 MB)."""
        site_id = self._ensure_site_id(site_id)
        folder_path = self._ensure_folder(folder_path)
        folder_path = self._normalize_drive_path(folder_path, site_id)
        target = f"{folder_path}/{filename}".strip("/") if folder_path else filename
        encoded = urllib.parse.quote(target, safe="/")
        site_param = self._site_param(site_id)
        create_url = f"{self.GRAPH_BASE_URL}/sites/{site_param}/drive/root:/{encoded}:/createUploadSession"

        headers = {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json",
        }
        response = requests.post(
            create_url, headers=headers, json={"item": {"@microsoft.graph.conflictBehavior": "rename"}},
            **self._request_kwargs(),
        )
        response.raise_for_status()
        upload_url = response.json()["uploadUrl"]

        file_size = os.path.getsize(local_path)
        if file_size == 0:
            raise ValueError(f"Arquivo vazio: {local_path}")

        result: dict = {}
        with open(local_path, "rb") as f:
            start = 0
            while start < file_size:
                chunk = f.read(self._UPLOAD_CHUNK_SIZE)
                end = start + len(chunk) - 1
                resp = requests.put(
                    upload_url,
                    headers={
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {start}-{end}/{file_size}",
                    },
                    data=chunk,
                    **self._request_kwargs(timeout=120),
                )
                resp.raise_for_status()
                result = resp.json()
                if resp.status_code in (200, 201):
                    return result
                start = end + 1
        return result

class fc_teams_enumMessageType(Enum):
    Info = "Accent"
    Success = "Good"
    Warning = "Warning"
    Error = "Attention"
class fc_teams:
       
    def sendCardChatBot(
        messageTo: str,
        messageType: fc_teams_enumMessageType,
        messageTitle: str,
        messageBody: str,
        messageLink: str = "",
        messageLinkText: str = "",
        sections: dict[str, list[dict]] = {}
    ):
        try:

            if messageType == fc_teams_enumMessageType.Info:
                messageTitle = "ℹ️ " + messageTitle
            elif messageType == fc_teams_enumMessageType.Success:
                messageTitle = "✅ " + messageTitle
            elif messageType == fc_teams_enumMessageType.Error:
                messageTitle = "❌ " + messageTitle
            elif messageType == fc_teams_enumMessageType.Warning:
                messageTitle = "⚠️ " + messageTitle
            
            load_dotenv()
            bot_name = os.getenv("bot_name")

            body_items = [
                {
                    "type": "TextBlock",
                    "text": messageTitle,
                    "wrap": "true",
                    "size": "Large",
                    "weight": "Bolder"
                },
                {
                    "type": "TextBlock",
                    "text": messageBody,
                    "wrap": "true",
                    "spacing": "Small"
                }
            ]

            for section_name, facts in sections.items():
                if facts:
                    body_items.append({
                        "type": "TextBlock",
                        "text": section_name,
                        "wrap": True,
                        "spacing": "Large",
                        "weight": "Bolder",
                        "color": "Accent"
                    })
                    body_items.append({
                        "type": "FactSet",
                        "facts": facts
                    })

            card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [{
                    "type": "Container",
                        "type": "Container",
                        "style": messageType.value,
                        "showBorder": "true",
                        "roundedCorners": "true",
                    "items": body_items
                },
                {
                "type": "Container",
                "spacing": "Medium",
                "items": [
                    {
                    "type": "ColumnSet",
                    "columns": [
                        {
                        "type": "Column",
                        "width": "auto",
                        "items": [
                            {
                            "type": "TextBlock",
                            "text": bot_name,
                            "isSubtle": "true",
                            "wrap": "true",
                            "size": "Small",
                            "horizontalAlignment": "Left"
                            }
                        ]
                        },
                        {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                            "type": "TextBlock",
                            "text": f"Enviado em: {fc_utils.formated_now()}",
                            "isSubtle": "true",
                            "wrap": "true",
                            "size": "Small",
                            "horizontalAlignment": "Right"
                            }
                        ]
                        }
                    ]
                    }
                ]
                }]
            }

            if messageLink:
                card["actions"] = [{
                    "type": "Action.OpenUrl",
                    "title": messageLinkText,
                    "url": messageLink
                }]

            payload = {
                'to': messageTo,
                'cardjson': json.dumps(card)
            }

            headers = {'Content-Type': 'application/json'}
            
            fc_maestro = BotcityMaestroConnection()
            teamCredentials = fc_maestro.get_credentials("TEAMS")

            teamsUrl = teamCredentials["TeamURLSendBotChat"]
            proxyUrl = teamCredentials["proxyUrl"]
            proxyDomain = teamCredentials["proxyDomain"]
            proxyUser = teamCredentials["proxyUser"]
            proxyPassword = teamCredentials["proxyPassword"]

            session = requests.Session()
            session.proxies = {
                'http': proxyUrl,
                'https': proxyUrl
            }
            session.auth = HttpNtlmAuth(f'{proxyDomain}\\{proxyUser}', proxyPassword)

            response = session.post(teamsUrl, json=payload, headers=headers, timeout=(5,30))


            print(response.status_code)
            print(response.content)
        except Exception as ex:
            print(str(ex))

    # # Dicionário de seções com listas de fatos
    # sections = {
    #     "Resumo": [
    #         {"title": "Itens com sucesso:", "value": "🟢 70"},
    #         {"title": "Itens com erro:", "value": "🔴 0"}
    #     ],
    #     "Parâmetros": [
    #         {"title": "MB51", "value": "N"},
    #         {"title": "COOISPI", "value": "S"},
    #         {"title": "LX2 - BKPC_NEP", "value": "S"},
    #         {"title": "COR3_SubOrdem", "value": "N"},
    #         {"title": "COR3_Lotes", "value": "N"}
    #     ]
    # }

    # # Chamando o método
    # sendCardChatBot(
    #     messageTo="usuario@empresa.com"
    #     messageType=enumStatus.Success,
    #     messageTitle="✅ Sucesso! - Dados processados",
    #     messageBody="Todos os relatórios foram executados com sucesso.",
    #     messageLink="https://bayer.botcity.dev/task/8333756",  # pode deixar vazio "" se não quiser exibir botão
    #     sections=sections
    # )    