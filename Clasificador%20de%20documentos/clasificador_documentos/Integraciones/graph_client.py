"""
GraphClient — correos via Graph API REST + SharePoint via Graph API.
Sin win32com — no depende de Outlook abierto.
"""

import tempfile
import time
import uuid
import msal
import requests
from pathlib import Path
from urllib.parse import quote
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_SCOPES = ["https://graph.microsoft.com/.default"]
_TOKEN_REFRESH_MARGEN_SEG = 300  # renovar 5 minutos antes de que expire

from configuracion.ajustes import (
    MODO_MOCK,
    TENANT_ID, CLIENT_ID, CLIENT_SECRET,
    OUTLOOK_EMAIL,
    SHAREPOINT_DRIVE_ID,
    SHAREPOINT_CARPETA_OCS,
    EXTENSIONES_BLOQUEADAS,
)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClient:

    def __init__(self):
        self.modo_mock = MODO_MOCK
        if self.modo_mock:
            print("[GraphClient] Modo MOCK activo — sin conexión real a Graph API")
        else:
            self._inicializar_real()

    # ------------------------------------------------------------------
    # AUTENTICACIÓN
    # ------------------------------------------------------------------

    def _inicializar_real(self):
        self._msal_app = msal.ConfidentialClientApplication(
            client_id=CLIENT_ID,
            client_credential=CLIENT_SECRET,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        )
        self._token_expira_en = 0  # fuerza renovación inmediata
        self._renovar_token_si_necesario()
        self.session = self._crear_sesion()
        print("[GraphClient] Conectado correctamente.")

    def _renovar_token_si_necesario(self):
        if time.time() < self._token_expira_en:
            return
        resultado = self._msal_app.acquire_token_for_client(scopes=_SCOPES)
        if "access_token" not in resultado:
            raise RuntimeError(
                f"Error al obtener token de Graph API: {resultado.get('error_description')}"
            )
        self._token = resultado["access_token"]
        self._headers = {"Authorization": f"Bearer {self._token}"}
        expires_in = resultado.get("expires_in", 3600)
        self._token_expira_en = time.time() + expires_in - _TOKEN_REFRESH_MARGEN_SEG

    def _crear_sesion(self):
        session = requests.Session()
        retries = Retry(
            total=5,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "POST", "PUT"],
        )
        adapter = HTTPAdapter(max_retries=retries, pool_connections=20, pool_maxsize=20)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    def _get(self, url: str) -> dict:
        self._renovar_token_si_necesario()
        r = self.session.get(url, headers=self._headers, timeout=30)
        if not r.ok:
            try:
                err = r.json().get("error", {})
                raise RuntimeError(
                    f"Graph API {r.status_code} [{err.get('code')}]: {err.get('message')}"
                )
            except (ValueError, AttributeError):
                r.raise_for_status()
        return r.json()

    def _post(self, url: str, json: dict) -> dict:
        self._renovar_token_si_necesario()
        r = self.session.post(url, headers=self._headers, json=json, timeout=30)
        r.raise_for_status()
        return r.json()

    def _put_bytes(self, url: str, data: bytes):
        self._renovar_token_si_necesario()
        r = self.session.put(
            url,
            headers={**self._headers, "Content-Type": "application/octet-stream"},
            data=data,
            timeout=60,
        )
        r.raise_for_status()
        return r.json()

    # ------------------------------------------------------------------
    # CORREOS (Graph API REST — sin win32com)
    # ------------------------------------------------------------------

    def obtener_correos_nuevos(self, procesados=None, dir_temporal=None, max_correos=None) -> list[dict]:
        if self.modo_mock:
            return []
        return self._obtener_correos_win32(procesados, dir_temporal, max_correos)

    def _obtener_correos_win32(self, procesados=None, dir_temporal=None, max_correos=None) -> list[dict]:
        import win32com.client
        import pywintypes
        import pythoncom

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except pywintypes.com_error:
                raise RuntimeError("Outlook no está abierto. Ábrelo antes de ejecutar el clasificador.")

            namespace = outlook.GetNamespace("MAPI")
            bandeja = self._buscar_bandeja(namespace)
            if bandeja is None:
                raise RuntimeError(f"No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")

            mensajes = bandeja.Items
            mensajes.Sort("[ReceivedTime]", True)

            # Restrict a últimos 90 días — evita iterar los 17k mensajes completos
            from datetime import datetime, timedelta
            fecha_limite = (datetime.now() - timedelta(days=90)).strftime("%m/%d/%Y")
            mensajes = mensajes.Restrict(f"[ReceivedTime] >= '{fecha_limite}'")
            mensajes.Sort("[ReceivedTime]", True)

            tope = (max_correos * 15) if max_correos else 300

            correos = []
            total = mensajes.Count
            print(f"  [Outlook] {total} mensajes en los últimos 90 días")

            for idx in range(1, min(tope, total) + 1):
                try:
                    msg = mensajes[idx]
                    entry_id = msg.EntryID
                    # Internet Message-ID es estable aunque Exchange reconstruya el OST
                    # EntryID puede cambiar entre sesiones y causar reprocesamiento
                    try:
                        msg_id = msg.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
                        )
                        if msg_id and msg_id.strip():
                            entry_id = msg_id.strip()
                    except Exception:
                        pass  # fallback al EntryID si la propiedad no está disponible

                    if procesados and entry_id in procesados:
                        continue

                    if msg.Attachments.Count == 0:
                        continue

                    adjuntos = self._guardar_adjuntos(msg, dir_temporal)
                    if not adjuntos:
                        continue

                    correos.append({
                        "id":             entry_id,
                        "asunto":         msg.Subject or "",
                        "cuerpo":         msg.Body or "",
                        "remitente":      msg.SenderEmailAddress or "",
                        "nombre_remitente": msg.SenderName or "",
                        "adjuntos":       adjuntos,
                        "_msg_obj":       msg,
                    })

                    if max_correos and len(correos) >= max_correos:
                        break

                except Exception as e:
                    print(f"  [WARN] Error leyendo mensaje #{idx}: {e}")

            return correos

        finally:
            pythoncom.CoUninitialize()

    def _buscar_bandeja(self, namespace):
        email_lower = OUTLOOK_EMAIL.lower()
        for store in namespace.Stores:
            try:
                if email_lower in store.DisplayName.lower():
                    carpeta = store.GetDefaultFolder(6)
                    print(f"[Outlook] Bandeja: {store.DisplayName}")
                    return carpeta
            except Exception:
                continue
        try:
            for cuenta in namespace.Accounts:
                try:
                    if cuenta.SmtpAddress.lower() == email_lower:
                        carpeta = cuenta.DeliveryStore.GetDefaultFolder(6)
                        print(f"[Outlook] Bandeja vía Account: {cuenta.SmtpAddress}")
                        return carpeta
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def _guardar_adjuntos(self, msg, dir_temporal) -> list[dict]:
        adjuntos = []
        for i in range(1, msg.Attachments.Count + 1):
            try:
                adj = msg.Attachments.Item(i)
                nombre = adj.FileName
                if not nombre:
                    continue
                ext = Path(nombre).suffix.lower()
                if ext in EXTENSIONES_BLOQUEADAS:
                    continue
                if dir_temporal:
                    ruta_tmp = str(Path(dir_temporal) / f"adj_{uuid.uuid4().hex[:8]}{ext}")
                else:
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
                    tmp.close()
                    ruta_tmp = tmp.name
                adj.SaveAsFile(ruta_tmp)
                adjuntos.append({"nombre": nombre, "ruta_local": ruta_tmp})
            except Exception as e:
                print(f"  [WARN] Error guardando adjunto: {e}")
        return adjuntos

    # ------------------------------------------------------------------
    # SHAREPOINT
    # ------------------------------------------------------------------

    def buscar_carpeta_oc(self, numero_oc_formateado: str) -> str:
        if self.modo_mock:
            return numero_oc_formateado

        if not hasattr(self, "_cache_carpetas_oc"):
            self._cache_carpetas_oc = {}

        if numero_oc_formateado in self._cache_carpetas_oc:
            return self._cache_carpetas_oc[numero_oc_formateado]

        carpeta_base = quote(SHAREPOINT_CARPETA_OCS, safe="/")
        filtro = quote(f"startswith(name,'{numero_oc_formateado}')", safe="',()")
        url = (
            f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
            f"/root:/{carpeta_base}:/children"
            f"?$select=name&$filter={filtro}"
        )
        try:
            data = self._get(url)
            items = data.get("value", [])
            if items:
                nombre = items[0]["name"]
                self._cache_carpetas_oc[numero_oc_formateado] = nombre
                print(f"Carpeta encontrada: {nombre}")
                return nombre
        except Exception as e:
            print(f"  [WARN] buscar_carpeta_oc: {e}")

        self._cache_carpetas_oc[numero_oc_formateado] = numero_oc_formateado
        return numero_oc_formateado

    def crear_carpeta_si_no_existe(self, ruta: str):
        if self.modo_mock:
            return

        partes = Path(ruta).parent.parts
        ruta_actual = ""
        for parte in partes:
            ruta_actual = f"{ruta_actual}/{parte}" if ruta_actual else parte
            ruta_encoded = quote(ruta_actual, safe="/")
            url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_encoded}:"

            try:
                self._renovar_token_si_necesario()
                r = self.session.get(url, headers=self._headers, timeout=30)
                if r.status_code == 200:
                    continue
            except Exception:
                pass

            parent = str(Path(ruta_actual).parent)
            nombre = Path(ruta_actual).name
            if parent != ".":
                parent_encoded = quote(parent, safe="/")
                crear_url = (
                    f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
                    f"/root:/{parent_encoded}:/children"
                )
            else:
                crear_url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root/children"

            self._post(crear_url, {"name": nombre, "folder": {}})
            print(f"Carpeta creada: {ruta_actual}")
            time.sleep(0.2)

    def archivo_existe(self, ruta: str) -> bool:
        if self.modo_mock:
            return False
        self._renovar_token_si_necesario()
        ruta_encoded = quote(ruta, safe="/")
        url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_encoded}:"
        try:
            r = self.session.get(url, headers=self._headers, timeout=30)
            return r.status_code == 200
        except Exception:
            return False

    def resolver_ruta_versionada(self, ruta: str) -> str:
        if not self.archivo_existe(ruta):
            return ruta
        p = Path(ruta)
        base = str(p.parent)
        stem = p.stem
        suf = p.suffix
        v = 2
        while True:
            nueva = f"{base}/{stem}_v{v}{suf}"
            if not self.archivo_existe(nueva):
                return nueva
            v += 1

    def subir_archivo(self, ruta_local, ruta_sharepoint: str):
        if self.modo_mock:
            print(f"  [MOCK] Subido: {Path(ruta_local).name} -> {ruta_sharepoint}")
            return

        data = Path(ruta_local).read_bytes()
        url = (
            f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
            f"/root:/{quote(ruta_sharepoint, safe='/')}:/content"
        )

        if len(data) <= 4 * 1024 * 1024:
            self._put_bytes(url + "?@microsoft.graph.conflictBehavior=replace", data)
        else:
            session_info = self._post(
                f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
                f"/root:/{quote(ruta_sharepoint, safe='/')}:/createUploadSession",
                {"item": {"@microsoft.graph.conflictBehavior": "replace"}},
            )
            upload_url = session_info["uploadUrl"]
            chunk = 5 * 1024 * 1024
            for i in range(0, len(data), chunk):
                trozo = data[i : i + chunk]
                fin = min(i + chunk - 1, len(data) - 1)
                self.session.put(
                    upload_url,
                    headers={"Content-Range": f"bytes {i}-{fin}/{len(data)}"},
                    data=trozo,
                    timeout=60,
                ).raise_for_status()

        print(f"  Subido: {Path(ruta_local).name} -> {ruta_sharepoint}")

    def enviar_correo(self, destinatario: str, asunto: str, cuerpo: str) -> bool:
        """
        Envia un correo via Graph API en nombre de OUTLOOK_EMAIL.
        Requiere permiso Mail.Send (application) en el App Registration de Azure AD.
        Retorna True si fue exitoso.
        """
        if self.modo_mock:
            print(f"  [MOCK] Correo a {destinatario}: {asunto}")
            return True

        url     = f"{_GRAPH_BASE}/users/{OUTLOOK_EMAIL}/sendMail"
        payload = {
            "message": {
                "subject": asunto,
                "body": {
                    "contentType": "Text",
                    "content": cuerpo,
                },
                "toRecipients": [
                    {"emailAddress": {"address": destinatario}}
                ],
            },
            "saveToSentItems": True,
        }
        try:
            r = self.session.post(url, headers=self._headers, json=payload, timeout=30)
            if r.status_code == 202:
                return True
            print(f"  [WARN] enviar_correo: {r.status_code} {r.text[:200]}")
            return False
        except Exception as e:
            print(f"  [ERROR] enviar_correo: {e}")
            return False

