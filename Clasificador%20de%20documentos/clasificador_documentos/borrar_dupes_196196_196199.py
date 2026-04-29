"""
Script temporal — borra los duplicados de las POs 196196, 196197, 196198, 196199.
Elimina la segunda subida (15:22-15:30) de cada archivo.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from urllib.parse import quote
from configuracion.ajustes import SHAREPOINT_DRIVE_ID
from Integraciones.graph_client import GraphClient, _GRAPH_BASE

RUTAS_A_BORRAR = [
    # --- 196197 (segunda subida 15:22) ---
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/4.02 Factura Definitiva/adj_3kxud1q3_01_INVOICE_EARINV622755.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/adj_3kxud1q3_02_BL_MEDUYJ146651.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/4.27 Packing list definitivo/adj_3kxud1q3_03_PACKING LIST_SC613370.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/OTROS/adj_3kxud1q3_04_QUALITY CERTIFICATE_ADS-197-26.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/4.10 Certificado Origen definitivo (COO)/adj_3kxud1q3_05_CO_SC613370.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196197-ABLE000/4. DOCUMENTACION/4.18 Certifi Zoos Origen/adj_3kxud1q3_06_HEALTH CERTIFICATE_SC613370.pdf",
    # --- 196199 (segunda subida 15:24) ---
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/4.02 Factura Definitiva/adj_ep83cdst_01_INVOICE_EARINV622754.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/adj_ep83cdst_02_BL_MEDUYJ141595.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/4.27 Packing list definitivo/adj_ep83cdst_03_PACKING LIST_SC613368.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/OTROS/adj_ep83cdst_04_QUALITY CERTIFICATE_ADS-197-26.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/4.10 Certificado Origen definitivo (COO)/adj_ep83cdst_05_CO_SC613368.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196199-ABLE000/4. DOCUMENTACION/4.18 Certifi Zoos Origen/adj_ep83cdst_06_HEALTH CERTIFICATE_SC613368.pdf",
    # --- 196198 (segunda subida 15:27) ---
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/4.02 Factura Definitiva/adj_5_fyxu96_01_INVOICE_EARINV622753.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/adj_5_fyxu96_02_BL_MEDUYJ145489.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/4.27 Packing list definitivo/adj_5_fyxu96_03_PACKING LIST_SC613367.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/OTROS/adj_5_fyxu96_04_QUALITY CERTIFICATE_ADS-189-26ADS-177-26.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/4.10 Certificado Origen definitivo (COO)/adj_5_fyxu96_05_CO_SC613367.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196198-ABLE000/4. DOCUMENTACION/4.18 Certifi Zoos Origen/adj_5_fyxu96_06_HEALTH CERTIFICATE_SC613367.pdf",
    # --- 196196 (segunda subida 15:30) ---
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/4.02 Factura Definitiva/adj_zz6rle45_01_INVOICE_EARINV622752.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/adj_zz6rle45_02_BL_MEDUYJ141512.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/4.27 Packing list definitivo/adj_zz6rle45_03_PACKING LIST_SC613366.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/OTROS/adj_zz6rle45_04_QUALITY CERTIFICATE_ADS-197-26.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/4.10 Certificado Origen definitivo (COO)/adj_zz6rle45_05_CO_SC613366.pdf",
    "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s/cmer-OC-00196196-ABLE000/4. DOCUMENTACION/4.18 Certifi Zoos Origen/adj_zz6rle45_06_HEALTH CERTIFICATE_SC613366.pdf",
]


def main():
    cliente = GraphClient()
    ok = 0
    errores = 0
    for ruta in RUTAS_A_BORRAR:
        url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{quote(ruta, safe='/')}:"
        try:
            cliente._renovar_token_si_necesario()
            resp = cliente.session.delete(url, headers=cliente._headers, timeout=30)
            if resp.status_code in (204, 404):
                estado = "BORRADO" if resp.status_code == 204 else "NO ENCONTRADO"
                print(f"  [{estado}] {ruta.split('/')[-1]}")
                ok += 1
            else:
                print(f"  [ERROR {resp.status_code}] {ruta.split('/')[-1]} — {resp.text[:200]}")
                errores += 1
        except Exception as e:
            print(f"  [EXCEPCION] {ruta.split('/')[-1]} — {e}")
            errores += 1

    print(f"\nListo: {ok} borrados/no encontrados, {errores} errores.")


if __name__ == "__main__":
    main()
