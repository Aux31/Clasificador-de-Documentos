"""
Prueba unitaria de las validaciones del Shipping Agreement (SA).
Verifica que cada regla SA implementada en validador_campos.py dispara
correctamente con texto positivo y negativo, SIN llamar a Claude ni
necesitar archivos externos.

Uso:
    python Pruebas/prueba_validaciones_sa.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Nucleo.validador_campos import validar_campos

# ---------------------------------------------------------------------------
# Helpers de presentación
# ---------------------------------------------------------------------------
_VERDE  = "\033[92m"
_ROJO   = "\033[91m"
_RESET  = "\033[0m"
_NEGRITA= "\033[1m"

_ok  = 0
_fail = 0

def check(nombre_prueba: str, resultado: list[dict], espera_campo: str | None):
    """
    Verifica que el resultado contenga (o no) la inconsistencia esperada.
    espera_campo = None  → se espera que NO haya inconsistencias.
    espera_campo = str   → se espera al menos una inconsistencia con ese campo.
    """
    global _ok, _fail
    campos_encontrados = [i["campo"] for i in resultado]

    if espera_campo is None:
        paso = len(resultado) == 0
    else:
        paso = any(espera_campo in c for c in campos_encontrados)

    estado = f"{_VERDE}PASS{_RESET}" if paso else f"{_ROJO}FAIL{_RESET}"
    print(f"  {estado}  {nombre_prueba}")

    if not paso:
        if espera_campo:
            print(f"         Esperaba campo con '{espera_campo}', encontré: {campos_encontrados or '(ninguno)'}")
        else:
            print(f"         Esperaba sin inconsistencias, encontré: {campos_encontrados}")
        _fail += 1
    else:
        _ok += 1

    if resultado:
        for inc in resultado:
            sev = inc.get("severidad", "?").upper()
            print(f"         [{sev}] {inc['campo']}: {inc['descripcion'][:90]}...")


# ---------------------------------------------------------------------------
# TEXTOS DE PRUEBA
# ---------------------------------------------------------------------------

# ── V-01: Factura en español ─────────────────────────────────────────────────
INVOICE_EN_ESPANOL = """
FACTURA COMERCIAL
Proveedor: ABC Trading Co.
Cliente: MERCADEO DE ARTICULOS DE CONSUMO S.A.
Descripcion: Mercancias varias
Cantidad: 500 unidades
Precio unitario: USD 10.00
Peso bruto: 1200 KGS
Peso neto: 1000 KGS
Total: USD 5000.00
"""

INVOICE_EN_INGLES = """
COMMERCIAL INVOICE
Shipper: ABC Trading Co.
Consignee: MERCADEO DE ARTICULOS DE CONSUMO S.A.
Description of Goods: Various merchandise
Quantity: 500 units
Unit Price: USD 10.00
Amount: USD 5000.00
"""

# ── V-02: Consignatario en BL ────────────────────────────────────────────────
BL_CONSIGNATARIO_CORRECTO = """
BILL OF LADING
SHIPPER: ABC Trading Co., Shanghai, China
CONSIGNEE: MERCADEO DE ARTICULOS DE CONSUMO S.A.
CARRETERA INTERAMERICANA CARTAGO, LEGAL ID: 3-101-137584
PORT OF LOADING: SHANGHAI
PORT OF DISCHARGE: LIMON, COSTA RICA
FREIGHT PREPAID
"""

BL_CONSIGNATARIO_FALTANTE = """
BILL OF LADING
SHIPPER: ABC Trading Co., Shanghai, China
CONSIGNEE: GRUPO INTECA S.A.
PORT OF LOADING: SHANGHAI
PORT OF DISCHARGE: LIMON, COSTA RICA
FREIGHT PREPAID
"""

BL_SIN_ID_JURIDICO = """
BILL OF LADING
SHIPPER: ABC Trading Co., Shanghai, China
CONSIGNEE: MERCADEO DE ARTICULOS DE CONSUMO S.A.
CARRETERA INTERAMERICANA CARTAGO
PORT OF LOADING: SHANGHAI
PORT OF DISCHARGE: LIMON, COSTA RICA
FREIGHT PREPAID
"""

# ── V-03: Flete en BL ────────────────────────────────────────────────────────
BL_CON_FLETE = """
BILL OF LADING
SHIPPER: ABC Trading Co.
CONSIGNEE: MERCADEO DE ARTICULOS DE CONSUMO S.A. LEGAL ID: 3-101-137584
PORT OF LOADING: NINGBO
PORT OF DISCHARGE: LIMON
FREIGHT PREPAID
CONTAINER: TCKU1234560
"""

BL_SIN_FLETE = """
BILL OF LADING
SHIPPER: ABC Trading Co.
CONSIGNEE: MERCADEO DE ARTICULOS DE CONSUMO S.A. LEGAL ID: 3-101-137584
PORT OF LOADING: NINGBO
PORT OF DISCHARGE: LIMON
CONTAINER: TCKU1234560
"""

# ── V-04: Sello en Packing List ──────────────────────────────────────────────
PL_CON_SELLO = """
PACKING LIST
Container: TCKU1234560
Total Cartons: 200
Total Gross Weight: 5000 KGS
Total Net Weight: 4500 KGS
Seal No.: CR123456
"""

PL_SIN_SELLO = """
PACKING LIST
Container: TCKU1234560
Total Cartons: 200
Total Gross Weight: 5000 KGS
Total Net Weight: 4500 KGS
"""

# ── V-05/V-06: Contenedor ISO 6346 (preexistente) ───────────────────────────
BL_CONTENEDOR_INVALIDO = """
BILL OF LADING
SHIPPER: ABC Co.
CONSIGNEE: MERCADEO DE ARTICULOS DE CONSUMO S.A. LEGAL ID: 3-101-137584
PORT OF LOADING: SHANGHAI
PORT OF DISCHARGE: LIMON
FREIGHT PREPAID
CONTAINER NO: TCKA1234568
"""

# ── V-07: Peso neto > peso bruto (preexistente) ──────────────────────────────
PL_PESOS_INVERTIDOS = """
PACKING LIST
Container: TCKU1234560
Total Cartons: 200
Total Gross Weight: 4000 KGS
Total Net Weight: 5000 KGS
Seal No.: CR123456
"""


# ---------------------------------------------------------------------------
# EJECUCIÓN DE PRUEBAS
# ---------------------------------------------------------------------------

print(f"\n{'═'*65}")
print(f"  {_NEGRITA}PRUEBAS VALIDACIONES SHIPPING AGREEMENT — validador_campos.py{_RESET}")
print(f"{'═'*65}\n")

# ── V-01 Factura en español ──────────────────────────────────────────────────
print(f"{_NEGRITA}V-01 · Idioma de la factura [SA cláusula 5.2]{_RESET}")
check("Factura en español → sin inconsistencia",
      validar_campos("INVOICE", INVOICE_EN_ESPANOL), None)
check("Factura solo en inglés → inconsistencia de idioma",
      validar_campos("INVOICE", INVOICE_EN_INGLES), "Idioma de la factura")
print()

# ── V-02 Consignatario en BL ─────────────────────────────────────────────────
print(f"{_NEGRITA}V-02 · Consignatario en el BL [SA Annex 1]{_RESET}")
check("BL con consignatario correcto → sin inconsistencia",
      validar_campos("BL", BL_CONSIGNATARIO_CORRECTO), None)
check("BL con consignatario incorrecto → inconsistencia",
      validar_campos("BL", BL_CONSIGNATARIO_FALTANTE), "Consignatario")
check("BL sin cédula jurídica → inconsistencia",
      validar_campos("BL", BL_SIN_ID_JURIDICO), "Consignatario")
print()

# ── V-03 Flete en BL ─────────────────────────────────────────────────────────
print(f"{_NEGRITA}V-03 · Flete impreso en el BL [SA cláusula 4 punto 3]{_RESET}")
check("BL con flete indicado → sin inconsistencia",
      validar_campos("BL", BL_CON_FLETE), None)
check("BL sin flete indicado → inconsistencia",
      validar_campos("BL", BL_SIN_FLETE), "Condición de flete")
print()

# ── V-04 Sello en Packing List ───────────────────────────────────────────────
print(f"{_NEGRITA}V-04 · Número de sello en Packing List [SA cláusula 3.3]{_RESET}")
check("Packing List con sello → sin inconsistencia",
      validar_campos("PACKING LIST", PL_CON_SELLO), None)
check("Packing List sin sello → inconsistencia",
      validar_campos("PACKING LIST", PL_SIN_SELLO), "sello")
print()

# ── V-05/V-06 Contenedor ISO 6346 (preexistente) ────────────────────────────
print(f"{_NEGRITA}V-05/V-06 · Número de contenedor ISO 6346 (preexistente){_RESET}")
check("BL con contenedor categoría inválida → inconsistencia",
      validar_campos("BL", BL_CONTENEDOR_INVALIDO), "contenedor")
print()

# ── V-07 Pesos invertidos (preexistente) ─────────────────────────────────────
print(f"{_NEGRITA}V-07 · Peso neto > peso bruto (preexistente){_RESET}")
check("Packing List con pesos invertidos → inconsistencia",
      validar_campos("PACKING LIST", PL_PESOS_INVERTIDOS), "Pesos totales")
print()

# ── Tipos sin texto ───────────────────────────────────────────────────────────
print(f"{_NEGRITA}Casos borde{_RESET}")
check("Texto vacío → sin inconsistencia",
      validar_campos("BL", ""), None)
check("Tipo desconocido → sin inconsistencia",
      validar_campos("OTROS", BL_CON_FLETE), None)
print()

# ── Resumen ───────────────────────────────────────────────────────────────────
total = _ok + _fail
print(f"{'═'*65}")
print(f"  Resultado: {_ok}/{total} pruebas pasaron", end="")
if _fail == 0:
    print(f"  {_VERDE}✓ Todo correcto{_RESET}")
else:
    print(f"  {_ROJO}✗ {_fail} fallo(s){_RESET}")
print(f"{'═'*65}\n")
