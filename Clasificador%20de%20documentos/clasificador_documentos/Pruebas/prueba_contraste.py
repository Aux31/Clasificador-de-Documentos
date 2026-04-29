"""
Prueba aislada de contrastar_sugerencia_vs_inconsistencias().
No llama a la API de Claude — solo verifica la lógica de contraste local.

Casos probados:
  1. Borrador cubre todo → sin cambios
  2. Borrador omite una inconsistencia → se agrega [AGREGADO]
  3. Borrador vacío → devuelve vacío sin error
  4. Sin inconsistencias → devuelve borrador sin cambios

Uso:
    python pruebas/prueba_contraste.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Nucleo.clasificador_claude import contrastar_sugerencia_vs_inconsistencias

SEP = "-" * 60


def caso(n, titulo, sugerencia, documentos):
    print(f"\n{'='*60}")
    print(f"CASO {n}: {titulo}")
    print(SEP)
    resultado = contrastar_sugerencia_vs_inconsistencias(sugerencia, documentos)
    print(resultado if resultado else "(vacío)")
    agregado = "[AGREGADO" in resultado
    print(f"\n→ ¿Se agregó contenido? {'SÍ' if agregado else 'NO'}")


# ---------------------------------------------------------------------------
# Caso 1: Claude cubrió todo — no debe agregar nada
# ---------------------------------------------------------------------------
caso(
    1,
    "Borrador cubre TODAS las inconsistencias",
    sugerencia=(
        "*** This is an automated notification ***\n"
        "Dear Supplier,\n"
        "Please review the following issues:\n\n"
        "INVOICE:\n"
        "  - The unit price does not match the purchase order. Please reissue the invoice with the correct price.\n"
        "  - The container number TCKU3456789 is missing. Please provide the correct container reference.\n\n"
        "We look forward to your prompt response."
    ),
    documentos=[{
        "nombre_archivo": "INVOICE_PO196893.pdf",
        "tipo": "INVOICE",
        "inconsistencias": [
            {"campo": "unit price", "descripcion": "price does not match purchase order", "severidad": "alta"},
            {"campo": "container number", "descripcion": "container TCKU3456789 missing from invoice", "severidad": "media"},
        ],
    }],
)

# ---------------------------------------------------------------------------
# Caso 2: Claude omitió una inconsistencia — debe agregar [AGREGADO]
# ---------------------------------------------------------------------------
caso(
    2,
    "Borrador OMITE una inconsistencia",
    sugerencia=(
        "*** This is an automated notification ***\n"
        "Dear Supplier,\n\n"
        "INVOICE:\n"
        "  - The unit price does not match the purchase order. Please reissue.\n\n"
        "Thank you for your cooperation."
    ),
    documentos=[{
        "nombre_archivo": "INVOICE_PO196893.pdf",
        "tipo": "INVOICE",
        "inconsistencias": [
            {"campo": "unit price", "descripcion": "price does not match purchase order", "severidad": "alta"},
            # Esta no aparece en el borrador:
            {"campo": "net weight", "descripcion": "net weight exceeds gross weight declared", "severidad": "alta"},
        ],
    }],
)

# ---------------------------------------------------------------------------
# Caso 3: Borrador vacío (Claude falló) — no debe lanzar error
# ---------------------------------------------------------------------------
caso(
    3,
    "Borrador vacío (fallo de API simulado)",
    sugerencia="",
    documentos=[{
        "nombre_archivo": "BL_PO196893.pdf",
        "tipo": "BL",
        "inconsistencias": [
            {"campo": "port of discharge", "descripcion": "port not in Costa Rica (Limón/Caldera)", "severidad": "media"},
        ],
    }],
)

# ---------------------------------------------------------------------------
# Caso 4: Sin inconsistencias — debe devolver el borrador sin cambios
# ---------------------------------------------------------------------------
caso(
    4,
    "Sin inconsistencias (lista vacía)",
    sugerencia="Dear Supplier, everything looks good.",
    documentos=[],
)

# ---------------------------------------------------------------------------
# Caso 5: Múltiples documentos, uno cubierto y otro no
# ---------------------------------------------------------------------------
caso(
    5,
    "Múltiples documentos — solo uno parcialmente cubierto",
    sugerencia=(
        "*** This is an automated notification ***\n"
        "Dear Supplier,\n\n"
        "INVOICE:\n"
        "  - The invoice total does not match packing list subtotals. Please reissue.\n\n"
        "Thank you."
    ),
    documentos=[
        {
            "nombre_archivo": "INVOICE_PO196893.pdf",
            "tipo": "INVOICE",
            "inconsistencias": [
                {"campo": "invoice total", "descripcion": "total does not match packing list subtotals", "severidad": "alta"},
            ],
        },
        {
            "nombre_archivo": "BL_PO196893.pdf",
            "tipo": "BL",
            "inconsistencias": [
                # Estas no aparecen en el borrador:
                {"campo": "shipper name", "descripcion": "shipper name missing from bill of lading", "severidad": "media"},
                {"campo": "notify party", "descripcion": "notify party address incomplete", "severidad": "baja"},
            ],
        },
    ],
)

print(f"\n{'='*60}")
print("Prueba finalizada.")
