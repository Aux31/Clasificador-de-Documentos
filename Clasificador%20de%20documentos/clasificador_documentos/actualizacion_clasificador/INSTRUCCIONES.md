# Instrucciones de Actualización — Clasificador de Documentos

Fecha: 2026-04-24

## Qué contiene esta carpeta

Tres archivos modificados y uno nuevo para el proyecto `clasificador_documentos`.
No se crearon módulos nuevos ni se cambiaron interfaces. Solo se reemplazan estos archivos.

```
actualizacion_clasificador/
├── configuracion/
│   └── ajustes.py          ← REEMPLAZAR
├── Nucleo/
│   ├── validador_campos.py ← REEMPLAZAR
│   └── clasificador.py     ← REEMPLAZAR
├── Pruebas/
│   └── prueba_validaciones_sa.py  ← ARCHIVO NUEVO (copiar)
└── INSTRUCCIONES.md
```

---

## Pasos para aplicar la actualización

### 1. Hacer una copia de seguridad (recomendado)

Antes de reemplazar, guardar los archivos originales:

```
clasificador_documentos/configuracion/ajustes.py          → ajustes.py.bak
clasificador_documentos/Nucleo/validador_campos.py         → validador_campos.py.bak
clasificador_documentos/Nucleo/clasificador.py             → clasificador.py.bak
```

### 2. Reemplazar los archivos modificados

Copiar los siguientes archivos desde esta carpeta hacia la ubicación correspondiente
dentro de `clasificador_documentos/`:

| Archivo en esta carpeta | Destino en el proyecto |
|---|---|
| `configuracion/ajustes.py` | `clasificador_documentos/configuracion/ajustes.py` |
| `Nucleo/validador_campos.py` | `clasificador_documentos/Nucleo/validador_campos.py` |
| `Nucleo/clasificador.py` | `clasificador_documentos/Nucleo/clasificador.py` |

### 3. Copiar el archivo nuevo

| Archivo en esta carpeta | Destino en el proyecto |
|---|---|
| `Pruebas/prueba_validaciones_sa.py` | `clasificador_documentos/Pruebas/prueba_validaciones_sa.py` |

### 4. Verificar la instalación

Abrir una terminal, ir a la carpeta `clasificador_documentos/` y correr:

```
python Pruebas/prueba_validaciones_sa.py
```

El resultado esperado al final es:

```
Resultado: 13/13 pruebas pasaron  ✓ Todo correcto
```

Si aparecen fallos, restaurar los archivos .bak y reportar el error.

---

## Resumen de qué cambió en cada archivo

### configuracion/ajustes.py
- Se agregó la constante `ESTRUCTURA_CARPETAS_OC` con la estructura completa de
  carpetas por OC en SharePoint (12 secciones, más de 50 subcarpetas).
- No se modificó ninguna variable existente.

### Nucleo/validador_campos.py
- Se agregaron 4 funciones nuevas de validación basadas en el Shipping Agreement:
  - `_validar_invoice_en_espanol()` — detecta facturas exclusivamente en inglés
  - `_validar_consignatario_bl()` — verifica nombre legal y cédula de MERCASA en el BL
  - `_validar_flete_en_bl()` — verifica que el flete esté impreso en el BL
  - `_validar_sello_en_packing()` — verifica número de marchamo en el Packing List
- Se actualizó el registro `_VALIDACIONES` para incluir las nuevas funciones
  en los tipos correspondientes (BL, MBL, HBL, AWB, INVOICE, PACKING LIST, PL+INV).
- Las validaciones preexistentes (contenedor ISO 6346, pesos) no se modificaron.

### Nucleo/clasificador.py
- Se agregó el import de `validar_campos` al inicio del archivo.
- Se agregó un bloque al final de `procesar_adjunto()` que fusiona las
  inconsistencias del SA con las que ya devuelve Claude, para que aparezcan
  en los reportes Word generados en `Registros/`.

### Pruebas/prueba_validaciones_sa.py (archivo nuevo)
- Script de prueba unitaria con 13 casos que verifica cada validación SA.
- No afecta el funcionamiento del sistema en producción.
- Resultado verificado: 13/13 pruebas pasaron.
