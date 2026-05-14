import re

LOG_PATH = r"c:\Users\aux22.gg\Documents\GitHub\Clasificador%20de%20documentos\clasificador_documentos\registros_subidas.log"
OUT_PATH = r"c:\Users\aux22.gg\Documents\GitHub\Clasificador%20de%20documentos\clasificador_documentos\registros_subidas_normalizado.log"

FECHA_RE = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$')
ID_RE = re.compile(r'^GINT\w+$')

def es_fecha(s):
    return bool(FECHA_RE.match(s.strip()))

def es_id(s):
    return bool(ID_RE.match(s.strip()))

def normalizar_ruta(ruta):
    return ruta.replace('->', '/')

def normalizar_linea(linea):
    linea = linea.rstrip('\n')
    partes = [p.strip() for p in linea.split(' | ')]

    if len(partes) < 4:
        return linea  # no reconocida, dejar igual

    num = partes[0]

    # Detectar si el campo 1 es un ID (GINT...)
    if es_id(partes[1]):
        id_ = partes[1]
        # partes: num | ID | fecha | email | [referencia |] ruta
        if len(partes) == 5:
            # num | ID | fecha | email | ruta
            fecha, email, ruta = partes[2], partes[3], partes[4]
        elif len(partes) >= 6:
            # num | ID | fecha | email | referencia | ruta
            fecha, email, ruta = partes[2], partes[3], partes[-1]
        else:
            return linea
        ruta = normalizar_ruta(ruta)
        return f"{num} | {id_} | {fecha} | {email} | {ruta}"
    else:
        # Sin ID: num | fecha | email | ruta (posiblemente ruta con ->)
        if len(partes) == 4:
            fecha, email, ruta = partes[1], partes[2], partes[3]
        else:
            return linea
        ruta = normalizar_ruta(ruta)
        return f"{num} | {fecha} | {email} | {ruta}"

with open(LOG_PATH, 'r', encoding='utf-8') as f:
    lineas = f.readlines()

normalizadas = [normalizar_linea(l) for l in lineas]

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write('\n'.join(normalizadas))

print(f"Listo. {len(normalizadas)} entradas normalizadas -> {OUT_PATH}")
