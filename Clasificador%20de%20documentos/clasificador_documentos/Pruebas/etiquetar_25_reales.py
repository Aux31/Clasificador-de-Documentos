"""
Toma los 25 correos únicos más recientes del log real,
inserta GINTXXXXXZ en el cuerpo del correo en Outlook,
y añade el ID en cada línea del log correspondiente.
"""
import sys
import io
import pythoncom
import win32com.client
from pathlib import Path
from datetime import datetime, timedelta

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

OUTLOOK_EMAIL = "documentacion@grupointeca.com"
LOG_PATH = Path(r"c:\Users\aux22.gg\Documents\GitHub\Clasificador%20de%20documentos\clasificador_documentos\registros_subidas.log")

# ---------------------------------------------------------------------------
def agrupar_log(lineas):
    """Devuelve lista de ((rem, fecha_dia), [indices]) ordenada por más reciente."""
    grupos = {}
    for i, l in enumerate(lineas):
        partes = l.split("|")
        if len(partes) < 3:
            continue
        rem = partes[2].strip().lower()
        ts_str = partes[1].strip()[:10]
        if not rem or not ts_str or ts_str == "":
            continue
        key = (rem, ts_str)
        if key not in grupos:
            grupos[key] = []
        grupos[key].append(i)
    return sorted(grupos.items(), key=lambda x: x[1][-1], reverse=True)

def encontrar_correo(items, total, rem_buscar, fecha_str):
    fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d")
    desde = fecha_dt - timedelta(days=1)
    hasta = fecha_dt + timedelta(days=2)
    idx = 1
    while idx <= total:
        try:
            msg = items[idx]
            if msg.Class == 43:
                rem = (msg.SenderEmailAddress or "").lower()
                if rem == rem_buscar:
                    rt = msg.ReceivedTime.replace(tzinfo=None)
                    if desde <= rt <= hasta:
                        return msg
        except Exception:
            pass
        idx += 1
    return None

def insertar_tag(msg, etiqueta):
    try:
        html = msg.HTMLBody or ""
        if etiqueta not in html:
            bloque = f'<div style="color:#ffffff;font-size:1px;mso-hide:all">{etiqueta}</div>'
            if "</body>" in html.lower():
                i = html.lower().rfind("</body>")
                html = html[:i] + bloque + html[i:]
            else:
                html = html + bloque
            msg.HTMLBody = html
    except Exception as e:
        print(f"    [WARN HTML] {e}")
    try:
        body = msg.Body or ""
        if etiqueta not in body:
            msg.Body = body + f"\n\n{etiqueta}"
    except Exception as e:
        print(f"    [WARN Body] {e}")
    msg.Save()

# ---------------------------------------------------------------------------
def main():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    bandeja = None
    for store in ns.Stores:
        try:
            if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                bandeja = store.GetDefaultFolder(6)
                break
        except Exception:
            continue

    if not bandeja:
        print(f"[ERROR] No se encontró la cuenta {OUTLOOK_EMAIL}")
        return

    lineas = LOG_PATH.read_text(encoding="utf-8").splitlines()
    grupos = agrupar_log(lineas)[:25]

    items = bandeja.Items
    items.Sort("[ReceivedTime]", True)
    total = items.Count

    lineas_mod = list(lineas)
    encontrados = 0
    no_encontrados = 0

    for contador, ((rem, fecha), indices) in enumerate(grupos, start=1):
        etiqueta = f"GINT{contador:05d}Z"
        msg = encontrar_correo(items, total, rem, fecha)

        if msg:
            insertar_tag(msg, etiqueta)
            asunto = (msg.Subject or "").strip()[:55]
            print(f"  {etiqueta} | {fecha} | {rem[:38]} | {asunto}")
            for i in indices:
                partes = lineas_mod[i].split("|")
                # Insertar ID después del consecutivo: "N | GINTXXXXXZ | ts | ..."
                lineas_mod[i] = partes[0] + "| " + etiqueta + " |" + "|".join(partes[1:])
            encontrados += 1
        else:
            print(f"  {etiqueta} | {fecha} | {rem[:38]} | [NO ENCONTRADO EN OUTLOOK]")
            no_encontrados += 1

    # Backup antes de sobreescribir
    backup = LOG_PATH.with_suffix(".log.bak")
    import shutil
    shutil.copy2(LOG_PATH, backup)

    LOG_PATH.write_text("\n".join(lineas_mod) + "\n", encoding="utf-8")
    print(f"\nLog actualizado. Encontrados: {encontrados} | No encontrados: {no_encontrados}")
    print(f"Backup: {backup.name}")

if __name__ == "__main__":
    main()
