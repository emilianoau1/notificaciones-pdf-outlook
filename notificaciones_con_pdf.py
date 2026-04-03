"""
notificaciones_con_pdf.py
--------------------------
Automatización de notificaciones por correo con adjuntos PDF individuales.

Lee un Excel con datos de clientes y busca automáticamente el PDF
correspondiente a cada registro (por número de unidad/folio).
Envía o guarda como borrador en Outlook según configuración.
Al finalizar, genera un reporte Excel del resultado de cada envío.

Autor: [Tu nombre]
Herramientas: Python, pandas, pathlib, win32com (Outlook)
"""

import os
import sys
import traceback
from pathlib import Path

import pandas as pd

try:
    import win32com.client as win32
except ImportError:
    print("Falta la librería pywin32. Instálala con: pip install pywin32")
    sys.exit(1)


# ─────────────────────────────────────────────
# CONFIGURACIÓN — Ajusta estas rutas y opciones
# ─────────────────────────────────────────────

EXCEL_PATH  = r'C:\ruta\a\tu\archivo\CLIENTES.xlsx'
PDF_FOLDER  = r'C:\ruta\a\tu\carpeta\PDFS'

# Correos en copia (supervisores, administración, etc.)
CC_EMAILS = "supervisor@tuempresa.com; gerencia@tuempresa.com"

# False → crea borradores para revisión antes de enviar
# True  → envía los correos directamente
SEND_DIRECTLY = False


# ─────────────────────────────────────────────
# PLANTILLA DEL CORREO
# ─────────────────────────────────────────────

BODY_TEMPLATE = """
<html>
  <body style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
    <p>Estimado(a) cliente,</p>

    <p>Esperamos que se encuentre muy bien.</p>

    <p>
      Por medio del presente, le informamos que como parte de nuestro proceso de revisión
      al término del contrato, hemos identificado condiciones que generan cargos adicionales
      conforme a lo establecido en las condiciones pactadas.
    </p>

    <p>
      En el documento adjunto encontrará el detalle de las condiciones identificadas
      y los cargos correspondientes, los cuales se verán reflejados en su facturación
      en un plazo de 30 días naturales.
    </p>

    <p>
      Agradecemos su comprensión y confianza. Estamos a su disposición para
      aclarar cualquier duda.
    </p>

    <p>Reciba un cordial saludo,<br>[Nombre de tu empresa]</p>
  </body>
</html>
"""


# ─────────────────────────────────────────────
# FUNCIONES
# ─────────────────────────────────────────────

def normalizar_texto(valor):
    """Convierte a string y elimina espacios. Devuelve '' si es NaN."""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def buscar_pdf(carpeta_pdfs: Path, identificador: str):
    """
    Busca el PDF que corresponde a un identificador (número de unidad, folio, etc.).

    Estrategia:
    1. Busca coincidencia exacta: IDENTIFICADOR.pdf
    2. Si no existe, busca sin distinción de mayúsculas/minúsculas

    Retorna el Path del archivo si lo encuentra, o None si no existe.
    """
    # Intento 1: nombre exacto
    for extension in ['.pdf', '.PDF']:
        candidato = carpeta_pdfs / f"{identificador}{extension}"
        if candidato.exists():
            return candidato

    # Intento 2: búsqueda sin distinción de mayúsculas
    for archivo in carpeta_pdfs.iterdir():
        if archivo.is_file() and archivo.suffix.lower() == '.pdf':
            if archivo.stem.strip().upper() == identificador.upper():
                return archivo

    return None


def validar_columnas(df: pd.DataFrame, columnas_requeridas: list):
    """Verifica que el Excel tenga las columnas necesarias. Lanza error si faltan."""
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en el Excel: {', '.join(faltantes)}")


def crear_correo(outlook, destinatario, cc, asunto, cuerpo_html, ruta_pdf=None):
    """
    Crea un correo en Outlook con los datos proporcionados.
    Si se indica ruta_pdf, adjunta el archivo.
    Retorna el objeto mail sin enviarlo ni guardarlo.
    """
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.CC = cc
    mail.Subject = asunto
    mail.HTMLBody = cuerpo_html

    if ruta_pdf is not None:
        mail.Attachments.Add(str(ruta_pdf))

    return mail


# ─────────────────────────────────────────────
# PROCESO PRINCIPAL
# ─────────────────────────────────────────────

def main():
    excel_file = Path(EXCEL_PATH)
    pdf_folder = Path(PDF_FOLDER)

    # Validar que existan las rutas configuradas
    if not excel_file.exists():
        raise FileNotFoundError(f"No se encontró el Excel: {excel_file}")
    if not pdf_folder.exists():
        raise FileNotFoundError(f"No se encontró la carpeta de PDFs: {pdf_folder}")

    print("Leyendo Excel...")
    df = pd.read_excel(excel_file)

    # Validar columnas requeridas
    columnas_requeridas = ["ID_UNIDAD", "GRUPO_CLIENTE", "EJECUTIVO", "CORREO"]
    validar_columnas(df, columnas_requeridas)

    # Normalizar texto en todas las columnas requeridas
    for col in columnas_requeridas:
        df[col] = df[col].apply(normalizar_texto)

    # Filtrar filas sin ID o sin correo
    df = df[(df["ID_UNIDAD"] != "") & (df["CORREO"] != "")].copy()

    if df.empty:
        print("No hay registros válidos con ID y correo.")
        return

    outlook = win32.Dispatch("Outlook.Application")

    enviados  = 0
    borradores = 0
    errores   = 0
    reporte   = []

    total = len(df)
    print(f"Se procesarán {total} registros...\n")

    for idx, row in df.iterrows():
        id_unidad      = row["ID_UNIDAD"]
        grupo_cliente  = row["GRUPO_CLIENTE"]
        ejecutivo      = row["EJECUTIVO"]
        correo_destino = row["CORREO"]

        try:
            pdf_encontrado = buscar_pdf(pdf_folder, id_unidad)

            asunto = f"Condiciones de retorno — Unidad {id_unidad} ({grupo_cliente})"
            mail   = crear_correo(
                outlook,
                destinatario = correo_destino,
                cc           = CC_EMAILS,
                asunto       = asunto,
                cuerpo_html  = BODY_TEMPLATE,
                ruta_pdf     = pdf_encontrado
            )

            if pdf_encontrado:
                if SEND_DIRECTLY:
                    mail.Send()
                    estado  = "ENVIADO"
                    detalle = str(pdf_encontrado)
                    enviados += 1
                    print(f"[{idx+1}/{total}] ENVIADO   — {id_unidad} → {correo_destino}")
                else:
                    mail.Save()
                    estado  = "BORRADOR"
                    detalle = str(pdf_encontrado)
                    borradores += 1
                    print(f"[{idx+1}/{total}] BORRADOR  — {id_unidad} → {correo_destino}")
            else:
                # Sin PDF: guarda borrador para adjuntar manualmente
                mail.Save()
                estado  = "BORRADOR SIN PDF"
                detalle = "PDF no encontrado — adjuntar manualmente"
                borradores += 1
                print(f"[{idx+1}/{total}] SIN PDF   — {id_unidad} (borrador guardado)")

            reporte.append({
                "ID_UNIDAD":     id_unidad,
                "GRUPO_CLIENTE": grupo_cliente,
                "CORREO":        correo_destino,
                "ESTADO":        estado,
                "DETALLE":       detalle
            })

        except Exception as e:
            errores += 1
            print(f"[{idx+1}/{total}] ERROR     — {id_unidad}: {e}")
            reporte.append({
                "ID_UNIDAD":     id_unidad,
                "GRUPO_CLIENTE": grupo_cliente,
                "CORREO":        correo_destino,
                "ESTADO":        "ERROR",
                "DETALLE":       str(e)
            })

    # Generar reporte final
    reporte_df  = pd.DataFrame(reporte)
    ruta_reporte = excel_file.parent / "reporte_envios.xlsx"
    reporte_df.to_excel(ruta_reporte, index=False)

    print("\n─────────────────────────────")
    print("Proceso finalizado.")
    print(f"  Total procesados : {total}")
    print(f"  Enviados         : {enviados}")
    print(f"  Borradores       : {borradores}")
    print(f"  Errores          : {errores}")
    print(f"  Reporte guardado : {ruta_reporte}")
    print("─────────────────────────────")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Error general:")
        print(str(e))
        traceback.print_exc()
