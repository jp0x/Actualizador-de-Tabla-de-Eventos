import imaplib
import email
import pandas as pd
import os
from datetime import datetime
from email.header import decode_header
from dotenv import load_dotenv

# === CARGAR VARIABLES DEL ENTORNO (.env) ===
load_dotenv()

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
ARCHIVO_ENTRADA = os.getenv("ARCHIVO_ENTRADA")
ARCHIVO_SALIDA = os.getenv("ARCHIVO_SALIDA")

# === CONEXIÓN IMAP ===
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)
mail.select("inbox")

# === CARGAR CLIENTES Y FILTROS ===
df = pd.read_excel(ARCHIVO_ENTRADA)
df["Último evento"] = ""

# === FUNCIÓN DE DECODIFICACIÓN DE ASUNTO ===
def decode_subject(msg):
    subject = msg["Subject"]
    decoded, charset = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(charset or "utf-8", errors="ignore")
    return decoded

# === PROCESAR CADA CLIENTE ===
for index, row in df.iterrows():
    remitente = str(row["Remitente_clave"]).strip()
    asunto_inicio = str(row["Asunto_comienza_con"]).strip()

    if not remitente or not asunto_inicio:
        df.at[index, "Último evento"] = "Falta filtro"
        continue

    filtro = f'(FROM "{remitente}" SUBJECT "{asunto_inicio}")'
    try:
        result, data = mail.search(None, filtro)
        correos = data[0].split()
        if correos:
            ultimo_id = correos[-1]
            result, data = mail.fetch(ultimo_id, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            fecha = email.utils.parsedate_to_datetime(msg["Date"])
            df.at[index, "Último evento"] = fecha.strftime("%d-%m-%Y %H:%M")
        else:
            df.at[index, "Último evento"] = "No encontrado"
    except Exception as e:
        df.at[index, "Último evento"] = f"Error: {str(e)}"

mail.logout()

# === GUARDAR RESULTADO ===
df.to_excel(ARCHIVO_SALIDA, index=False)
print(f"[✔] Archivo actualizado guardado como {ARCHIVO_SALIDA}")
