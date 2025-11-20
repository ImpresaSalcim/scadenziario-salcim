import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formatdate
from datetime import datetime, timedelta
import os
import json
import requests

# === LETTURA CONFIGURAZIONE ===
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

email_mittente = config["email_mittente"]
password_app = config["password_app"]
smtp_server = config["smtp_server"]
smtp_port = config["smtp_port"]
email_destinatario = config["email_destinatario"]
google_file_id = config["google_drive_file_id"]

# === SCARICA FILE EXCEL DA GOOGLE DRIVE ===
def scarica_excel_google():
    url_download = f"https://drive.google.com/uc?export=download&id={google_file_id}"
    nome_file = "scadenze.xlsx"

    print("📥 Scarico file Excel da Google Drive...")

    r = requests.get(url_download)
    if r.headers.get("Content-Type") != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        print("❌ Il file scaricato non è un Excel valido!")
        print("Content-Type:", r.headers.get("Content-Type"))
        return None

    with open(nome_file, "wb") as f:
        f.write(r.content)

    print("✅ File Excel scaricato correttamente!")
    return nome_file

# === INVIO EMAIL ===
def invia_email(oggetto, html):
    msg = MIMEMultipart("related")
    msg["Subject"] = oggetto
    msg["From"] = email_mittente
    msg["To"] = email_destinatario
    msg["Date"] = formatdate(localtime=True)

    parte_html = MIMEText(html, "html")
    msg.attach(parte_html)

    # Logo opzionale
    logo_path = "logo_salcim.png"
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            msg_image = MIMEImage(f.read())
            msg_image.add_header("Content-ID", "<logo>")
            msg.attach(msg_image)

    # Invio via SMTP SSL
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
        server.login(email_mittente, password_app)
        server.send_message(msg)

# === FUNZIONE PRINCIPALE ===
def controlla_scadenze_excel():
    file_excel = scarica_excel_google()
    if not file_excel:
        return

    df = pd.read_excel(file_excel)

    # Trova colonna data scadenza
    colonna_data = None
    for col in df.columns:
        if "scadenza" in col.lower():
            colonna_data = col
            break

    if not colonna_data:
        print("❌ Nessuna colonna che contenga 'scadenza' trovata!")
        print("Colonne:", df.columns)
        return

    df.rename(columns={colonna_data: "Data Scadenza"}, inplace=True)
    df["Data Scadenza"] = pd.to_datetime(df["Data Scadenza"], errors="coerce").dt.date

    oggi = datetime.now().date()

    # Categorie
    scadute = df[df["Data Scadenza"] < oggi].sort_values("Data Scadenza")
    imminenti = df[(df["Data Scadenza"] >= oggi) & (df["Data Scadenza"] <= oggi + timedelta(days=7))]
    lungo_termine = df[(df["Data Scadenza"] > oggi + timedelta(days=7)) & (df["Data Scadenza"] <= oggi + timedelta(days=30))]

    for g in (scadute, imminenti, lungo_termine):
        if not g.empty:
            g["Data Scadenza"] = g["Data Scadenza"].apply(lambda x: x.strftime("%d/%m/%Y"))

    def crea_tabella_html(df, colore):
        if df.empty:
            return "<p style='color:gray;'>Nessuna scadenza.</p>"
        return df.to_html(index=False, border=0).replace(
            "<table",
            f"<table style='width:100%;border-collapse:collapse;color:{colore};font-size:14px;'"
        )

    html = f"""
    <h2>RIEPILOGO SCADENZE al {oggi.strftime("%d/%m/%Y")}</h2>

    <h3 style="color:red;">Scadenze scadute</h3>
    {crea_tabella_html(scadute, "red")}

    <h3 style="color:orange;">Scadenze imminenti (0-7 giorni)</h3>
    {crea_tabella_html(imminenti, "orange")}

    <h3 style="color:green;">Scadenze a lungo termine (8-30 giorni)</h3>
    {crea_tabella_html(lungo_termine, "green")}
    """

    invia_email(f"RIEPILOGO SCADENZE al {oggi.strftime('%d/%m/%Y')} - Salcim", html)
    print("📧 Email inviata correttamente!")

# === AVVIO ===
if __name__ == "__main__":
    controlla_scadenze_excel()
