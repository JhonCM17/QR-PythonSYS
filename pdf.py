import tkinter as tk
from tkinter import filedialog, messagebox
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image as ReportLabImage, Spacer
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import requests
import urllib.parse

# Función para leer los datos de Google Sheets
def leer_datos():
    try:
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        SERVICE_ACCOUNT_FILE = 'key.json'
        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        SPREADSHEET_ID = '1U5_FZw1kasmdaskndasknn-12ea'
        RANGE_NAME = 'Hoja 1!A1:D'

        service = build('sheets', 'v4', credentials=credentials)
        sheet = service.spreadsheets()

        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])

        normalized_values = [row + [''] * (4 - len(row)) for row in values[1:]]
        df = pd.DataFrame(normalized_values, columns=values[0])
        df['QR'] = df['LINK'].apply(generar_enlace_qr)
        
        messagebox.showinfo("Éxito", "Datos leídos correctamente.")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer los datos: {e}")

# Función para generar los PDFs
def generar_pdf(data, background_path):
    for index, row in data.iterrows():
        nombre = row['NOMBRE']
        qr_link = row['QR']

        archivo_pdf = f"Invitacion_{nombre}.pdf"
        pdf = SimpleDocTemplate(archivo_pdf, pagesize=A4)
        contenido = []

        estilos = getSampleStyleSheet()
        titulo = Paragraph(f"Invitación para {nombre}", estilos['Title'])
        contenido.append(titulo)
        contenido.append(Spacer(1, 220))

        img_path = f"{nombre}_qr.png"
        if qr_link and qr_link.startswith("http"):
            try:
                response = requests.get(qr_link)
                if response.status_code == 200:
                    with open(img_path, "wb") as f:
                        f.write(response.content)
                    img = ReportLabImage(img_path, 2*inch, 2*inch)
                    contenido.append(img)
                else:
                    contenido.append(Paragraph("No se pudo descargar el código QR.", estilos['BodyText']))
            except Exception as e:
                contenido.append(Paragraph(f"Error al descargar el QR: {e}", estilos['BodyText']))
        else:
            contenido.append(Paragraph("No hay código QR disponible.", estilos['BodyText']))

        estilo_nombre = ParagraphStyle(name='EstiloNombre', fontSize=20, textColor='white', alignment=1)
        nombre_paragraph = Paragraph(nombre, estilo_nombre)
        contenido.append(Spacer(1, 30))
        contenido.append(nombre_paragraph)

        pdf.build(contenido, onFirstPage=lambda c, d: draw_background(c, d, background_path),
                  onLaterPages=lambda c, d: draw_background(c, d, background_path))
    
    messagebox.showinfo("Éxito", "PDFs generados correctamente.")

# Función para generar los enlaces QR
def generar_enlace_qr(texto):
    base_url = "https://quickchart.io/qr?text="
    return base_url + urllib.parse.quote(texto)

# Función para dibujar el fondo del PDF
def draw_background(canvas, doc, background_path):
    canvas.drawImage(background_path, 0, 0, width=doc.width + 2 * inch, height=doc.height + 2 * inch)

# Crear la ventana principal con Tkinter
root = tk.Tk()
root.title("Generador de PDFs con QR")

# Botones y acciones
btn_leer_datos = tk.Button(root, text="Leer Datos de Google Sheets", command=lambda: leer_datos())
btn_leer_datos.pack(pady=10)

btn_generar_pdf = tk.Button(root, text="Generar PDFs", command=lambda: generar_pdf(leer_datos(), "1.png"))
btn_generar_pdf.pack(pady=10)

# Ejecutar la aplicación
root.mainloop()
