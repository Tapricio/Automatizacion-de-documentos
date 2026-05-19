import os
from tkinter import Tk, filedialog
import win32com.client

# ---------------------------------
# Seleccionar carpeta origen
# ---------------------------------
root = Tk()
root.withdraw()
root.attributes('-topmost', True)

carpeta_origen = filedialog.askdirectory(
    title="Selecciona carpeta con DOCX"
)

if not carpeta_origen:
    print("No seleccionaste carpeta origen")
    exit()

# ---------------------------------
# Seleccionar carpeta destino
# ---------------------------------
carpeta_destino_base = filedialog.askdirectory(
    title="Selecciona carpeta destino PDFs"
)

if not carpeta_destino_base:
    print("No seleccionaste carpeta destino")
    exit()

# ---------------------------------
# Normalizar rutas
# ---------------------------------
carpeta_origen = os.path.abspath(
    carpeta_origen
)

carpeta_destino_base = os.path.abspath(
    carpeta_destino_base
)

# ---------------------------------
# Crear carpeta PDFs
# ---------------------------------
nombre_base = (
    os.path.basename(carpeta_origen)
    + " pdf"
)

carpeta_pdf = os.path.join(
    carpeta_destino_base,
    nombre_base
)

contador = 1
base_original = carpeta_pdf

while os.path.exists(carpeta_pdf):

    carpeta_pdf = (
        f"{base_original} ({contador})"
    )

    contador += 1

os.makedirs(carpeta_pdf)

# ---------------------------------
# Abrir Word UNA sola vez
# ---------------------------------
word = win32com.client.DispatchEx(
    "Word.Application"
)

word.Visible = False

# formato PDF
wdFormatPDF = 17

# ---------------------------------
# Obtener DOCX
# ---------------------------------
archivos = [
    f for f in os.listdir(carpeta_origen)
    if f.lower().endswith(".docx")
]

total = len(archivos)

# ---------------------------------
# Convertir
# ---------------------------------
for i, archivo in enumerate(archivos, start=1):

    ruta_docx = os.path.abspath(
        os.path.join(carpeta_origen, archivo)
    )

    nombre_pdf = (
        os.path.splitext(archivo)[0]
        + ".pdf"
    )

    ruta_pdf = os.path.abspath(
        os.path.join(carpeta_pdf, nombre_pdf)
    )

    try:

        doc = word.Documents.Open(
            ruta_docx,
            ReadOnly=True
        )

        doc.SaveAs(
            ruta_pdf,
            FileFormat=wdFormatPDF
        )

        doc.Close(False)

        print(
            f"[{i}/{total}] Convertido: {nombre_pdf}"
        )

    except Exception as e:

        print(
            f"\nError en:\n{ruta_docx}\n"
        )

        print(e)

# ---------------------------------
# Cerrar Word
# ---------------------------------
word.Quit()

print("\nConversión terminada")
print(f"\nPDFs guardados en:\n{carpeta_pdf}")