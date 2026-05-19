import pandas as pd
from docx import Document
from tkinter import Tk, filedialog
import os

# ---------------------------------
# Selección archivos
# ---------------------------------
root = Tk()
root.withdraw()
root.attributes('-topmost', True)

excel_path = filedialog.askopenfilename(
    title="Selecciona Excel",
    filetypes=[("Excel", "*.xlsx")]
)

template_path = filedialog.askopenfilename(
    title="Selecciona Word",
    filetypes=[("Word", "*.docx")]
)

output_folder = filedialog.askdirectory(
    title="Selecciona carpeta destino"
)

filaPrint = 1

# ---------------------------------
# Validaciones
# ---------------------------------
if not excel_path:
    print("No seleccionaste Excel")
    exit()

if not template_path:
    print("No seleccionaste Word")
    exit()

if not output_folder:
    print("No seleccionaste carpeta")
    exit()

# ---------------------------------
# Crear carpeta única
# ---------------------------------
nombre_base = "documentos generados"

carpeta_generados = os.path.join(
    output_folder,
    nombre_base
)

contador = 1

# si existe, crear otra
while os.path.exists(carpeta_generados):

    carpeta_generados = os.path.join(
        output_folder,
        f"{nombre_base} ({contador})"
    )

    contador += 1

# crear carpeta final
os.makedirs(carpeta_generados)

# ---------------------------------
# Leer Excel
# ---------------------------------
df = pd.read_excel(excel_path)

# ---------------------------------
# Mostrar columnas
# ---------------------------------
print("\nColumnas disponibles:\n")

for i, columna in enumerate(df.columns):
    print(f"{i + 1}. {columna}")

# ---------------------------------
# Elegir columnas nombre archivo
# ---------------------------------
seleccion = input(
    "\nEscribe números separados por coma para el nombre del archivo.\nEjemplo: 1,2\n\n> "
)

indices = [
    int(x.strip()) - 1
    for x in seleccion.split(",")
]

columnas_nombre = [
    df.columns[i]
    for i in indices
]

# ---------------------------------
# Texto adicional opcional
# ---------------------------------
texto_adicional = input(
    "\nTexto adicional opcional para el nombre del archivo (ENTER para omitir):\n\n> "
).strip()

print("\nNombre archivo usará:")
print(columnas_nombre)

if texto_adicional:
    print(f"Texto adicional: {texto_adicional}")

# ---------------------------------
# Función reemplazo
# Mantiene formatos Word
# ---------------------------------
def reemplazar_en_parrafo(parrafo, reemplazos):

    for run in parrafo.runs:

        texto_run = run.text

        # placeholders largos primero
        for key, value in sorted(
            reemplazos.items(),
            key=lambda x: len(x[0]),
            reverse=True
        ):

            texto_run = texto_run.replace(
                str(key),
                str(value)
            )

        # mantener formato original
        run.text = texto_run

# ---------------------------------
# Generar documentos
# ---------------------------------
for index, fila in df.iterrows():

    doc = Document(template_path)

    reemplazos = {}

    # crear diccionario reemplazos
    for columna in df.columns:
        reemplazos[str(columna)] = fila[columna]

    # ---------------------------------
    # Párrafos normales
    # ---------------------------------
    for p in doc.paragraphs:
        reemplazar_en_parrafo(
            p,
            reemplazos
        )

    # ---------------------------------
    # Tablas
    # ---------------------------------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_en_parrafo(
                        p,
                        reemplazos
                    )

    # ---------------------------------
    # Construir nombre archivo
    # ---------------------------------
    partes_nombre = []

    for col in columnas_nombre:

        valor = str(fila[col]).strip()

        partes_nombre.append(valor)

    # agregar texto adicional
    if texto_adicional:
        partes_nombre.insert(0, texto_adicional)

    nombre_archivo = " ".join(partes_nombre)

    # limpiar caracteres inválidos
    caracteres_invalidos = r'\/:*?"<>|'

    for c in caracteres_invalidos:
        nombre_archivo = nombre_archivo.replace(c, "")

    nombre_archivo += ".docx"

    # ---------------------------------
    # Ruta final
    # ---------------------------------
    ruta_salida = os.path.join(
        carpeta_generados,
        nombre_archivo
    )

    # ---------------------------------
    # Guardar documento
    # ---------------------------------
    doc.save(ruta_salida)

    print(f"Generado: {filaPrint} | Nombre archivo: {nombre_archivo}")

    filaPrint += 1

print("\nProceso terminado")