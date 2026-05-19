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

filaPrint=1


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
# Crear carpeta "documentos generados"
# ---------------------------------
carpeta_generados = os.path.join(
    output_folder,
    "documentos generados"
)

os.makedirs(
    carpeta_generados,
    exist_ok=True
)

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
# Elegir nombre archivo
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

print("\nNombre archivo usará:")
print(columnas_nombre)

# ---------------------------------
# Función reemplazo
# ---------------------------------
def reemplazar_en_parrafo(parrafo, reemplazos):

    texto_completo = parrafo.text

    # reemplazar placeholders largos primero
    for key, value in sorted(
        reemplazos.items(),
        key=lambda x: len(x[0]),
        reverse=True
    ):

        texto_completo = texto_completo.replace(
            str(key),
            str(value)
        )

    # limpiar runs
    for run in parrafo.runs:
        run.text = ""

    # escribir texto nuevo
    if parrafo.runs:
        parrafo.runs[0].text = texto_completo
    else:
        parrafo.add_run(texto_completo)

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
    filaPrint+=1

print("\nProceso terminado")