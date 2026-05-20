import re
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

nombre_base = input(
    "\nEscribe el nombre de la carpeta donde guardarás los documentos:\n\n> "
).strip() or "documentos generados"

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
carpeta_generados = os.path.join(
    output_folder,
    nombre_base
)

contador = 1

while os.path.exists(carpeta_generados):

    carpeta_generados = os.path.join(
        output_folder,
        f"{nombre_base} ({contador})"
    )

    contador += 1

os.makedirs(carpeta_generados)

# ---------------------------------
# Leer Excel
# ---------------------------------
df = pd.read_excel(excel_path)

# ---------------------------------
# Crear columna Adjuntos como texto
# ---------------------------------
if "Adjuntos" not in df.columns:
    df["Adjuntos"] = ""

df["Adjuntos"] = df["Adjuntos"].astype(str)

# ---------------------------------
# Columnas plantilla
# ---------------------------------
columnas_plantilla = [
    c for c in df.columns
    if c != "Adjuntos"
]

# ---------------------------------
# Mostrar columnas
# ---------------------------------
print("\nColumnas disponibles:\n")

for i, columna in enumerate(columnas_plantilla):
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
    columnas_plantilla[i]
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
# Funciones formato
# ---------------------------------
def obtener_formato_por_caracter(parrafo):

    formato_por_caracter = []

    for run in parrafo.runs:

        formato = {
            "bold": run.bold,
            "italic": run.italic,
            "underline": run.underline,
            "style": run.style,
            "font_name": run.font.name,
            "font_size": run.font.size,
            "highlight_color": run.font.highlight_color,
            "color_rgb": run.font.color.rgb if run.font.color else None,
            "strike": run.font.strike,
            "subscript": run.font.subscript,
            "superscript": run.font.superscript,
        }

        formato_por_caracter.extend(
            [formato] * len(run.text)
        )

    return formato_por_caracter


def aplicar_formato_run(run, formato):

    run.bold = formato["bold"]
    run.italic = formato["italic"]
    run.underline = formato["underline"]

    if formato["style"] is not None:
        run.style = formato["style"]

    if formato["font_name"] is not None:
        run.font.name = formato["font_name"]

    if formato["font_size"] is not None:
        run.font.size = formato["font_size"]

    run.font.highlight_color = formato["highlight_color"]
    run.font.strike = formato["strike"]
    run.font.subscript = formato["subscript"]
    run.font.superscript = formato["superscript"]

    if formato["color_rgb"] is not None:
        run.font.color.rgb = formato["color_rgb"]


# ---------------------------------
# Reemplazo conservando formato
# ---------------------------------
def reemplazar_en_parrafo(parrafo, reemplazos):

    texto_original = "".join(
        run.text for run in parrafo.runs
    )

    if not texto_original:
        return

    claves_ordenadas = sorted(
        reemplazos.keys(),
        key=lambda x: len(str(x)),
        reverse=True
    )

    patron = re.compile(
        "|".join(
            re.escape(str(k))
            for k in claves_ordenadas
        )
    )

    coincidencias = list(
        patron.finditer(texto_original)
    )

    if not coincidencias:
        return

    formato_por_caracter = obtener_formato_por_caracter(
        parrafo
    )

    segmentos = []

    ultima_pos = 0

    for coincidencia in coincidencias:

        inicio = coincidencia.start()
        fin = coincidencia.end()

        # texto antes placeholder
        if inicio > ultima_pos:

            segmentos.append(
                (
                    texto_original[ultima_pos:inicio],
                    formato_por_caracter[ultima_pos:inicio]
                )
            )

        # valor reemplazo
        valor_reemplazo = str(
            reemplazos[
                coincidencia.group(0)
            ]
        )

        # usar formato exacto del placeholder
        formato_reemplazo = (
            formato_por_caracter[inicio]
            if inicio < len(formato_por_caracter)
            else formato_por_caracter[-1]
        )

        segmentos.append(
            (
                valor_reemplazo,
                formato_reemplazo
            )
        )

        ultima_pos = fin

    # texto restante
    if ultima_pos < len(texto_original):

        segmentos.append(
            (
                texto_original[ultima_pos:],
                formato_por_caracter[ultima_pos:]
            )
        )

    # limpiar runs
    for run in list(parrafo.runs):
        run._element.getparent().remove(
            run._element
        )

    # reconstruir
    for texto_segmento, formato_segmento in segmentos:

        if not texto_segmento:
            continue

        # placeholder reemplazado
        if isinstance(formato_segmento, dict):

            run = parrafo.add_run(
                texto_segmento
            )

            aplicar_formato_run(
                run,
                formato_segmento
            )

            continue

        acumulado = ""
        ultimo_formato = formato_segmento[0]

        for caracter, formato_caracter in zip(
            texto_segmento,
            formato_segmento
        ):

            if (
                formato_caracter != ultimo_formato
                and acumulado
            ):

                run = parrafo.add_run(
                    acumulado
                )

                aplicar_formato_run(
                    run,
                    ultimo_formato
                )

                acumulado = caracter
                ultimo_formato = formato_caracter

            else:
                acumulado += caracter

        if acumulado:

            run = parrafo.add_run(
                acumulado
            )

            aplicar_formato_run(
                run,
                ultimo_formato
            )

# ---------------------------------
# Generar documentos
# ---------------------------------
for index, fila in df.iterrows():

    doc = Document(template_path)

    reemplazos = {}

    # crear reemplazos
    for columna in df.columns:

        valor = fila[columna]

        reemplazos[str(columna)] = (
            ""
            if pd.isna(valor)
            else valor
        )

    # ---------------------------------
    # Párrafos
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
    # Nombre archivo
    # ---------------------------------
    partes_nombre = []

    for col in columnas_nombre:

        valor = str(
            fila[col]
        ).strip()

        partes_nombre.append(
            valor
        )

    if texto_adicional:

        partes_nombre.insert(
            0,
            texto_adicional
        )

    nombre_archivo = " ".join(
        partes_nombre
    )

    # limpiar caracteres inválidos
    caracteres_invalidos = r'\/:*?"<>|'

    for c in caracteres_invalidos:

        nombre_archivo = nombre_archivo.replace(
            c,
            ""
        )

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

    # guardar SOLO nombre archivo
    df.at[index, "Adjuntos"] = nombre_archivo

    print(
        f"Generado: {filaPrint} | Nombre archivo: {nombre_archivo}"
    )

    filaPrint += 1

# ---------------------------------
# Guardar Excel actualizado
# ---------------------------------
df.to_excel(
    excel_path,
    index=False
)

print("\nProceso terminado")