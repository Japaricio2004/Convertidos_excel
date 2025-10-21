from flask import Flask, render_template, request, jsonify
import pandas as pd
import io
import os
import traceback
from openpyxl import load_workbook

app = Flask(__name__)

# Memoria simple para almacenar el último libro cargado como dict de DataFrames
_LAST_BOOK = {}

@app.route("/")
def index():
    return render_template("index.html")


def _choose_engine(filename: str) -> str:
    """Selecciona el engine correcto según la extensión del archivo."""
    fname = filename.lower()
    if fname.endswith((".xlsx", ".xlsm")):
        return "openpyxl"
    if fname.endswith(".xls"):
        return "xlrd"
    if fname.endswith(".xlsb"):
        return "pyxlsb"
    # Por defecto, intentar openpyxl
    return "openpyxl"


def _promote_header(df: pd.DataFrame) -> pd.DataFrame:
    """
    Promueve la fila más informativa como encabezado para evitar columnas 'Unnamed'.
    Estrategia:
    - Encontrar la fila con mayor cantidad de valores no vacíos.
    - Usarla como encabezado, limpiar espacios, reemplazar NaN por cadena vacía.
    - Remover filas hasta ese encabezado y resetear índice.
    """
    if df.empty:
        return df

    # Reemplazar todo-NaN por vacío para contar mejor
    counts = df.notna().sum(axis=1)
    header_row_idx = int(counts.idxmax())

    # Si la primera fila ya parece ser la mejor, igual normalizamos
    new_header = df.iloc[header_row_idx].fillna("")
    # Convertir a string y limpiar
    new_columns = [str(c).strip() for c in new_header]

    # Si todos están vacíos, no cambiamos nada
    if all(col == "" for col in new_columns):
        return df.fillna("")

    # Crear nuevo dataframe sin la fila de encabezado
    body = df.iloc[header_row_idx + 1 :].reset_index(drop=True)
    body.columns = new_columns

    # Limpiar NaN
    body = body.fillna("")

    return body


@app.route("/compute", methods=["POST"])
def compute():
    try:
        data = request.get_json(silent=True) or {}
        sheet = data.get("sheet")
        formulas = data.get("formulas", [])  # [{"name": "NuevaCol", "expr": "Kilos * 2"}, ...]
        group_by = data.get("group_by")  # ["Exportadora", ...]
        aggregates = data.get("aggregates")  # {"Kilos": "sum", ...}

        if not _LAST_BOOK:
            return jsonify({"error": "Primero sube un archivo Excel."}), 400
        if not sheet or sheet not in _LAST_BOOK:
            return jsonify({"error": "Hoja no encontrada. Especifica una hoja válida."}), 400

        df = _LAST_BOOK[sheet].copy()

        # Aplicar fórmulas de columnas nuevas
        allowed_funcs = {
            'abs': abs,
            'round': round,
        }
        if formulas:
            for item in formulas:
                name = item.get("name")
                expr = item.get("expr")
                if not name or not expr:
                    continue
                try:
                    # pandas.eval para expresiones vectorizadas
                    df[name] = pd.eval(expr, engine='python', parser='pandas', local_dict={**allowed_funcs, **df.to_dict(orient='series')})
                except Exception as e:
                    return jsonify({"error": f"Error en fórmula '{name}': {e}"}), 400

        # Agregaciones por grupo si se solicita
        if group_by and aggregates:
            try:
                result = df.groupby(group_by).agg(aggregates).reset_index()
            except Exception as e:
                return jsonify({"error": f"Error en agregación: {e}"}), 400
        else:
            result = df

        html = result.to_html(classes="table table-bordered table-striped", index=False, escape=False)
        return jsonify({
            "sheet": sheet,
            "columns": list(result.columns),
            "rows": len(result),
            "html": html
        })
    except Exception as e:
        print(f"💥 ERROR /compute: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        # Verificar que se envió un archivo
        if 'file' not in request.files:
            return jsonify({"error": "No se envió ningún archivo"}), 400
        
        file = request.files["file"]
        
        # Verificar que se seleccionó un archivo
        if file.filename == "":
            return jsonify({"error": "No se seleccionó archivo"}), 400

        # Verificar extensión
        allowed_ext = (".xlsx", ".xls", ".xlsm", ".xlsb")
        if not file.filename.lower().endswith(allowed_ext):
            return jsonify({"error": "Solo se permiten archivos Excel (.xlsx, .xls, .xlsm, .xlsb)"}), 400

        print(f"📁 Procesando archivo: {file.filename}")
        
        # Leer Excel directamente desde la memoria
        file_bytes = io.BytesIO(file.read())
        
        print("🔍 Leyendo archivo Excel...")

        engine = _choose_engine(file.filename)
        print(f"   ↳ Engine seleccionado: {engine}")

        # Estrategia: usar la PRIMERA FILA como encabezado exactamente como en Excel, expandiendo merges si aplica.
        html_sheets = {}
        fname = file.filename.lower()
        
        # Limpiar libro previo
        global _LAST_BOOK
        _LAST_BOOK = {}

        if engine == "openpyxl":
            try:
                file_bytes.seek(0)
                wb = load_workbook(filename=file_bytes, data_only=True)
                print(f"✅ Excel leído con openpyxl. Hojas: {wb.sheetnames}")

                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]

                    # Construir mapa de celdas fusionadas
                    merged_ranges = list(ws.merged_cells.ranges)

                    def expand_merged_value(row, col):
                        for mr in merged_ranges:
                            if (row, col) in mr.cells:
                                tl_row, tl_col = mr.min_row, mr.min_col
                                return ws.cell(tl_row, tl_col).value
                        return ws.cell(row, col).value

                    max_col = ws.max_column

                    # Encabezado de dos niveles (fila 1 y fila 2), expandiendo merges
                    top_header = []
                    sub_header = []
                    for c in range(1, max_col + 1):
                        v1 = expand_merged_value(1, c)
                        v2 = expand_merged_value(2, c)
                        top_header.append("" if v1 is None else str(v1).strip())
                        sub_header.append("" if v2 is None else str(v2).strip())

                    # Combinar nombres de columnas segun reglas:
                    # - si sub_header tiene valor => usar "top - sub" si top existe, si no usar solo sub
                    # - si sub_header está vacío => usar top
                    combined_cols = []
                    for t, s in zip(top_header, sub_header):
                        if s and t:
                            combined_cols.append(f"{t} - {s}")
                        elif s and not t:
                            combined_cols.append(s)
                        else:
                            combined_cols.append(t)

                    # Datos desde la fila 3
                    data = []
                    for r in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=max_col, values_only=True):
                        data.append(["" if val is None else val for val in r])

                    df = pd.DataFrame(data, columns=combined_cols)
                    df = df.fillna("")
                    print(f"📊 Procesando hoja: {sheet_name} - Forma: {df.shape}")
                    print(f"   Columnas detectadas (2 filas combinadas): {list(df.columns)}")

                    # Guardar en memoria
                    _LAST_BOOK[sheet_name] = df.copy()

                    html_sheets[sheet_name] = df.to_html(
                        classes="table table-bordered table-striped",
                        index=False,
                        escape=False
                    )
            except Exception as e:
                print(f"❌ Error leyendo con openpyxl personalizado (2 niveles): {e}")
                return jsonify({"error": f"No se pudo leer el archivo .xlsx: {e}"}), 400
        else:
            # Para .xls, .xlsb: usar header=[0,1] y combinar niveles con la misma regla
            try:
                file_bytes.seek(0)
                excel_data = pd.read_excel(file_bytes, sheet_name=None, engine=engine, header=[0,1])
                print(f"✅ Excel leído con {engine}. Hojas: {list(excel_data.keys())}")
                for sheet_name, df in excel_data.items():
                    # df.columns es un MultiIndex (top, sub)
                    new_cols = []
                    for t, s in df.columns:
                        t = "" if pd.isna(t) else str(t).strip()
                        s = "" if pd.isna(s) else str(s).strip()
                        if s and t:
                            new_cols.append(f"{t} - {s}")
                        elif s and not t:
                            new_cols.append(s)
                        else:
                            new_cols.append(t)
                    df.columns = new_cols
                    df = df.fillna("")
                    print(f"📊 Procesando hoja: {sheet_name} - Forma: {df.shape}")
                    print(f"   Columnas detectadas (2 filas combinadas): {list(df.columns)}")

                    # Guardar en memoria
                    _LAST_BOOK[sheet_name] = df.copy()

                    html_sheets[sheet_name] = df.to_html(
                        classes="table table-bordered table-striped",
                        index=False,
                        escape=False
                    )
            except Exception as e:
                print(f"❌ Error leyendo con {engine}: {e}")
                return jsonify({"error": f"No se pudo leer el archivo con el motor '{engine}': {e}"}), 400

        print("🎉 Procesamiento completado exitosamente")
        return jsonify(html_sheets)

        print("🎉 Procesamiento completado exitosamente")
        return jsonify(html_sheets)

    except Exception as e:
        print(f"💥 ERROR: {str(e)}")
        print("TRACEBACK:")
        traceback.print_exc()
        
        return jsonify({"error": f"Error al procesar el archivo: {str(e)}"}), 500

if __name__ == "__main__":
    print("🚀 Servidor Flask iniciado en http://127.0.0.1:5000")
    print("📊 Listo para recibir archivos Excel...")
    app.run(debug=True)