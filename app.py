from flask import Flask, render_template, request, jsonify
import pandas as pd
import io
import traceback

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

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
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({"error": "Solo se permiten archivos Excel (.xlsx, .xls)"}), 400

        print(f"📁 Procesando archivo: {file.filename}")
        
        # Leer Excel directamente desde la memoria
        file_bytes = io.BytesIO(file.read())
        
        print("🔍 Leyendo archivo Excel...")
        
        # Leer con manejo de errores específico
        try:
            excel_data = pd.read_excel(file_bytes, sheet_name=None, engine='openpyxl')
        except Exception as e:
            print(f"❌ Error con openpyxl: {e}")
            # Intentar con otro engine
            file_bytes.seek(0)  # Resetear el buffer
            excel_data = pd.read_excel(file_bytes, sheet_name=None, engine='xlrd')
        
        print(f"✅ Excel leído correctamente. Hojas: {list(excel_data.keys())}")
        
        # Convertir a HTML
        html_sheets = {}
        for sheet_name, df in excel_data.items():
            print(f"📊 Procesando hoja: {sheet_name} - Forma: {df.shape}")
            
            # Limpiar NaN values
            df_clean = df.fillna('')
            
            # Verificar nombres de columnas
            print(f"   Columnas: {list(df_clean.columns)}")
            
            html_sheets[sheet_name] = df_clean.to_html(
                classes="table table-bordered table-striped", 
                index=False,
                escape=False
            )

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