"""Generador de ejemplo para Convertidor/Convertidos_excel
Ejecuta este script para crear `sample.xlsx` con datos ficticios que puedes subir a la app.
Requiere pandas y openpyxl (ya están en requirements.txt).

Uso:
    python scripts\generate_sample.py

Esto generará `sample.xlsx` en la raíz del proyecto.
"""
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "sample.xlsx"

data = {
    "Nombre": ["Ana", "Luis", "María", "Carlos", "Sofía", "Jorge", "Lucía", "Diego"],
    "Apellido": ["Pérez", "Gómez", "Rodríguez", "López", "Hernández", "Martínez", "Duarte", "Vargas"],
    "Kilos": [12.5, 8.0, 15.2, 7.8, 9.0, 11.1, 6.4, 20.0],
    "Envases": [2, 1, 3, 1, 2, 2, 1, 4],
}

df = pd.DataFrame(data)
# Añadir una columna calculada para ejemplo
df['Total'] = df['Kilos'] * 1.0

if __name__ == '__main__':
    try:
        df.to_excel(OUT, index=False, engine='openpyxl')
        print(f"✔ Archivo de ejemplo creado en: {OUT}")
    except Exception as e:
        print("Error creando el archivo de ejemplo:", e)
