# 📊 Excel Viewer Pro — ¡Hazlo fácil, hazlo chévere!

Una pequeña app web en Flask para subir, visualizar y trabajar con archivos Excel (.xlsx/.xls/.xlsm/.xlsb). Pensada para análisis rápido: previsualiza hojas, evalúa fórmulas estilo Excel y exporta resultados con un solo clic.

![Excel Viewer Pro](https://img.shields.io/badge/Flask-%2335126B?style=flat&logo=flask)
![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Pandas](https://img.shields.io/badge/Pandas-2.x-brightgreen)

---

Por qué es chévere:
- Interfaz moderna con drag & drop.
- Detecta y combina encabezados (cuando hay 2 filas de cabecera).
- Evalúa fórmulas por fila o como total, con soporte en español e inglés.
- Exporta resultados directo a Excel.

## ✨ Novedades rápidas

- Mejor manejo de encabezados "Unnamed" y columnas con caracteres especiales.
- Panel de fórmulas con helpers (operaciones rápidas A ± B, funciones SUM/AVERAGE, SI, CONCATENAR, DIAS.LAB, etc.).

## Requisitos

- Python 3.8+ (recomendado 3.10+)
- Instalar dependencias desde `requirements.txt`.

## Inicio rápido (PowerShell)

Abre PowerShell en la carpeta del proyecto y ejecuta:

```powershell
# Crear y activar entorno virtual (recomendado)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar la app (modo desarrollo)
python .\app.py

# Abrir en el navegador: http://127.0.0.1:5000/
```

Tip: si tienes el launcher `py`, usa `py -3 .\app.py`.

## Cómo usar (en 3 pasos)

1. Arrastra o selecciona tu archivo Excel (.xlsx/.xls).
2. Elige la hoja y usa el panel de fórmulas para crear columnas nuevas o calcular totales.
3. Exporta el resultado a Excel si lo necesitas.

## Endpoints (útiles para integración)

- GET `/` — UI principal.
- POST `/upload` — Subir archivo (multipart/form-data con clave `file`).
- GET `/columns?sheet=<name>` — Lista las columnas detectadas en la hoja.
- POST `/formula` — Evaluar fórmula en memoria y devolver HTML/resultado. Ejemplo de body JSON:

```json
{
	"sheet": "Hoja1",
	"expr": "SUM([Kilos],[Envases])",
	"name": "Total Kilos",
	"mode": "row",
	"format": "number:0.00"
}
```

- POST `/formula/export` — Igual que `/formula` pero descarga un archivo Excel con el resultado.
- POST `/compute` — Aplicar varias fórmulas y/o agregaciones por grupo.

## Ejemplos útiles

- Evaluar porcentaje por fila: `( [Kilos] / [Total] ) * 100` — usar `format: "percent:0.00"` para vista amigable.
- Concatenar: `CONCATENAR([Nombre], " ", [Apellido])`.

## Sugerencias de uso y buenas prácticas

- Usa referencias entre corchetes cuando los nombres de columna tienen espacios: `[Nombre Columna]`.
- Para archivos grandes, abre solo las hojas que necesites o considera preprocesar con scripts offline.

## Docker (opcional)

Si quieres ponerlo en un contenedor rápido, puedes usar un Dockerfile simple (ejemplo mínimo):

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 5000
ENV FLASK_APP=app.py
CMD ["python", "app.py"]
```

## Troubleshooting rápido

- Error leyendo `.xls`/`.xlsb`: asegúrate de tener `xlrd` para `.xls` o instala `pyxlsb` si trabajas con `.xlsb`.
- Problemas de dependencia: crea un entorno limpio y reinstala usando el `requirements.txt` provisto.

## Contribuye

Si quieres mejorar la app: crea un issue con tu idea o un PR con pruebas mínimas. Sugiero añadir un archivo `LICENSE` (p. ej. MIT) si quieres compartir libremente.

## Créditos

Desarrollado por Jorge Aparicio — diseñado para agilizar análisis rápidos de Excel.

---

¿Quieres que deje el README aún más visual (con capturas pequeñas, GIF o ejemplos de Excel)? Puedo añadirlo con un archivo `assets/` y una pequeña demo.

