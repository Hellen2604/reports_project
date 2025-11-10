from flask import Flask, send_file, render_template_string
import os
import reporte_vencimiento  # tu script actual

app = Flask(__name__)

# En Render, no hay escritorio: usa una ruta temporal
RUTA_SALIDA = "/tmp/Reporte_Cartera"

# Crear la carpeta si no existe
os.makedirs(RUTA_SALIDA, exist_ok=True)

@app.route("/")
def index():
    html = """
    <html>
    <head><title>Generación de Reportes</title></head>
    <body style="font-family:sans-serif;">
        <h2>Reporte de Análisis de Vencimiento</h2>
        <form action="/generar" method="get">
            <button type="submit">Generar Reporte</button>
        </form>
    </body>
    </html>
    """
    return render_template_string(html)

@app.route("/generar")
def generar():
    try:
        # Ejecuta la función principal de tu script
        cols, rows = reporte_vencimiento.conectar_y_ejecutar(reporte_vencimiento.CONSULTA)

        # Genera el archivo dentro del contenedor
        out_file = os.path.join(RUTA_SALIDA, "reporte_vencimiento.html")
        reporte_vencimiento.generar_html(cols, rows, out_file)

        # Devuelve el archivo al navegador
        return send_file(out_file)
    except Exception as e:
        return f"<h3>Error al generar reporte:</h3><pre>{str(e)}</pre>"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
