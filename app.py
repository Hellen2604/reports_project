from flask import Flask, send_file, render_template_string
import os
import reporte_vencimiento  # tu script actual

app = Flask(__name__)

# Carpeta de salida: Desktop del usuario actual + carpeta Reporte_Cartera
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
RUTA_SALIDA = os.path.join(desktop_path, 'Reporte_Cartera')

# Crear la carpeta si no existe
os.makedirs(RUTA_SALIDA, exist_ok=True)

@app.route("/")
def index():
    # Página simple con botón para generar reporte
    html = """
    <html>
    <head><title>Generación de Reportes</title></head>
    <body>
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
        out_file = os.path.join(RUTA_SALIDA, "reporte_vencimiento.html")
        reporte_vencimiento.generar_html(cols, rows, out_file)
        return send_file(out_file)
    except Exception as e:
        return f"<h3>Error al generar reporte:</h3><p>{str(e)}</p>"

if __name__ == "__main__":
    # host="0.0.0.0" hace que sea accesible desde toda la LAN
    app.run(host="0.0.0.0", port=5000)
