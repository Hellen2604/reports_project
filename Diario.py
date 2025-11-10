
import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyodbc
from datetime import datetime
import re
import sys
import subprocess
import traceback

# Configuración de la conexión
server = '192.168.100.2'
database = 'SOFTLAND'
username = 'reporte'
password = 'reporte2016'

def verificar_odbc_drivers():
    """Verifica qué controladores ODBC están disponibles en el sistema"""
    try:
        drivers_disponibles = []
        drivers_comprobados = [
            'ODBC Driver 18 for SQL Server',
            'ODBC Driver 17 for SQL Server',
            'ODBC Driver 13 for SQL Server',
            'SQL Server Native Client 11.0',
            'SQL Server Native Client 10.0',
            'SQL Server'
        ]
        
        # Obtener lista de drivers instalados
        drivers_instalados = [x for x in pyodbc.drivers() if 'SQL Server' in x or 'SQL Server' in x]
        print("Drivers ODBC instalados:", drivers_instalados)
        
        return drivers_instalados
    except Exception as e:
        print("Error al verificar drivers ODBC:", str(e))
        return []
def obtener_esquema_desde_consulta(consulta):
    """
    Analiza la consulta SQL para determinar si hace referencia a AGANORSA o MATCASA
    """
    consulta = consulta.upper()
    
    # Buscar referencias a tablas con esquemas
    patron_esquema = r'\b(?:AGANORSA|MATCASA)\.\w+\b'
    esquemas_encontrados = re.findall(patron_esquema, consulta)
    
    # Contar ocurrencias de cada esquema
    contador_aganorsa = sum(1 for esquema in esquemas_encontrados if 'AGANORSA' in esquema)
    contador_matcasa = sum(1 for esquema in esquemas_encontrados if 'MATCASA' in esquema)
    
    # Determinar el esquema principal basado en qué tablas se referencian
    if contador_aganorsa > contador_matcasa:
        return "AGANORSA"
    elif contador_matcasa > contador_aganorsa:
        return "MATCASA"
    else:
        # Si hay igual número o ninguno, buscar en las condiciones WHERE
        if 'AGANORSA' in consulta:
            return "AGANORSA"
        elif 'MATCASA' in consulta:
            return "MATCASA"
        else:
            # Por defecto, usar MATCASA
            return "MATCASA"
def validar_consulta(consulta):
    consulta = consulta.strip().upper()
    if not consulta: 
        return False, "La consulta no puede estar vacía"
    if not consulta.startswith("SELECT"): 
        return False, "Solo se permiten consultas SELECT"
    
    palabras_peligrosas = ["DROP", "DELETE", "UPDATE", "INSERT", "ALTER", "EXEC", "EXECUTE", 
                          "TRUNCATE", "CREATE", "SHUTDOWN", "GRANT", "REVOKE"]
    
    for palabra in palabras_peligrosas:
        if palabra in consulta: 
            return False, f"No se permite la palabra '{palabra}' en la consulta"
    
    return True, "Consulta válida"
def ejecutar_consulta_sql(consulta):
    # Obtener drivers disponibles
    drivers_disponibles = verificar_odbc_drivers()
    
    if not drivers_disponibles:
        raise Exception("No se encontró ningún controlador ODBC compatible. Por favor instale el ODBC Driver for SQL Server.")
    
    # Intentar con cada driver disponible
    for driver in drivers_disponibles:
        try:
            conn_str = f'DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password};Encrypt=yes;TrustServerCertificate=yes;'
            conn = pyodbc.connect(conn_str, timeout=10)
            cursor = conn.cursor()
            cursor.execute(consulta)
            columnas = [column[0] for column in cursor.description]
            datos = cursor.fetchall()
            conn.close()
            return columnas, datos
        except Exception as e:
            print(f"Error con driver {driver}: {str(e)}")
            # Continuar con el siguiente driver
    
    # Si todos los drivers fallaron
    raise Exception(f"No se pudo conectar con ninguno de los controladores disponibles: {drivers_disponibles}")
def formato_numero(valor):
    try:
        if valor is None: 
            return ""
        if isinstance(valor, str):
            valor = valor.replace(",", "").strip()
        valor = float(valor) if valor else 0.0
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError): 
        return str(valor) if valor else ""
def formato_fecha(fecha):
    if fecha is None: 
        return ""
    if isinstance(fecha, str):
        try: 
            fecha = datetime.strptime(fecha, '%Y-%m-%d')
        except: 
            return fecha
    return fecha.strftime('%d/%m/%Y') if hasattr(fecha, 'strftime') else str(fecha)
def generar_html_asientos(columnas, datos, esquema):
    if not datos: 
        return "<p>No hay datos para mostrar.</p>"
    
    asientos = {}
    for fila in datos:
        asiento = fila[0]
        if asiento not in asientos: 
            asientos[asiento] = []
        asientos[asiento].append(fila)
    
    html_output = ""
    for asiento, filas in asientos.items():
        primera_fila = filas[0]
        fecha, fecha_creacion, notas, usuario = formato_fecha(primera_fila[1]), formato_fecha(primera_fila[2]), primera_fila[3] or "", primera_fila[11] or ""
        
        # Determinar información de la empresa según el esquema
        if esquema == "AGANORSA":
            nombre_empresa = "AGANORSA"
            direccion = "Dir: KM 15.5 CARRETERA NUEVA LEON"
            telefono = "226996487 /22699490"
        else:
            nombre_empresa = "MATCASA"
            direccion = "Dir: KM 15.5 CARRETERA NUEVA LEON"
            telefono = "226996487 /22699490"
        
        html_output += f"""
<div class="document" id="document-{asiento}">
    <div class="page-content">
        <div class="header">
            <div class="company-info">
                <h1>{nombre_empresa}</h1>
                <p>{direccion}</p>
                <p><span style="margin-left: 3em;">800 MTS AL NORTE</span></p>
                <p><span style="margin-left: 3em;">{telefono}</span></p>
            </div>
            <div class="date-info">
                <p>Fecha: <span class="current-date"></span></p>
                <p>Hora: <span class="current-time"></span></p>
                <p>Página: 1</p>
            </div>
        </div>
        <div class="document-details">
            <p><span style="margin-left: 3em;">Asiento: {asiento}</span></p>
            <p><span style="margin-left: 3em;">Fecha: {fecha}</span></p>
            <p><span style="margin-left: 1em;">Fecha Creación: {fecha_creacion}</span></p>
            <p>Última Modificación: <span class="last-modification"></span></p>
            <p>Notas: {notas}</p>
        </div>
        <div class="table-container">
            <table class="accounting-table">
                <thead>
                    <tr>
                        <th>Centro Costo</th><th>Cuenta Contable</th><th>Descripción</th>
                        <th>Débito C$</th><th>Crédito C$</th><th>Débito $</th><th>Crédito $</th>
                    </tr>
                </thead>
                <tbody>
"""
        total_debito_local, total_credito_local, total_debito_dolar, total_credito_dolar = 0, 0, 0, 0
        for fila in filas:
            centro_costo, cuenta_contable, descripcion = fila[4] or "", fila[5] or "", fila[6] or ""
            debito_local, credito_local = float(fila[7] or 0), float(fila[8] or 0)
            debito_dolar, credito_dolar = float(fila[9] or 0), float(fila[10] or 0)
            total_debito_local += debito_local
            total_credito_local += credito_local
            total_debito_dolar += debito_dolar
            total_credito_dolar += credito_dolar
            
            html_output += f"""
                    <tr>
                        <td>{centro_costo}</td><td>{cuenta_contable}</td><td>{descripcion}</td>
                        <td class="number-cell">{formato_numero(debito_local) if debito_local != 0 else ''}</td>
                        <td class="number-cell">{formato_numero(credito_local) if credito_local != 0 else ''}</td>
                        <td class="number-cell">{formato_numero(debito_dolar) if debito_dolar != 0 else ''}</td>
                        <td class="number-cell">{formato_numero(credito_dolar) if credito_dolar != 0 else ''}</td>
                    </tr>
"""
        html_output += f"""
                </tbody>
                <tfoot>
                    <tr class="totals-row">
                        <td colspan="3" style="text-align: right; padding-right: 5px;"><span style="font-weight: bold;">Totales:</span></td>
                        <td class="number-cell"><strong>{formato_numero(total_debito_local)}</strong></td>
                        <td class="number-cell"><strong>{formato_numero(total_credito_local)}</strong></td>
                        <td class="number-cell"><strong>{formato_numero(total_debito_dolar)}</strong></td>
                        <td class="number-cell"><strong>{formato_numero(total_credito_dolar)}</strong></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>
    <div class="signatures-footer">
        <div class="signature-box">
            <span class="signature-label">{usuario}</span><div class="signature-line"></div><span class="signature-label">ELABORADO POR</span>
        </div>
        <div class="signature-box">
            <span class="signature-label">&nbsp;</span><div class="signature-line"></div><span class="signature-label">REVISADO POR</span>
        </div>
        <div class="signature-box">
            <span class="signature-label">&nbsp;</span><div class="signature-line"></div><span class="signature-label">AUTORIZADO POR</span>
        </div>
    </div>
</div>
"""
    return html_output
def guardar_html(html_content, filename="reporte_contable.html", esquema="MATCASA"):
    html_template = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Reporte Contable - {esquema}</title>
    <style>
        body {{
            font-family: 'Arial', monospace; font-size: 12px; line-height: 1.2;
            color: #000; background-color: #f5f5f5; padding: 20px;
        }}
        .document {{
            background-color: white; max-width: 800px; margin: 0 auto;
            padding: 30px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            min-height: 10.9in;
            display: flex; flex-direction: column;
        }}
        .page-content {{ flex-grow: 1; }}
        .header {{ display: flex; justify-content: space-between; margin-bottom: 30px; }}
        .header h1 {{ margin: 0; font-size: 17px; font-weight: normal; }}
        .company-info p {{ margin: 3px 0; font-size: 14px; }}
        .date-info {{ text-align: right; font-size: 15px; }}
        .date-info p {{ margin: 2px 0; }}
        .document-details {{ margin: 20px 0; font-size: 11px; }}
        .document-details p {{ margin: 3px 0; }}
        .accounting-table {{
            width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 11px;
        }}
        .accounting-table thead, .accounting-table tfoot {{ display: table-row-group; }}
        .accounting-table thead tr {{
            border-top: 1px solid #000; border-bottom: 1px solid #000; background-color: #f0f0f0;
        }}
        .accounting-table th {{ padding: 1px; font-weight: bold; text-align: center; }}
        .accounting-table td {{ padding: 5px; text-align: left; }}
        .number-cell {{ text-align: right; }}
        /* --- ESTILO CORREGIDO Y RESTAURADO --- */
        .totals-row td:nth-child(n+2) {{
            border-top: 1px solid #000;
            border-bottom: 1px solid #000;
            font-weight: bold;
        }}
        .signatures-footer {{
            display: flex; justify-content: space-between; width: 100%;
            padding-top: 50px; page-break-inside: avoid;
        }}
        .signature-box {{ text-align: center; width: 200px; }}
        .signature-line {{ border-top: 1px solid #000; margin: 5px 0; height: 5px; width: 100%; }}
        .signature-label {{ display: block; font-size: 14px; font-weight: normal; }}
        .controls {{ text-align: center; margin-bottom: 20px; }}
        .btn {{ background-color: #007bff; color: white; border: none; padding: 10px 20px; font-size: 14px; cursor: pointer; }}
        @media print {{
            body {{ background-color: white; padding: 0; margin: 0; }}
            .controls, .document > br {{ display: none; }}
            .document {{
                box-shadow: none; padding: 0; margin: 0; max-width: 100%;
                min-height: 95vh; page-break-after: always;
            }}
            .document:last-child {{ page-break-after: avoid; }}
        }}
    </style>
</head>
<body>
    <div class="controls"><button class="btn" onclick="window.print()">Imprimir</button></div>
    {html_content}
    <script>
        function updateDateTime() {{
            const now = new Date();
            const day = String(now.getDate()).padStart(2, '0');
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            const currentDate = `${{day}}/${{month}}/${{year}}`;
            let hours = now.getHours();
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const ampm = hours >= 12 ? 'PM' : 'AM';
            hours = hours % 12 || 12;
            const currentTime = `${{hours}}:${{minutes}} ${{ampm}}`;
            document.querySelectorAll('.current-date').forEach(el => el.textContent = currentDate);
            document.querySelectorAll('.current-time').forEach(el => el.textContent = currentTime);
            document.querySelectorAll('.last-modification').forEach(el => el.textContent = currentDate);
        }}
        window.onload = updateDateTime;
    </script>
</body>
</html>
"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_template)
        return filename
    except Exception as e:
        raise Exception(f"Error al guardar el archivo HTML: {e}")
def imprimir_consulta():
    consulta = t.get("1.0", tk.END).strip()
    es_valida, mensaje = validar_consulta(consulta)
    if not es_valida: 
        messagebox.showerror("Error", mensaje)
        return
    
    # Determinar el esquema desde la consulta
    esquema = obtener_esquema_desde_consulta(consulta)
    
    try:
        columnas, datos = ejecutar_consulta_sql(consulta)
        if not datos: 
            messagebox.showinfo("Información", "No se encontraron datos.")
            return
        
        html_content = generar_html_asientos(columnas, datos, esquema)
        filename = guardar_html(html_content, f"reporte_contable_{esquema}.html", esquema)
        
        try: 
            # Abrir el archivo en el navegador predeterminado
            if sys.platform == "win32":
                os.startfile(filename)
            elif sys.platform == "darwin":
                subprocess.run(["open", filename])
            else:
                subprocess.run(["xdg-open", filename])
        except Exception as e:
            messagebox.showinfo("Éxito", f"Reporte generado: {filename}\n\nError al abrir: {str(e)}")
            
    except Exception as e: 
        messagebox.showerror("Error", f"Error al ejecutar la consulta: {e}")
        
        # Ofrecer opción para instalar el driver si no se encuentra
        if "controlador ODBC" in str(e).lower() or "driver" in str(e).lower():
            respuesta = messagebox.askyesno(
                "Controlador ODBC no encontrado", 
                "¿Desea abrir la página de descarga del controlador ODBC para SQL Server?"
            )
            if respuesta:
                import webbrowser
                webbrowser.open("https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server")
# Función para verificar dependencias al inicio
def verificar_dependencias():
    try:
        import pandas as pd
        import pyodbc
        import tkinter as tk
        print("✓ Todas las dependencias están disponibles")
        return True
    except ImportError as e:
        print(f"✗ Error de dependencia: {e}")
        return False
# Crear la interfaz gráfica
def crear_interfaz():
    global main, t
    
    main = tk.Tk()
    main.title("Sistema de Consultas SQL")
    main.config(bg="#a7d9d1")
    main.geometry("700x400")
    
    t = scrolledtext.ScrolledText(master=main, wrap=tk.WORD)
    t.config(bg="#ffffff", fg="#000000", font=("Courier New", 10))
    t.place(x=20, y=60, width=660, height=277)
    
    button = tk.Button(master=main, text="GENERAR", command=imprimir_consulta)
    button.config(bg="#71768b", fg="#000", font=("Times", 16))
    button.place(x=300, y=350, height=40, width=100)
    
    l = tk.Label(master=main, text="DIGITA LA CONSULTA SQL")
    l.config(bg="#4dceae", fg="#000", font=("Times", 16))
    l.place(x=220, y=14, height=30)
    
    # Verificar drivers al inicio
    try:
        drivers = verificar_odbc_drivers()
        if not drivers:
            messagebox.showwarning(
                "Controladores no encontrados", 
                "No se encontraron controladores ODBC para SQL Server. "
                "La aplicación intentará usar cualquier controlador disponible, pero si falla, "
                "deberá instalar manualmente el ODBC Driver for SQL Server."
            )
        else:
            print(f"Controladores ODBC disponibles: {drivers}")
    except Exception as e:
        print(f"Error al verificar controladores: {e}")
# Punto de entrada principal
if __name__ == "__main__":
    # Verificar dependencias primero
    if not verificar_dependencias():
        print("Instalando dependencias faltantes...")
        try:
            import pip
            pip.main(['install', 'pandas', 'pyodbc'])
        except:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'pyodbc'])
    
    # Crear y ejecutar la interfaz
    crear_interfaz()
    main.mainloop()
