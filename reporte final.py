#!/usr/bin/env python3
# reporte_vencimiento.py
# Ejecuta consulta de "Análisis de Vencimiento" y genera HTML + PDF + Excel en ruta específica (sin logo)

import pyodbc
import os
import sys
import subprocess
from datetime import datetime
import xlwt
import pdfkit
import base64

# --- CONFIGURACIÓN ---
SERVER = '192.168.100.2'
DATABASE = 'SOFTLAND'
USERNAME = 'reporte'
PASSWORD = 'reporte2016'

# Ruta de salida
RUTA_SALIDA = r"C:\Users\javier\Desktop\Reporte_Cartera"
os.makedirs(RUTA_SALIDA, exist_ok=True)

# Ruta wkhtmltopdf (modifica si tu instalación es distinta)
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

# --- CONSULTA ---
CONSULTA = """
-- tu consulta original aquí
SELECT 
    A.CLIENTE,
    B.NOMBRE,
    B.LIMITE_CREDITO,
    SUM(
        CASE 
            WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                THEN ABS(A.SALDO_LOCAL) 
            ELSE -ABS(A.SALDO_LOCAL) 
        END
    ) AS SALDO_ACTUAL,

    SUM(
        CASE 
            WHEN A.FECHA_VENCE >= GETDATE() THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS NO_VENCIDOS,

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 1 AND 30 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [1_30_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 31 AND 60 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [31_60_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 61 AND 90 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [61_90_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 91 AND 120 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [91_120_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 121 AND 150 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [121_150_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) BETWEEN 151 AND 500 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [151_500_DIAS],

    SUM(
        CASE 
            WHEN DATEDIFF(DAY, A.FECHA_VENCE, GETDATE()) > 500 THEN 
                CASE 
                    WHEN A.TIPO IN ('FAC','INT','L/C','N/D','O/D','RED','B/V','RHP') 
                        THEN ABS(A.SALDO_LOCAL)
                    ELSE -ABS(A.SALDO_LOCAL)
                END
            ELSE 0 
        END
    ) AS [MAS_500_DIAS]

FROM MATCASA.DOCUMENTOS_CC AS A
LEFT JOIN MATCASA.CLIENTE AS B ON A.CLIENTE = B.CLIENTE
WHERE A.SALDO_LOCAL <> 0
GROUP BY A.CLIENTE, B.NOMBRE, B.LIMITE_CREDITO
ORDER BY B.NOMBRE ASC;
"""

# --- FUNCIONES ---
def obtener_driver_preferido():
    preferidos = [
        'ODBC Driver 18 for SQL Server',
        'ODBC Driver 17 for SQL Server',
        'ODBC Driver 13 for SQL Server',
        'SQL Server Native Client 11.0',
        'SQL Server'
    ]
    instalados = pyodbc.drivers()
    for d in preferidos:
        if d in instalados:
            return d
    for d in instalados:
        if 'SQL' in d.upper():
            return d
    return None

def conectar_y_ejecutar(consulta):
    driver = obtener_driver_preferido()
    if not driver:
        raise RuntimeError("No se encontró un driver ODBC para SQL Server.")
    conn_str = f"DRIVER={{{driver}}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};Encrypt=yes;TrustServerCertificate=yes;"
    conn = pyodbc.connect(conn_str, timeout=30)
    cur = conn.cursor()
    cur.execute(consulta)
    cols = [c[0] for c in cur.description]
    rows = cur.fetchall()
    conn.close()
    return cols, rows

def formato_numero(valor):
    if valor is None:
        return ""
    try:
        v = float(valor)
    except Exception:
        return str(valor)
    if abs(v) < 0.005:
        return ""
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s

# --- GENERAR HTML ---
def generar_html(result_cols, result_rows, filename="reporte_vencimiento.html"):
    ahora = datetime.now()
    fecha_corte = ahora.strftime("%d/%m/%Y")
    hora = ahora.strftime("%H:%M:%S")

    html = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Análisis de Vencimiento - Cuentas por Cobrar</title>
<style>
body {{ font-family: Arial, sans-serif; background:#fff; color:#000; margin:0; padding:0; }}
.container {{ width:100%; margin:10px auto; }}
.header {{ display:flex; justify-content:space-between; align-items:flex-start; padding:5px 0; }}
.company {{ font-size:16px; font-weight:700; line-height:1.3; }}
.company .addr {{ font-weight:400; font-size:13px; }}
.meta {{ text-align:right; font-size:13px; }}
.title {{ text-align:center; margin:2px 0; font-size:18px; font-weight:700; }}
.subtitle {{ text-align:center; margin-bottom:8px; font-size:13px; }}
.table-wrapper {{ border-top:1px solid #000; margin-top:4px; }}
table {{ width:100%; border-collapse: separate; border-spacing: 0px 0; font-size:12px; }}
thead th {{ border-bottom:1px solid #000; padding:6px 4px; font-weight:700; background:#f2f2f2; }}
tbody td {{ padding:5px 4px; border-bottom:1px dotted #ccc; }}
.col-right {{ text-align:right; padding-right:12px; }}
.total-row {{ font-weight:700; background:#f5f5f5; }}
button#btnPrint {{ position:fixed; top:10px; right:10px; padding:8px 10px; font-size:14px; z-index:800; }}
@media print {{
    button#btnPrint {{ display:none; }}
    thead {{ display:table-header-group; }}
    tfoot {{ display:table-footer-group; }}
}}
</style>
</head>
<body>
<button id="btnPrint" onclick="window.print()">Imprimir</button>
<div class="container">
<div class="header">
    <div class="company">
        MATCASA<br>
        <span class="addr">KM 15.5 CARRETERA NUEVA A LEON<br>800 MTS AL NORTE<br>22696487 / 22699490</span>
    </div>
    <div class="meta">
        Fecha : {fecha_corte}<br>
        Hora : {hora}
    </div>
</div>
<div class="title">Análisis de Vencimiento en Moneda Local - Cuentas por Cobrar</div>
<div class="subtitle">
    Fecha de corte: {fecha_corte}<br>
    Solo Documentos en Moneda: COR
</div>
<div class="table-wrapper">
<table>
<thead>
<tr>
<th>Cliente</th><th>Nombre</th><th>Límite Crédito</th><th>Saldo Actual</th><th>No Vencidos</th>
<th>1-30 días</th><th>31-60 días</th><th>61-90 días</th><th>91-120 días</th>
<th>121-150 días</th><th>151-500 días</th><th>Más de 500 días</th>
</tr>
</thead>
<tbody>
"""

    totales = [0.0]*9
    for row in result_rows:
        cliente = row[0] or ""
        nombre = row[1] or ""
        limite_credito = formato_numero(row[2])
        nums_raw=[]
        nums_fmt=[]
        for i in range(3,12):
            val = row[i]
            vnum = float(val) if val is not None else 0.0
            nums_raw.append(vnum)
            nums_fmt.append(formato_numero(val))
        for idx,v in enumerate(nums_raw):
            totales[idx]+=v

        html+=f"<tr><td>{cliente}</td><td>{nombre}</td><td class='col-right'>{limite_credito}</td>"
        for n in nums_fmt:
            html+=f"<td class='col-right'>{n}</td>"
        html+="</tr>"

    html+="<tr class='total-row'><td></td><td></td><td class='col-right'>TOTALES :</td>"
    for t in totales:
        html+=f"<td class='col-right'>{formato_numero(t)}</td>"
    html+="</tr>"

    html+="</tbody></table></div></div></body></html>"

    filepath = os.path.join(RUTA_SALIDA, filename)
    with open(filepath,"w",encoding="utf-8") as f:
        f.write(html)
    return filepath, html

# --- GENERAR PDF ---
def generar_pdf(ruta_html, ruta_pdf):
    options = {
        'page-size': 'Letter',
        'margin-top': '10mm',
        'margin-right': '10mm',
        'margin-bottom': '15mm',
        'margin-left': '10mm',
        'encoding': 'UTF-8',
        'footer-right': 'Página [page] de [toPage]',
        'footer-font-size': '9',
        'print-media-type': '',
    }
    try:
        pdfkit.from_file(ruta_html, ruta_pdf, options=options, configuration=config)
        print("PDF generado:", ruta_pdf)
    except Exception as e:
        print(f"Error al generar PDF: {e}")

# --- GENERAR EXCEL ---
def generar_excel(result_rows, filename="reporte_vencimiento.xls"):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Reporte")
    bold = xlwt.easyxf('font: bold 1; align: horiz center')

    ws.write_merge(1,1,0,11,"MATCASA", bold)
    ws.write_merge(2,2,0,11,"KM 15.5 CARRETERA NUEVA A LEON", bold)
    ws.write_merge(3,3,0,11,"800 MTS AL NORTE", bold)
    ws.write_merge(4,4,0,11,"22696487 / 22699490", bold)
    ahora=datetime.now()
    fecha_corte=ahora.strftime("%d/%m/%Y")
    hora=ahora.strftime("%H:%M:%S")
    ws.write_merge(5,5,0,11,f"Fecha : {fecha_corte}   Hora : {hora}", bold)
    ws.write_merge(6,6,0,11,"Análisis de Vencimiento en Moneda Local - Cuentas por Cobrar", bold)
    ws.write_merge(7,7,0,11,f"Fecha de corte: {fecha_corte}", bold)
    ws.write_merge(8,8,0,11,"Solo Documentos en Moneda: COR", bold)

    headers = ["Cliente","Nombre","Límite Crédito","Saldo Actual","No Vencidos",
               "1-30 días","31-60 días","61-90 días","91-120 días","121-150 días","151-500 días","Más de 500 días"]
    for col,h in enumerate(headers):
        ws.write(9,col,h,bold)

    totales=[0.0]*9
    for r_idx,row in enumerate(result_rows,start=10):
        for c_idx,val in enumerate(row):
            ws.write(r_idx,c_idx,val if val is not None else "")
        for i in range(3,12):
            totales[i-3]+=float(row[i]) if row[i] is not None else 0.0

    total_row=len(result_rows)+10
    ws.write(total_row,2,"TOTALES :",bold)
    for i,val in enumerate(totales):
        ws.write(total_row,i+3,val,bold)

    filepath = os.path.join(RUTA_SALIDA, filename)
    wb.save(filepath)
    print("Excel generado:", filepath)

# --- Abrir HTML ---
def abrir_html_en_navegador(filepath):
    try:
        if sys.platform.startswith("win"):
            os.startfile(filepath)
        elif sys.platform=="darwin":
            subprocess.run(["open",filepath])
        else:
            subprocess.run(["xdg-open",filepath])
    except Exception as e:
        print("No se pudo abrir el navegador automáticamente:", e)
        print("Archivo generado en:",os.path.abspath(filepath))

# --- FLUJO PRINCIPAL ---
def main():
    print("Conectando a la base de datos...")
    try:
        cols, rows = conectar_y_ejecutar(CONSULTA)
    except Exception as e:
        print("Error al ejecutar la consulta:", e)
        sys.exit(1)

    if not rows:
        print("No hay datos para mostrar")
        empty_html = "<!doctype html><html><body><p>No hay datos para mostrar.</p></body></html>"
        out = os.path.join(RUTA_SALIDA, "reporte_de_cartera.html")
        with open(out,"w",encoding="utf-8") as f:
            f.write(empty_html)
        abrir_html_en_navegador(out)
        return

    ahora = datetime.now()
    fecha_str = ahora.strftime("%d%m%Y")  # formato DDMMYYYY

    # Nombres de archivos
    html_filename = f"reporte_de_cartera_{fecha_str}.html"
    pdf_filename = f"reporte_de_cartera_{fecha_str}.pdf"
    excel_filename = f"reporte_de_cartera_{fecha_str}.xls"

    html_file, html_content = generar_html(cols, rows, filename=html_filename)
    print("HTML generado:", html_file)
    abrir_html_en_navegador(html_file)

    ruta_pdf = os.path.join(RUTA_SALIDA, pdf_filename)
    generar_pdf(html_file, ruta_pdf)
    generar_excel(rows, filename=excel_filename)


if __name__=="__main__":
    main()
