#!/usr/bin/env python3
# reporte_vencimiento.py
# Genera:
#  - reporte_vencimiento.html
#  - reporte_vencimiento.pdf
#  - reporte_vencimiento.xls

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
RUTA_SALIDA = r"C:\Users\Jose Darce\Desktop\Reporte_Cartera"
os.makedirs(RUTA_SALIDA, exist_ok=True)

# Ruta wkhtmltopdf (modifica si tu instalación es distinta)
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

# --- CONSULTA ---
CONSULTA = """
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
        'SQL Server Native Client 11.0',
        'SQL Server'
    ]
    instalados = pyodbc.drivers()
    for d in preferidos:
        if d in instalados:
            return d
    return instalados[0] if instalados else None

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

def imagen_base64(ruta_imagen):
    with open(ruta_imagen, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def generar_html(result_cols, result_rows, logo_b64, filename="reporte_vencimiento.html"):
    ahora = datetime.now()
    fecha_corte = ahora.strftime("%d/%m/%Y")
    hora = ahora.strftime("%H:%M:%S")

    html = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Análisis de Vencimiento</title>
<style>
body {{ font-family: Arial, Helvetica, sans-serif; color:#000; margin:0; padding:0; }}
.container {{ width:100%; margin:20px auto; }}
.header {{ display:flex; justify-content:space-between; align-items:flex-start; padding:15px 0; }}
.header-left {{ display:flex; align-items:flex-start; }}
.logo-container img {{ height:80px; }}
.company {{ font-size:16px; font-weight:700; }}
.company .addr {{ font-weight:400; font-size:13px; }}
.meta {{ text-align:right; font-size:13px; }}
.title {{ text-align:center; margin:6px 0; font-size:18px; font-weight:700; }}
.subtitle {{ text-align:center; margin-bottom:12px; font-size:13px; }}
.table-wrapper {{ border-top:1px solid #000; margin-top:7px; }}
table {{ width:100%; border-collapse: collapse; font-size:12px; }}
thead th {{ border-bottom:1px solid #000; padding:6px 4px; font-weight:700; background:#f2f2f2; }}
tbody td {{ padding:5px 4px; border-bottom:1px dotted #ccc; }}
.col-right {{ text-align:right; padding-right:12px; }}
.total-row {{ font-weight:700; background:#f5f5f5; }}
</style>
</head>
<body>
<div class="container">
<div class="header">
  <div class="header-left">
    <div class="logo-container">
      <img src="data:image/png;base64,{logo_b64}" alt="Logo Matadero Cacique">
    </div>
    <div class="company">
      MATCASA<br>
      <span class="addr">KM 15.5 CARRETERA NUEVA A LEON<br>800 MTS AL NORTE<br>22696487 / 22699490</span>
    </div>
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
    html+="</tr></tbody></table></div></div></body></html>"

    filepath = os.path.join(RUTA_SALIDA, filename)
    with open(filepath,"w",encoding="utf-8") as f:
        f.write(html)
    return filepath

def generar_pdf(html_path, pdf_path):
    options = {
        'page-size': 'Letter',
        'orientation': 'Landscape',
        'margin-top': '15mm',
        'margin-right': '8mm',
        'margin-bottom': '10mm',
        'margin-left': '8mm',
        'encoding': 'UTF-8',
        'dpi': 300,
        'zoom': 1.0,
    }
    pdfkit.from_file(html_path, pdf_path, configuration=config, options=options)
    print("PDF generado:", pdf_path)

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

def main():
    print("Conectando a la base de datos...")
    try:
        cols, rows = conectar_y_ejecutar(CONSULTA)
    except Exception as e:
        print("Error al ejecutar la consulta:", e)
        sys.exit(1)

    if not rows:
        print("No hay datos para mostrar")
        return

    logo_path = os.path.join(RUTA_SALIDA, "cacique.png")
    logo_b64 = imagen_base64(logo_path) if os.path.exists(logo_path) else ""

    html_path = generar_html(cols, rows, logo_b64)
    print("HTML generado:", html_path)

    pdf_path = os.path.join(RUTA_SALIDA, "reporte_vencimiento.pdf")
    generar_pdf(html_path, pdf_path)

    generar_excel(rows)

if __name__ == "__main__":
    main()
