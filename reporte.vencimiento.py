#!/usr/bin/env python3
# reporte_vencimiento.py
# Ejecuta una consulta de "Análisis de Vencimiento" y genera un HTML con formato de reporte.
# Requisitos: python, pyodbc, pandas (opcional). Ajusta server/database/credenciales abajo.

import pyodbc
import os
import sys
import subprocess
from datetime import datetime

# --- CONFIGURACIÓN de conexión ---
SERVER = '192.168.100.2'
DATABASE = 'SOFTLAND'
USERNAME = 'reporte'
PASSWORD = 'reporte2016'

# Consulta (la que nos proporcionaste)
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

    -- No vencidos (fecha vencimiento >= hoy)
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

    -- 1-30 días
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

    -- 31-60 días
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

    -- 61-90 días
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

    -- 91-120 días
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

    -- 121-150 días
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

    -- 151-500 días
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

    -- Más de 500 días
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




# --- UTILIDADES ---
def obtener_driver_preferido():
    """Devuelve el primer driver ODBC compatible encontrado en orden de preferencia."""
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
    # fallback: cualquier driver que contenga 'SQL'
    for d in instalados:
        if 'SQL' in d.upper():
            return d
    return None

def conectar_y_ejecutar(consulta):
    driver = obtener_driver_preferido()
    if not driver:
        raise RuntimeError("No se encontró un driver ODBC para SQL Server. Instala 'ODBC Driver for SQL Server'.")
    conn_str = f"DRIVER={{{driver}}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};Encrypt=yes;TrustServerCertificate=yes;"
    conn = pyodbc.connect(conn_str, timeout=30)
    cur = conn.cursor()
    cur.execute(consulta)
    cols = [c[0] for c in cur.description]
    rows = cur.fetchall()
    conn.close()
    return cols, rows

def formato_numero(valor):
    """Formatea número como 1.234.567,89 (puntos miles, coma decimales).
       Devuelve cadena vacía si valor es None o aproximadamente 0."""
    if valor is None:
        return ""
    try:
        v = float(valor)
    except Exception:
        return str(valor)
    # Si es 0 -> no mostrar
    if abs(v) < 0.005:
        return ""
    # Formateo con separación de miles y coma decimal
    s = f"{v:,.2f}"
    # Garantizar puntos para miles y coma para decimales
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

def generar_html(result_cols, result_rows, filename="reporte_vencimiento.html"):
    """Genera HTML del reporte con el formato pedido y fila de totales."""
    # Fecha actual
    ahora = datetime.now()
    fecha_corte = ahora.strftime("%d/%m/%Y")
    hora = ahora.strftime("%H:%M:%S")

    # Cabecera del HTML (estilos adaptados al ejemplo)
    html = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Análisis de Vencimiento - Cuentas por Cobrar</title>
<style>
    body {{ font-family: Arial, sans-serif; background:#fff; color:#000; }}
    .container {{ width: 1250px; margin: 70px auto; }}
    .header {{ display:flex; justify-content:space-between; align-items:flex-start; padding:15px 0; }}
    .company {{ font-size:16px; font-weight:700; }}
    .company .addr {{ font-weight:400; font-size:13px; }}
    .title {{ text-align:center; margin-top:6px; margin-bottom:6px; font-size:18px; font-weight:700; }}
    .subtitle {{ text-align:center; margin-bottom:12px; font-size:13px; }}
    .meta {{ text-align:right; font-size:13px; }}
    .table-wrapper {{ border-top:1px solid #000; margin-top:7px; }}
    table {{ width:100%; border-collapse:separate;border-spacing: 0px 0; font-size:12px; }}
    thead th {{ border-bottom:1px solid #000; padding:6px 4px; font-weight:700; background:#f2f2f2; }}
    tbody td {{ padding:5px 4px; border-bottom: 1px dotted #ccc; }}
    .col-center {{ text-align:center; }}
    .col-right {{ text-align:right; padding-right: 12px;}}
    .small {{ font-size:11px; color:#333; }}
    .footer-note {{ margin-top:12px; font-size:11px; }}
    .total-row {{ font-weight:700; background:#f5f5f5; }}
    @media print {{
        .container {{ width: auto; margin:0; }}
    }}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <div class="company">
            <!-- Logo: colócalo cuando quieras: <img src="cacique.png" alt="logo" style="height:60px"> -->
            MATCASA<br>
            <span class="addr">KM 15.5 CARRETERA NUEVA A LEON<br>800 MTS AL NORTE<br>22696487 / 22699490</span>
        </div>
        <div class="meta">
            Fecha : {fecha_corte}<br>
            Hora : {hora}<br>
            Página : 1
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
                <th style="width:8%;">Cliente</th>
                <th style="width:22%;">Nombre</th>
                <th style="width:10%;">Límite Crédito</th>
                <th style="width:10%;">Saldo Actual</th>
                <th style="width:7%;">No Vencidos</th>
                <th style="width:7%;">1-30 días</th>
                <th style="width:7%;">31-60 días</th>
                <th style="width:7%;">61-90 días</th>
                <th style="width:7%;">91-120 días</th>
                <th style="width:7%;">121-150 días</th>
                <th style="width:7%;">151-500 días</th>
                <th style="width:7%;">Más de 500 días</th>
            </tr>
        </thead>
        <tbody>
"""

    # Inicializar totales (desde SALDO_ACTUAL hasta MAS_500_DIAS -> 9 columnas)
    totales = [0.0] * 9

    # Recorremos filas y construimos el html, acumulando totales
    for row in result_rows:
        # row es un pyodbc.Row; indexamos por posición
        cliente = row[0] or ""
        nombre = row[1] or ""
        limite_credito_raw = row[2]
        limite_credito = formato_numero(limite_credito_raw)

        # columnas numéricas: índices 3..11 (9 columnas)
        nums_raw = []
        nums_fmt = []
        for i in range(3, 12):
            val = row[i]
            # acumular para totales con conversión a float segura (ignorar None)
            try:
                vnum = float(val) if val is not None else 0.0
            except Exception:
                # si no es convertible, tratar como 0
                vnum = 0.0
            nums_raw.append(vnum)
            nums_fmt.append(formato_numero(val))
        # sumar en totales
        for idx, v in enumerate(nums_raw):
            totales[idx] += v

        # construir fila
        html += f"""
            <tr>
                <td class="small">{cliente}</td>
                <td class="small">{nombre}</td>
                <td class="col-right">{limite_credito}</td>
                <td class="col-right">{nums_fmt[0]}</td>
                <td class="col-right">{nums_fmt[1]}</td>
                <td class="col-right">{nums_fmt[2]}</td>
                <td class="col-right">{nums_fmt[3]}</td>
                <td class="col-right">{nums_fmt[4]}</td>
                <td class="col-right">{nums_fmt[5]}</td>
                <td class="col-right">{nums_fmt[6]}</td>
                <td class="col-right">{nums_fmt[7]}</td>
                <td class="col-right">{nums_fmt[8]}</td>
            </tr>
"""

    # Fila de totales: la palabra TOTALES debe quedar debajo de "Límite Crédito",
    # por eso dejamos vacío Cliente y Nombre, y ponemos "TOTALES" en la columna Límite Crédito.
    html += f"""
        <tr class="total-row">
            <td></td>
            <td></td>
            <td class="col-right">TOTALES :</td>
            <td class="col-right">{formato_numero(totales[0])}</td>
            <td class="col-right">{formato_numero(totales[1])}</td>
            <td class="col-right">{formato_numero(totales[2])}</td>
            <td class="col-right">{formato_numero(totales[3])}</td>
            <td class="col-right">{formato_numero(totales[4])}</td>
            <td class="col-right">{formato_numero(totales[5])}</td>
            <td class="col-right">{formato_numero(totales[6])}</td>
            <td class="col-right">{formato_numero(totales[7])}</td>
            <td class="col-right">{formato_numero(totales[8])}</td>
        </tr>
"""

    html += """
        </tbody>
    </table>
    </div>
</div>
</body>
</html>
"""
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)
    return filename

def abrir_html_en_navegador(filepath):
    try:
        if sys.platform.startswith("win"):
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.run(["open", filepath])
        else:
            subprocess.run(["xdg-open", filepath])
    except Exception as e:
        print("No se pudo abrir el navegador automáticamente:", e)
        print("Archivo generado en:", os.path.abspath(filepath))

# --- FLUJO PRINCIPAL ---
def main():
    print("Conectando a la base de datos y ejecutando consulta...")
    try:
        cols, rows = conectar_y_ejecutar(CONSULTA)
    except Exception as exc:
        print("ERROR al ejecutar la consulta:", exc)
        sys.exit(1)

    if not rows:
        print("La consulta no devolvió filas. Se generará un HTML indicando 'No hay datos'.")
        # Generar HTML vacío con mensaje
        empty_html = """<!doctype html><html><body><p>No hay datos para mostrar.</p></body></html>"""
        out = "reporte_vencimiento.html"
        with open(out, "w", encoding="utf-8") as f:
            f.write(empty_html)
        abrir_html_en_navegador(out)
        return

    print(f"Filas obtenidas: {len(rows)}. Generando HTML...")
    out_file = generar_html(cols, rows, "reporte_vencimiento.html")
    print("HTML generado:", out_file)
    abrir_html_en_navegador(out_file)
    print("Listo.")

if __name__ == "__main__":
    main()
