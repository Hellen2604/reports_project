import os
import sys
import subprocess
import pyodbc
import re
from datetime import datetime
from tkinter import messagebox

# --- Funciones auxiliares necesarias ---

def verificar_odbc_drivers():
    """Verifica qué controladores ODBC están disponibles en el sistema"""
    try:
        drivers_instalados = [x for x in pyodbc.drivers() if 'SQL Server' in x]
        return drivers_instalados
    except Exception as e:
        print("Error al verificar drivers ODBC:", str(e))
        return []

def obtener_esquema_desde_consulta(consulta):
    consulta = consulta.upper()
    patron_esquema = r'\b(?:AGANORSA|MATCASA)\.\w+\b'
    esquemas_encontrados = re.findall(patron_esquema, consulta)
    contador_aganorsa = sum(1 for esquema in esquemas_encontrados if 'AGANORSA' in esquema)
    contador_matcasa = sum(1 for esquema in esquemas_encontrados if 'MATCASA' in esquema)
    if contador_aganorsa > contador_matcasa:
        return "AGANORSA"
    elif contador_matcasa > contador_aganorsa:
        return "MATCASA"
    else:
        if 'AGANORSA' in consulta: return "AGANORSA"
        if 'MATCASA' in consulta: return "MATCASA"
        return "MATCASA"

def validar_consulta(consulta):
    consulta = consulta.strip().upper()
    if not consulta: 
        return False, "La consulta no puede estar vacía"
    if not consulta.startswith("SELECT"): 
        return False, "Solo se permiten consultas SELECT"
    palabras_peligrosas = ["DROP", "DELETE", "UPDATE", "INSERT", "ALTER", 
                           "EXEC", "EXECUTE", "TRUNCATE", "CREATE", "SHUTDOWN", "GRANT", "REVOKE"]
    for palabra in palabras_peligrosas:
        if palabra in consulta: 
            return False, f"No se permite la palabra '{palabra}' en la consulta"
    return True, "Consulta válida"

def ejecutar_consulta_sql(consulta, server, database, username, password):
    drivers_disponibles = verificar_odbc_drivers()
    if not drivers_disponibles:
        raise Exception("No se encontró ningún controlador ODBC compatible.")

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
            continue
    raise Exception(f"No se pudo conectar con ninguno de los drivers disponibles: {drivers_disponibles}")

# --- Función principal para el botón "Imprimir/Generar" ---
def boton_generar_consulta(consulta_texto, server, database, username, password):
    """
    Valida la consulta, ejecuta SQL y retorna resultados (columnas, datos, esquema).
    """
    # Validar consulta
    es_valida, mensaje = validar_consulta(consulta_texto)
    if not es_valida:
        messagebox.showerror("Error", mensaje)
        return None, None, None

    # Determinar esquema
    esquema = obtener_esquema_desde_consulta(consulta_texto)

    try:
        columnas, datos = ejecutar_consulta_sql(consulta_texto, server, database, username, password)
        if not datos:
            messagebox.showinfo("Información", "No se encontraron datos.")
            return columnas, datos, esquema
        return columnas, datos, esquema
    except Exception as e:
        messagebox.showerror("Error al ejecutar la consulta", str(e))
        return None, None, esquema

# --- Uso de ejemplo ---
if __name__ == "__main__":
    server = '192.168.100.2'
    database = 'SOFTLAND'
    username = 'reporte'
    password = 'reporte2016'
    
    consulta = "SELECT * FROM MATCASA.DOCUMENTOS_CC"  # ejemplo
    columnas, datos, esquema = boton_generar_consulta(consulta, server, database, username, password)
    if datos:
        print(f"Esquema: {esquema}, Filas obtenidas: {len(datos)}")
