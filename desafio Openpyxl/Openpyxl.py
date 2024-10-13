import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime

archivo = "informe_gastos.xlsx"
hoja = "Gastos"

def crear_archivo():
    if not os.path.exists(archivo):
        hoja_de_trabajo = Workbook()
        libro = hoja_de_trabajo.active
        libro.title = hoja
        libro.append(["Fecha", "Descripción", "Monto"])
        hoja_de_trabajo.save(archivo)
        print(f"Archivo '{archivo}' creado con la hoja '{hoja}'.")
    else:
        print(f"El archivo '{archivo}' ya existe.")

def datos_gastos():
    gastos = []
    while True:
        print("Ingrese los detalles del gasto (o ingrese 'finalizar' para terminar):")
        
        while True:
            fecha = input("Fecha (YYYY-MM-DD): ")
            if fecha.lower() == 'finalizar':
                return gastos
            try:
                fecha = datetime.strptime(fecha, "%Y-%m-%d").date()
                break
            except ValueError:
                print("Formato de fecha erroneo. Por favor use YYYY-MM-DD.")
        
        descripcion = input("Descripción: ")
        if descripcion.lower() == 'finalizar':
            return gastos
        
        while True:
            monto = input("Monto: ")
            if monto.lower() == 'finalizar':
                return gastos
            try:
                monto = float(monto)
                if monto < 0:
                    print("El monto no puede ser negativo. Intente otra vez")
                    continue
                break
            except ValueError:
                print("Monto inválido. Por favor, ingrese un número.")
        
        gastos.append({
            "fecha": fecha,
            "descripcion": descripcion,
            "monto": monto
        })
        print("Gasto agregado correctamente.")

def guardar_datos(gastos):
    hoja_de_trabajo = openpyxl.load_workbook(archivo)
    libro = hoja_de_trabajo[hoja]
    
    for gasto in gastos:
        libro.append([gasto["fecha"].strftime("%Y-%m-%d"), gasto["descripcion"], gasto["monto"]])
    
    hoja_de_trabajo.save(archivo)
    print(f"Gastos guardados en '{archivo}' exitosamente.")

def mostrar_resumen(gastos):
    if not gastos:
        print("No se ingresaron gastos.")
        return
    
    total_gastos = sum(gasto["monto"] for gasto in gastos)
    numero_gastos = len(gastos)
    gasto_mas_caro = max(gastos, key=lambda x: x["monto"])
    gasto_mas_barato = min(gastos, key=lambda x: x["monto"])
    
    print("--- Resumen de Gastos ---")
    print(f"Número total de gastos: {numero_gastos}")
    print(f"Gasto más caro: {gasto_mas_caro['fecha']} - {gasto_mas_caro['descripcion']} - Q{gasto_mas_caro['monto']:.2f}")
    print(f"Gasto más barato: {gasto_mas_barato['fecha']} - {gasto_mas_barato['descripcion']} - Q{gasto_mas_barato['monto']:.2f}")
    print(f"Monto total de gastos: Q{total_gastos:.2f}")

def main():
    print("=== Gestión de Informe de Gastos ===")
    crear_archivo()
    gastos = datos_gastos()
    
    if gastos:
        guardar_datos(gastos)
        mostrar_resumen(gastos)
    else:
        print("No se ingresaron gastos para guardar.")
    print("Tarea completada. El informe de gastos se ha guardado en 'informe_gastos.xlsx'.")
    
main()