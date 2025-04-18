import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import cv2
import threading
import os  # Para verificar si el archivo existe
from openpyxl import load_workbook  # Importar para ajustar columnas

# Diccionario para almacenar los registros de los empleados
registros = {}

# Tarifa por hora para calcular el monto a pagar
tarifa_por_hora = 10.0  # Puedes ajustar este valor según sea necesario

# Función para registrar la entrada
def registrar_entrada():
    id_empleado = entry_id.get().strip()
    if id_empleado in registros:
        messagebox.showwarning("Advertencia", f"El empleado {id_empleado} ya tiene una entrada registrada.")
    else:
        hora_entrada = datetime.now()
        registros[id_empleado] = {'entrada': hora_entrada}
        messagebox.showinfo("Éxito", f"Entrada registrada para el empleado {id_empleado} a las {hora_entrada}")

# Función para registrar la salida
def registrar_salida():
    id_empleado = entry_id.get().strip()
    if id_empleado in registros and 'salida' not in registros[id_empleado]:
        hora_salida = datetime.now()
        registros[id_empleado]['salida'] = hora_salida
        messagebox.showinfo("Éxito", f"Salida registrada para el empleado {id_empleado} a las {hora_salida}")
    else:
        messagebox.showwarning("Advertencia", f"No se encontró una entrada registrada para el empleado {id_empleado} o ya tiene una salida registrada.")

# Función para guardar los registros en un archivo Excel sin sobrescribir
def guardar_en_excel():
    archivo_excel = "registros_empleados.xlsx"  # Nombre del archivo Excel

    # Cargar datos existentes si el archivo ya existe
    if os.path.exists(archivo_excel):
        try:
            df_existente = pd.read_excel(archivo_excel)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo existente: {e}")
            return
    else:
        # Si no existe, crea un DataFrame vacío
        df_existente = pd.DataFrame(columns=["ID Empleado", "Hora de Entrada", "Hora de Salida", "Horas de Trabajo", "Monto a Pagar"])

    # Crear nuevos datos para agregar
    data_nueva = []
    for id_empleado, registro in registros.items():
        entrada = registro['entrada'].strftime("%Y-%m-%d %H:%M:%S")
        salida = registro.get('salida', 'No registrada')
        if salida != 'No registrada':
            salida = salida.strftime("%Y-%m-%d %H:%M:%S")
            # Calcular horas de trabajo si hay entrada y salida
            horas_trabajo = registro['salida'] - registro['entrada']
            horas_trabajo_horas = horas_trabajo.total_seconds() / 3600.0  # Convertir a horas
            horas_trabajo_texto = str(horas_trabajo)  # Convertir a texto para la tabla
            # Calcular monto a pagar
            monto_a_pagar = round(horas_trabajo_horas * tarifa_por_hora, 2)  # Redondear a dos decimales
        else:
            horas_trabajo_texto = 'No disponible'
            monto_a_pagar = 'No disponible'

        data_nueva.append([id_empleado, entrada, salida, horas_trabajo_texto, monto_a_pagar])

    # Crear un DataFrame para los nuevos datos
    df_nuevo = pd.DataFrame(data_nueva, columns=["ID Empleado", "Hora de Entrada", "Hora de Salida", "Horas de Trabajo", "Monto a Pagar"])

    # Combinar el DataFrame existente con los nuevos datos
    df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)

    try:
        # Guardar el DataFrame combinado en el archivo Excel
        df_final.to_excel(archivo_excel, index=False)

        # Ajustar el ancho de las columnas con openpyxl
        ajustar_columnas_excel(archivo_excel)

        messagebox.showinfo("Éxito", f"Registros guardados en '{archivo_excel}'")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

# Función para ajustar el ancho de columnas
def ajustar_columnas_excel(archivo_excel):
    try:
        workbook = load_workbook(archivo_excel)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Obtiene la letra de la columna
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2  # Ajustar el tamaño
            sheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(archivo_excel)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ajustar el tamaño de las columnas: {e}")

# Función para escanear el código QR
def escanear_qr():
    def capturar_qr():
        qrCode = cv2.QRCodeDetector()
        cap = cv2.VideoCapture(0)  # Cambia el número si es necesario (0, 1, 2...)

        if not cap.isOpened():
            messagebox.showerror("Error", "No se puede abrir la cámara")
            return

        while True:
            ret, frame = cap.read()
            if ret:
                ret_qr, decoded_info, points, _ = qrCode.detectAndDecodeMulti(frame)
                if ret_qr:
                    for info in decoded_info:
                        if info:
                            entry_id.delete(0, tk.END)
                            entry_id.insert(0, info)
                            cap.release()
                            cv2.destroyAllWindows()
                            return
                cv2.imshow('Escaneo de QR', frame)

            # Salir del escaneo con la tecla 'ESC'
            if cv2.waitKey(1) & 0xFF == 27:  # Código ASCII para 'ESC'
                break

        cap.release()
        cv2.destroyAllWindows()

    thread = threading.Thread(target=capturar_qr)
    thread.start()

# Función para calcular el monto acumulado
def calcular_acumulado():
    id_empleado = entry_id.get().strip().lower()  # Quitar espacios y usar minúsculas
    archivo_excel = "registros_empleados.xlsx"  # Archivo donde están los registros

    if os.path.exists(archivo_excel):
        try:
            df = pd.read_excel(archivo_excel)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")
            return
    else:
        messagebox.showwarning("Advertencia", "No hay datos guardados. Asegúrate de registrar y guardar primero.")
        return

    # Limpieza y estandarización de los IDs
    df["ID Empleado"] = df["ID Empleado"].astype(str).str.strip().str.lower()

    # Filtrar por el ID del empleado
    registros_empleado = df[df["ID Empleado"] == id_empleado]

    if registros_empleado.empty:
        messagebox.showinfo("Información", f"No se encontraron registros para el empleado {id_empleado}.")
    else:
        # Reemplazar "No disponible" por 0 para realizar la suma correctamente
        registros_empleado["Monto a Pagar"] = pd.to_numeric(
            registros_empleado["Monto a Pagar"], errors="coerce"
        ).fillna(0)

        # Sumar únicamente los montos disponibles
        acumulado = registros_empleado["Monto a Pagar"].sum()
        messagebox.showinfo(
            "Acumulado", f"El monto acumulado para el empleado {id_empleado} es: ${acumulado:.2f}"
        )

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Registro de Entrada y Salida de Empleados")

# Etiqueta y campo de entrada para el ID del empleado
label_id = tk.Label(root, text="ID del Empleado:")
label_id.grid(row=0, column=0, padx=10, pady=10)
entry_id = tk.Entry(root)
entry_id.grid(row=0, column=1, padx=10, pady=10)

# Botón para registrar entrada
btn_entrada = tk.Button(root, text="Registrar Entrada", command=registrar_entrada)
btn_entrada.grid(row=1, column=0, padx=10, pady=10)

# Botón para registrar salida
btn_salida = tk.Button(root, text="Registrar Salida", command=registrar_salida)
btn_salida.grid(row=1, column=1, padx=10, pady=10)

# Botón para guardar en Excel
btn_guardar = tk.Button(root, text="Guardar en Excel", command=guardar_en_excel)
btn_guardar.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Botón para escanear QR
btn_qr = tk.Button(root, text="Escanear QR", command=escanear_qr)
btn_qr.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Botón para calcular acumulado
btn_acumulado = tk.Button(root, text="Acumulado", command=calcular_acumulado)
btn_acumulado.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Iniciar la aplicación
root.mainloop()