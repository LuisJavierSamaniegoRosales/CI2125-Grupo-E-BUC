import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import cv2
import threading

# Diccionario para almacenar los registros de los empleados
registros = {}

# Función para registrar la entrada
def registrar_entrada():
    id_empleado = entry_id.get()
    if id_empleado in registros:
        messagebox.showwarning("Advertencia", f"El empleado {id_empleado} ya tiene una entrada registrada.")
    else:
        hora_entrada = datetime.now()
        registros[id_empleado] = {'entrada': hora_entrada}
        messagebox.showinfo("Éxito", f"Entrada registrada para el empleado {id_empleado} a las {hora_entrada}")

# Función para registrar la salida
def registrar_salida():
    id_empleado = entry_id.get()
    if id_empleado in registros and 'salida' not in registros[id_empleado]:
        hora_salida = datetime.now()
        registros[id_empleado]['salida'] = hora_salida
        messagebox.showinfo("Éxito", f"Salida registrada para el empleado {id_empleado} a las {hora_salida}")
    else:
        messagebox.showwarning("Advertencia", f"No se encontró una entrada registrada para el empleado {id_empleado} o ya tiene una salida registrada.")

# Función para guardar los registros en un archivo Excel
def guardar_en_excel():
    if not registros:
        messagebox.showwarning("Advertencia", "No hay registros para guardar.")
        return

    data = []
    for id_empleado, registro in registros.items():
        entrada = registro['entrada'].strftime("%Y-%m-%d %H:%M:%S")
        salida = registro.get('salida', 'No registrada')
        if salida != 'No registrada':
            salida = salida.strftime("%Y-%m-%d %H:%M:%S")
        data.append([id_empleado, entrada, salida])

    df = pd.DataFrame(data, columns=["ID Empleado", "Hora de Entrada", "Hora de Salida"])

    try:
        df.to_excel("registros_empleados.xlsx", index=False)
        messagebox.showinfo("Éxito", "Registros guardados en 'registros_empleados.xlsx'")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

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

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

    thread = threading.Thread(target=capturar_qr)
    thread.start()

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

# Iniciar la aplicación
root.mainloop()

