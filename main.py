import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl
import os


class AppPreparadurias:
    def __init__(self, root):
        self.root = root
        self.root.title("Registro de Estudiantes - Preparadurías")
        self.root.geometry("400x300")
        
        # Crear el archivo de Excel si no existe
        self.nombre_archivo = "estudiantes.xlsx"
        self.crear_archivo_excel()

        # Matriz de los Horarios
        self.horario = self.generar_horario()

        # Guardar el preparador y la materia actuales
        self.preparador, self.materia = self.verificar_preparador()

        # Interfaz gráfica
        self.crear_interfaz()

    def generar_horario(self):
        matriz = []  # Crear una lista vacía para almacenar las filas
        with open('horarios.txt', 'r') as file:
            lineas = file.readlines()  # Leer todas las líneas del archivo
            for linea in lineas:
                # Eliminar espacios en blanco y saltos de línea al final de cada línea
                fila = linea.strip().split('-')  # Separar por el guion '-'
                matriz.append(fila)  # Agregar la fila a la matriz
        
        # Mostrar la matriz para ver el resultado (opcional)
        for fila in matriz:
            print(fila)
        return matriz        

    def verificar_preparador(self):
        # Obtener el día y hora actual
        dia_semana_actual = datetime.now().strftime('%A')
        hora_actual = datetime.now().strftime('%H:%M')

        # Diccionario para traducir los días de la semana
        dias_semana = {
            "Monday": "Lunes",
            "Tuesday": "Martes",
            "Wednesday": "Miercoles",
            "Thursday": "Jueves",
            "Friday": "Viernes",
            "Saturday": "Sabado",
            "Sunday": "Domingo"
        }

        # Traducir el día de la semana
        dia_semana_actual_es = dias_semana.get(dia_semana_actual, "Dia No Existente")

        # Buscar en la matriz `self.horario`
        for fila in self.horario:
            dia, nombre_preparador, materia, hora_inicio, hora_fin = fila
            
            # Comparar día y rango de horas
            if dia == dia_semana_actual_es and hora_inicio <= hora_actual <= hora_fin:
                print(f"Preparador actual: {nombre_preparador} para {materia} de {hora_inicio} a {hora_fin}")
                return nombre_preparador, materia

        print("No hay preparador asignado en este horario.")
        return "No asignado", "No asignado"

    def crear_archivo_excel(self):
        """Crea el archivo de Excel si no existe."""
        if not os.path.exists(self.nombre_archivo):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Estudiantes"
            ws.append(["Nombre y Apellido", "Cédula", "Materia", "Preparador"])
            wb.save(self.nombre_archivo)

    def crear_interfaz(self):
        """Crea los elementos de la interfaz gráfica."""
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        # Etiqueta y entrada para Nombre y Apellido
        tk.Label(frame, text="Nombre y Apellido:").grid(row=0, column=0, sticky="w")
        self.nombre_entry = tk.Entry(frame, width=30)
        self.nombre_entry.grid(row=0, column=1)

        # Etiqueta y entrada para Cédula
        tk.Label(frame, text="Cédula:").grid(row=1, column=0, sticky="w")
        self.cedula_entry = tk.Entry(frame, width=30)
        self.cedula_entry.grid(row=1, column=1)

        # Botón para registrar
        self.registrar_btn = tk.Button(
            frame, text="Registrar", command=self.registrar_estudiante, bg="blue", fg="white"
        )
        self.registrar_btn.grid(row=3, column=0, columnspan=2, pady=10)

    def registrar_estudiante(self):
        """Registra al estudiante en el archivo Excel."""
        nombre = self.nombre_entry.get().strip()
        cedula = self.cedula_entry.get().strip()

        if not nombre or not cedula:
            messagebox.showerror("Error", "Todos los campos deben estar llenos.")
            return

        try:
            # Guardar en el archivo Excel
            wb = openpyxl.load_workbook(self.nombre_archivo)
            ws = wb.active
            ws.append([nombre, cedula, self.materia, self.preparador])
            wb.save(self.nombre_archivo)

            messagebox.showinfo("Éxito", f"Estudiante {nombre} registrado en {self.materia} con el preparador {self.preparador}.")
            self.limpiar_campos()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el registro: {e}")

    def limpiar_campos(self):
        """Limpia los campos del formulario."""
        self.nombre_entry.delete(0, tk.END)
        self.cedula_entry.delete(0, tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = AppPreparadurias(root)
    root.mainloop()
