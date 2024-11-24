import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from datetime import datetime
from PIL import Image, ImageTk
import openpyxl
import os


class AppPreparadurias:
    def __init__(self, root):
        self.root = root
        self.root.title("Inicio - Registro de Estudiantes")
        self.root.geometry("800x450")
        self.root.configure(bg="white")

        # Crear el archivo de Excel si no existe
        self.nombre_archivo = "estudiantes.xlsx"
        self.crear_archivo_excel()
        
        # Matriz de los Horarios
        self.horario = self.generar_horario()
        
        self.preparadores = []
        
        for fila in self.horario:
            self.preparadores.append(fila[1])
        
        # Guardar el preparador y la materia actuales
        self.preparador, self.materia = self.verificar_preparador()

        # Crear un contenedor para la interfaz (permite destruir contenido fácilmente)
        self.contenedor = None
        self.crear_pantalla_inicio()

    def limpiar_contenedor(self):
        """Elimina cualquier contenido previo del contenedor principal."""
        if self.contenedor:
            self.contenedor.destroy()
            
    def generar_horario(self):
        matriz = []  # Crear una lista vacía para almacenar las filas
        with open('horarios.txt', 'r') as file:
            lineas = file.readlines()  # Leer todas las líneas del archivo
            for linea in lineas:
                # Eliminar espacios en blanco y saltos de línea al final de cada línea
                fila = linea.strip().split('-')  # Separar por el guion '-'
                matriz.append(fila)  # Agregar la fila a la matriz
        
        # Mostrar la matriz para ver el resultado (opcional)
        # for fila in matriz:
        #     print(fila)
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

        # Buscar en la matriz self.horario
        for fila in self.horario:
            dia, nombre_preparador, materia, hora_inicio, hora_fin = fila
            
            # Comparar día y rango de horas
            if dia == dia_semana_actual_es and hora_inicio <= hora_actual <= hora_fin:
                print(f"Preparador actual: {nombre_preparador} para {materia} de {hora_inicio} a {hora_fin}")
                return nombre_preparador, materia

        print("No hay preparador asignado en este horario.")
        return "No asignado", "No asignado"   

    def crear_pantalla_inicio(self):
        """Crea la ventana de inicio con un header, título y footer."""
        # Limpiar cualquier contenido previo
        self.limpiar_contenedor()
        
        self.root.title("Inicio - Registro de Estudiantes")

        # Crear un nuevo contenedor
        self.contenedor = tk.Frame(self.root, bg="white")
        self.contenedor.pack(fill="both", expand=True)

        # Header
        header = tk.Frame(self.contenedor, bg="#1e293b", height=70, padx=20, pady=15)
        header.pack(fill="x", padx=25)

        # Título dentro del header (lado izquierdo)
        titulo_header = tk.Label(
            header,
            text="Sistema de Registro de Estudiantes",
            bg="#1e293b",
            fg="white",
            font=("Roboto", 16, "bold"),
            anchor="w"
        )
        titulo_header.pack(side="left", padx=10)

        # Logo dentro del header (lado derecho)
        try:
            # Cargar la imagen del logo
            logo_image = Image.open("LogoUnetBlanco.png")  # Reemplaza con la ruta de tu logo
            logo_image = logo_image.resize((50, 50), Image.LANCZOS)  # Redimensionar el logo
            logo_photo = ImageTk.PhotoImage(logo_image)

            # Crear un Label con la imagen del logo
            logo_label = tk.Label(header, image=logo_photo, bg="#1e293b")
            logo_label.image = logo_photo  # Referencia para evitar que se elimine por el recolector de basura
            logo_label.pack(side="right", padx=10)
        except Exception as e:
            print(f"Error al cargar el logo: {e}")

        # Título principal
        titulo = tk.Label(
            self.contenedor,
            text="Bienvenido",
            bg="#3182ce",
            fg="white",
            font=("Roboto", 20, "bold"),
            pady=40,
        )
        titulo.pack(fill="x", padx=25)
        
        # Tabla de Horarios (Lunes a Viernes)
        self.crear_tabla_horario()

        # Botón para ir al formulario
        boton_formulario = tk.Button(
            self.contenedor,
            text="Registrar estudiante",
            command=self.abrir_formulario,
            bg="#3182ce",
            fg="white",
            font=("Arial", 14),
            padx=20,
            pady=10,
            borderwidth=0,
        )
        boton_formulario.pack()

        # Footer
        footer = tk.Frame(self.contenedor, bg="#1e293b", height=50, padx=20, pady=10)
        footer.pack(side="bottom", fill="x", padx=25)
        tk.Label(
            footer,
            text="© Gustavo Morillo & Victoria Ballesteros",
            bg="#1e293b",
            fg="white",
            font=("Arial", 10),
            anchor="e",
        ).pack(fill="x", padx=10, pady=5)

    def crear_tabla_horario(self):
        # Crear el Frame para la tabla (con un fondo blanco)
        tabla_frame = tk.Frame(self.contenedor, bg="white")
        tabla_frame.pack(pady=20)

        # Crear un estilo para el Treeview
        style = ttk.Style()

        # Configurar el estilo del encabezado
        style.configure(
            "Custom.Treeview.Heading",  # Nombre del estilo
            font=("Arial", 12),  # Fuente
            background="#3182ce",       # Color de fondo
            foreground="white",         # Color del texto
            borderwidth=1,              # Ancho del borde
            relief="flat",              # Sin relieve
        )
        # Configurar el estilo de las filas
        style.configure(
            "Custom.Treeview", 
            font=("Arial", 11), 
            rowheight=35,
            background="white", 
            fieldbackground="white",  # Fondo del campo editable
        )
        
        style.map(
            "Custom.Treeview.Heading",
            background=[("active", "#3182ce")],  # Mantener el mismo color
            foreground=[("active", "white")],
        )

        # Crear el Treeview (tabla)
        columns = ('Hora', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes')
        tabla = ttk.Treeview(tabla_frame, columns=columns, show='headings', style="Custom.Treeview", height=2)

        # Configurar los encabezados de la tabla
        tabla.heading('Hora', text='Hora')
        tabla.heading('Lunes', text='Lunes')
        tabla.heading('Martes', text='Martes')
        tabla.heading('Miércoles', text='Miércoles')
        tabla.heading('Jueves', text='Jueves')
        tabla.heading('Viernes', text='Viernes')

        # Configurar el ancho de las columnas
        tabla.column('Hora', width=145, anchor="center")
        tabla.column('Lunes', width=120, anchor="center")
        tabla.column('Martes', width=120, anchor="center")
        tabla.column('Miércoles', width=120, anchor="center")
        tabla.column('Jueves', width=120, anchor="center")
        tabla.column('Viernes', width=120, anchor="center")

        # Insertar las filas con los nombres de los preparadores
        tabla.insert('', 'end', values=('8:00 AM - 10:00 AM', self.preparadores[0], self.preparadores[2], self.preparadores[4], self.preparadores[6], self.preparadores[8]))
        tabla.insert('', 'end', values=('10:00 AM - 12:00 PM', self.preparadores[1], self.preparadores[3], self.preparadores[5], self.preparadores[7], self.preparadores[9]))

        # Colocar la tabla en el Frame
        tabla.pack()

    def abrir_formulario(self):
        """Abre la ventana con el formulario de registro."""
        # Limpiar cualquier contenido previo
        self.limpiar_contenedor()

        # Crear un nuevo contenedor
        self.contenedor = tk.Frame(self.root, bg="white")
        self.contenedor.pack(fill="both", expand=True)

        self.root.title("Formulario - Registro de Estudiantes")
        # Header
        header = tk.Frame(self.contenedor, bg="#1e293b", height=70, padx=25, pady=15)
        header.pack(fill="x", padx=25)
        tk.Label(
            header,
            text="Sistema de Registro de Estudiantes",
            bg="#1e293b",
            fg="white",
            font=("Arial", 16, "bold"),
        ).pack(side="left")

        # Contenedor para el logo y el botón
        derecha_frame = tk.Frame(header, bg="#1e293b")
        derecha_frame.pack(side="right")

        # Logo dentro del header (lado derecho)
        try:
            # Cargar la imagen del logo
            logo_image = Image.open("LogoUnetBlanco.png")  # Reemplaza con la ruta de tu logo
            logo_image = logo_image.resize((50, 50), Image.LANCZOS)  # Redimensionar el logo
            logo_photo = ImageTk.PhotoImage(logo_image)

            # Crear un Label con la imagen del logo
            logo_label = tk.Label(derecha_frame, image=logo_photo, bg="#1e293b")
            logo_label.image = logo_photo  # Referencia para evitar que se elimine por el recolector de basura
            logo_label.pack(side="right")
        except Exception as e:
            print(f"Error al cargar el logo: {e}")
            
        # Botón para volver al inicio
        boton_volver = tk.Button(
            derecha_frame,
            text="Volver",
            command=self.crear_pantalla_inicio,
            bg="#3182ce",
            fg="white",
            font=("Arial", 12),
            padx=15,
            pady=10,
            borderwidth=0,
        )
        boton_volver.pack(side="right", padx=10)  # Espacio entre el botón y el logo
        
        # Título del formulario
        titulo = tk.Label(
            self.contenedor,
            text="Complete los campos presentados a continuación",
            bg="#3182ce",
            fg="white",
            font=("Arial", 20, "bold"),
            pady=40,
        )
        titulo.pack(fill="x", padx=25)

        # Contenedor principal
        frame = tk.Frame(self.contenedor, bg="white", padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        # Configuración del grid para centrar elementos
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        # Etiqueta y entrada para Nombre y Apellido
        tk.Label(frame, text="Nombre y Apellido:", bg="white", font=("Arial", 12)).grid(
            row=0, column=0, sticky="e", pady=5
        )
        self.nombre_entry = tk.Entry(frame, width=30, font=("Arial", 12))
        self.nombre_entry.grid(row=0, column=1, pady=5)

        # Etiqueta y entrada para Cédula
        tk.Label(frame, text="Cédula:", bg="white", font=("Arial", 12)).grid(
            row=1, column=0, sticky="e", pady=5
        )
        self.cedula_entry = tk.Entry(frame, width=30, font=("Arial", 12))
        self.cedula_entry.grid(row=1, column=1, pady=5)
        
        # Etiqueta y entrada para Sección
        tk.Label(frame, text="Sección:", bg="white", font=("Arial", 12)).grid(
            row=2, column=0, sticky="e", pady=5
        )
        self.seccion_entry = tk.Entry(frame, width=30, font=("Arial", 12))
        self.seccion_entry.grid(row=2, column=1, pady=5)

        # Botón para registrar
        registrar_btn = tk.Button(
            frame,
            text="Registrar",
            command=self.registrar_estudiante,
            bg="#3182ce",
            fg="white",
            font=("Arial", 12),
            borderwidth=0,
            pady=10,
            padx=15
        )
        registrar_btn.grid(row=3, column=0, columnspan=2, pady=(15, 0))

        # Footer
        footer = tk.Frame(self.contenedor, bg="#1e293b", height=50, padx=20, pady=10)
        footer.pack(side="bottom", fill="x", padx=25)
        tk.Label(
            footer,
            text="© Gustavo Morillo y Victoria Ballesteros",
            bg="#1e293b",
            fg="white",
            font=("Arial", 10),
            anchor="e",
        ).pack(fill="x", padx=10, pady=5)

    def crear_archivo_excel(self):
        """Crea el archivo de Excel si no existe."""
        if not os.path.exists(self.nombre_archivo):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Estudiantes"
            ws.append(["Nombre y Apellido", "Cédula", "Materia", "Sección", "Preparador"])
            wb.save(self.nombre_archivo)

    def registrar_estudiante(self):
        """Registra al estudiante en el archivo Excel."""
        nombre = self.nombre_entry.get().strip()
        cedula = self.cedula_entry.get().strip()
        seccion = self.seccion_entry.get().strip()

        if not nombre or not cedula or not seccion:
            messagebox.showerror("Error", "Todos los campos deben estar llenos.")
            return
        
        try:
            seccion = int(seccion)
            if seccion <= 0 or seccion > 17:
                messagebox.showerror("Error", "La sección debe ser un número entre 1 y 17.")
                return
        except ValueError:
            messagebox.showerror("Error", "La sección debe ser un número.")
            return

        try:
            # Guardar en el archivo Excel
            wb = openpyxl.load_workbook(self.nombre_archivo)
            ws = wb.active
            self.preparador, self.materia = self.verificar_preparador()
            ws.append([nombre, cedula, self.materia, seccion, self.preparador])
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
