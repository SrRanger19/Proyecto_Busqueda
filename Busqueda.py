import openpyxl
import tkinter as tk
from tkinter import ttk

def buscar_correos_por_dominio(nombre_archivo, dominio):
    # Limpiar resultados anteriores
    for i in tree.get_children():
        tree.delete(i)

    try:
        libro = openpyxl.load_workbook(nombre_archivo)
        hoja_activa = libro.active
        for fila in hoja_activa.iter_rows(min_row=2, values_only=True):  # Comenzamos desde la segunda fila
            if isinstance(fila[2], str) and dominio in fila[2]:  # Verificamos si el dominio está en la columna de correo
                tree.insert("", tk.END, values=fila)
    except FileNotFoundError:
        tree.insert("", tk.END, values=[f"El archivo {nombre_archivo} no fue encontrado."])
    except Exception as e:
        tree.insert("", tk.END, values=[f"Ocurrió un error: {str(e)}"])

def buscar():
    dominio_a_buscar = dominio_entry.get()
    if dominio_a_buscar:
        buscar_correos_por_dominio(nombre_archivo, dominio_a_buscar)
    else:
        for i in tree.get_children():
            tree.delete(i)
        tree.insert("", tk.END, values=["Ingrese un dominio válido."])

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Buscador de Correos por Dominio")

# Configurar el Treeview
columnas = ("FOLIO", "NOMBRE", "CORREO", "TELEFONO")  # Reemplaza con los nombres de tus columnas
tree = ttk.Treeview(ventana, columns=columnas, show="headings")
for columna in columnas:
    tree.heading(columna, text=columna)
    tree.column(columna, anchor=tk.CENTER)

# Widgets
etiqueta = ttk.Label(ventana, text="Ingrese el dominio a buscar:")
dominio_entry = ttk.Entry(ventana)
boton_buscar = ttk.Button(ventana, text="Buscar", command=buscar)

# Colocar widgets en la ventana
etiqueta.grid(row=0, column=0, padx=10, pady=10)
dominio_entry.grid(row=0, column=1, padx=10, pady=10)
boton_buscar.grid(row=0, column=2, padx=10, pady=10)
tree.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

nombre_archivo = 'Datos.xlsx'

# Iniciar el bucle de la interfaz gráfica
ventana.mainloop()
