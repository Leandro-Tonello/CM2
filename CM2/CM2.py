import re
import openpyxl
import pandas as pd
import datetime
import random
import string
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from tkinter import ttk
from Clases.Tarjeta import Tarjeta
from Clases.Persona import Persona

# Especifica la ruta del archivo Excel
archivo_excel = r'\Users\Leandro\Desktop\Pendiente Fer\Book[1].xlsx'

# Lee el archivo Excel en un DataFrame de pandas
df = pd.read_excel(archivo_excel, header=None)  # No asume que la primera fila es el encabezado

# Creo el filtro que separa una persona de otra
filas_apellido_nombre = df[df.eq('Apellido y Nombre').any(axis=1)].index

# Inicializa una lista para almacenar todas las instancias de Persona
todas_las_personas = []

# Itera sobre las filas con "Apellido y Nombre" y crea instancias de Persona

for fila in filas_apellido_nombre:
    nombre = df.loc[fila + 1, 1]  # Se asume que los nombres están en la segunda columna (columna 1)
    dni_raw = df.loc[fila + 1, 2]  # Se asume que los DNI están en la tercera columna (columna 2)
  
    # Busquedas desde la celda Celular
    fila_celular = df[df.eq('Celular').any(axis=1)].index[df[df.eq('Celular').any(axis=1)].index > fila].min()
    celular = df.loc[fila_celular, 1]  # Se asume que los números de celular están en la segunda columna (columna 1)
    telefono = df.loc[fila_celular - 1, 1]  
    email = df.loc[fila_celular + 1, 1]  

    # Busquedas desde la celda Nro. Tarjeta CBU
    fila_tarjeta = df[df.eq('Nro. Tarjeta / CBU').any(axis=1)].index[df[df.eq('Nro. Tarjeta / CBU').any(axis=1)].index > fila].min()
    tarjeta1 = df.loc[fila_tarjeta + 1, 1]
    tarjeta2 = df.loc[fila_tarjeta + 2, 1]
    try:
        tarjeta3 = df.loc[fila_tarjeta + 3, 1]  # Accede a la celda que está tres filas debajo de "Tarjeta"
    except KeyError:
        tarjeta3 = None 
    try:
        tarjeta4 = df.loc[fila_tarjeta + 4, 1]  # Accede a la celda que está tres filas debajo de "Tarjeta"
    except KeyError:
        tarjeta4 = None

    if pd.notna(nombre) and pd.notna(dni_raw) and pd.notna(celular) and pd.notna(telefono) and pd.notna(email):
        dni = dni_raw.replace('DNI.', '').strip()
        persona = Persona(nombre, dni, celular, telefono, email, tarjeta1, tarjeta2, tarjeta3, tarjeta4)
        todas_las_personas.append(persona)



# Imprime todas las instancias de Persona
for persona in todas_las_personas:
    print(persona)