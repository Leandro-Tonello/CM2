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


 # Especifica la ruta del archivo Excel
archivo_excel = r'\Users\Leandro\Desktop\Pendiente Fer\Book[1].xlsx'

class Persona:
    contador_personas = 0

    def __init__(self, nombre, dni, celular, telefono, email, t1, t2, t3, t4):
        Persona.contador_personas += 1
        self.numero_persona = Persona.contador_personas
        self.nombre = nombre
        self.dni =  dni.replace('DNI.','').strip()
        self.celular = self.extraer_numeros(celular)
        self.telefono = self.extraer_numeros(telefono)
        self.email = email
        self.t1 = self.extraer_numeros(t1)
        self.t2 = self.extraer_numeros(t2)
        self.t3 = self.extraer_numeros(t3)
        self.t4 = self.extraer_numeros(t4)
     
    def extraer_numeros(self, texto):
        numeros = re.sub(r'\D','', str(texto))
        return numeros

    def __str__(self):
        return f"Nro {self.numero_persona}: {self.nombre}, Dni:{self.dni}, Cel:{self.celular}, Tel:{self.telefono}, Email:{self.email},  T1:{self.t1}, T2:{self.t2}, T3:{self.t3}, T4:{self.t4}"

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

# Muestra el DataFrame
print(df)