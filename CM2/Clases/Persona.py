from asyncio.windows_events import NULL
from Clases.Tarjeta import Tarjeta
import re

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
        self.v1 = Tarjeta(str(self.t1)).ValidarTarjeta()
        self.t2 = self.extraer_numeros(t2)
        self.v2 = Tarjeta(str(self.t2)).ValidarTarjeta()
        self.t3 = self.extraer_numeros(t3)
        self.v3 = Tarjeta(str(self.t3)).ValidarTarjeta()
        self.t4 = self.extraer_numeros(t4)
        self.v4 = Tarjeta(str(self.t4)).ValidarTarjeta()
     
    def extraer_numeros(self, texto):
        numeros = re.sub(r'\D','', str(texto))
        return numeros

    def __str__(self):
        return f"Nro {self.numero_persona}: {self.nombre}, Dni:{self.dni}, Cel:{self.celular}, Tel:{self.telefono}, Email:{self.email},  T1:{self.t1},?:{self.v1}, T2:{self.t2}, ?:{self.v2}, T3:{self.t3}, ?:{self.v3}, T4:{self.t4}, ?:{self.v4}"