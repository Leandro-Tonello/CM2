from Clases.Tarjeta import Tarjeta

# Definicion de la clase Persona
class Persona:
    def __init__(self, numero_cliente, nombre, telefono, documento, tarjeta):
        self.numero_cliente = numero_cliente
        self.nombre = nombre
        self.telefono = telefono
        self.documento = documento
        self.tarjeta = tarjeta
        self.validateTarjeta = Tarjeta(str(self.tarjeta)).ValidarTarjeta()
        
