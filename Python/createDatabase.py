import os
from tkinter import messagebox
from tkinter import *
import pyodbc 
import json
import sys


def leer_configuracion(fichero_config):
    with open(fichero_config, 'r') as archivo:
        configuracion = json.load(archivo)
    return configuracion

if getattr(sys,'frozen',False):
    PathName=os.path.dirname(sys.executable)
else:
    PathName=os.path.dirname(os.path.abspath(__file__))

#PathName=os.path.dirname(__file__)
fileConfig = PathName + '/config.json'
#fileConfig="C:/AppGestorRequisitos/config.json"

configuracion=leer_configuracion(fileConfig)    

# Acceder a los datos
direccion_bd = configuracion['database']['host']
servidor=direccion_bd.replace("/","\\")
database=configuracion['database']['database_name']
username=configuracion['database']['username']
tablename=configuracion['database']['table_name']
password='Prueb@sManuel23'


# Cadena de conexión
connection_string = f'DRIVER={{SQL Server}};SERVER={servidor};DATABASE={database};UID={username};PWD={password}'

# Establecer la conexión
connection = pyodbc.connect(connection_string)

cursor = connection.cursor()


try:
    cursor.execute('''
        CREATE TABLE T_REQUISITOS (
        ID int IDENTITY(1,1) PRIMARY KEY,
        ID_REQ varchar(20),
        DESC_REQ varchar(8000),
        RESP varchar(2),
        COMMENT varchar(4000),
        VEHICULO varchar(50),
        CLIENTE varchar(50),
        FECHA_ACT date,
        COMENT_INT varchar(4000),
        VERS int,
        PROYECTO_ORIGEN varchar(50),
        FICHERO_ORIGEN varchar(100),
        ENTREGABLE_CBC varchar(20))''')
    connection.commit()
    messagebox.showinfo("BBDD","LA BASE DE DATOS SE HA CREADO CORRECTAMENTE")
except Exception as e:
    messagebox.showinfo("¡ATENCIÓN!","NO SE HA PODIDO CREAR LA TABLA EN LA BASE DE DATOS: " + str(e))

cursor.close()
connection.close()



'''
 CREATE TABLE T_REQUISITOS (
        ID int IDENTITY(1,1) PRIMARY KEY,
        ID_REQ varchar(20),
        DESC_REQ varchar(4000),
        RESP varchar(1),
        COMMENT varchar(4000),
        VEHICULO varchar(50),
        CLIENTE varchar(50),
        FECHA_ACT varchar(20),
        COMENT_INT varchar(4000),
        VERS int,
        PROYECTO_ORIGEN varchar(50),
        FICHERO_ORIGEN varchar(100),
        ENTREGABLE_CBC varchar(20))

'''