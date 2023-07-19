import os
import sqlite3
from tkinter import messagebox
from tkinter import *

fileName=os.path.basename(__file__)
PathName=os.path.dirname(__file__)

def conexionBBDD():

    miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
    miCursor=miConexion.cursor()

    try:
        miCursor.execute('''
            CREATE TABLE REQUISITOS (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            ID_REQ VARCHAR(20),
            DESC_REQ VARCHAR(4000),
            RESP VARCHAR(1),
            COMMENT VARCHAR(4000),
            VEHICULO VARCHAR(20),
            CLIENTE VARCHAR(20),
            FECHA_ACT VARCHAR(20),
            COMENT_INT VARCHAR(4000),
            VERSION INTEGER)''')
        
        messagebox.showinfo("BBDD","LA BASE DE DATOS SE HA CREADO CORRECTAMENTE")
    except:
        messagebox.showinfo("¡ATENCIÓN!","LA BASE DE DATOS YA EXISTE")


conexionBBDD()

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
			FICHERO_ORIGEN varchar(100))

'''