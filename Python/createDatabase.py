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
            CREATE TABLE CBC (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            FAM_REQ VARCHAR(20),
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