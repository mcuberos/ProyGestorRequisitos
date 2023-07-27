'''
ESTA FUNCIÓN DEBE RECIBIR POR PARÁMETROS EL FICHERO EXCEL A CARGAR EN LA BASE DE DATOS, ADEMÁS DE LAS COLUMNAS Y FILAS DE REFERENCIA 
LA FUNCIÓN DEBE RECORRER DICHO EXCEL E INSERTAR EN BASE DE DATOS TODAS LAS CLÁUSULAS COMPROBANDO SI LA CLÁUSULA YA ESTÁ EN BBDD

#Ejemplo: python LoadDatabase.py 'C:/Users/mcuberos/Desktop/AppGestorRequisitos_old/Python/D0000016800_EEFAE SALOON HVAC P2_Ed-HC_230517.xlsx' '7' 'G'''

import os
import sqlite3
import re
import pandas as pd
from tkinter import messagebox
import sys
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename 
from pandas import ExcelWriter
import pyodbc 
import json
from datetime import date


def leer_configuracion(fichero_config):
    with open(fichero_config, 'r') as archivo:
        configuracion = json.load(archivo)
    return configuracion

PathName=os.path.dirname(__file__)
fileConfig=PathName + "/config.json"

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

def InsertClauseFromTemp(clausula):
    """InsertClauseFromTemp INTENTA INSERTAR UNA CLÁUSULA EN BBDD."""
    global connection_string
  
    # Establecer la conexión
    connection = pyodbc.connect(connection_string)

    cursor = connection.cursor()
    cursor.execute("INSERT INTO T_REQUISITOS VALUES (?, ?, ?, ?,?,'CAF',?,NULL,1,?,?,?)",(clausula[0],clausula[1],clausula[2],clausula[3],clausula[6],date.today().strftime('%Y-%m-%d'),clausula[4],clausula[5],clausula[7]))
    connection.commit()

print("SELECCIONE EL FICHERO ExcelTemp A CARGAR")
fileName=askopenfilename()
df=pd.read_excel(fileName, sheet_name="Requirements Temp",keep_default_na=FALSE)


#RECORRO EL DATAFRAME EN LA COLUMNA colClause, OMITIENDO LAS FILAS QUE SE CORRESPONDEN A TÍTULOS 
for kk in range(len(df)):
        clausula=(df.iloc[kk][0],df.iloc[kk][1],df.iloc[kk][7],df.iloc[kk][8],df.iloc[kk][11],df.iloc[kk][12],df.iloc[kk][13],df.iloc[kk][14])
        InsertClauseFromTemp(clausula)


messagebox.showinfo("PROCESO FINALIZADO","SE HAN INSERTADO LAS CLÁUSULAS EN LA BASE DE DATOS")

