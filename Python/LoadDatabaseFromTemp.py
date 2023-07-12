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

dftemp = pd.DataFrame({'Id_Req': [],
                   'Desc_Req': [],
                   'ID_Req_BBDD': [],
                   'Descr_Req_BBDD':[],
                   '%Coincidencia':[],
                   'Resp_BBDD':[],
                   'Coment_BBDD':[],
                   'Nueva_Respuesta':[],
                   'Nuevos_comentarios':[]})

def InsertClauseFromTemp(clausula):
    """InsertClauseFromTemp INTENTA INSERTAR UNA CLÁUSULA EN BBDD."""
    PathName=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd
    miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
    miCursor=miConexion.cursor()
    miCursor.execute("INSERT INTO REQUISITOS VALUES (NULL, ?, ?, ?, ?,'TRANVÍA','CAF',NULL,NULL,1)",(clausula[0],clausula[1],clausula[2],clausula[3]))
    miConexion.commit()

fileName=askopenfilename()
df=pd.read_excel(fileName, sheet_name="Requirements Temp",keep_default_na=FALSE)


#RECORRO EL DATAFRAME EN LA COLUMNA colClause, OMITIENDO LAS FILAS QUE SE CORRESPONDEN A TÍTULOS 
for kk in range(len(df)):
        clausula=(df.iloc[kk][0],df.iloc[kk][1],df.iloc[kk][7],df.iloc[kk][8])
        InsertClauseFromTemp(clausula)


messagebox.showinfo("PROCESO FINALIZADO","SE HAN INSERTADO LAS CLÁUSULAS EN LA BASE DE DATOS")

