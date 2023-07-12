######## FUNCIÓN PARA CARGAR UN CBC Y AUTORELLENAR LAS CLÁUSULAS EN BASE A LO QUE HAYA EN BASE DE DATOS, GUARDANDO LOS RESULTADOS EN UN EXCEL TEMPORAL


'''
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

def FillClause(clausula):
    """Recorre la bbdd buscando la cláusula para informar el excel original"""

    global dftemp
    PathName=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd
    miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
    miCursor=miConexion.cursor()
    miCursor.execute("SELECT * FROM REQUISITOS")
    listadoRequisitos=miCursor.fetchall()
    ExisteClausula=FALSE
    best_accuracy=0

    for requisito in listadoRequisitos:
        if clausula[1]==requisito[1] or clausula[2]==requisito[2]:
            #EXISTE LA CLAUSULA TAL CUAL EN BASE DE DATOS. SE DEBE TOMAR ESE VALOR
           # print("La cláusula situada en la fila", clausula[0]+1  ,"con identificador ",clausula[4]," ya existe en bbdd con el id ",requisito[0],"-ID_REQ:",requisito[1])
            ExisteClausula=TRUE
            nueva_fila = pd.Series([clausula[1], clausula[2], requisito[1],requisito[2],100,requisito[3],requisito[4],"",""], index=dftemp.columns)
            dftemp = dftemp._append(nueva_fila, ignore_index=True)
        else:
            accuracy=CheckClause(clausula[2],requisito[2])
            if accuracy>best_accuracy:
                nueva_fila = pd.Series([clausula[1], clausula[2], requisito[1],requisito[2],accuracy,requisito[3],requisito[4],"",""], index=dftemp.columns)
                best_accuracy=accuracy
        
    if ExisteClausula==FALSE:
        #SI NO SE HA ENCONTRADO EXACTAMENTE LA CLAUSULA, AÑADO NUEVA_FILA, QUE ES EL REQUISITO QUE HA ENCONTRADO CON MAYOR ACCURACY.
        dftemp = dftemp._append(nueva_fila, ignore_index=True)
    
    return dftemp


def CheckClause(newClause,requirement):
    """ESTA FUNCIÓN COMPRUEBA SI UN REQUISITO NUEVO ES IGUAL A OTRO GUARDADO EN BBDD.
        DEVUELVE UN % DE COINDICENCIA ENTRE LOS DOS REQUISITOS"""
    if newClause==requirement:
        accuracy=100
    else:
        accuracy=0
        longParcial=30
        numTramos=int(len(newClause)/longParcial)+1
        #longParcial=int(len(newClause)/15)
        for aux in range(numTramos):
            cadena=newClause[aux*longParcial:longParcial*(aux+1)]
            if re.search(re.escape(cadena),requirement):
                accuracy=accuracy+99/numTramos
                    
    return accuracy


fileName=askopenfilename()
filaHeader=input("INDIQUE LA FILA DONDE SE ENCUENTRA LA CABECERA DEL CBC ")
colIdReq=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LOS IDs DEL REQUISITO (A,B,C,D,...) ")
colClause=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS DESCRIPCIONES DE LOS REQUISITOS (A,B,C,D,...) ") #para convertir la columna (letra) a número: ord(colClause.lower())-96
colResp=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS RESPUESTAS A LOS REQUISITOS - C/NC/NA (A,B,C,D,...) ")
colComments=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LOS COMENTARIOS (A,B,C,D,...) ")

'''
filaHeader="7"
colClause="g"
colResp="Q"
colComments="r"
colFamReq="a"
'''

df=pd.read_excel(fileName, sheet_name="Requirements",header=int(filaHeader)-1,keep_default_na=FALSE)


#RECORRO EL DATAFRAME EN LA COLUMNA colClause, OMITIENDO LAS FILAS QUE SE CORRESPONDEN A TÍTULOS 
for kk in range(len(df)):
    if (df.iloc[kk][(ord(colResp.lower())-97)]!="" and len(df.iloc[kk][(ord(colClause.lower())-97)])>5):
        #defino la variable clausula como una tupla que contiene: id_requisito, descripción de la clausula, la respuesta, comentarios
        clausula=(kk,df.iloc[kk][(ord(colIdReq.lower())-97)],df.iloc[kk][(ord(colClause.lower())-97)],df.iloc[kk][(ord(colResp.lower())-97)],df.iloc[kk][(ord(colComments.lower())-97)])
        FillClause(clausula)


if len(dftemp)>0:
    Ruta=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd
    excelTemp=(Ruta + "/CBCTemporal.xlsx")
    
    writer = ExcelWriter(excelTemp)
    dftemp.to_excel(writer, 'CBC Temporal', index=False)
    writer.close()

    messagebox.showinfo("EXCEL TEMPORAL CREADO","SE HA CREADO UN EXCEL TEMPORAL CON LAS COINCIDENCIAS DE LAS CLÁUSULAS ENCONTRADAS EN LA RUTA " + Ruta)   


messagebox.showinfo("FIN PROCESO","EJECUCIÓN FINALIZADA CON ÉXITO")