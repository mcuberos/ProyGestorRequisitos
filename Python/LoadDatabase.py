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



def InsertClause(clausula):
    """InsertClause INTENTA INSERTAR UNA CLÁUSULA EN BBDD.
    ANTES DE INSERTARLA DEBE COMBROBAR SI LA CLÁUSULA YA EXISTE LLAMANDO A LA FUNCIÓN CHECKCLAUSE"""
    PathName=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd
    miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
    #miConexion=sqlite3.connect("C:/Users/mcuberos/OneDrive - Internacional Hispacold/Documentos/git/ProyGestorRequisitos/Python/BBDD_CBCs")
    miCursor=miConexion.cursor()
    miCursor.execute("SELECT * FROM REQUISITOS")
    listadoRequisitos=miCursor.fetchall()
    ExisteClausula=FALSE

    for requisito in listadoRequisitos:
        print(requisito[2])
        accuracy=CheckClause(clausula[0],requisito[2]) #La salida de la función debe ser un % de coincidencia (accuracy) entre los dos requisitos. if accuracy<60%, insert requisito en bbdd
        if accuracy>80:
      #  if clausula[0]==requisito[2]:
            ExisteClausula=TRUE
            print("La cláusula ya existe en bbdd con el id ",requisito[0], "con un porcentaje de coincidencia del ",accuracy," por ciento")
            break
        
    if ExisteClausula==FALSE:
        #miCursor.execute("INSERT INTO REQUISITOS VALUES (NULL,'FAM_REQ','DESCRICPION DEL REQUISITO','C','COMENTARIO DE PRUEBA','TRANVIA','CAF',NULL,'COMENTARIO INTERNO',1)")#ejemplo
        miCursor.execute("INSERT INTO REQUISITOS VALUES (NULL, ?, ?, ?, ?,'TRANVÍA','CAF',NULL,NULL,1)",(clausula[3],clausula[0],clausula[1],clausula[2]))
        miConexion.commit()

def CheckClause(newClause,requirement):
    """ESTA FUNCIÓN COMPRUEBA SI UN REQUISITO NUEVO ES IGUAL A OTRO GUARDADO EN BBDD.
        DEVUELVE UN % DE COINDICENCIA ENTRE LOS DOS REQUISITOS"""
    if newClause==requirement:
        accuracy=100
    else:
        accuracy=0
        longParcial=int(len(newClause)/15)
        for aux in range(15):
            cadena=newClause[aux*longParcial:longParcial*(aux+1)]
            if re.search(re.escape(cadena),requirement):
                accuracy=accuracy+100/15
                    
    return accuracy


'''CREO UNA PANTALLA PARA SELECCIONAR LOS PARÁMETROS
fileName=""
filaHeader=""
colClause=""
colResp=""
colComments=""
colFamReq=""


raiz = Tk()

raiz.title("SELECCIONE EL CBC")
miFrame2=Frame(raiz,width=1200,height=60)
miFrame=Frame(raiz, width=1200,height=600)
miFrame.pack()

def abreFichero():
    fichero=filedialog.askopenfile(title="Abrir",filetypes=(("Ficheros de Excel","*.xlsx"),("Ficheros de Texto","*.txt"),("Todos los ficheros","*.*")))
    print (fichero.name)
    fileName=fichero.name
    print (fileName)
    
Button(raiz,text="Abrir Fichero", command=abreFichero).pack()


cuadroCabecera=Entry(miFrame)
cuadroCabecera.grid(row=1,column=1,padx=10,pady=10)
CabeceraLabel = Label(miFrame,text="FILA DE LA CABECERA:")
CabeceraLabel.grid(row=1,column=0,sticky="w",padx=10,pady=10)

cuadroReq=Entry(miFrame)
cuadroReq.grid(row=2,column=1,padx=10,pady=10)
ReqLabel = Label(miFrame,text="COLUMNA DESCR. REQUISITO:")
ReqLabel.grid(row=2,column=0,sticky="w",padx=10,pady=10)

cuadroRespuesta=Entry(miFrame)
cuadroRespuesta.grid(row=3,column=1,padx=10,pady=10)
RespuestaLabel = Label(miFrame,text="COLUMNA RESPUESTA REQUISITO:")
RespuestaLabel.grid(row=3,column=0,sticky="w",padx=10,pady=10)

ComentariosLabel = Label(miFrame,text="COLUMNA COMENTARIOS:")
ComentariosLabel.grid(row=4,column=0,sticky="w",padx=10,pady=10)
CuadroComentarios=Entry(miFrame)
CuadroComentarios.grid(row=4,column=1,padx=10,pady=10)

FamReqLabel = Label(miFrame,text="COLUMNA FAMILIA REQUISITOS:")
FamReqLabel.grid(row=5,column=0,sticky="w",padx=10,pady=10)
CuadroFamReq=Entry(miFrame)
CuadroFamReq.grid(row=5,column=1,padx=10,pady=10)

def codigoBoton():
    print("valor de filename")
    print(fileName)
    print("valor de fila cabecera")
    print(filaHeader)
   
    filaHeader=cuadroCabecera
    colClause=cuadroReq
    colResp=cuadroRespuesta
    colComments=CuadroComentarios
    colFamReq=CuadroFamReq
    print("valor de filename")
    print(fileName)
    print("valor de fila cabecera")
    print(filaHeader)

botonEnvio=Button(raiz,text="ENVIAR",command=codigoBoton) #indicamos que al pulsar al botón llame a la funcion codigoBoton
botonEnvio.pack()

raiz.mainloop()

'''


fileName=askopenfilename()
filaHeader=input("INDIQUE LA FILA DONDE SE ENCUENTRA LA CABECERA DEL CBC")
colClause=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS DESCRIPCIONES DE LOS REQUISITOS (A,B,C,D,...)") #para convertir la columna (letra) a número: ord(colClause.lower())-96
colResp=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS RESPUESTAS A LOS REQUISITOS - C/NC/NA (A,B,C,D,...)")
colComments=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LOS COMENTARIOS (A,B,C,D,...)")
colFamReq=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LA FAMILIA DEL REQUISITO (A,B,C,D,...)")


df=pd.read_excel(fileName, sheet_name="Requirements",header=int(filaHeader)-1)#, on_bad_lines='skip')
#excel="C:\Users\mcuberos\Desktop\AppGestorRequisitos_old\Python\D0000016800_EEFAE SALOON HVAC P2_Ed-HC_230517.xlsx"


'''
#ESTE CÓDIGO RECORRE LA COLUMNA "CLAUSE" Y LA VA IMPRIMIENDO CELDA A CELDA
for etiqueta, contenido in df.items():
    if etiqueta=="Clause":
        j=1
        for elemento in contenido:
            print("INICIO ELEMENTO " + str(j))
            print(elemento)
            j+=1
           # print('Nombre de la etiqueta: ', etiqueta)
           # print('Contenido: ', contenido, sep='\n')
'''

#RECORRO EL DATAFRAME EN LA COLUMNA colClause
for kk in range(len(df)):
    #defino la variable clausula como una tupla que contiene la descripción de la clausula, la respuesta, comentarios y req. familia
    clausula=(df.iloc[kk][(ord(colClause.lower())-97)],df.iloc[kk][(ord(colResp.lower())-97)],df.iloc[kk][(ord(colComments.lower())-97)],df.iloc[kk][(ord(colFamReq.lower())-97)])
    InsertClause(clausula)
