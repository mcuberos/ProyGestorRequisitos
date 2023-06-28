'''
ESTA FUNCIÓN DEBE RECIBIR POR PARÁMETROS EL FICHERO EXCEL A CARGAR EN LA BASE DE DATOS, además 
LA FUNCIÓN DEBE RECORRER DICHO EXCEL E INSERTAR EN BASE DE DATOS TODAS LAS CLÁUSULAS COMPROBANDO SI LA CLÁUSULA YA ESTÁ EN BBDD

#Ejemplo: python LoadDatabase.py 'C:/Users/mcuberos/Desktop/AppGestorRequisitos_old/Python/D0000016800_EEFAE SALOON HVAC P2_Ed-HC_230517.xlsx' '7' 'G'''

import os
import sqlite3
import pandas as pd
from tkinter import messagebox
import sys
from tkinter import *
from tkinter.filedialog import askopenfilename 

def InsertClause(clausula):
    PathName=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd
    miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
    miCursor=miConexion.cursor()
    miCursor.execute("SELECT * FROM CBCs")
    listadoRequisitos=miCursor.fetchall()
    ExisteClausula=FALSE

    for requisito in listadoRequisitos:
        if clausula[0]==requisito( [1]):
            ExisteClausula=TRUE
        
    if ExisteClausula==FALSE:
        miCursor.execute("INSERT INTO REQUISITOS VALUES (NULL,clausula(3),clausula(0),clausula(1),clausula(2),'TRANVÍA','CAF',NULL,NULL,1)")
        miConexion.commit()

      

def EvaluaRepetida():
    pass
    


#fileName=os.path.basename(__file__)
PathName=os.path.dirname(__file__) #busco la ruta donde se debe encontrar la bbdd

fileName=askopenfilename()
filaHeader=input("INDIQUE LA FILA DONDE SE ENCUENTRA LA CABECERA DEL CBC")
colClause=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS DESCRIPCIONES DE LOS REQUISITOS") #para convertirlo a número: ord(colClause.lower())-96
colResp=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LAS RESPUESTAS A LOS REQUISITOS (C/NC/NA)")
colComments=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LOS COMENTARIOS")
colFamReq=input("INDIQUE LA COLUMNA DONDE SE ENCUENTRAN LA FAMILIA DEL REQUISITO")


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
    #defino la variable clausula como una tupla que contiene la propia clausula, la respuesta, comentarios y req. familia
    clausula=(df.iloc[kk][(ord(colClause.lower())-97)],df.iloc[kk][(ord(colResp.lower())-97)],df.iloc[kk][(ord(colComments.lower())-97)],df.iloc[kk][(ord(colFamReq.lower())-97)])
    #clausula=df.iloc[kk][(ord(colClause.lower())-97)]
    print(df.iloc[kk][(ord(colClause.lower())-97)])
    input("introduzca una tecla para continuar")









'''
for elem in df[0,:]:
    print(elem)

for elemento in range(len(df)):
    print(df.iloc[elemento][int(colClause)-1])

root=Tk()
root.title("CARGA BASE DE DATOS")
print(len(sys.argv))
print(sys.argv[0])
miFrame=Frame(root, width=600, height=600).pack()


miConexion=sqlite3.connect(PathName + "/BBDD_CBCs")
miCursor=miConexion.cursor()


root.mainloop()'''