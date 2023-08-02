from tkinter import messagebox
import getpass

action=input("Seleccione una opción (inserte solo el número): \n1. Cargar un CBC a la Base de Datos (SOLO ADMIN). \n2. Cargar un fichero temporal a la base de datos (SOLO ADMIN). \n3. Cargar un CBC para generar un excel temporal con la propuesta de autorrellenado. \n4. Cargar un excel temporal para autorellenar un CbC \n")
'''
action=input("Seleccione una opción (inserte solo el número): \n1. Crear una Base de Datos (SOLO ADMIN). \n2. Cargar un CBC a la Base de Datos (SOLO ADMIN). \n3. Cargar un fichero temporal a la base de datos (SOLO ADMIN). \n4. Cargar un CBC para generar un excel temporal con la propuesta de autorrellenado. \n5. Cargar un excel temporal para autorellenar un CbC \n")
if action=="1":
    pwd=getpass.getpass("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import createDatabase
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")
'''

if action=="1":
    pwd=getpass.getpass("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import LoadDatabase
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")

elif action=="2":
    pwd=getpass.getpass("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import LoadDatabaseFromTemp
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")

elif action=="3":
    import CreateCBCTemp

elif action=="4":
    import AutoFillFromCBCTemp

else:
    messagebox.showinfo("WARNING!","AN INVALID OPTION HAS BEEN SELECTED. PLEASE RUN THE PROGRAM AGAIN.")

input("PRESIONE ENTER PARA SALIR")