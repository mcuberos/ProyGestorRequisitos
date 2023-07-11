from tkinter import messagebox
action=input("Seleccione una opción (inserte solo el número): \n1. Crear una Base de Datos. \n2. Cargar un CBC a la Base de Datos. \n3. Cargar un fichero temporal a la base de datos \n")

if action=="1":
    import createDatabase

elif action=="2":
    import LoadDatabase

elif action=="3":
    import LoadDatabaseFromTemp

else:
    messagebox.showinfo("WARNING!","AN INVALID OPTION HAS BEEN SELECTED. PLEASE RUN THE PROGRAM AGAIN.")

input("PRESIONE ENTER PARA SALIR")