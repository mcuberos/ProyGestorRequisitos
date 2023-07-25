from tkinter import messagebox
action=input("Seleccione una opción (inserte solo el número): \n1. Crear una Base de Datos (SOLO ADMIN). \n2. Cargar un CBC a la Base de Datos (SOLO ADMIN). \n3. Cargar un fichero temporal a la base de datos (SOLO ADMIN). \n4. Cargar un CBC para generar un excel temporal con la propuesta de autorrellenado. \n5. Cargar un excel temporal para autorellenar un CbC \n")

if action=="1":
    pwd=input("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import createDatabase
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")

elif action=="2":
    pwd=input("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import LoadDatabase
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")

elif action=="3":
    pwd=input("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import LoadDatabaseFromTemp
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")

elif action=="4":
    import CreateCBCTemp

elif action=="5":
    import AutoFillFromCBCTemp

else:
    messagebox.showinfo("WARNING!","AN INVALID OPTION HAS BEEN SELECTED. PLEASE RUN THE PROGRAM AGAIN.")

input("PRESIONE ENTER PARA SALIR")