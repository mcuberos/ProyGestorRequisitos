from tkinter import messagebox
action=input("Choose C to create a new database. Choose L to load a CBC file to a existing database ")

if action=="C":
    import createDatabase

elif action=="L":
    import LoadDatabase
    
else:
    messagebox.showinfo("WARNING!","AN INVALID OPTION HAS BEEN SELECTED. PLEASE RUN THE PROGRAM AGAIN.")

input("PRESIONE ENTER PARA SALIR")