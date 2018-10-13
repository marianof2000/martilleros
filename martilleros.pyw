import tkinter
from tkinter import *
import tkinter.ttk as ttk
import os.path
import xlrd
from datetime import datetime

def get_platform():
    platforms = {
        "linux" : "Linux",
        'linux1' : 'Linux',
        'linux2' : 'Linux',
        'darwin' : 'OS X',
        'win32' : 'Windows'
    }
    if sys.platform not in platforms:
        return sys.platform
    
    return platforms[sys.platform]

def estructurafecha(a, f):
    #hoy=str(datetime.now())[:10]
    #dia=str(datetime.now())[8:10] 
    #mes=str(datetime.now())[5:7] 
    anio=int(str(datetime.now())[:4])
    if int(a[:4])>anio or int(a[:4])<1901:
        etiqueta1.config(text="El año está mal fila: "+str(f), fg="red")
        return False
    elif int(a[4:6])>12 or int(a[4:6])<1:
        etiqueta1.config(text="El mes está mal fila: "+str(f), fg="red")
        return False
    elif int(a[6:8])>31 or int(a[6:8])<1:
        etiqueta1.config(text="El día está mal fila: "+str(f), fg="red")
        return False
    else:
        return True


def conectar():
    try:
        book = xlrd.open_workbook('padron2.xls')
    except FileNotFoundError:
        etiqueta1.config(text="El archivo no existe", fg="red")
    except:
        etiqueta1.config(text="Error!!", fg="red")
    else:
        hoja = book.sheet_by_index(0)
        if hoja.ncols != 22:
            etiqueta1.config(text="Cantidad de columnas incorrectas", fg="red")
        else:
            etiqueta1.config(text="Archivo correcto")
            fechamax=combo1.get()+limite[int(combo2.get())-1]
            etiqueta1.config(text=fechamax)
            for fila in range(2,hoja.nrows):
                #print(str(fila) + " " +str(hoja.cell_value(fila,4))[:8])
                if not estructurafecha(str(hoja.cell_value(fila,4))[:8],fila):
                    break

def generar():
    pass

anios=["2017","2018","2019","2020","2021","2022","2023"]
meses=["1", "2", "3"]
limite=['0430','0831','1231']

raiz=Tk()
raiz.title("Colegio de Martilleros de Dolores - AICP 2 - v0.6 "+get_platform())
raiz.resizable(0,0)
raiz.geometry("520x250")
if get_platform()=="Windows":
    raiz.iconbitmap("aicp2.ico")

etiqueta1=Label(raiz, text="Padron2.xls", font=("Arial",11))
etiqueta1.place(x=40,y=10)

etiqueta2=Label(raiz, text="Período -------------------------", font=("Arial",11))
etiqueta2.place(x=250,y=10)

boton1=Button(raiz, text="--Conectar XLS--", font=("Arial",14), command=conectar)
boton1.place(x=30,y=30)

combo1=ttk.Combobox(raiz, state='readonly', values=anios, font=("Arial",14))
combo1.place(x=220, y=30, width=120, height=35)
combo1.current(0)

combo2=ttk.Combobox(raiz, state='readonly', values=meses, font=("Arial",14))
combo2.place(x=360, y=30, width=120, height=35)
combo2.current(0)

barra1=ttk.Progressbar(raiz)
barra1.place(x=40, y=100, width=440, height=50)
# barra1.start()

boton2=Button(raiz, text="Generar", font=("Arial",14), command=generar)
boton2.place(x=390,y=190)

etiqueta3=Label(raiz, text="", 
    font=("Arial",14), fg="red")
etiqueta3.place(x=60,y=160)

etiqueta4=Label(raiz, text="Martilleros AICP 2 - Aplicativo ARBA\nMariano Francisco (2011 - 2018)", 
    font=("Arial",10), justify=LEFT)
etiqueta4.place(x=30,y=190)

raiz.mainloop()
