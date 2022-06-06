from cProfile import label
from functools import partial
import os
from openpyxl import load_workbook
import numpy as np
from tkinter import Button, Label, Tk, filedialog


root = Tk()
root.title('Factura Auditoria')
root.geometry('450x450+300+200')

def abrir():    
    archivo = filedialog.askopenfilename(title="abrir", filetypes=(('Libro Excel (*.xlsx)','*.xlsx'),))
    wb = load_workbook(archivo)

    ws_datos = wb[wb.sheetnames[1]]

    cell_range = ws_datos['A':'B']
    factura = [[],[]]
    llaves = []
    ventas = []
    for l,v in zip(cell_range[0][:], cell_range[1][:]):
        llaves.append(l.value)
        ventas.append(v.value)

    indice_datos = 0
    venta = 0

    for llave in range(min(llaves),max(llaves)+1):
        while indice_datos <=(len(llaves)-1) and llaves[indice_datos] == llave: 
            venta = venta + ventas[indice_datos]
            indice_datos = indice_datos + 1

        if llaves[indice_datos-1] == llave:
            factura[0].append(llave)
            factura[1].append(venta)
            venta = 0

    for i in range(0,len(factura[1])):
        if factura[1][i] > asignacion:
            factura[1][i] = asignacion

    total_factura =sum(factura[1])

    a = np.array(factura)
    a = a.transpose()
    factura = a.tolist()

    wb.create_sheet('FACTURA')
    ws_factura = wb['FACTURA']

    ws_factura.append(['#LLAVE', 'CONSUMO'])
    for r in factura:
        ws_factura.append(r)

    ws_factura.append([None, total_factura])

    BotonAbrirExcel.destroy()
    labelIndicaciones.destroy()
    global LabelFactura
    LabelFactura= Label(root,text=f'Total factura: {total_factura} pesos',font=("Arial", 15))
    LabelFactura.pack()
    global BotonGuardarExcel
    BotonGuardarExcel = Button(root,text='Guardar Excel', command=partial(guardar,wb), fg='green',font=("Arial", 25))  
    BotonGuardarExcel.pack()
    

def guardar(wb):
    archivo = filedialog.asksaveasfilename(title="guardar", filetypes=(('Libro Excel (*.xlsx)','*.xlsx'),))
    wb.save(archivo+'.xlsx') 
    BotonGuardarExcel.destroy()   
    botonNuevaFactura.pack() 
    botonSalir.pack()

def go(valorAsignado: int):
    BotonAsignacion21500.destroy()
    BotonAsignacion55000.destroy()
    myLabel.destroy()
    global LabelValorAsignado
    LabelValorAsignado = Label(root,text=f"Asignacion: {valorAsignado}",font=("Arial", 17))
    LabelValorAsignado.pack()
    global asignacion
    asignacion = valorAsignado
    labelIndicaciones.pack()
    BotonAbrirExcel.pack() 

def close():
   root.quit()   

def nuevaFactura():
    botonSalir.destroy()
    LabelValorAsignado.destroy()
    botonNuevaFactura.destroy()
    LabelFactura.destroy()
    calcularFactura()

def calcularFactura():
    global myLabel
    myLabel = Label(root,text="Seleccione el valor de la asignacion:",font=("Arial", 15))
    global labelIndicaciones
    labelIndicaciones =  Label(root,text="En la 2da hoja del excel deben estar los registros\n de las ventas en la forma LLAVE/VALOR.",fg='goldenrod',font=("Arial", 14))
    global BotonAsignacion21500
    BotonAsignacion21500 = Button(root,text='21500', command=partial(go,21500), fg='green',font=("Arial", 25))  
    global BotonAsignacion55000
    BotonAsignacion55000 = Button(root,text='55000', command=partial(go,55000), fg='green',font=("Arial", 25))   
    global BotonAbrirExcel 
    BotonAbrirExcel = Button(root,text='Abrir Excel', command=abrir, fg='green',font=("Arial", 25))  
    global botonSalir 
    botonSalir =  Button(root,text='Salir', command=close, fg='blue',font=("Arial", 25))
    global botonNuevaFactura
    botonNuevaFactura = Button(root,text='Calcular otra factura', command=nuevaFactura, fg='green',font=("Arial", 25))  
    myLabel.pack()
    BotonAsignacion21500.pack()
    BotonAsignacion55000.pack()

calcularFactura()
root.mainloop()