from tkinter import * 
from tkinter import filedialog 
from tkinter import messagebox
import os
import wave
import pandas as pd
from datetime import date
import time


def openFolder():
    dir = filedialog.askdirectory()
    if dir != "":
        updateMessageLabel()
        os.chdir(dir)
        path = os.getcwd()
        #print(path)

        content = os.listdir(dir)
        diccionarioAudio = {}

        for item in content:
            if item.endswith(".wav"):
                #print("AUDIO")
                waveFile = wave.open(item, "r")
                params = waveFile.getparams()
                frames = params[3]
                rate = params[2]
                duration = float( frames/rate )
                #print(item + " duracion: " + "0" + str(duration))
                diccionarioAudio[item] = duration
                waveFile.close()
        
        if len(diccionarioAudio) == 0:
            messagebox.showinfo("Reporte de audios", "No se encontraron archivos de audio por analizar")
        else:
            writeExcel(path, diccionarioAudio)
            messagebox.showinfo("Reporte de audios", "El reporte de audio a finalizado")


def updateMessageLabel():
    label.configure(text="Se detecto un directorio")


def writeExcel(path, data):
    dataKeys = []
    dataValues = []
    for i, val in data.items():
        dataKeys.append(i)
        dataValues.append(val)

    timeFormat = timer(dataValues)
    """
    print( dataKeys )
    print( timeFormat )
    """
    today = date.today()
    df = pd.DataFrame({
        #'Archivo': ['file_1', 'file_2', 'file_3'],
        #'Duracion': ['00:00:40','00:00:33','00:00:30']
        'Archivo': dataKeys,
        'Duracion': timeFormat
        })
    #df = df[['Archivo', 'Duracion']]
    write = pd.ExcelWriter(str(path)+"/"+str(today)+".xlsx", engine='xlsxwriter')
    df.to_excel(write, 'Hoja de reportes', index=False)
    write.save()


def timer(arg):
    duration = []
    newFormat = ""
    hh = ""
    mm = ""
    ss = ""
    for item in arg:
        hour, rem = divmod(item, 3600)
        minutes, seconds = divmod(rem, 60)
        hh = int(hour) if int(hour) >= 10 else "0" + str( int(hour) )
        mm = int(minutes) if int(minutes) >= 10 else "0" + str( int(minutes) )
        ss = int(seconds) if int(seconds) >= 10 else "0" + str( int(seconds) )
        newFormat =  str(hh) + ":" + str(mm) + ":" + str(ss) 
        #print(newFormat)
        duration.append(newFormat)
    return duration


window  = Tk()
window.title("Welcome")
window.geometry("350x200")
Button(text="Elegir carpeta", command=openFolder).place(x=15, y=20)
#Button(text="Elegir carpeta", bg="deep sky blue", command=openFolder).pack()
label = Label(window, text="No se ha seleccionado ningún directorio")
label.place(x=15, y=60)
#label.pack()
window.mainloop()