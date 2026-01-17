#https://docs.python.org/es/3/library/tkinter.html

from processor import process_excel_safe
import tkinter as tk
from tkinter import filedialog,messagebox

def seleccionar_excel():
    return filedialog.askopenfilename(
    title="seleccionar Archivo Excel",
    filetypes=[("Archivo Excel", "*.xlsx")]
    )
def on_click_proccessor():
    archivo = seleccionar_excel()
    exito,mensaje = process_excel_safe(archivo)
    if exito:
        messagebox.showinfo("Proceso Completo",mensaje)
    else:
        messagebox.showerror("Error ",mensaje)  
def iniciar_app():
    root = tk.Tk()
    root.title("Procesador de archivos Excel")     
    root.geometry ("400x400") 
    root.resizable(False,False)

    boton = tk.Button(
    root,
    text="Seleccionar Archivo Excel",
    command=on_click_proccessor,
    width=30,
    height=2
    )
    boton.pack(pady=60)
    root.mainloop()