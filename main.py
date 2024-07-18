#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import traceback
import sys
import time
# importamos la dependencia General
import General_Prod

class App(tk.Tk):
    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        return os.path.join(self.base_path, relative_path)
    
    def __init__(self):
        super().__init__()
        self.title("Sell Out to Web Service")
        self.geometry("950x700")
        self.resizable(0, 0)
        
        # Inicializar variables
        self.ruta = tk.StringVar()
        self.num_distri = tk.StringVar()
        self.name_distri = tk.StringVar()
        self.selected = tk.StringVar()
        # Variables de texto para cada entrada
        self.cod_syngenta_var = tk.StringVar()
        self.cod_ext_var = tk.StringVar()
        self.name_prod_distr_var = tk.StringVar()
        self.cod_ext_syc_var = tk.StringVar()
        
        self.configure(bg='#f0f0f0')  # Cambiar el color de fondo de la ventana
        
        self.iniciar_app() 
    
    def create_widgets(self):
        entry_font = {'font': ('Helvetica', 12)}
        
        # Códigos de producto con clave Syngenta
        self.CodSygenta = ttk.Label(self, text="Ingresa el nombre de la columna los códigos Syngenta", style='TLabel')
        self.CodSygenta_entry = ttk.Entry(self, textvariable=self.cod_syngenta_var, **entry_font)
        self.CodSygenta_entry.config(width=30)
        self.B5 = ttk.Button(self, text="Insertar", command=self.obtener_CodSycDise)
        
        # Códigos de productos con clave externa
        self.CodExt = ttk.Label(self, text="Ingresa el nombre de la columna los códigos Externos", style='TLabel')
        self.CodExt_entry = ttk.Entry(self, textvariable=self.cod_ext_var, **entry_font)
        self.CodExt_entry.config(width=30)
        self.B7 = ttk.Button(self, text="Insertar", command=self.obtener_CodExt)
        
        # Descripciones de los productos
        self.NameProdDistr = ttk.Label(self, text="Ingresa el nombre de la columna las descripciones", style='TLabel')
        self.NameProdDistr_entry = ttk.Entry(self, textvariable=self.name_prod_distr_var, **entry_font)
        self.NameProdDistr_entry.config(width=30)
        self.B6 = ttk.Button(self, text="Insertar", command=self.obtener_NameProdDistr)
        
        # Codigos de productos con clave externa cuando de origen viene con claves Syngenta
        self.CodExt_Syc = ttk.Label(self, text="Ingresa el nombre de la columna los códigos Externos", style='TLabel')
        self.CodExt_entry_Syc = ttk.Entry(self, textvariable=self.cod_ext_syc_var, **entry_font)
        self.CodExt_entry_Syc.config(width=30)
        self.B8 = ttk.Button(self, text="Insertar", command=self.obtener_CodExt_Syc)

        # Inicialmente ocultar todos los elementos
        self.hide_all()

    def create_greeting_message(self):
        # Obtener el valor de Tipo
        Tipo = self.selected.get()
        
        # Mostrar/Ocultar elementos según el valor de Tipo
        if Tipo == "0":
            # Mostrar elementos de tipo Syngenta
            self.show_syngenta_elements()
            # Ocultar elementos de tipo Externo
            self.hide_external_elements()
        else:
            # Mostrar elementos de tipo Externo
            self.show_external_elements()
            # Ocultar elementos de tipo Syngenta
            self.hide_syngenta_elements()

    def show_syngenta_elements(self):
        self.CodSygenta.place(relx=0.01, rely=0.7)
        self.CodSygenta_entry.place(relx=0.45, rely=0.7)
        self.B5.place(relx=0.88, rely=0.7)
        self.NameProdDistr.place(relx=0.01, rely=0.75)
        self.NameProdDistr_entry.place(relx=0.45, rely=0.75)
        self.B6.place(relx=0.88, rely=0.75)
        self.CodExt_Syc.place(relx=0.01, rely=0.8)
        self.CodExt_entry_Syc.place(relx=0.45, rely=0.8)
        self.B8.place(relx=0.88, rely=0.8)

    def hide_syngenta_elements(self):
        self.CodSygenta.place_forget()
        self.CodSygenta_entry.place_forget()
        self.B5.place_forget()
        # self.NameProdDistr.place_forget()
        # self.NameProdDistr_entry.place_forget()
        # self.B6.place_forget()
        self.CodExt_Syc.place_forget()
        self.CodExt_entry_Syc.place_forget()
        self.B8.place_forget()

    def show_external_elements(self):
        self.CodExt.place(relx=0.01, rely=0.7)
        self.CodExt_entry.place(relx=0.45, rely=0.7)
        self.B7.place(relx=0.88, rely=0.7)
        self.NameProdDistr.place(relx=0.01, rely=0.75)
        self.NameProdDistr_entry.place(relx=0.45, rely=0.75)
        self.B6.place(relx=0.88, rely=0.75)

    def hide_external_elements(self):
        self.CodExt.place_forget()
        self.CodExt_entry.place_forget()
        self.B7.place_forget()
        # self.NameProdDistr.place_forget()
        # self.NameProdDistr_entry.place_forget()
        # self.B6.place_forget()

    def hide_all(self):
        self.hide_syngenta_elements()
        self.hide_external_elements()
        
    def iniciar_app(self):
        self.deiconify()
        
        paddings = {'padx': 10, 'pady': 10}
        entry_font = {'font': ('Helvetica', 12)}
        
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        # Estilos
        self.style = ttk.Style(self)
        self.style.configure('TLabel', font=('Helvetica', 12), foreground='#333333', background='#f0f0f0')
        self.style.configure('TButton', font=('Helvetica', 12), background='#28536b', foreground='black')
        self.style.map('TButton', background=[('active', '#45a049')], foreground=[('active', 'black')])  # Color al pasar el ratón

        
        # Aplicar transparencia
        self.attributes('-alpha', 0.95)  # Valor entre 0 (completamente transparente) y 1 (opaco)
        
        # Establecemos la ruta absoluta:
        self.ruta_ico_s = self.resource_path("ConAgro_icon_small.png")
        self.ruta_ico_b = self.resource_path("ConAgro_icon_big.png")
        
        # configure icon
        self.icon_big = tk.PhotoImage(file=self.ruta_ico_b)
        self.icon_small = tk.PhotoImage(file=self.ruta_ico_s)
        self.iconphoto(False, self.icon_big, self.icon_small)
        self.iconbitmap('ConAgro.ico')
        
        label = ttk.Label(self, text="Bienvenido a la Aplicación para Homologación de catálogos ConAgro", style='Title.TLabel')
        label.place(relx=0.5, rely=0.03, anchor='center')
        
        ruta_label = ttk.Label(self, text="Selecciona el catálogo a Homologar:", style='TLabel')
        ruta_label.place(relx=0.1, rely=0.1)
        
        self.browse_button = ttk.Button(self, text="Seleccionar archivo", command=self.seleccionar_archivo)
        self.browse_button.place(relx=0.4, rely=0.1)
        
        self.ruta_label = ttk.Label(self, text="Ruta del archivo seleccionado: \n"+"", style='TLabel')
        self.ruta_label.place(relx=0.01, rely=0.2)
        
        ## Establecemos ruta absoluta de archivo con datos de distribuidores:
        self.ruta_CatalogoDistribuidores = self.resource_path("Catalogo Distribuidores.xlsx")
        self.Catalogo = pd.read_excel(self.ruta_CatalogoDistribuidores, sheet_name=0)
        aux_Num = list(self.Catalogo['Sold to'].sort_values().unique())
        aux_Name = list(self.Catalogo['Descripción'].unique())
        
        # Entradas para el numero de distribuidor
        distri_label = ttk.Label(self, text="Ingresa el Número de cliente:", style='TLabel')
        distri_label.place(relx=0.1, rely=0.3)
        
        self.distri_entry = ttk.Combobox(self, textvariable=self.num_distri, font=('Helvetica', 11), state='normal')
        self.distri_entry['values'] = aux_Num
        self.distri_entry.place(relx=0.35, rely=0.3)
        self.distri_entry.config(width=40)
        
        
        self.B3 = ttk.Button(self, text="Insertar", command=self.obtener_num)
        self.B3.place(relx=0.78, rely=0.3)
        
        name_label = ttk.Label(self, text="Ingresa el nombre del cliente:", style='TLabel')
        name_label.place(relx=0.1, rely=0.4)
        
        # Entrada para el nombre del distribuidor
        self.name_entry = ttk.Combobox(self, textvariable=self.name_distri, font=('Helvetica', 11), state='normal')
        self.name_entry['values'] = aux_Name
        self.name_entry.place(relx=0.35, rely=0.4)
        self.name_entry.config(width=40)
        
        self.B4 = ttk.Button(self, text="Insertar", command=self.obtener_nombre)
        self.B4.place(relx=0.78, rely=0.4)
        
        radio_label = ttk.Label(self, text="Selecciona el tipo de producto con el que cuenta el catálogo:", style='TLabel')
        radio_label.place(relx=0.05, rely=0.6)
        
        self.RB1 = ttk.Radiobutton(self, text="Syngenta", variable=self.selected, value="0", command=self.create_greeting_message)
        self.RB2 = ttk.Radiobutton(self, text="Externo", variable=self.selected, value="1", command=self.create_greeting_message)
        
        self.RB1.place(relx=0.6, rely=0.6)
        self.RB2.place(relx=0.7, rely=0.6)
        
        self.B2 = ttk.Button(self, text="Homologar catálogo", command=self.mostrar_mensaje)
        self.B2.place(relx=0.85, rely=0.9, anchor='center')
        
        self.create_widgets()

    # Método para cerrar la venta App.
    def cerrar_ventana(self):
        self.destroy()
    
    def seleccionar_archivo(self):
        filetypes = (('Archivo de Excel', '*.xlsx'), ('All files', '*.*'))
        filename = filedialog.askopenfilename(title='Abrir archivo', initialdir='/', filetypes=filetypes)
        self.ruta.set(filename)
        self.ruta_label.config(text="Ruta del archivo seleccionado: \n"+filename)
    # Ventana de error si el número de cliente ingresado no existe.
    def Cliente_Not_foud(self):  # sourcery skip: class-extract-method
        messagebox.showerror("Error", "El cliente que estás ingresando no está registrado en ConAgro")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Ventana que muestra error si la ruta del archivo es incorrecta.
    def mostrar_error(self):
        messagebox.showerror("Error", "La ruta del archivo o el número de distribuidor no son válidos. Por favor vuelve a ejecutar el programa e ingresa los valores adecuador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Venta que se muestra si el archivo ingresado es de un tipo incorrecto o simplemente no existe.
    def error_Archivo(self):
        messagebox.showerror("Error", "Archivo no permitido, por favor vuelve a ejecutar el programa ingresando un archivo de tipo .xlsx")
        self.after(30000, self.cerrar_ventana)
        # sys.exit()   
    # Ventana que se muestra si el layout de excel no es el correcto.
    def error_columnas(self):
        messagebox.showerror("Error", "El nombre de los encabezados no coincide, reportar al administrador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    
    def obtener_num(self):
        numero = self.num_distri.get()
        messagebox.showinfo("Número de cliente", f"Se ha ingresado el número de cliente: {numero}")
        self.B3.config(state=tk.DISABLED)
        
    
    def obtener_nombre(self):
        nombre = self.name_distri.get()
        messagebox.showinfo("Nombre de cliente", f"Se ha ingresado el nombre de cliente: {nombre}")
        self.B4.config(state=tk.DISABLED)
    
    def obtener_CodSycDise(self):
        cod_syngenta = self.cod_syngenta_var.get()
        messagebox.showinfo("Código Syngenta", f"Se ha ingresado el código Syngenta: {cod_syngenta}")
        self.B5.config(state=tk.DISABLED)
    
    def obtener_CodExt(self):
        cod_ext = self.cod_ext_var.get()
        messagebox.showinfo("Código Externo", f"Se ha ingresado el código Externo: {cod_ext}")
        self.B7.config(state=tk.DISABLED)
    
    def obtener_NameProdDistr(self):
        name_prod_distr = self.name_prod_distr_var.get()
        messagebox.showinfo("Nombre de Producto Distribuidor", f"Se ha ingresado el nombre del producto distribuidor: {name_prod_distr}")
        self.B6.config(state=tk.DISABLED)
    
    def obtener_CodExt_Syc(self):
        cod_ext_syc = self.cod_ext_syc_var.get()
        messagebox.showinfo("Código Externo de Syngenta", f"Se ha ingresado el código externo de Syngenta: {cod_ext_syc}")
        self.B8.config(state=tk.DISABLED)

    def format_dataframe_as_table(self, df):
        # Convertir DataFrame a texto con formato de tabla
        # Ajustar el ancho de las columnas basado en el contenido
        col_widths = [max(df[col].astype(str).map(len).max(), len(col)) for col in df.columns]
        
        table = ''
        # Crear encabezado
        header = ' | '.join(f"{col:{col_widths[i]}}" for i, col in enumerate(df.columns))
        table += header + '\n'
        table += '-' * len(header) + '\n'
        
        # Crear filas
        for index, row in df.iterrows():
            row_text = ' | '.join(f"{str(row[col]):{col_widths[i]}}" for i, col in enumerate(df.columns))
            table += row_text + '\n'
        
        return table
    
    def close_toplevel(self):
        if hasattr(self, 'top'):
            self.top.destroy()
            self.destroy()
            sys.exit()
    
    def mostrar_mensaje(self):
        try:
            DF_Proces = self.Procesamiento(
                Tipo=int(self.selected.get()),
                Num_Distri=self.num_distri.get(),
                Name_Distri=self.name_distri.get(),
                ruta=self.ruta.get(),
                CodSyngenta=self.cod_syngenta_var.get(),
                NomDistriProd=self.name_prod_distr_var.get(),
                CodDistriProd_Syc=self.cod_ext_syc_var.get(), 
                CodDistriProd=self.cod_ext_var.get()
            )
            
            # Crear ventana Toplevel
            self.top = tk.Toplevel(self)
            self.top.title("Información de homologación")
            self.top.geometry("800x400")
            
            # Crear Treeview
            tree = ttk.Treeview(self.top, columns=list(DF_Proces.columns), show='headings')
            tree.pack(fill='both', expand=True)
            
            # Configurar columnas del Treeview
            for col in DF_Proces.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor='center', width=100)
            
            # Insertar datos en el Treeview
            for index, row in DF_Proces.iterrows():
                tree.insert("", "end", values=list(row))
            
            # Crear y mostrar información adicional
            info_label = tk.Label(self.top, text=(
                "Ruta: " + self.ruta.get() +
                "\nNúmero de distribuidor: " + self.num_distri.get() +
                "\nNombre del distribuidor: " + self.name_distri.get() +
                "\nTipo de producto: " + self.selected.get()
            ), padx=10, pady=10)
            info_label.pack(side='top', anchor='w')
            
            self.close_button = ttk.Button(self.top, text="Terminar Homologación", command=self.close_toplevel)
            self.close_button.pack(pady=20)
        
        except FileNotFoundError as e1:
            traceback.print_exc()
            self.error_Archivo()
        except IndexError as e2:
            traceback.print_exc()
            self.error_columnas()
        except KeyError as e3:
            traceback.print_exc()
            self.error_columnas()
        except TypeError as e4:
            self.mostrar_error()
        
    def Procesamiento(self, Tipo, Num_Distri, Name_Distri, ruta, CodDistriProd_Syc, CodSyngenta, NomDistriProd, CodDistriProd):
        Distribuidor_H = General_Prod.Catalogo(v_Num_Distri=Num_Distri,v_Name_Distri=Name_Distri, v_ruta=ruta, v_CodDistriProd_Syc=CodDistriProd_Syc, v_CodSyngenta=CodSyngenta, v_NomDistriProd=NomDistriProd, v_CodDistriProd=CodDistriProd)
        Distri_df, ruta_destino=Distribuidor_H.get_Catalogo(Distribuidor_H.ruta, v_Name_Distri=Distribuidor_H.v_Name_Distri)
        print(Distri_df.columns)
        print("Tipo de catalogo:", Tipo)
        if Tipo == 1:
            H_output=Distribuidor_H.Tipo_Externo(Distribuidor=Distri_df, v_Name_Distri=Name_Distri, v_Num_Distri=Num_Distri, CodDistriProd=CodDistriProd, NomDistriProd=NomDistriProd)
            # print(H_output)
            Distribuidor_H.Output_Homologacion(ruta_last=ruta_destino, Distribuidor=H_output)
            print("Homologación de con claves externar terminada con exito.")
        else:
            print("Tipo código Syngenta")
        return H_output

if __name__ == "__main__":
    app = App()
    app.mainloop()