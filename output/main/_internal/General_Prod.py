#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
import pandas as pd
import numpy as np
import difflib
import os
import sys

# Manipulación de rutas:

class Catalogo:
    # Obtención de la ruta adaptada para Pyinstaller
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        return os.path.join(self.base_path, relative_path)
    
    def __init__(self, v_Num_Distri, v_ruta, v_Name_Distri, v_CodDistriProd, v_NomDistriProd, v_CodSyngenta, v_CodDistriProd_Syc):
        self.ruta = v_ruta
        self.v_Num_Distri = v_Num_Distri
        self.v_Name_Distri = v_Name_Distri
        self.v_CodDistriProd = v_CodDistriProd
        self.v_NomDistriProd = v_NomDistriProd
        self.v_CodSyngenta = v_CodSyngenta
        self.v_CodDistriProd_Syc = v_CodDistriProd_Syc
        

    def get_ruta(self):
        # Obtener la ruta absoluta del archivo actual
        absFilePath = os.path.abspath(__file__)
        # Obtener el directorio del archivo actual
        Antecesor = os.path.dirname(absFilePath)
        return Antecesor

    # Carga del catálogo de productos elegido por el usuario.
    def get_Catalogo(self, ruta, v_Name_Distri):
        # Obtener el directorio del archivo padre
        Homologado = os.path.abspath(os.path.join(ruta, os.pardir, os.pardir))
        # Construir la ruta completa al archivo "Homologado/<v_Name_Distri>.xlsx"
        ruta_last = os.path.join(Homologado, "Homologado", v_Name_Distri + ".xlsx")
        print(ruta_last)
        
        # Leer el archivo Excel en un DataFrame de pandas
        Distribuidor = pd.read_excel(ruta, sheet_name=0)
        row = Distribuidor.shape[0]
        colum = Distribuidor.shape[1]
        print("Dimensiones del conjunto original:")
        print(f"[{row} rows x {colum} columns]")
        return Distribuidor, ruta_last
    
    def Tipo_Syngenta(self, Distribuidor, v_Name_Distri, v_Num_Distri, NomDistriProd, CodDistriProd, CodSyngenta):
        # Carga de catalogo de materiales.
        ruta_CatalogoSyc = os.path.join(self.get_ruta(), 'Catalogo Base.xlsx')

        Materiales = pd.read_excel(ruta_CatalogoSyc)
        
        # Separamos el catalogo por año de producto.
        Materiales24 = Materiales[Materiales['Año']==2024]
        Materiales23 =  Materiales[Materiales['Año']==2023]
        Materiales22 =  Materiales[Materiales['Año']==2022]
        Materiales21 =  Materiales[Materiales['Año']==2021]
        Materiales20 =  Materiales[Materiales['Año']==2020]
        
        # Ordemos los catalogos de materiales por Producto:
        Materiales23.sort_values(['Producto'])
        Materiales24.sort_values(['Producto'])
        Materiales22.sort_values(['Producto'])
        Materiales21.sort_values(['Producto'])
        Materiales20.sort_values(['Producto'])
        
        # Cambiamos los nombres propios del catalogo de distribuidor al formato ConAgro
        Materiales.rename(columns={'Producto':'DescSyngenta', 'SKU': 'CodSyngenta'}, inplace=True)
        Distribuidor.rename(columns={CodDistriProd:'CodDistriProd', NomDistriProd: 'NomDistriProd', CodSyngenta: 'CodSyngenta'}, inplace=True)
        
        # Empezamos con el mappíng para añadir los códigos de productos correspondientes.
        
        Materiales=Materiales[['CodSyngenta', 'DescSyngenta']]
        DescSyngenta = Materiales['DescSyngenta']
        CodSyngenta = Materiales['CodSyngenta']
        d_Materiales = dict(zip(CodSyngenta,DescSyngenta))
        Distribuidor['DescSyngenta']=Distribuidor['CodSyngenta'].map(d_Materiales)
        
        # Añadimos las columnas faltantes.
        Distribuidor['ClaveDistri'] = v_Num_Distri
        Distribuidor['DescDistri'] = v_Name_Distri
        Distribuidor['Impuesto']  = 0
        Distribuidor['Pais']= 'MEX'
        Distribuidor['Presentacion'] = np.nan
        Columas = ['CodSyngenta', 'DescSyngenta', 'ClaveDistri', 'DescDistri', 'CodDistriProd', 'NomDistriProd', 'Presentacion', 'Impuesto', 'Pais']
        Distribuidor = Distribuidor[Columas]
        
        # Validamos que no tengamos información diplicada.
        
        print(Distribuidor[Distribuidor['DescSyngenta'].isnull()])
        
        # Validación para saber que se conserve el mismo número de columnas.
        row=Distribuidor.shape[0]
        colum=Distribuidor.shape[1]
        print("Dimensiones del conjunto original:")
        print("["+str(row)+" rows x "+str(colum)+" columns]")
        print("Dimensiones del conjunto homologado",len(Distribuidor))
        
        return Distribuidor
    
    def Tipo_Externo(self, Distribuidor, v_Name_Distri, v_Num_Distri, NomDistriProd, CodDistriProd):
        # Carga de catalogo de materiales.
        ruta_CatalogoSyc = os.path.join(self.get_ruta(), 'Catalogo Base.xlsx')


        Materiales = pd.read_excel(ruta_CatalogoSyc)
        
        # Separamos el catalogo por año de producto.
        Materiales24 = Materiales[Materiales['Año']==2024]
        Materiales23 =  Materiales[Materiales['Año']==2023]
        Materiales22 =  Materiales[Materiales['Año']==2022]
        Materiales21 =  Materiales[Materiales['Año']==2021]
        Materiales20 =  Materiales[Materiales['Año']==2020]
            
        # Ordemos los catalogos de materiales por SKU:
        Materiales23.sort_values(['SKU'])
        Materiales24.sort_values(['SKU'])
        Materiales22.sort_values(['SKU'])
        Materiales21.sort_values(['SKU'])
        Materiales20.sort_values(['SKU'])
        
        # Encontramos las coincidencias usando logica difusa
        Distribuidor['DescSyngenta'] = Distribuidor[NomDistriProd].str.rstrip().apply(
        lambda x: (difflib.get_close_matches(x.upper(), Materiales24['Producto'], cutoff=0.7)[:1] or 
                difflib.get_close_matches(x.upper(), Materiales23['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales22['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales21['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales20['Producto'], cutoff=0.6)[:1] or [None])[0]
        )
        
        # Validamos el número de coincidencias y el total
        efectividad = len( Distribuidor[Distribuidor['DescSyngenta'].notnull()])
        Per = efectividad/len(Distribuidor['DescSyngenta'])
        porcentaje_formateado = "{:.2%}".format(Per)
        print("Numero de coincidencias: ",efectividad)
        print("Porcentaje de efectividad: ", porcentaje_formateado)
        
        # Cambiamos los nombres propios del catalogo de distribuidor al formato ConAgro
        Materiales.rename(columns={'Producto':'DescSyngenta', 'SKU': 'CodSyngenta'}, inplace=True)
        Distribuidor.rename(columns={CodDistriProd:'CodDistriProd', NomDistriProd: 'NomDistriProd'}, inplace=True)
        
        # Empezamos con el mappíng para añadir los códigos de productos correspondientes.
        Materiales=Materiales[['DescSyngenta', 'CodSyngenta']]
        DescSyngenta = Materiales['DescSyngenta']
        CodSyngenta = Materiales['CodSyngenta']
        d_Materiales = dict(zip(DescSyngenta, CodSyngenta))
        # print(d_Materiales)
        Distribuidor['CodSyngenta']=Distribuidor['DescSyngenta'].map(d_Materiales)
        
        
        # Añadimos las columnas faltantes.
        Distribuidor['ClaveDistri'] = v_Num_Distri
        Distribuidor['DescDistri'] = v_Name_Distri
        Distribuidor['Impuesto']  = 0
        Distribuidor['Pais']= 'MEX'
        Distribuidor['Presentacion'] = np.nan
        Columas = ['CodSyngenta', 'DescSyngenta', 'ClaveDistri', 'DescDistri', 'CodDistriProd', 'NomDistriProd', 'Presentacion', 'Impuesto', 'Pais']
        Distribuidor = Distribuidor[Columas]
        
        # Validamos que no tengamos información diplicada.
        print(Distribuidor[Distribuidor['DescSyngenta'].duplicated()])
        
        # Validación para saber que se conserve el mismo número de columnas.
        row=Distribuidor.shape[0]
        colum=Distribuidor.shape[1]
        print("Dimensiones del conjunto original:")
        print("["+str(row)+" rows x "+str(colum)+" columns]")
        
        return Distribuidor
    
    def Output_Homologacion(self, ruta_last, Distribuidor):
        # Guardamos el resultado en la carpeta correspondiente.
        Distribuidor.to_excel(ruta_last, "Homologado",index=False)