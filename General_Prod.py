#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
import pandas as pd
import numpy as np
from pathlib import Path
import difflib
from pathlib import Path
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
        self.absFilePath = Path(__file__)
        self.Antecesor = self.absFilePath.parent
        return self.Antecesor
    
    # Carga del catalogo de productos elegido por el usuario.
    def get_Catalogo(self, ruta, v_Name_Distri):
        ruta = Path(ruta)
        Homologado = ruta.parent.parent
        ruta_last = Path(Homologado,"Homologado",v_Name_Distri+".xlsx")
        print(ruta_last)
        Distribuidor = pd.read_excel(ruta, sheet_name=0)
        row=Distribuidor.shape[0]
        colum=Distribuidor.shape[1]
        print("Dimensiones del conjunto original:")
        print("["+str(row)+" rows x "+str(colum)+" columns]")
        return Distribuidor, ruta_last
    
    def Tipo_Syngenta(self, Distri, v_Name_Distri, v_Num_Distri):
        # Carga de catalogo de materiales.
        ruta_CatalogoSyc=Path(self.get_ruta(),'Catalogo Base.xlsx')

        Materiales = pd.read_excel(ruta_CatalogoSyc)
        
        #  Filtramos el catalogo base de materiales
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
        Distri['DescSyngenta'] = Distri[v_Name_Distri].str.rstrip().apply(
        lambda x: (difflib.get_close_matches(x.upper(), Materiales24['Producto'], cutoff=0.7)[:1] or 
                difflib.get_close_matches(x.upper(), Materiales23['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales22['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales21['Producto'], cutoff=0.6)[:1] or
                difflib.get_close_matches(x.upper(), Materiales20['Producto'], cutoff=0.6)[:1] or [None])[0]
        )
        
        # Validamos el número de coincidencias y el total
        efectividad = len( Distri[Distri['DescSyngenta'].notnull()])
        Per = efectividad/len(Distri['DescSyngenta'])
        porcentaje_formateado = "{:.2%}".format(Per)
        print("Numero de coincidencias: ",efectividad)
        print("Porcentaje de efectividad: ", porcentaje_formateado)
        
        # Cambiamos los nombres propios del catalogo de distribuidor al formato ConAgro
        Materiales.rename(columns={'Producto':'DescSyngenta', 'SKU': 'CodSyngenta'}, inplace=True)
        Distri.rename(columns={v_Num_Distri:'v_Num_Distri', v_Name_Distri: 'v_Name_Distri'}, inplace=True)
        
        # Empezamos con el mappíng para añadir los códigos de productos correspondientes.
        Materiales=Materiales[['DescSyngenta', 'CodSyngenta']]
        DescSyngenta = Materiales['DescSyngenta']
        CodSyngenta = Materiales['CodSyngenta']
        d_Materiales = dict(zip(DescSyngenta, CodSyngenta))
        # print(d_Materiales)
        Distri['CodSyngenta']=Distri['DescSyngenta'].map(d_Materiales)
        
        
        # Añadimos las columnas faltantes.
        Distri['ClaveDistri'] = v_Num_Distri
        Distri['DescDistri'] = v_Name_Distri
        Distri['Impuesto']  = 0
        Distri['Pais']= 'MEX'
        Distri['Presentacion'] = np.nan
        Columas = ['CodSyngenta', 'DescSyngenta', 'ClaveDistri', 'DescDistri', 'v_Num_Distri', 'v_Name_Distri', 'Presentacion', 'Impuesto', 'Pais']
        Distri = Distri[Columas]
        
        # Validamos que no tengamos información diplicada.
        print(Distri[Distri['DescSyngenta'].duplicated()])
        
        # Validación para saber que se conserve el mismo número de columnas.
        row=Distri.shape[0]
        colum=Distri.shape[1]
        print("Dimensiones del conjunto original:")
        print("["+str(row)+" rows x "+str(colum)+" columns]")
        
        return Distri
    
    def Tipo_Externo(self, Distribuidor, v_Name_Distri, v_Num_Distri, NomDistriProd, CodDistriProd):
        # Carga de catalogo de materiales.
        ruta_CatalogoSyc=Path(self.get_ruta(),'Catalogo Base.xlsx')

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