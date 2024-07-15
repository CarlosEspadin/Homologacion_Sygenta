#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
import pandas as pd
import numpy as np
from pathlib import Path
import difflib
from pathlib import Path

# Manipulación de rutas:

absFilePath = Path(__file__)
Antecesor = absFilePath.parent.parent
print(Antecesor)
                                    
## Condiciones impuestas por el usuario.
ruta = input("Ingresa la ruta del archivo: ")
print("Esta es la ruta: ", ruta)

ruta = Path(ruta)
Homologado = ruta.parent.parent
print(Homologado)


Num_Distri = input("Ingresa el número del distribuidor: ")
print("Número de distribuidor: ", Num_Distri)

NameDistri = input("Ingresa el nombre del distribuidor: ")
print("Nombre de distribuidor: ", NameDistri)

# Carga del catalogo de productos elegido por el usuario.
Distribuidor = pd.read_excel(ruta, sheet_name=0)
row=Distribuidor.shape[0]
colum=Distribuidor.shape[1]
print("Dimensiones del conjunto original:")
print("["+str(row)+" rows x "+str(colum)+" columns]")


# Carga de catalogo de materiales.
ruta_CatalogoSyc=Path(Antecesor,'Catalogo Base.xlsx' )

Materiales = pd.read_excel(ruta_CatalogoSyc)

# Separamos el catalogo por año de producto.
Materiales24 = Materiales[Materiales['Año']==2024]
Materiales23 =  Materiales[Materiales['Año']==2023]
Materiales22 =  Materiales[Materiales['Año']==2022]
Materiales21 =  Materiales[Materiales['Año']==2021]
Materiales20 =  Materiales[Materiales['Año']==2020]

# Ruta para guardar el resultado final

ruta_last = Path(Homologado,"Homologado",NameDistri+".xlsx")
print(ruta_last)

# Nombres de Columnas
print("Nombres de columnas del archivo", Distribuidor.columns)

Tipo = input("¿El catalogo cuenta con claves Syngenta? \n Ingresa 1 para si y 0 en caso contrario.")

if Tipo == '0':
    CodDistriProd = input("Ingresa el nombre de la columna que contiene los códigos de productos de los distribuidores: ")
    print("CodDistriProd ->", CodDistriProd)

    NomDistriProd = input("Ingresa el nombre de la columna que contiene las descripciones de los productos de los distribuidores: ")
    print("NomDistriProd ->", NomDistriProd)
    
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
    Distribuidor['ClaveDistri'] = Num_Distri
    Distribuidor['DescDistri'] = NameDistri
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
else:
    CodDistriProd = input("Ingresa el nombre de la columna que contiene los códigos de productos de los distribuidores: ")
    print("CodDistriProd ->", CodDistriProd)

    CodSyngenta = input("Ingresa el nombre de la columna que contiene los códigos de los productos de Syngenta: ")
    print("CodSyngenta ->", CodSyngenta)
    
    NomDistriProd = input("Ingresa el nombre de la columna que contiene las descripciones de los productos de los distribuidores: ")
    print("NomDistriProd ->", NomDistriProd)
    
    
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
    Distribuidor['ClaveDistri'] = Num_Distri
    Distribuidor['DescDistri'] = NameDistri
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
    
# Guardamos el resultado en la carpeta correspondiente.
Distribuidor.to_excel(ruta_last, "Homologado",index=False)