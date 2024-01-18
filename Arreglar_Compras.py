import pandas as pd
import numpy as np
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename


# Ruta del archivo Excel con la lista de proveedores
ruta_proveedores_excel = 'PROVEEDORES.xlsx'

# Leer las columnas ORIGEN y DESTINO de la hoja 1 del archivo Excel
df_proveedores_Origen_Destino = pd.read_excel(ruta_proveedores_excel, sheet_name='Hoja1', usecols=['CUIT', 'ORIGEN', 'DESTINO'])

df_proveedores_Origen_Destino['ORIGEN'] = df_proveedores_Origen_Destino['ORIGEN'].astype(str)
df_proveedores_Origen_Destino['DESTINO'] = df_proveedores_Origen_Destino['DESTINO'].astype(str)
df_proveedores_Origen_Destino['CUIT'] = df_proveedores_Origen_Destino['CUIT'].astype(np.int64)



# Ruta donde se encuentra al archivo CSV de AFIP
ruta_completa = askopenfilename(title='Seleccione el archivo CSV de AFIP', filetypes=[('CSV', '*.csv')])

# DataFrame con datos originales
df = pd.read_csv(ruta_completa, sep=';', header=0)



# si el CUIT de df se encuentra en el CUIT de df_proveedores_Origen_Destino, entonces reemplazar el valor de la columna ORIGEN del df por el de la columna DESTINO del df (ORIGEN y DESTINO se encuentran en df_proveedores_Origen_Destino)
# Combinar DataFrames en función del CUIT
df_merged = pd.merge(df, df_proveedores_Origen_Destino, left_on='Nro. Doc. Vendedor', right_on='CUIT', how='left')

df_merged["CUIT"].fillna(0, inplace = True)

for row in range(len(df_merged)):
    # Asignar valores de las columnas de ORIGEN y DESTINO a variables
    Origen = df_merged['ORIGEN'][row]
    Destino = df_merged['DESTINO'][row]
    Cuit = df_merged['CUIT'][row]
    # reemplazar el valor de la columna de Destino por el de la columna de Origen cuando el valor de la columna de Destino sea nulo y el valor de la columna de Origen no sea nulo
    if Cuit != 0:
        df_merged.loc[row, Destino] = df_merged.loc[row, Origen]
        df_merged.loc[row, Origen] = np.nan

# Eliminar columnas ORIGEN, DESTINO y CUIT
df_merged.drop(['ORIGEN', 'DESTINO', 'CUIT'], axis=1, inplace=True)

# Guardar el DataFrame actualizado en un nuevo archivo CSV si es necesario
df_merged.to_csv(ruta_completa.replace(".csv" , "_Actualizado.csv"), sep=';', index=False , decimal=',')

showinfo('Información', 'El archivo se ha actualizado correctamente')
