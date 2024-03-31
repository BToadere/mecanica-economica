import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.interpolate import interp1d
import datetime
import plotly.graph_objects as go
import os
import datetime
import matplotlib.dates as mdates

def pinta(tablas, columnas):

    carpeta_exportaciones = os.path.join('temp','graficas')
    if not os.path.exists(carpeta_exportaciones):
        os.makedirs(carpeta_exportaciones)
    
    # Inicializar la figura
    fig, ax = plt.subplots() 
    max_longitud = 0
    
    for i, tabla in enumerate(tablas):
        for j, columna in enumerate(columnas[i]):
            # print('columna:\n', columna)
            # Extraer datos de la tabla
            x = pd.to_datetime(tabla['time']).apply(lambda x: x.timestamp()).to_numpy()
            y = tabla[columna].to_numpy()  # Columna a graficar

            # Eliminar elementos nulos
            x = x[~np.isnan(y)]
            y = y[~np.isnan(y)]

            # Indicador dominio mas garnde entre las columnas
            longitud_dominio = len(x)
            
            if longitud_dominio > max_longitud:
                max_longitud = longitud_dominio
                x_eje = x

            if len(y) == 0:
                print(f'No se ha podido graficar la {os.environ.get("nombre_evento")}_{columna}')
                continue
    
            dom_ini = x[0]
            dom_fin = x[-1]
            
            # Generar nuevos puntos para la interpolación
            xs = np.linspace(dom_ini, dom_fin, 100)

            try:    
                # Intentar la interpolación cuadrática
                f_interp = interp1d(x, y, kind='quadratic')     # Puedes cambiar 'linear' a 'cubic', 'quadratic', etc.
                ys_interp = f_interp(xs)
                ys_interp = f_interp(xs)
            except ValueError:
                # En caso de error, realizar la interpolación lineal
                f_interp = interp1d(x, y, kind='linear')
                ys_interp = f_interp(xs)
    
            # Graficar los resultados
            ax.plot(x, y, 'o')
            ax.plot(xs, ys_interp, label=f'{columna}')

    # Agregar malla horizontal
    ax.grid(linestyle='--', alpha=0.4, which='both')  

    # Modificar límites de los ejes
    # plt.xlim(x[0], x[-1])
    # plt.ylim(-1.5, 1.5)
 
    # Establecer los nuevos nombres de los ticks en el eje x
    num_maximo_ticks = 13
    if len(x_eje)< num_maximo_ticks:    
        fec_x = pd.to_datetime(x_eje, unit='s')
        fec_x_formateada = [fecha.strftime('%d/%m/%Y') for fecha in fec_x]
        x_ticks = x_eje
        
    else:
        indices = np.linspace(0, len(x_eje) - 1, num_maximo_ticks, dtype=int)
        x_ticks = [x_eje[i] for i in indices]
        fec_x = pd.to_datetime(x_ticks, unit='s')
        fec_x_formateada = [fecha.strftime('%d/%m/%Y') for fecha in fec_x]
    
    ax.set_xticks(x_ticks)
    ax.set_xticklabels(fec_x_formateada)

    # Añadir leyenda
    ax.legend()
    plt.xticks(rotation=25)

    # Guardar figura
    lista_todas_columnas = [elem for columna in columnas for elem in columna]
    ruta_figura = os.path.join(carpeta_exportaciones, f'{os.environ.get("nombre_evento")}-{"-".join(lista_todas_columnas)}.png')
    plt.savefig(ruta_figura)
    
    # Mostrar la gráfica
    plt.show()

def pinta_px(tablas, columnas):
    # Crear la carpeta 'Exportaciones' si no existe
    carpeta_exportaciones = os.path.join('temp','graficas')
    if not os.path.exists(carpeta_exportaciones):
        os.makedirs(carpeta_exportaciones)
    
    # Inicializar la figura
    go.Figure()    
    
    for i in range(len(tablas)):
        tabla = tablas[i]
        # Extraer el nombre de la tabla
        nombre_tabla = [nombre for nombre, valor in globals().items() if valor is tabla][0]

        # Crear la ruta completa para guardar la figura
        ruta_figura = os.path.join(carpeta_exportaciones, f'{nombre_tabla}_{"_".join(columnas[i])}.png')
        
        for columna in columnas[i]:
            # Extraer datos de la tabla
            x = pd.to_datetime(tabla['time']).apply(lambda x: x.timestamp()).to_numpy()
            # print('columna:', columna)
            y = tabla[columna].to_numpy()  # Columna a graficar
    
            # Eliminar elementos nulos
            x = x[~np.isnan(y)]
            y = y[~np.isnan(y)]
            print(len(x), len(y))
            print(x, y)
            if len(y) == 0:
                print(f'No se ha podido graficar la {nombre_tabla}-{columna}')
                continue
    
            dom_ini = x[0]
            dom_fin = x[-1]
            
            # Generar nuevos puntos para la interpolación
            xs = np.linspace(dom_ini, dom_fin, 100)
            fec_x = [datetime.datetime.fromtimestamp(ts) for ts in x]
            fec_xs = [datetime.datetime.fromtimestamp(ts) for ts in xs]

            try:    
                # Intentar la interpolación cuadrática
                f_interp = interp1d(x, y, kind='quadratic')     # Puedes cambiar 'linear' a 'cubic', 'quadratic', etc.
                ys_interp = f_interp(xs)
                ys_interp = f_interp(xs)
            except ValueError:
                # En caso de error, realizar la interpolación lineal
                # print(f'Interpolación cuadrática fallida para {nombre_tabla}_{columna}. Realizando interpolación lineal.')
                f_interp = interp1d(x, y, kind='linear')
                ys_interp = f_interp(xs)
    
            # Graficar los resultados
            plt.plot(fec_x, y, 'o')#, label=f'Data- {columna}')         #creo que hay qeu cambiarlo aqui
            plt.plot(fec_xs, ys_interp, label=f'{columna}')
        
    # Personalizar ejes
    # plt.xlabel('Eje X')  # Nombre del eje X
    # plt.ylabel('Eje Y')  # Nombre del eje Y
    plt.title('\n'.join(columnas[i]))                  # Título de la gráfica
    # plt.axhline(0, color='black',linewidth=0.5)     # Línea horizontal en y=0
    # plt.axvline(25, color='black',linewidth=0.5)     # Línea vertical en x=0
    plt.grid(axis='y', linestyle='--', alpha=0.7, which='both')  # Agregar malla horizontal
    
    # Modificar límites de los ejes
    # plt.xlim(-2*np.pi, 2*np.pi)
    # plt.ylim(-1.5, 1.5)
    
    # Añadir leyenda
    plt.legend()
    
    # Guardar figura
    plt.savefig(ruta_figura)
    
    # Mostrar la gráfica
    plt.show()
