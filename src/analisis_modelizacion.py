# -*- coding: utf-8 -*-
"""
Created on Sun Nov 26 22:18:02 2023

@author: botoa
"""
# from IPython import get_ipython
# get_ipython().magic('clear')

import sys
import pandas as pd
import numpy as np
import copy
from scipy.interpolate import *
import matplotlib.pyplot as plt
import plotly.express as px

import os
import datetime

class evento(pd.DataFrame):
    def __init__(self, data_or_path=None, hoja=None, numero_registros=0):
        if isinstance(data_or_path, pd.DataFrame):
            super().__init__(data_or_path)
        elif isinstance(data_or_path, str):
            super().__init__(pd.read_excel(data_or_path, sheet_name=hoja))
            self.columns= ['time', 'value']
        else:
            print('ERROR: No se ha inicializado correctamente la clase evento')
        # print('numero_registros: ', numero_registros)
        # print(numero_registros," > ",len(self))
        if numero_registros > 0:
            self.truncar(numero_registros)
            if numero_registros > len(self):
                print(f"No se ha truncado nada, numero({numero_registros}) > registros({len(self)})")

    def truncar(self, numero_registros):    
        self.drop(self.index[numero_registros:], inplace = True)
        # self = self.truncate(after=numero_registros-1)

    def volatilidad(self, column_name, grupo):
        columna = self[column_name]
        N = len(columna)
        result = np.empty(N)
        result[:] = np.nan
        if N < grupo:
            print(f'No se ha podido calcular la volatilidad en el rango de {grupo} hay solamente {N} elementos.')
            return result

        if not isinstance(columna, np.ndarray):
            columna = columna.to_numpy()

        # Aplicar logaritmo a la columna
        columna = np.log(columna)

        # Calcular la diferencia entre elementos consecutivos (analogo formula articulo)
        columna = np.diff(columna)#np.insert(, 0, 0)
        # print(N, len(columna))
        
        # print(result)
        
        for i in range(grupo, N):
            # print(i, range(i-grupo, i))
            result[i] = np.std(columna[i-grupo:i+grupo])
        # print(len(result), result)
        self[f'volatility_{grupo}'] = result
        return
        
    def entropia_shannon(self, nombre_columna, grupo=None):
        columna = self[nombre_columna]
        N = len(columna)
        if grupo is None:
            grupo = N
        elif grupo > N:
            print(f'WARNING[ENTROPIA SHANNON]: El rango seleccionado {grupo} es mayor al numero de elementos en la columna ("{nombre_columna}")')
        elif grupo is not None:
            sobran = N % grupo
            columna = columna[sobran:]
        
        if np.max(columna) == 1:
            print(f'WARNING[ENTROPIA SHANNON]: El orden del valor maximo es {orden_max_valor}, la columna introducida ("{nombre_columna}") esta normalizada.')

        entropia_por_grupo = []
        for i in range(0, N, grupo):
            lista_grupo = columna.iloc[i:i+grupo].to_numpy()
 
            max_valor = np.max(lista_grupo)
            normal = lista_grupo/max_valor

            orden_max_valor = int(np.log10(max_valor))
            orden = 10**orden_max_valor

            normal = normal*orden
            normal = np.round(normal)

            result = np.empty(grupo)
            result[:] = np.nan
            for i in range(grupo):
                cantidad = np.count_nonzero(normal == normal[i])
                result[i] = cantidad
            
            frecuencia = result / grupo
            entropia = np.log(frecuencia)*frecuencia
            valor_entropia = np.abs(entropia.sum())
            entropia_por_grupo.append(valor_entropia)

        return entropia_por_grupo

    

class serie_x(evento):
    def __init__(self, data_or_path, hoja=None, numero_registros=0, val_p=0):
        super().__init__(data_or_path, hoja, numero_registros)
        self._val_p = val_p  # Propiedad para val_p
        self._normalizar_derivar()

    def _normalizar_derivar(self):
        tiempo = self['time'].to_list()
        valor = self['value'].to_numpy()

        normal = valor / np.max(valor)
        var_abs = np.insert(np.abs(np.diff(normal)), 0, 0)
        norm_filtr = copy.copy(normal)
        var_abs_filtr = copy.copy(var_abs)
        # print('va_p = ', self._val_p)
        # Crear una ind booleano >= val_p
        ind_valp = var_abs_filtr >= self._val_p
        
        # Sustituir los valores menores que val_p por 0
        norm_filtr[~ind_valp] = 0
        var_abs_filtr[~ind_valp] = 0
        
        self['normal'] = normal
        self['var_abs'] = var_abs
        self['ind_valp'] = ind_valp
        self['norm_valp'] = norm_filtr
        self['var_valp'] = var_abs_filtr

    def nodos(self):
        nodos_df = self[self['ind_valp'] == 1]
        # print('tipo final', type(serie_x.from_dataframe(nodos_df, numero_registros=len(nodos_df), val_p=self._val_p)))
        x = serie_x(nodos_df, numero_registros=len(nodos_df), val_p=self._val_p)
        return x

    @property
    def val_p(self):
        return self._val_p

    @val_p.setter
    def val_p(self, nuevo_valor):
        self._val_p = nuevo_valor
        self._normalizar_derivar()  # Vuelve a calcular al cambiar el límite de val_p


class serie_z(serie_x):
    def __init__(self, data_or_path, hoja=None, numero_registros=0, val_p=0, val_r=1, boton=0):
        super().__init__(data_or_path, hoja, numero_registros)
        self._val_p = val_p     # Propiedad para val_p
        self._val_r = val_r     # Propiedad para val_r
        self._boton = boton     
        super()._normalizar_derivar()
        self.resultado = self._agregar_parametrizar()
        
        
    def _agregar_parametrizar(self):
        # Tu código para calcular los resultados
        
        tiempo_original = list(self['time'])
        norm_rang = self['norm_valp'].to_numpy()
        var_rang = self['var_valp'].to_numpy()
        ind_ini = self['ind_valp'].to_numpy()
        N = len(ind_ini)

        self.drop(self.index, inplace=True)
        columnas = self.columns.tolist()  
        for nombre in columnas:
            if nombre != 'time':
                del self[nombre]
                
        # self.drop(columns=self.columns.difference(['time']))

        # Ajustar el valor de val_r si el botón es 1
        if self._boton == 1:  # Corregido: usar self._boton en lugar de boton
            self._val_r = int(N / self._val_r)

        sobran = N % self._val_r
        if sobran >= 0:
            norm_rang = norm_rang[sobran:]
            var_rang = var_rang[sobran:]
            ind_ini = ind_ini[sobran:]

        # Reshape para dividir en grupos de val_r
        norm_rang = norm_rang.reshape(-1, self._val_r)
        ind_ini = ind_ini.reshape(-1, self._val_r)

        # Contar la cantidad de unos en cada grupo
        count_nodo = np.sum(ind_ini, axis=1)
        
        # Calcular densidad, diferencias y promedios
        densidad = count_nodo / np.max(count_nodo)
        den_diff = np.insert(np.diff(densidad), 0, None)
        val_prom = np.mean(norm_rang, axis=1)
        val_diff = np.diff(val_prom)
        val_diff2 = np.insert(np.diff(val_diff), 0, None)

        # Añadir NaN's por el tamaño
        val_diff = np.insert(val_diff, 0, None)
        val_diff2 = np.insert(val_diff2, 0, None)

        # Calcular velocidad y sus diferencias
        velocidad = densidad * val_diff
        vel_diff = np.diff(velocidad)
        vel_diff2 = np.insert(np.diff(vel_diff), 0, None)

        # Añadir NaN's por el tamaño
        vel_diff = np.insert(vel_diff, 0, None)
        vel_diff2 = np.insert(vel_diff2, 0, None)

        # Calcular presión y sus diferencias
        presion = -densidad * velocidad ** 2
        pres_diff = np.insert(np.diff(presion), 0, None)

        # Calcular término de viscosidad y viscosidad
        viscosidad = np.empty(len(densidad))
        term_visc = np.empty(len(densidad))
        viscosidad_ABE = np.empty(len(densidad))
        viscosidad[:] = np.nan
        term_visc[:] = np.nan
        viscosidad_ABE[:] = np.nan
        m = 1
        for i in range(0, len(densidad)):
            if presion[i]*val_diff[i] != 0:
                term_visc[i] = velocidad[i]*(vel_diff[i]/val_diff[i])+(m/densidad[i])*((pres_diff[i]/val_diff[i])-(presion[i]/densidad[i])*(den_diff[i]/val_diff[i]))
                # print(visc_comun)
                if i + 1 < len(densidad): # and visc_comun != 0
                    viscosidad[i] = ((velocidad[i+1]-velocidad[i])+velocidad[i]*(vel_diff[i]/val_diff[i])+(m/densidad[i])*((pres_diff[i]/val_diff[i])-(presion[i]/densidad[i])*(den_diff[i]/val_diff[i])))*densidad[i]/(((2*(val_diff[i-1]*velocidad[i]-(val_diff[i-1]+val_diff[i-2])*velocidad[i-1]+val_diff[i-2]*velocidad[i-2])/(val_diff[i-1]*val_diff[i-2]*(val_diff[i-1]+val_diff[i-2])))))
                
                viscosidad_ABE[i] = term_visc[i]*densidad[i]/((2*(val_diff[i-1]*velocidad[i]-(val_diff[i-1]+val_diff[i-2])*velocidad[i-1]+val_diff[i-2]*velocidad[i-2])/(val_diff[i-1]*val_diff[i-2]*(val_diff[i-1]+val_diff[i-2]))))

        # Diferencia entre viscosidad y viscosidad ABE
        substrac_visco = viscosidad_ABE - viscosidad
        # print(substrac_visco, len(substrac_visco))
        self['time'] = tiempo_original[self._val_r-1::self._val_r]
        self['range'] = list(range(self._val_r, N+1, self._val_r))
        self['density'] = densidad
        self['dif density'] = den_diff
        self['dif values'] = val_diff
        self['velocity'] = velocidad
        self['dif vel'] = vel_diff
        self['pressure'] = presion
        self['dif pressure'] = pres_diff
        self['viscosity'] = viscosidad
        self['viscosity ABE'] = viscosidad_ABE
        self['subs visco'] = substrac_visco
         
        
        # Agrega más resultados si es necesario

