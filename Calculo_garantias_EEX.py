# -*- coding: utf-8 -*-
"""
Created on Thu Nov 18 16:10:46 2021

@author: hlopez
"""

# Librerias
import pandas as pd
import numpy as np
import sys
# import os
# import requests
from datetime import date, timedelta, datetime
from itertools import permutations
from calendar import monthrange
pd.options.mode.chained_assignment = None


############################################################################### VARIABLES MODIFICABLES

# Porcentaje de mayoracion
mayoracion = 1.2

###############################################################################

# Inicializa el cronometro
inicio = datetime.now()

# Carpetas
# carpeta_destino = ''
carpeta_destino = ''
carpeta_clientes = ''
carpeta_posiciones_local = ''
carpeta_posiciones_servidor = ''
# carpeta_span = ''
carpeta_span = ''
carpeta_garantias_cliente = carpeta_destino + 'Garantias por cliente\\'
carpeta_garantias_totales = carpeta_destino + 'Garantias totales\\'
carpeta_garantias_it = carpeta_destino + 'Garantias IT\\'

# Inicializa el log
stdout_console = sys.stdout
sys.stdout = open(carpeta_destino+'log_manual.txt', 'w')
print('Ejecucion empezada ', inicio)

# Fechas
fecha_hoy = date.today()
# fecha_hoy = date(2022, 2, 16)
fecha_correo = fecha_hoy.strftime('%d-%m-%Y')
ano_actual = fecha_hoy.year
mes_actual = fecha_hoy.month
mes_siguiente = fecha_hoy.month + 1
fechas_restantes = pd.date_range(fecha_hoy + timedelta(days=1), date(ano_actual, 12, 31))
trimestres_restantes = list(np.unique(fechas_restantes.quarter)+1)
meses_restantes = list(np.unique(fechas_restantes.month)+1)
semanas_restantes = list(np.unique(fechas_restantes.isocalendar().week)+1)
dias_restantes = len(pd.date_range(fecha_hoy + timedelta(days=1), date(ano_actual, mes_actual, monthrange(ano_actual, mes_actual)[1])))
fechas_producto_dia = [(fecha_hoy + timedelta(days=d-3)).strftime('%d%Y%m') for d in range(4)]
lista_anos = list(range(ano_actual, ano_actual+11))
lista_codigo_anos = list(range(ano_actual+1-2000, ano_actual+11-2000)) # para los productos cal
lista_trimestres = list(range(1,5))
lista_meses = list(range(1,13))
lista_semanas = list(range(1,53))
lista_dias = list(range(1,31))
nombre_meses = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
if (fecha_hoy - timedelta(days=2)).weekday() == 6: # si la fecha de ayer cae en domingo cambia fecha al viernes
    fecha_old = (fecha_hoy - timedelta(days=4)).strftime('%Y%m%d')
else:
    fecha_old = (fecha_hoy - timedelta(days=2)).strftime('%Y%m%d')
if (fecha_hoy - timedelta(days=1)).weekday() == 6: # si la fecha de ayer cae en domingo cambia fecha al viernes
    fecha0_span = (fecha_hoy - timedelta(days=3)).strftime('%Y%m%d')
else:
    fecha0_span = (fecha_hoy - timedelta(days=1)).strftime('%Y%m%d')
fecha1_span = fecha_hoy.strftime('%Y%m%d')
if (fecha_hoy + timedelta(days=1)).weekday() == 5: # si la fecha de mañana cae en sabado cambia fecha al lunes
    fecha2_span = (fecha_hoy + timedelta(days=3)).strftime('%Y%m%d')
else:
    fecha2_span = (fecha_hoy + timedelta(days=1)).strftime('%Y%m%d')

# Archivos
# Nombre de los archivos de posiciones tanto en formato xlsx como csv
archivo_posiciones_xlsx = 'Posiciones_{}.xlsx'.format(fecha1_span)
archivo_posiciones_csv_old = 'Posiciones_{}.csv'.format(fecha0_span)
archivo_posiciones_csv = 'Posiciones_{}.csv'.format(fecha1_span)
# url de los parametros de garantias para el dia en curso y para el d-1 en caso de fallo
url_scan_ranges = 'https://public.eex-group.com/ecc/risk-management/reports-files/{}_scanningranges_{}.csv'.format(fecha1_span, fecha2_span)
url_ic_spreads = 'https://public.eex-group.com/ecc/risk-management/reports-files/{}_intercommodityspreads_{}.csv'.format(fecha1_span, fecha2_span)
url_scan_ranges_old = 'https://public.eex-group.com/ecc/risk-management/reports-files/{}_scanningranges_{}.csv'.format(fecha0_span, fecha1_span)
url_ic_spreads_old = 'https://public.eex-group.com/ecc/risk-management/reports-files/{}_intercommodityspreads_{}.csv'.format(fecha0_span, fecha1_span)
# nombre de los archivos de parametros con la fecha en curso y d-1 para usar en caso de fallo
archivo_scan_ranges_old = '{}_scanningranges_{}.csv'.format(fecha0_span, fecha1_span)
archivo_ic_spreads_old = '{}_intercommodityspreads_{}.csv'.format(fecha0_span, fecha1_span)
archivo_scan_ranges = '{}_scanningranges_{}.csv'.format(fecha1_span, fecha2_span)
archivo_ic_spreads = '{}_intercommodityspreads_{}.csv'.format(fecha1_span, fecha2_span)
archivo_garantias = '{}_Garantias_EEX_{}.xlsx'
archivo_spreads = '{}_Spreads_{}.xlsx'
archivo_clientes = 'Clientes.xlsx'
archivo_exportacion = 'Garantias_EEX_{}.csv'.format(fecha1_span)

# Variables vacias
dict_garantias = {}
tabla_garantias_totales = pd.DataFrame(columns=['CodCliente', 'Bolsa', 'Nombre', 'Alias', 'Valor', 'Code', 'Net',
                                                'Multiplicador', 'PriceScanRange', 'PriceScanRisk',
                                                'ActiveScenario', 'OriginalDelta(Gross)',
                                                'OriginalDelta(Net)', 'SCANContribution', 'IC_charge', 'Remaining_Delta',
                                                'SPAN_Requirements', 'Mayoracion', 'SPAN_Requirements_R4'])

###############################################################################
# Funciones

# Calculadora del contract value factor para los productos de EEX (multiplicador)
def calcula_cvf(cartera, mes_actual, fecha_hoy):
    '''
    Calculadora del contract value factor(cvf), que es el numero de horas que estan contenidas en cada contrato.
    Es equivalente al multiplicador de MEFF.
    
    Parameters
    ----------
    cartera : pd.DataFrame
        DESCRIPTION.

    Returns
    -------
    cartera : pd.DataFrame
        DESCRIPTION.
    '''
    lista_productos = list(np.unique(cartera.Code))
    dict_productos = dict()
    for producto in lista_productos:
        tipo = producto[2:4]
        ano = int(producto[4:8])
        mes = int(producto[8:])
        # Todos los calculos tienen control sobre el ano bisiesto
        try: # intenta convertir el tipo a numerico, si falla es porque NO es un producto dia
            int(tipo)
            # Calcula el dia de cambio de hora
            # Marzo
            ultimo_dia_marzo = date(ano, 3, 31)
            offset_marzo = (ultimo_dia_marzo.weekday() - 6) % 7
            ultimo_domingo_marzo = ultimo_dia_marzo - timedelta(days=offset_marzo)
            # Octubre
            ultimo_dia_octubre = date(ano, 10, 30)
            offset_octubre = (ultimo_dia_octubre.weekday() - 6) % 7
            ultimo_domingo_octubre = ultimo_dia_octubre - timedelta(days=offset_octubre)
            # Establece el numero de horas en funcion del dia
            if (mes == 3) & (int(tipo) == ultimo_domingo_marzo.day): numero_horas = 23
            elif (mes == 10) & (int(tipo) == ultimo_domingo_octubre.day): numero_horas = 25
            else: numero_horas = 24
        except ValueError:
            if (tipo == 'BM') & (mes == 3): numero_horas = (24 * monthrange(ano, mes)[1]) - 1 # Se quita una hora en marzo
            elif (tipo == 'BM') & (mes == 10): numero_horas = (24 * monthrange(ano, mes)[1]) + 1 # Se añade una hora en octubre
            elif (tipo == 'BM'): numero_horas = 24 * monthrange(ano, mes)[1]
            elif (tipo == 'BQ') & (mes == 1): numero_horas = 24 * (monthrange(ano, mes)[1] + monthrange(ano, mes+1)[1] + monthrange(ano, mes+2)[1]) - 1 # Se quita una hora en marzo
            elif (tipo == 'BQ') & (mes == 10): numero_horas = 24 * (monthrange(ano, mes)[1] + monthrange(ano, mes+1)[1] + monthrange(ano, mes+2)[1]) + 1 # Se añade una hora en octubre
            elif (tipo == 'BQ'): numero_horas = 24 * (monthrange(ano, mes)[1] + monthrange(ano, mes+1)[1] + monthrange(ano, mes+2)[1])
            elif (tipo == 'BY'): numero_horas = 24 * len(pd.date_range(date(ano, 1, 1), date(ano, 12, 31)))
            # A parte comprobamos si el producto es del mes en curso, en ese caso quitamos las horas gastadas
            if (tipo == 'BM') & (mes == mes_actual):
                numero_horas_gastadas = fecha_hoy.day * 24
                numero_horas -= numero_horas_gastadas
                # En caso de que un dia con cambio de hora provoque un numero negativo, lo truncamos en 0
                if numero_horas < 0: numero_horas = 0
        # Mete los valores en el diccionario
        dict_productos.update({producto: numero_horas})
    # Pasa el diccionario a dataframe para luego juntar en funcion del producto
    tabla_factores = pd.DataFrame.from_dict(dict_productos, orient='index', columns=['Multiplicador'])
    cartera = cartera.join(tabla_factores, on='Code')
    # Devuelve el resultado
    return cartera


# Calculo de los perfect spreads
def calculo_perfect_spreads(garantias):
    '''
    Parameters
    ----------
    garantias : pd.DataFrame
        Tabla que tiene que tener al menos estas columnas: Net, PriceScanRange.
        Ademas, el indice tiene que ser los codigos numericos de productos. 
    Returns
    -------
    garantias : pd.DataFrame
        DESCRIPTION.
    productos : list
        lista con todos los productos.
    compras : list
        lista con todos los productos de compra.
    ventas : list
        lista con todos los productos de venta.
    '''
    # Coge las compras y las combina con las ventas
    compras = list(garantias.index[garantias.Net>0])
    ventas = list(garantias.index[garantias.Net<0])
    productos = list(garantias.index)
    # 1. Años
    # Saca todos los productos año que pueden ser cascadeables
    productos_cascadeables_ano = [p for p in productos if 'BY' in p]
    if len(productos_cascadeables_ano)>0:
        for p in productos_cascadeables_ano:
            # Saca el año del producto
            ano = p[4:8]
            # Comprueba si el producto es compra o venta y saca los candidatos del año
            if p in compras:
                candidatos = [v for v in ventas if ano in p]
            else:
                candidatos = [c for c in compras if ano in p]
            # Ordena los valores para luego compararlos
            candidatos.sort()
            # Saca los cascadeos si los candidatos son suficientes (minimo 4 trimestres)
            if len(candidatos) >= 4: # Minimo 4 trimestres
                # Cascadeo en trimestres
                cascadeo_at = [p[:2] + 'BQ' + p[4:8] + '%02d'%(m) for m in [1, 4, 7, 10]]
                cascadeo_at.sort()
                candidatos_at = [s for s in candidatos if s in cascadeo_at]
                # Cascadeo para el primer trimestre en meses y el resto en trimestres
                cascadeo_amt = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in [1, 2, 3]] + \
                               [p[:2] + 'BQ' + p[4:8] + '%02d'%(m) for m in [4, 7, 10]]
                cascadeo_amt.sort()
                candidatos_amt = [s for s in candidatos if s in cascadeo_amt]
                # Cascadeo en meses
                cascadeo_am = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in range(1, 13)]
                cascadeo_am.sort()
                candidatos_am = [s for s in candidatos if s in cascadeo_am]
                # Compara las listas de cascadeos con la de los candidatos filtrados
                # Asi tenemos la certeza de que la lista que coincida exactamente forma un perfect spread
                # Esto se hace asi para eliminar posiciones que no sirvan para el perfect spread
                if (cascadeo_at == candidatos_at): perfect_spreads = cascadeo_at
                elif (cascadeo_amt == candidatos_amt): perfect_spreads = cascadeo_amt
                elif (cascadeo_am == candidatos_am): perfect_spreads = cascadeo_am
                else: continue
                # Añade el producto de mayor jerarquia en la lista de perfect_spreads
                perfect_spreads = perfect_spreads + [p]
                # Actualiza las garantias
                garantias.loc[perfect_spreads, 'IC_charge'] = 0.99 * garantias.loc[perfect_spreads, 'Net'].abs().min() * garantias.loc[perfect_spreads, 'PriceScanRange']
                # Actualiza los productos
                # Una vez actualizados los productos, compras y ventas no hace falta hacerlo con productos_cascadeables porque no va a encontrar candidatos
                productos = list(set(productos) - set(perfect_spreads))
                compras = list(set(compras) - set(perfect_spreads))
                ventas = list(set(ventas) - set(perfect_spreads))
            # Si no hay suficientes candidatos continua el bucle
            else: continue
    # 2. Trimestres
    # Saca todos los productos trimestre que pueden ser cascadeables
    productos_cascadeables_tri = [p for p in productos if 'BQ' in p]
    if len(productos_cascadeables_tri)>0:
        for p in productos_cascadeables_tri:
            # Saca el año del producto
            ano = p[4:8]
            # Comprueba si el producto es compra o venta y saca los candidatos del año
            if p in compras:
                candidatos = [v for v in ventas if ano in p]
            else:
                candidatos = [c for c in compras if ano in p]
            # Ordena los valores para luego compararlos
            candidatos.sort()
            # Saca los cascadeos si los candidatos son suficientes (minimo 4 trimestres)
            if len(candidatos) >= 3: # Minimo 3 meses
                # Saca los productos cascadeados de cada trimestre
                cascadeo_t1 = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in range(1, 4)]
                cascadeo_t2 = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in range(4, 7)]
                cascadeo_t3 = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in range(7, 10)]
                cascadeo_t4 = [p[:2] + 'BM' + p[4:8] + '%02d'%(m) for m in range(10, 13)]
                #Filtra los posibles candidatos
                candidatos_t1 = [s for s in candidatos if s in cascadeo_t1]
                candidatos_t2 = [s for s in candidatos if s in cascadeo_t2]
                candidatos_t3 = [s for s in candidatos if s in cascadeo_t3]
                candidatos_t4 = [s for s in candidatos if s in cascadeo_t4]
                # Compara las listas de cascadeos con la de los candidatos filtrados
                # Asi tenemos la certeza de que la lista que coincida exactamente forma un perfect spread
                # Esto se hace asi para eliminar posiciones que no sirvan para el perfect spread
                if (cascadeo_t1 == candidatos_t1): perfect_spreads = cascadeo_t1
                elif (cascadeo_t2 == candidatos_t2): perfect_spreads = cascadeo_t2
                elif (cascadeo_t3 == candidatos_t3): perfect_spreads = cascadeo_t3
                elif (cascadeo_t4 == candidatos_t4): perfect_spreads = cascadeo_t4
                else: continue
                # Añade el producto de mayor jerarquia en la lista de perfect_spreads
                perfect_spreads = perfect_spreads + [p]
                # Actualiza las garantias
                garantias.loc[perfect_spreads, 'IC_charge'] = 0.99 * garantias.loc[perfect_spreads, 'Net'].abs().min() * garantias.loc[perfect_spreads, 'PriceScanRange']
                # Actualiza los productos
                # Una vez actualizados los productos, compras y ventas no hace falta hacerlo con productos_cascadeables porque no va a encontrar candidatos
                productos = list(set(productos) - set(perfect_spreads))
                compras = list(set(compras) - set(perfect_spreads))
                ventas = list(set(ventas) - set(perfect_spreads))
            # Si no hay suficientes candidatos continua el bucle
            else: continue
    # Devuelve el resultado
    return garantias, productos, compras, ventas


# Actualizacion del IC credit
def calcula_ic_credit(df):
    '''
    Esta funcion actualiza mediante un bucle los valores de la delta y recalcula despues el IC credit de cada par de commodities 
    
    Parameters
    ----------
    df : pd.DataFrame
        DESCRIPTION.

    Returns
    -------
    df : pd.DataFrame
        DESCRIPTION.

    '''
    for f in range(len(df)):
        com_a = df['CommodityA'][f]
        com_b = df['CommodityB'][f]
        # Hay que actualizar las delas de los productos tanto en commodity A como B
        df['OriginalDelta(Net)_A'][f+1:][df['CommodityA'] == com_a] = df['Remaining_Delta_A'][f]
        df['OriginalDelta(Net)_B'][f+1:][df['CommodityB'] == com_b] = df['Remaining_Delta_B'][f]
        df['OriginalDelta(Net)_B'][f+1:][df['CommodityB'] == com_a] = df['Remaining_Delta_A'][f]
        df['OriginalDelta(Net)_A'][f+1:][df['CommodityA'] == com_b] = df['Remaining_Delta_B'][f]
        # Recalcula los IC charges
        df['IC_charge_A'] = abs((df[['OriginalDelta(Net)_A', 'PriceScanRange_A', 'Credit']].prod(axis=1)) / df['Multiplicador_A'])
        df['IC_charge_B'] = abs((df[['OriginalDelta(Net)_B', 'PriceScanRange_B', 'Credit']].prod(axis=1)) / df['Multiplicador_B'])
        df['IC_charge'] = df[['IC_charge_A', 'IC_charge_B']].min(axis=1).round(0)
        # Recalcula los remaining delta
        df['Remaining_Delta_A'] = df['OriginalDelta(Net)_A'].subtract(np.sign(df['OriginalDelta(Net)_A']) * df[['IC_charge', 'Multiplicador_A']].prod(axis=1) / df[['PriceScanRange_A', 'Credit']].prod(axis=1)).round(4)
        df['Remaining_Delta_B'] = df['OriginalDelta(Net)_B'].subtract(np.sign(df['OriginalDelta(Net)_B']) * df[['IC_charge', 'Multiplicador_B']].prod(axis=1) / df[['PriceScanRange_B', 'Credit']].prod(axis=1)).round(4)
    # Devuelve el resultado
    return df



#%% CELDA 1

# POSICIONES CLIENTES

# Carga la tabla de clientes (codigo bolas, nombre y alias)
tabla_clientes = pd.read_excel(carpeta_clientes+archivo_clientes, index_col=0)

# Carga la tabla de posiciones de clientes
try: # fichero del dia en curso
    try: # intenta en la carpeta del servidor
        tabla_posiciones = pd.read_csv(carpeta_posiciones_servidor+archivo_posiciones_csv, sep=';', thousands='.', decimal=',', encoding='latin-1', skipfooter=2, engine='python')
        print('Obtenido archivo de posiciones del dia en curso desde el servidor')
    except FileNotFoundError: # busca en la carpeta local
        print('No encontrado archivo de posiciones del dia en curso en el servidor')
        tabla_posiciones = pd.read_csv(carpeta_posiciones_local+archivo_posiciones_csv, sep=';', thousands='.', decimal=',', encoding='latin-1', skipfooter=2, engine='python')
        print('Obtenido archivo de posiciones del dia en curso desde local')
except FileNotFoundError: # fichero del dia anterior
    print('No encontrado el archivo de posiciones del dia en curso ni en el servidor ni en local')
    try: # intenta en la carpeta del servidor
        tabla_posiciones = pd.read_csv(carpeta_posiciones_servidor+archivo_posiciones_csv_old, sep=';', thousands='.', decimal=',', encoding='latin-1', skipfooter=2, engine='python')
        print('Obtenido archivo de posiciones del dia anterior desde el servidor')
    except FileNotFoundError: # busca en la carpeta local
        print('No encontrado archivo de posiciones del dia anterior en el servidor')
        tabla_posiciones = pd.read_csv(carpeta_posiciones_local+archivo_posiciones_csv_old, sep=';', thousands='.', decimal=',', encoding='latin-1', skipfooter=2, engine='python')
        print('Obtenido archivo de posiciones del dia anterior desde local')
        
# Tabla scanning ranges +  intercommodity spreads
try: # intenta sacar los archivos del dia en curso 
    try: # desde la url
        tabla_scan_ranges = pd.read_csv(url_scan_ranges, sep=';', decimal=',', thousands='.')
        tabla_ic_spreads = pd.read_csv(url_ic_spreads, sep=';', decimal='.')
        print('Obtenidos parametros de span del dia en curso')
        # Exporta los archivos para tenerlos al dia siguiente en caso de fallo
        tabla_scan_ranges.to_csv(carpeta_span + archivo_scan_ranges, sep=';', decimal=',', index=False)
        tabla_ic_spreads.to_csv(carpeta_span + archivo_ic_spreads, sep=';', decimal='.', index=False)
        print('Exportados parametros de span del dia en curso desde la url')
    except: # desde la carpeta del servidor
        tabla_scan_ranges = pd.read_csv(carpeta_span + archivo_scan_ranges, sep=';', decimal=',', thousands='.')
        tabla_ic_spreads = pd.read_csv(carpeta_span + archivo_ic_spreads, sep=';')
        print('Obtenidos parametros de span del dia en curso desde archivos csv')
except:
    print('No existen los parametros de span del dia en curso')
    try: # si no existen intenta con los del dia anterior
        tabla_scan_ranges = pd.read_csv(url_scan_ranges_old, sep=';', decimal=',', thousands='.')
        tabla_ic_spreads = pd.read_csv(url_ic_spreads_old, sep=';', decimal='.')
        print('Obtenidos parametros de span del dia dia anterior desde la url')
    except: # si estos tampoco existen, intenta los del dia anterior desde la carpeta para asegurarnos de que hay algun archivo de parametros
        tabla_scan_ranges = pd.read_csv(carpeta_span + archivo_scan_ranges_old, sep=';', decimal=',', thousands='.')
        tabla_ic_spreads = pd.read_csv(carpeta_span + archivo_ic_spreads_old, sep=';')
        print('Obtenidos parametros de span del dia anterior desde archivos csv')


        
#%% CELDA 2

# GARANTIAS ECC

# Estructura la tabla de posiciones
tabla_posiciones_ecc = tabla_posiciones[tabla_posiciones.BROKER =='ECC']
tabla_posiciones_ecc['Alias'] = list(tabla_clientes.loc[tabla_posiciones_ecc.CODIGO, 'Alias'])
tabla_posiciones_ecc['Product_ID'] = [i[:4] for i in tabla_posiciones_ecc.VALOR]
tabla_posiciones_ecc['Expiry_Year'] = ['20' + i[7:] for i in tabla_posiciones_ecc.VALOR]
tabla_posiciones_ecc['Expiry_Month'] = ['%02d'%(nombre_meses.index(i[4:7]) + 1) for i in tabla_posiciones_ecc.VALOR]
tabla_posiciones_ecc['Code'] = tabla_posiciones_ecc.Product_ID + tabla_posiciones_ecc.Expiry_Year + tabla_posiciones_ecc.Expiry_Month
tabla_posiciones_ecc['Net'] = tabla_posiciones_ecc.COMPRADAS - tabla_posiciones_ecc.VENDIDAS
tabla_posiciones_ecc = calcula_cvf(tabla_posiciones_ecc, mes_actual, fecha_hoy)
# Selecciona y ordena las columnas que necesitamos
tabla_posiciones_ecc = tabla_posiciones_ecc[['CODIGO', 'C. BOLSA', 'NOMBRE', 'Alias', 'VALOR', 'Code', 'Net', 'Multiplicador']]
# Cambia los nombres de las columnas
tabla_posiciones_ecc = tabla_posiciones_ecc.rename(columns={'CODIGO': 'CodCliente', 'C. BOLSA': 'Bolsa', 
                                                            'NOMBRE': 'Nombre', 'VALOR': 'Valor'})

# Clientes por numero de bolsa y por alias
lista_codigos = list(np.unique(tabla_posiciones_ecc['CodCliente']))

# 1. Calculo del scan risk general para cada combined commodity
# Genera el codigo de los productos
tabla_scan_ranges['Expiry_Month'] = ['%02d'%(m) for m in tabla_scan_ranges['Expiry_Month']]
tabla_scan_ranges['Code'] = tabla_scan_ranges.Product_ID + tabla_scan_ranges.Expiry_Year.astype(str) + tabla_scan_ranges.Expiry_Month.astype(str)
# Tabla de garantias
tabla_garantias = tabla_posiciones_ecc.join(tabla_scan_ranges[['Code', 'PriceScanRange']].set_index('Code'), on='Code')
tabla_garantias['PriceScanRisk'] = tabla_garantias['PriceScanRange'] / tabla_garantias['Multiplicador']
tabla_garantias['ActiveScenario'] = np.sign(tabla_garantias.Net) * tabla_garantias.PriceScanRange
tabla_garantias['OriginalDelta(Gross)'] = tabla_garantias[['Multiplicador', 'Net']].abs().prod(axis=1)
tabla_garantias['OriginalDelta(Net)'] = tabla_garantias[['Multiplicador', 'Net']].prod(axis=1)
tabla_garantias['SCANContribution'] = tabla_garantias[['Net', 'ActiveScenario']].prod(axis=1)
tabla_garantias['IC_charge'] = 0
tabla_garantias['Remaining_Delta'] = 0
# Revisa si hay productos dia del dia en curso, ya que estos no cuentan para garantias
elimina_prod_dia = [i for i in range(len(tabla_garantias)) if tabla_garantias['Code'][i][2:] in fechas_producto_dia]
tabla_garantias['SCANContribution'][elimina_prod_dia] = 0
print('1. Calculo del scan risk realizado con exito')

# 2. Calculo del intercommodity reduction
# 2.1. Tabla intercommodity spreads
tabla_ic_spreads = tabla_ic_spreads[~(tabla_ic_spreads.ExpiryYearA.isna() | tabla_ic_spreads.ExpiryYearB.isna())]
tabla_ic_spreads['ExpiryMonthA'] = ['%02d'%(int(m)) for m in tabla_ic_spreads['ExpiryMonthA']]
tabla_ic_spreads['ExpiryMonthB'] = ['%02d'%(int(m)) for m in tabla_ic_spreads['ExpiryMonthB']]
tabla_ic_spreads[['ExpiryYearA', 'ExpiryYearB']] = tabla_ic_spreads[['ExpiryYearA', 'ExpiryYearB']].astype(int).astype(str)
tabla_ic_spreads['IC_Code'] = tabla_ic_spreads[['CombinedCommodityA', 'ExpiryYearA', 'ExpiryMonthA']].agg(''.join, axis=1) \
                              + '-' \
                              + tabla_ic_spreads[['CombinedCommodityB', 'ExpiryYearB', 'ExpiryMonthB']].agg(''.join, axis=1)
                              
# 2.2. Combinaciones de productos
tuple_combinaciones = list(permutations(np.unique(tabla_garantias.Code), 2))
lista_combinaciones = [c[0]+'-'+c[1] for c in tuple_combinaciones]
tabla_intercommodity = pd.DataFrame({'CommodityA': [c[0] for c in tuple_combinaciones], 'CommodityB':  [c[1] for c in tuple_combinaciones]}, index=lista_combinaciones)
tabla_intercommodity = tabla_intercommodity.join(tabla_ic_spreads[['IC_Code', 'RatioA', 'RatioB', 'Credit']].set_index('IC_Code')).fillna(0)

# 2.3. Bucle por cliente (codigo)
for codigo in lista_codigos:
    '''
    codigo = 53637032
    '''    
    # Obtiene las garantias en funcion de la bolsa o del alias
    garantias = tabla_garantias[tabla_garantias['CodCliente'] == codigo]
    garantias.set_index('Code', inplace=True)
    # Evalua si hay posiciones contrarias y combina las compras con las ventas
    # 1. Perfect spreads
    if (garantias.Net>0).any() & (garantias.Net<0).any():
        garantias, productos, compras, ventas = calculo_perfect_spreads(garantias)
    # 2. Intercommodity spreads
    # Inicializa la tabla de intercommodity credit
    ic_credit = pd.DataFrame()
    if (garantias.Net>0).any() & (garantias.Net<0).any():
        '''
        # Coge las compras y las combina con las ventas
        compras = list(garantias.index[garantias.Net>0])
        ventas = list(garantias.index[garantias.Net<0])
        productos = list(garantias.index)
        '''
        lista_spreads = [p[0]+'-'+p[1] for p in permutations(productos, 2)]
        lista_perm_compras = [p[0]+'-'+p[1] for p in permutations(compras, 2)]
        lista_perm_ventas = [p[0]+'-'+p[1] for p in permutations(ventas, 2)]
        lista_eliminar = lista_perm_compras + lista_perm_ventas
        lista_spreads = list(set(lista_spreads) - set(lista_eliminar))
        tabla_spreads = pd.DataFrame(index=lista_spreads)
        # Busca en la tabla de intercommodity 
        tabla_spreads = tabla_spreads.join(tabla_intercommodity)
        tabla_spreads = tabla_spreads[tabla_spreads.Credit != 0]
        tabla_spreads = tabla_spreads.join(garantias.iloc[:, 5:], on='CommodityA')
        tabla_spreads = tabla_spreads.merge(garantias.iloc[:, 5:], left_on='CommodityB', right_index=True, suffixes=['_A', '_B'])
        tabla_spreads.sort_values('Credit', ascending=False, inplace=True)
        tabla_spreads['IC_charge_A'] = abs((tabla_spreads[['OriginalDelta(Net)_A', 'PriceScanRange_A', 'Credit']].prod(axis=1))/tabla_spreads['Multiplicador_A'])
        tabla_spreads['IC_charge_B'] = abs((tabla_spreads[['OriginalDelta(Net)_B', 'PriceScanRange_B', 'Credit']].prod(axis=1))/tabla_spreads['Multiplicador_B'])
        tabla_spreads['IC_charge'] = tabla_spreads[['IC_charge_A', 'IC_charge_B']].min(axis=1).round(0)
        tabla_spreads['Remaining_Delta_A'] = tabla_spreads['OriginalDelta(Net)_A'].subtract(np.sign(tabla_spreads['OriginalDelta(Net)_A'])*tabla_spreads[['IC_charge', 'Multiplicador_A']].prod(axis=1)/tabla_spreads[['PriceScanRange_A', 'Credit']].prod(axis=1)).round(4)
        tabla_spreads['Remaining_Delta_B'] = tabla_spreads['OriginalDelta(Net)_B'].subtract(np.sign(tabla_spreads['OriginalDelta(Net)_B'])*tabla_spreads[['IC_charge', 'Multiplicador_B']].prod(axis=1)/tabla_spreads[['PriceScanRange_B', 'Credit']].prod(axis=1)).round(4)
        tabla_spreads = calcula_ic_credit(tabla_spreads)
        # Agrupa y suma los spreads para meterlo en la tabla de garantias --> no tiene sentido porque hay productos en clase A y B
        # tabla_spreads = tabla_spreads[tabla_spreads['IC_charge'] != 0]
        # Mete los valores de IC credit y remaning delta en la tabla de garantias
        for producto in garantias.index:
            garantias['IC_charge'][producto] = tabla_spreads.loc[(tabla_spreads['CommodityA']==producto) | (tabla_spreads['CommodityB']==producto), 'IC_charge'].sum()
            garantias['Remaining_Delta'][producto] = np.sign(tabla_spreads.loc[(tabla_spreads['CommodityA']==producto) | (tabla_spreads['CommodityB']==producto), 'Remaining_Delta_A'].min()) * tabla_spreads.loc[(tabla_spreads['CommodityA']==producto) | (tabla_spreads['CommodityB']==producto), 'Remaining_Delta_A'].abs().min()
        # tabla_spreads.to_excel(r'C:\Users\hlopez\Documents\Pruebas Garantias EEX'+'\Spreads.xlsx')
    # Calcula las garantias totales = contribucion de cada producto - la minoracion de los intercommodity
    garantias['SPAN_Requirements'] = garantias['SCANContribution'] - garantias['IC_charge']
    # garantias['Mayoracion'] = np.where((garantias['Alias'] == 'Alcanzia') | (garantias['Alias'] == 'Estabanell'), 1.5, mayoracion)
    # garantias['SPAN_Requirements_R4'] = garantias[['SPAN_Requirements', 'Mayoracion']].prod(axis=1)
    
    # Mete en la tabla de garantias totales
    tabla_garantias_totales = tabla_garantias_totales.append(garantias.reset_index()).fillna(0)
    # Exportacion a excel
    garantias.to_excel(carpeta_garantias_cliente + archivo_garantias.format(fecha1_span, tabla_clientes['Alias'][codigo]))    

    # Guarda en un diccionario las garantias de cada bolsa
    try:
        dict_garantias.update({tabla_clientes['Alias'][codigo]: garantias})
    except:
        dict_garantias.update({codigo: garantias})
        
print('2. Calculo del intercommodity charge realizado con exito')

# Agrega las garantias por cliente
# tabla_exportacion = tabla_garantias_totales[['Bolsa', 'Alias', 'SPAN_Requirements', 'SPAN_Requirements_R4', 'Efectivo', 'Garantias+Efectivo']].reset_index(drop=True).groupby(['Bolsa', 'Alias']).sum()
tabla_exportacion = tabla_garantias_totales[['Bolsa', 'Alias', 'SPAN_Requirements']].reset_index(drop=True).groupby(['Bolsa', 'Alias']).sum()
tabla_exportacion.sort_index(level='Alias', inplace=True)

# Exporta las tablas de garantias
tabla_garantias_totales.to_excel(carpeta_garantias_totales + archivo_garantias.format(fecha1_span, 'Garantias_totales'))
tabla_exportacion.to_csv(carpeta_garantias_it + archivo_exportacion, sep=';', decimal=',')
print('Exportacion de garantias totales realizada con exito')



#%% CELDA 3

# 3. Calculo de las garantias para la cuenta omnibus
tabla_garantias_om = tabla_posiciones_ecc.drop(['CodCliente', 'Bolsa', 'Nombre', 'Alias', 'Multiplicador'], axis=1)
tabla_garantias_om = tabla_garantias_om.groupby(['Valor', 'Code']).sum().reset_index()
tabla_garantias_om = calcula_cvf(tabla_garantias_om, mes_actual, fecha_hoy)
tabla_garantias_om = tabla_garantias_om.join(tabla_scan_ranges[['Code', 'PriceScanRange']].set_index('Code'), on='Code')
tabla_garantias_om['PriceScanRisk'] = tabla_garantias_om['PriceScanRange']/ tabla_garantias_om['Multiplicador']
tabla_garantias_om['ActiveScenario'] = np.sign(tabla_garantias_om.Net) * tabla_garantias_om.PriceScanRange
tabla_garantias_om['OriginalDelta(Gross)'] = tabla_garantias_om[['Multiplicador', 'Net']].prod(axis=1).abs()
tabla_garantias_om['OriginalDelta(Net)'] = tabla_garantias_om[['Multiplicador', 'Net']].prod(axis=1)
tabla_garantias_om['SCANContribution'] = tabla_garantias_om[['Net', 'ActiveScenario']].prod(axis=1)
tabla_garantias_om.set_index('Code', inplace=True)
tabla_garantias_om['IC_charge'] = 0
tabla_garantias_om['Remaining_Delta'] = 0
# 1. Perfect spreads
if (tabla_garantias_om.Net>0).any() & (tabla_garantias_om.Net<0).any():
    garantias, productos, compras, ventas = calculo_perfect_spreads(tabla_garantias_om)
# Coge las compras y las combina con las ventas
'''
compras = list(tabla_garantias_om.index[tabla_garantias_om.Net>0])
ventas = list(tabla_garantias_om.index[tabla_garantias_om.Net<0])
productos = list(tabla_garantias_om.index)
'''
# 2. Intercommodity spreads
# Inicializa la tabla de intercommodity credit
ic_credit = pd.DataFrame()
if (garantias.Net>0).any() & (garantias.Net<0).any():
    lista_spreads = [p[0]+'-'+p[1] for p in permutations(productos, 2)]
    lista_perm_compras = [p[0]+'-'+p[1] for p in permutations(compras, 2)]
    lista_perm_ventas = [p[0]+'-'+p[1] for p in permutations(ventas, 2)]
    lista_eliminar = lista_perm_compras + lista_perm_ventas
    lista_spreads = list(set(lista_spreads) - set(lista_eliminar))
    tabla_spreads = pd.DataFrame(index=lista_spreads)
    tabla_spreads = tabla_spreads.join(tabla_intercommodity)
    # Busca en la tabla de intercommodity
    tabla_spreads = tabla_spreads[tabla_spreads.Credit != 0]
    tabla_spreads = tabla_spreads.join(tabla_garantias_om, on='CommodityA')
    tabla_spreads = tabla_spreads.merge(tabla_garantias_om, left_on='CommodityB', right_index=True, suffixes=['_A', '_B'])
    tabla_spreads.sort_values('Credit', ascending=False, inplace=True)
    tabla_spreads['IC_charge_A'] = abs((tabla_spreads[['OriginalDelta(Net)_A', 'PriceScanRange_A', 'Credit']].prod(axis=1)*100)/tabla_spreads['Multiplicador_A'])
    tabla_spreads['IC_charge_B'] = abs((tabla_spreads[['OriginalDelta(Net)_B', 'PriceScanRange_B', 'Credit']].prod(axis=1)*100)/tabla_spreads['Multiplicador_B'])
    tabla_spreads['IC_charge'] = tabla_spreads[['IC_charge_A', 'IC_charge_B']].min(axis=1).round(0)
    tabla_spreads['Remaining_Delta_A'] = tabla_spreads['OriginalDelta(Net)_A'].subtract(0.01*np.sign(tabla_spreads['OriginalDelta(Net)_A'])*tabla_spreads[['IC_charge', 'Multiplicador_A']].prod(axis=1)/tabla_spreads[['PriceScanRange_A', 'Credit']].prod(axis=1)).round(4)
    tabla_spreads['Remaining_Delta_B'] = tabla_spreads['OriginalDelta(Net)_B'].subtract(0.01*np.sign(tabla_spreads['OriginalDelta(Net)_B'])*tabla_spreads[['IC_charge', 'Multiplicador_B']].prod(axis=1)/tabla_spreads[['PriceScanRange_B', 'Credit']].prod(axis=1)).round(4)
    tabla_spreads = calcula_ic_credit(tabla_spreads)
    # Agrupa y suma los spreads para meterlo en la tabla de garantias
    tabla_spreads = tabla_spreads[tabla_spreads['IC_charge'] != 0]
    # Mete los valores de IC credit y remaning delta en la tabla de tabla_garantias_om
    for producto in tabla_garantias_om.index:
        if (tabla_spreads['CommodityA'] == producto).any():
            tabla_garantias_om['IC_charge'][producto] = tabla_spreads.loc[tabla_spreads['CommodityA'] == producto, 'IC_charge'].sum()
            tabla_garantias_om['Remaining_Delta'][producto] = np.sign(tabla_spreads.loc[tabla_spreads['CommodityA'] == producto, 'Remaining_Delta_A'].min()) * tabla_spreads.loc[tabla_spreads['CommodityA'] == producto, 'Remaining_Delta_A'].abs().min()
        elif (tabla_spreads['CommodityB'] == producto).any():
            tabla_garantias_om['IC_charge'][producto] = tabla_spreads.loc[tabla_spreads['CommodityB'] == producto, 'IC_charge'].sum()
            tabla_garantias_om['Remaining_Delta'][producto] = np.sign(tabla_spreads.loc[tabla_spreads['CommodityB'] == producto, 'Remaining_Delta_B'].min()) * tabla_spreads.loc[tabla_spreads['CommodityB'] == producto, 'Remaining_Delta_B'].abs().min()
        else: 
            continue
    
# Calcula las tabla_garantias_om totales = contribucion de cada producto - la minoracion de los intercommodity
tabla_garantias_om['SPAN_Requirements'] = tabla_garantias_om['SCANContribution'] - tabla_garantias_om['IC_charge']
print('3. Calculo de las garantias de la cuenta omnibus realizado con exito')

# Exportacion a excel
tabla_garantias_om.to_excel(carpeta_garantias_totales + archivo_garantias.format(fecha1_span, 'Omnibus'))
print('Exportacion de garantias omnibus realizada con exito')

# Tiempo de ejecucion
final = datetime.now()
tiempo_ejecucion = (final - inicio)
print('Ejecucion finalizada ', final)
print('Tiempo de ejecucion:', tiempo_ejecucion)

# Cierra el log
sys.stdout.close()
# Restaura sys.stdout al handler por defecto
sys.stdout = stdout_console



