# -*- coding: utf-8 -*-
import datetime as dt
import os 
import sys
#Paquetes de terceros

import pandas as pd
import xlwings as op


archivos = {}

AHORA = dt.datetime.today()
FECHA_ACTUAL = AHORA.strftime('%d-%m-%Y %H:%M')


def definir_rutas(Path_Base):
    '''Funcion para definir rutas en el archivo base del BOT'''
    ruta_base = Path_Base.replace('\\', '/')
    archivos['Creacion'] = ruta_base + '/Input/Creacion de cuentas CAV.xlsm'
    archivos['Query'] = ruta_base + '/Output/Salida.csv'
    archivos['Query1'] = ruta_base + '/Output/Salida1.csv'
    archivos['Resultado'] = ruta_base + '/Output/Resultadocruce CAV.xlsx'
    archivos['Log'] = ruta_base + '/Log.txt'


def log_error(info):
    '''Funcion para definir un log en caso de errores en la ejecucion del script Python'''
    exc_type, exc_obj, exc_tb = info
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    msg_error = f'''
    --- ERROR ---
    Fecha y hora: {FECHA_ACTUAL}
    Archivo : {fname}
    Línea : {exc_tb.tb_lineno}
    Descripción : {exc_obj}
    Detalles : {exc_type} - {exc_tb}
    '''
    with open(archivos['Log'],'a') as log:
        log.write(msg_error)


def Leer_archivos():
    '''Funcion para leer archivos utilizados en el cruce'''
    try:
        df_creacion= pd.read_excel(archivos['Creacion'])
        
        df_query= pd.read_csv(archivos['Query'], quotechar='"')
        
        df_query1= pd.read_csv(archivos['Query1'], quotechar='"')
        
        return df_creacion,df_query,df_query1
    except Exception:
        log_error(sys.exc_info())

def Definir_cruces(df_creacion,df_query,df_query1):
    '''Funcion para realizar cruces con los archivos utilizados. Ademas se eliminan columnas para dejar solo lo que se necesita en informe'''
    try:
     
        df_cruce= pd.merge(df_creacion, df_query, on='llave')
    
    
        df_cruce.drop(['rut'], axis=1, inplace=True)
        
        df_cruce1= pd.merge(df_creacion, df_cruce,how='outer' ,on='Rut', indicator = 'union')
        
        
        df_cruce1=df_cruce1[df_cruce1['union'] == 'left_only']
        
        df_cruce1.drop(['CONCATENADO_y'], axis=1, inplace=True)
        df_cruce1.drop(['DV_y'], axis=1, inplace=True)
        df_cruce1.drop(['APELLIDO PATERNO_y'], axis=1, inplace=True)
        df_cruce1.drop(['APELLIDO MATERNO_y'], axis=1, inplace=True)
        df_cruce1.drop(['NOMBRE_y'], axis=1, inplace=True)
        df_cruce1.drop(['FECHA AFILIACION_y'], axis=1, inplace=True)
        df_cruce1.drop(['FECHA DE INCORPORACION_y'], axis=1, inplace=True)
        df_cruce1.drop(['FECHA DE NACIMIENTO_y'], axis=1, inplace=True)
        df_cruce1.drop(['SEXO_y'], axis=1, inplace=True)
        df_cruce1.drop(['TIPO TRABAJADOR_y'], axis=1, inplace=True)
        df_cruce1.drop(['CLASE DE COTIZANTE_y'], axis=1, inplace=True)
        df_cruce1.drop(['CUENTA_y'], axis=1, inplace=True)
        df_cruce1.drop(['FONDO_y'], axis=1, inplace=True)
        df_cruce1.drop(['OPTO_y'], axis=1, inplace=True)
        df_cruce1.drop(['fecha de creacion_y'], axis=1, inplace=True)
        df_cruce1.drop(['llave_y'], axis=1, inplace=True)
        df_cruce1.drop(['union'], axis=1, inplace=True)
    
        df_query2=df_query1.rename (columns= {'rut':'Rut'})
        
        
        df_cruce2= pd.merge(df_cruce1, df_query2,how='outer' ,on='Rut', indicator = 'union')
        
        
        df_cuentasnocreadas=df_cruce2[df_cruce2['union'] == 'left_only']  
        
        df_cuentasnocreadas.drop(['FECHA AFILIACION_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['FECHA DE INCORPORACION_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['FECHA DE NACIMIENTO_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['SEXO_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['TIPO TRABAJADOR_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['CLASE DE COTIZANTE_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['OPTO_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['fecha de creacion_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['llave_x'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['llave'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['producto'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['fondo'], axis=1, inplace=True)   
        df_cuentasnocreadas.drop(['union'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['ind_acreditable'], axis=1, inplace=True)
        df_cuentasnocreadas.drop(['saldo_cuota_control'], axis=1, inplace=True)
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'CONCATENADO_x':'CONCATENADO'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'DV_x':'DV'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'APELLIDO PATERNO_x':'APELLIDO PATERNO'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'APELLIDO MATERNO_x':'APELLIDO_MATERNO'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'NOMBRE_x':'NOMBRE'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'CUENTA_x':'CUENTA'})
        df_cuentasnocreadas=df_cuentasnocreadas.rename (columns= {'FONDO_x':'FONDO'})
        
        # CRUCE NUEVO
        
        df_cruce3= pd.merge(df_creacion, df_query,how='outer' ,on='llave', indicator = 'union')
        df_incorrectas=df_cruce3[df_cruce3['union'] == 'left_only']
        df_incorrectas.drop(['union'], axis=1, inplace=True)
        df_cruce4= pd.merge(df_incorrectas, df_query2,how='outer' ,on='Rut', indicator = 'union')
        df_incorrectas1=df_cruce4[df_cruce4['union'] == 'both']
        df_incorrectas1.drop(['union'], axis=1, inplace=True)
        df_incorrectas2=df_incorrectas1.rename (columns= {'llave_x':'llave'})
        df_cruce5= pd.merge(df_incorrectas2, df_query2,how='outer' ,on='llave', indicator = 'union')
        df_incorrectas3=df_cruce5[df_cruce5['union'] == 'left_only']
        df_incorrectas3.drop(['union'], axis=1, inplace=True)
        df_incorrectas4=df_incorrectas3.rename (columns= {'llave_y':'llave1'})
        df_incorrectas4.drop(['llave1'], axis=1, inplace=True)
        df_cruce6= pd.merge(df_incorrectas4, df_cruce,how='outer' ,on='llave', indicator = 'union')
        df_incorrectas5=df_cruce6[df_cruce6['union'] == 'left_only']
        df_incorrectas5.drop(['FECHA AFILIACION_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['FECHA DE INCORPORACION_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['FECHA DE NACIMIENTO_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['SEXO_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['TIPO TRABAJADOR_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['CLASE DE COTIZANTE_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['OPTO_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['fecha de creacion_x'], axis=1, inplace=True)
        df_incorrectas5.drop(['llave'], axis=1, inplace=True)
        df_incorrectas5.drop(['rut'], axis=1, inplace=True)
        df_incorrectas5.drop(['union'], axis=1, inplace=True)
        df_incorrectas5.drop(['Rut'], axis=1, inplace=True)
        df_incorrectas5.drop(['FECHA AFILIACION_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['FECHA DE INCORPORACION_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['FECHA DE NACIMIENTO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['SEXO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['TIPO TRABAJADOR_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['CLASE DE COTIZANTE_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['OPTO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['fecha de creacion_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['Rut_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['producto_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['fondo_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['ind_acreditable_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['saldo_cuota_control_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['CONCATENADO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['DV_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['APELLIDO PATERNO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['APELLIDO MATERNO_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['NOMBRE_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['CUENTA_y'], axis=1, inplace=True)
        df_incorrectas5.drop(['FONDO_y'], axis=1, inplace=True)
        df_incorrectas5=df_incorrectas5.rename (columns= {'CONCATENADO_x':'CONCATENADO'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'Rut_x':'Rut'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'DV_x':'DV'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'APELLIDO PATERNO_x':'APELLIDO PATERNO'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'APELLIDO MATERNO_x':'APELLIDO_MATERNO'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'NOMBRE_x':'NOMBRE'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'CUENTA_x':'CUENTA'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'FONDO_x':'FONDO'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'producto_x':'PRODUCTO BASE DE DATOS'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'fondo_x':'FONDO_BASE_DE_DATOS'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'ind_acreditable_x':'IND_ACREDITABLE'})
        df_incorrectas5=df_incorrectas5.rename (columns= {'saldo_cuota_control_x':'SALDO_CUOTA_CONTROL'})
        
        #CRUCE NUEVO
        
        df_cuentascreadasincorrecta=df_cruce2[df_cruce2['union'] == 'both']
        
        df_cuentascreadasincorrecta.drop(['FECHA AFILIACION_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['FECHA DE INCORPORACION_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['FECHA DE NACIMIENTO_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['SEXO_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['TIPO TRABAJADOR_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['CLASE DE COTIZANTE_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['OPTO_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['fecha de creacion_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['llave_x'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['llave'], axis=1, inplace=True)
        df_cuentascreadasincorrecta.drop(['union'], axis=1, inplace=True)
        
        return df_cuentasnocreadas,df_incorrectas5,df_cruce
    except Exception:
        log_error(sys.exc_info())
    
    
def Confeccion_excel(df_cruce,df_cuentasnocreadas,df_incorrectas5,df_creacion):
    '''Funcion para confeccionar Excel que se envia a los encargados del proceso'''
    try:
    
        app = op.App(visible=False)
        book = op.Book(archivos['Resultado'])
        
        
        hoja = book.sheets['Cuentas creadas correctamente']
        hoja.range('2:1048576').api.Delete(op.constants.DeleteShiftDirection.xlShiftUp)
        hoja.range('A1').options(header=True, index=False).value = df_cruce
    
        hoja = book.sheets['Cuentas enviadas a crear']
        hoja.range('2:1048576').api.Delete(op.constants.DeleteShiftDirection.xlShiftUp)
        hoja.range('A1').options(header=True, index=False).value = df_creacion
        
        hoja = book.sheets['Cuentas creadas incorrectas']
        hoja.range('2:1048576').api.Delete(op.constants.DeleteShiftDirection.xlShiftUp)
        hoja.range('A1').options(header=True, index=False).value = df_incorrectas5
        
        hoja = book.sheets['Cuentas no creadas']
        hoja.range('2:1048576').api.Delete(op.constants.DeleteShiftDirection.xlShiftUp)
        hoja.range('A1').options(header=True, index=False).value = df_cuentasnocreadas
        
        book.save(archivos['Resultado'])
        
        app.quit()
    except Exception:
        log_error(sys.exc_info())

def Funcion_principal(Path_Base):
    '''Funcion para devolver valor a Python'''
    
    definir_rutas(Path_Base)
    df_creacion,df_query,df_query1 = Leer_archivos()
    df_cuentasnocreadas,df_incorrectas5,df_cruce = Definir_cruces(df_creacion,df_query,df_query1)
    Confeccion_excel(df_cruce,df_cuentasnocreadas,df_incorrectas5,df_creacion)