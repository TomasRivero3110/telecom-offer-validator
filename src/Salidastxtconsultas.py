# -*- coding: latin-1 -*-
#! / usr / bin / python
# vim: set fileencoding = latin-1:
import os, sys
import subprocess
import unicodedata
import tkinter.filedialog  #Libreria para crear cuadros de dialogo
import csv, sys, re, datetime, os #Librerias de python y del sistema
import xlwt, xlrd     #Librerias de gestion de archivos excel
import sys
import openpyxl
import pandas as pd
import datetime
from dateutil import parser
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Script_validacion_caracteristicas_de_paquetes')))
from ConsultaGeneralPaquetes import Paquetes   #impartando funciones de modulo de BD
from ConsultaBonosPorParent import BonosPorParent #impartando funciones de modulo de BD
from ConsultaPaquetesBono import BonosPorPaqueteDetallado #impartando funciones de modulo de BD
from openpyxl.workbook import Workbook as openpyxlWorkbook
from openpyxl.styles import PatternFill, Alignment, Font

def guardarArchivoText(nombre, objeto):
    '''Funcion guarda archivo en formato texto'''
    archivo = open(nombre,'w') 
    str1 = ',\n'.join(str(e) for e in objeto)
    archivo.write(str1)  
    archivo.close() 

if __name__ == '__main__':
    """Esta la funcion principal"""  
    
    #Consulta de los planes en su primera version en la BD TEST
    CRID = '9174576622413967267'
    lista_id_paquetes = Paquetes(CRID)
    lista_id_bonos = BonosPorParent(CRID)
    lista_id_bonos_paquetes = BonosPorPaqueteDetallado(CRID)
    #lista_id_loyalty = Loyalty(estado)
    now = datetime.datetime.now()
    fecha = now.strftime("%Y-%m-%d %H.%M")
    archivo3 = 'Lista de Paquetes en CR'+fecha+'.txt'
    #guardarArchivoText(archivo2, lista_id_equipos_Online)
    guardarArchivoText(archivo3, lista_id_paquetes)
    archivo2 = 'Lista de Bonos en CR'+fecha+'.txt'
    #guardarArchivoText(archivo2, lista_id_equipos_Online)
    guardarArchivoText(archivo2, lista_id_bonos)
    archivo1 = 'Lista de BonosPaquetes en CR'+fecha+'.txt'
    #guardarArchivoText(archivo2, lista_id_equipos_Online)
    guardarArchivoText(archivo1, lista_id_bonos_paquetes)
