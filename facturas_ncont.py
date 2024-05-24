# -*- coding: utf-8 -*-
"""
Created on Fri Jan 12 11:00:39 2024

@author: Ignacio Carvajal
"""

import os
import tkinter as tk
from tkinter import messagebox
import fitz  # PyMuPDF
import pandas as pd
import re


def quitar_guiones_y_espacios(cadena):

    if type(cadena)==list and len(cadena)>1:
        cadena = str(cadena[0])
        
    if type(cadena)==list:
        cadena = str(cadena)

    print(cadena)
    # Utilizamos la expresión regular para buscar guiones y espacios y reemplazarlos con una cadena vacía
    cadena_sin_guiones_espacios = re.sub(r'[-\s\[\]\'\']', '', cadena)

    return cadena_sin_guiones_espacios



def buscar_numero_contenedor(texto):
    # Utilizamos una expresión regular para encontrar secuencias alfanuméricas que coincidan con el formato del número de contenedor.
    patron_contenedor = re.compile(r'\b([A-Z]{4}\d{7})\b')
    patron_contenedor2 = re.compile(r'\b([A-Z]{4}\d{6}-\d)\b')
    patron_contenedor3 = re.compile(r'\b([A-Z]{4}  \d{7})\b')
    patron_contenedor4 = re.compile(r'\b([A-Z]{4} \d{6}-\d)\b')
    patron_contenedor5 = re.compile(r'\b([A-Z]{4} \d{7})\b')
    
    if len(re.findall(patron_contenedor, texto))>0:
        resultados = quitar_guiones_y_espacios(re.findall(patron_contenedor, texto))
        
    elif len(re.findall(patron_contenedor2, texto))>0:
        resultados = quitar_guiones_y_espacios(re.findall(patron_contenedor2, texto))
        
    elif len(re.findall(patron_contenedor3, texto))>0:
        
        resultados = re.findall(patron_contenedor3, texto)
       
    elif len(re.findall(patron_contenedor4, texto))>0:
        
        resultados = re.findall(patron_contenedor4, texto)[0]
        
    elif len(re.findall(patron_contenedor5, texto))>0:
        
        resultados = re.findall(patron_contenedor5, texto)[0]
    else:
        print(texto)
        resultados = "no se encontró"
        
    #print(len(re.findall(patron_contenedor, texto)))
    #print(len(re.findall(patron_contenedor2, texto)))
    #print(len(re.findall(patron_contenedor3, texto)))
    
    print(resultados)


    return resultados


def buscar_numero_factura(texto):
    # Utilizamos una expresión regular para encontrar el número de factura con el patrón específico.
    #patron_factura = re.compile(r'\b(\d+)\s*(?:Factura Electrónica|Factura no Afecta o Exenta Electrónica)\b')
    patron_factura = re.compile(r'\b(\d+)\s*(?:Factura Electrónica|Factura no Afecta o Exenta Electrónica)\b')


    # Buscar coincidencias en el texto
    resultados = re.findall(patron_factura, texto)

    return resultados

def procesar_archivo_pdf(ruta_archivo):
    #try:
    # Abrir el archivo PDF con PyMuPDF
    documento_pdf = fitz.open(ruta_archivo)
    
    for pagina_num in range(documento_pdf.page_count):
        
        # Obtener el texto de la página
        pagina = documento_pdf[pagina_num]
        texto_pagina = pagina.get_text()
        #print(texto_pagina)
       
        # Buscar números de contenedor en el texto
        numeros_contenedor = buscar_numero_contenedor(texto_pagina)
        
        numero_factura = buscar_numero_factura(texto_pagina)
        #print(numero_factura)
        # Imprimir resultados
       
        if numeros_contenedor and numero_factura:
            break
           # print(numeros_contenedor)
            #print(f"Encontrados en {ruta_archivo}, página {pagina_num + 1}: {numeros_contenedor}, nuemero_factura {numero_factura}")
        if numero_factura:
            
            break

    print('facts', numero_factura)
    return numeros_contenedor, numero_factura
    """
    except Exception as e:
        print(f"Error al procesar el archivo {ruta_archivo}: {e}")

    finally:
        # Cerrar el archivo PDF
        documento_pdf.close()
    """
def procesar_facturas(directorio_facturas):
    df_final = pd.DataFrame([])
    lista_contenedor, lista_factura = [], []
    try:
        
        # Cambiar al directorio de facturas
        os.chdir(directorio_facturas)
        
        # Listar archivos en el directorio actual
        archivos_pdf = [archivo for archivo in os.listdir() if archivo.endswith('.pdf')]
       
        # Procesar cada archivo PDF
        for archivo_pdf in archivos_pdf:
            
            ruta_archivo_pdf = os.path.join(directorio_facturas, archivo_pdf)
            
            n_cont, n_fact = procesar_archivo_pdf(ruta_archivo_pdf)
            print(n_fact,"hshs",ruta_archivo_pdf)
            lista_contenedor.append(quitar_guiones_y_espacios(n_cont))
            try:
                lista_factura.append(n_fact[0])
            except:
                lista_factura.append('no se encontro')
                print('saij')
       
        df_final['numero_factura'], df_final['numero_contenedor'] = lista_factura, lista_contenedor
   
        ruta_excel = os.path.join(os.pardir, 'resultados_facturas.xlsx')
        if os.path.exists(ruta_excel):
            os.remove(ruta_excel)
        df_final.to_excel(ruta_excel, index=False)
        
        print(df_final)
    except Exception as e:
        print(f"Error al procesar las facturas: {e}")


def procesar_facturas_desde_interfaz():
    directorio_actual = os.getcwd()
    directorio_facturas = os.path.join(directorio_actual, 'facturas')

    try:
        procesar_facturas(directorio_facturas)
        messagebox.showinfo("Proceso Completado", "Facturas procesadas correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error al procesar las facturas: {e}")

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Procesar Facturas")

# Función que se ejecuta al presionar el botón
def on_button_click():
    procesar_facturas_desde_interfaz()

# Crear un botón en la ventana
boton_procesar_facturas = tk.Button(ventana, text="Procesar Facturas", command=on_button_click)
boton_procesar_facturas.pack(pady=20)

# Iniciar el bucle principal de la interfaz gráfica
ventana.mainloop()





"""
if __name__ == "__main__":
    # Directorio actual
    directorio_actual = os.getcwd()

    # Directorio de facturas (puedes cambiarlo según tu estructura de carpetas)
    directorio_facturas = os.path.join(directorio_actual, 'facturas')

    # Procesar facturas
    procesar_facturas(directorio_facturas)
"""
