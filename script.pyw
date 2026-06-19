# -*- coding: utf-8 -*-

"""
PROYECTO: APR Receipt Normalizer
COPYRIGHT: (C) 2026 Vicente Ferrer
LICENCIA: PROPIETARIA - TODOS LOS DERECHOS RESERVADOS
"""

import pandas as pd
import pdfplumber
import os
import re
from unidecode import unidecode
from typing import List, Optional
from dotenv import load_dotenv

load_dotenv()

# --- CONFIGURACIÓN ---
EXCEL_FILE = os.getenv('EXCEL_FILE')
SHEET_NAME = os.getenv('SHEET_NAME')
PDF_FOLDER = os.getenv('PDF_FOLDER')
COL_MEDIDOR = os.getenv('COL_MEDIDOR')
COL_NOMBRE = os.getenv('COL_NOMBRE')
COL_APELLIDO1 = os.getenv('COL_APELLIDO1')
COL_APELLIDO2 = os.getenv('COL_APELLIDO2')


def normalize_text(text: str) -> str:
    """Convierte el texto a mayúsculas y elimina tildes/acentos para la búsqueda."""
    return unidecode(text).upper().strip()

def extract_name_from_pdf(pdf_path: str) -> Optional[str]:
    """
    Abre el PDF, extrae las dos líneas relevantes.
    PRIORIDAD 1: Busca 'CONSUMO' o 'CONSUMO AGUA'.
    PRIORIDAD 2: Si falla, busca la cadena 'Dirección: Santiago'.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()
            
            # --- 1. PRIORIDAD 1: Búsqueda de Consumo ---
            # Patrón flexible: CONSUMO o CONSUMO AGUA
            match_start = re.search(r'CONSUMO\s*(AGUA)?', text)
            
            keyword_found = True
            if not match_start:
                # --- 2. PRIORIDAD 2: Búsqueda de Respaldo (Dirección: Santiago) ---
                print("   [RESPALDO] No se encontró 'CONSUMO'. Buscando 'Dirección: Santiago'...")
                match_start = re.search(r'Dirección: Santiago', text)
                keyword_found = False # Marcamos que estamos usando el respaldo
            
            if not match_start:
                print(f"   [ERROR] No se encontró ninguna palabra clave de inicio en el PDF: {os.path.basename(pdf_path)}")
                return None
            
            text_after_keyword = text[match_start.end():]
            
            # Buscamos solo las líneas que no están vacías después de la palabra clave
            all_lines = [line.strip() for line in text_after_keyword.split('\n') if line.strip()]
            
            if not all_lines:
                 print(f"   [ERROR] No hay texto relevante después de la palabra clave en: {os.path.basename(pdf_path)}")
                 return None

            # 3. Capturar las dos primeras líneas y concatenarlas
            # (Se mantienen 2 líneas como se acordó previamente)
            lines_to_process = all_lines[:2] 
            potential_name_line = " ".join(lines_to_process)
            
            # --- 4. Limpieza del Nombre (igual que antes) ---
            
            # Limpieza de Separadores (Guion, Barra, Dos Puntos) 
            cleaned_name = potential_name_line.replace('-', ' ').replace('/', ' ').replace(':', ' ').strip()
            cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip() # Normaliza espacios múltiples
            
            # Limpieza de Honoríficos (SR, SRA, DON, DOÑA)
            honorific_pattern = r'^(SR|DO(N|Ñ)?)A?\.?\s*'
            cleaned_name = re.sub(honorific_pattern, '', cleaned_name, flags=re.IGNORECASE).strip()
            
            # Limpieza de Precios y Símbolos ($)
            price_pattern = r'\$?\s*[\d\.\,]+\s*\$?'
            cleaned_name = re.sub(r'^' + price_pattern, '', cleaned_name).strip()
            cleaned_name = re.sub(price_pattern + r'$', '', cleaned_name).strip()
            cleaned_name = re.sub(r'(\$\s*|\s*\$)', ' ', cleaned_name).strip()
            cleaned_name = re.sub(r'\d+', ' ', cleaned_name).strip() 
            cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip() 
            
            return cleaned_name
            
    except Exception as e:
        print(f"   [ERROR] Falló la lectura del PDF {os.path.basename(pdf_path)}: {e}")
        return None
    
def extract_id_from_pdf(pdf_path: str) -> Optional[int]:
    """
    Busca la etiqueta del medidor y extrae el número subsiguiente, 
    tolerando espacios irregulares tanto en la etiqueta como en el ID numérico.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text()
            
            match = re.search(r'(?:IDOR\s*N°|M\s*e\s*d\s*i\s*d\s*o\s*r)\s*[:\-]?\s*([\d\s]+)', text, re.IGNORECASE)

            if match:
                string_raw: str = match.group(1)
                string_clean: str = re.sub(r'\s+', '', string_raw)
                
                return int(string_clean) if string_clean.isdigit() else None
            
            print(f"   [ADVERTENCIA] No se encontró 'MEDIDOR N°' en {os.path.basename(pdf_path)}")
            return None
            
    except Exception as e:
        print(f"   [ERROR] Error al extraer ID del PDF {pdf_path}: {e}")
        return None

def find_medidor_and_rename(df: pd.DataFrame, pdf_folder: str):
    """
    Versión robusta para asegurar el match de medidores.
    """
    # 1. NORMALIZACIÓN CRÍTICA DEL EXCEL:
    # Convertimos a string, eliminamos el .0 (si es float) y quitamos espacios.
    df[COL_MEDIDOR] = (
        df[COL_MEDIDOR]
        .astype(str)
        .str.replace(r'\.0$', '', regex=True) # Elimina .0 al final
        .str.strip()
    )
    
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
    
    for filename in pdf_files:
        old_path = os.path.join(pdf_folder, filename)
        
        # 2. Extracción (Asegúrate que devuelva string)
        id_medidor_raw = extract_id_from_pdf(old_path)
        
        if id_medidor_raw is None:
            continue
            
        # 3. NORMALIZACIÓN DEL DATO DEL PDF:
        # Forzamos a string y quitamos espacios.
        id_medidor = str(id_medidor_raw).strip()
        
        # 4. BÚSQUEDA CON DEPURACIÓN (Debug)
        match = df[df[COL_MEDIDOR] == id_medidor]
        
        if match.empty:
            # Línea de depuración para ver qué está pasando realmente:
            print(f"   [DEBUG] Buscando medidor: '{id_medidor}' (tipo: {type(id_medidor)})")
            print(f"   [DEBUG] Primeros valores en Excel: {df[COL_MEDIDOR].head(3).tolist()}")
            print(f"   [NO ENCONTRADO] Medidor {id_medidor} no existe.\n")
            continue
            
        # 3. Obtener datos para el nuevo nombre (tomamos el primero si hay duplicados)
        row = match.iloc[0]
        nombre = normalize_text(str(row[COL_NOMBRE])).replace(' ', '_')
        apellido = normalize_text(str(row[COL_APELLIDO1])).replace(' ', '_')
        
        # 4. Formatear nombre de archivo: [Número de Medidor]_[Nombre]_[Primer Apellido]
        # Usamos :03d o similar si quieres ceros a la izquierda, o simplemente el string
        new_filename = f"{id_medidor}_{nombre}_{apellido}.pdf"
        new_path = os.path.join(pdf_folder, new_filename)
        
        # 5. Renombrar
        try:
            if os.path.exists(new_path):
                print(f"   [EXISTE] El archivo {new_filename} ya existe. Saltando...")
            else:
                os.rename(old_path, new_path)
                print(f"   ✅ {id_medidor}: {nombre} {apellido}")
        except Exception as e:
            print(f"   [ERROR] Al renombrar {filename}: {e}")

# --- FUNCIÓN PRINCIPAL ---
def main():
    print("--- 📄 INICIANDO SCRIPT DE ORGANIZACIÓN DE BOLETAS APR 💧 ---")
    
    # Verificar si el archivo Excel existe
    if not os.path.exists(EXCEL_FILE):
        print(f"🚨 ERROR: El archivo Excel '{EXCEL_FILE}' no se encontró en la carpeta raíz.")
        return

    # Verificar si la carpeta de PDFs existe
    if not os.path.exists(PDF_FOLDER):
        print(f"🚨 ERROR: La carpeta de PDFs '{PDF_FOLDER}' no se encontró en la carpeta raíz.")
        return

    try:
        # Abrir libro de Excel en la hoja seteada (paso 1)
        # Se asume que el archivo tiene una cabecera
        df_excel = pd.read_excel(
            EXCEL_FILE, 
            sheet_name=SHEET_NAME, 
            engine='openpyxl'
        )
        print(f"✅ Libro de Excel abierto exitosamente en la hoja '{SHEET_NAME}'.")
        
        # Validar que las columnas necesarias existan
        required_cols = [COL_MEDIDOR, COL_NOMBRE, COL_APELLIDO1, COL_APELLIDO2]
        for col in required_cols:
            if col not in df_excel.columns:
                print(f"🚨 ERROR: La columna requerida '{col}' no se encontró en la hoja de Excel. Revise la CONFIGURACIÓN.")
                print(df_excel.columns)
                return

        # Procesar los archivos
        find_medidor_and_rename(df_excel, PDF_FOLDER)
        
    except FileNotFoundError:
        print(f"🚨 ERROR: El archivo '{EXCEL_FILE}' o la hoja '{SHEET_NAME}' no se encontraron.")
    except Exception as e:
        print(f"🚨 Ocurrió un error inesperado: {e}")
        
    print("\n--- ✅ PROCESO COMPLETADO ---")

if __name__ == "__main__":
    main()