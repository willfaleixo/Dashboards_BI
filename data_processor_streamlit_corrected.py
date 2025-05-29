# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import time # Import time for retry delay
import shutil # Import shutil for file copy
import traceback # For detailed error logging

import streamlit as st

@st.cache_data(show_spinner=False)
def load_and_clean_data_streamlit_cached(file_path, is_csv=False):
    return load_and_clean_data_streamlit(file_path, is_csv)

def load_and_clean_data_streamlit(file_path, is_csv=False, retries=3, delay=3):
    """Loads data from the specified file (Excel or CSV), cleans it, and prepares it for the Streamlit dashboard."""
    print(f"File path to load: {file_path}")
    print(f"Is CSV: {is_csv}")

    df_base = None

    # Check if file exists before attempting to read
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return None
    
    # Check read permissions
    if not os.access(file_path, os.R_OK):
        print(f"Error: Read permission denied for file: {file_path}")
        return None

    for attempt in range(retries):
        try:
            if is_csv:
                print(f"Attempt {attempt + 1}/{retries}: Reading CSV file...")
                # Add low_memory=False for potentially mixed types
                df_base = pd.read_csv(file_path, low_memory=False)
                print(f"Dados carregados do CSV: {df_base.shape[0]} linhas, {df_base.shape[1]} colunas.")
            else: # Original Excel logic (kept for potential future use, but not currently used)
                print(f"Attempt {attempt + 1}/{retries}: Reading Excel file directly using openpyxl engine...")
                xls = pd.ExcelFile(file_path, engine="openpyxl")
                print(f"Planilhas encontradas no arquivo Excel: {xls.sheet_names}")
                df_base = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=0)
                print(f"Dados originais carregados da planilha '{xls.sheet_names[0]}': {df_base.shape[0]} linhas, {df_base.shape[1]} colunas.")
                print("Preview das primeiras linhas:")
                print(df_base.head())
            
            break # Exit loop if successful

        except FileNotFoundError:
            print(f"Error: File not found at {file_path} on attempt {attempt + 1}")
            return None # Stop if file not found
        except pd.errors.EmptyDataError:
             print(f"Error: The file at {file_path} is empty on attempt {attempt + 1}.")
             return None
        except Exception as e:
            print(f"An unexpected error occurred during file loading on attempt {attempt + 1}: {e}")
            traceback.print_exc() # Print detailed traceback
            if attempt < retries - 1:
                print(f"Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                print("Max retries reached. Failed to load file.")
                return None

    if df_base is None:
        print("Failed to load df_base after all attempts.")
        return None

    # --- Data Cleaning and Preparation --- 
    # Note: Column names/indices might differ slightly between Excel and CSV if headers weren't perfect
    # We need to map based on expected column names now, assuming CSV headers are correct
    try:
        # Define expected column names based on previous analysis
        expected_columns = {
            "Order Creation Date: Date": "DataCriacao",
            "Tipo Cliente": "TipoClienteY",
            "Nome Completo": "NomeCompletoZ",
            "CANAL": "CanalAA",
            "3P": "TresP_AH",
            "Customer By SO: Buying Group Name": "GrupoFranqueadoW",
            "Sales Organization Code": "SalesOrgE",
            "STATUS": "StatusKPI",
            "Orders - TOTAL Orders Qty": "QuantidadeKPI",
            "Orders - TOTAL Gross Amount (Document Currency)": "ValorFaturadoKPI",
            "Orders Detail - Order Document Number": "NumPedido",
            "Reject Reason Code": "MotivoRejeicao",
            "Customer By SO: Buying Group Name": "Franqueado",
            "Brand & Segment - Code": "BrandCode",          # Coluna F
            "PLM Attributes - Collection Mix Desc": "CollectionDesc",  # Coluna G
            "Brand & Segment - Category": "BrandCategory",  # Coluna H
            "Otico/Sport": "OticoSport",                    # Coluna I
            "Canal": "CanalBI"                              # Coluna AA
            # "Mês": "MesNomeOriginal", # AE - Not directly used, derived from DataCriacao
            # "Semana": "NumSemanaOriginal", # AF - Not directly used, derived from DataCriacao
        }
        
        # Check if expected columns exist in the loaded CSV
        actual_columns = df_base.columns.tolist()
        print(f"Colunas encontradas no arquivo Excel: {actual_columns}")  # Debug print all columns

        # Debug: mostrar valores únicos das colunas "Grupo Franqueado" e "CANAL BI" antes do rename
        if "Grupo Franqueado" in actual_columns:
            print(f"Valores únicos na coluna 'Grupo Franqueado': {df_base['Grupo Franqueado'].dropna().unique()[:10]}")
        else:
            print("Coluna 'Grupo Franqueado' não encontrada no arquivo original.")

        if "CANAL" in actual_columns:
            print(f"Valores únicos na coluna 'CANAL': {df_base['CANAL'].dropna().unique()[:10]}")
            print(f"Preview da coluna 'CANAL': {df_base['CANAL'].head(10)}")
        else:
            print("Coluna 'CANAL' não encontrada no arquivo original.")

        rename_map = {}
        missing_expected = []
        for excel_name, desired_name in expected_columns.items():
            if excel_name in actual_columns:
                rename_map[excel_name] = desired_name
            else:
                # Try finding based on index if name mismatch (less reliable)
                # This part needs careful mapping based on the actual CSV header vs Excel indices
                # For now, assume names match from CSV conversion
                missing_expected.append(excel_name)
                print(f"Warning: Expected column 	\"{excel_name}\" not found in CSV header.")

        if not rename_map: # If no columns could be mapped
             print("Error: Could not map any expected columns to the CSV headers. Check CSV file structure.")
             print(f"CSV Headers found: {actual_columns}")
             return None

        # Select and rename columns
        df = df_base[list(rename_map.keys())].copy()
        df.rename(columns=rename_map, inplace=True)
        print(f"Colunas renomeadas do CSV: {list(df.columns)}")

        # --- Continue with cleaning steps as before --- 
        df["DataCriacao"] = pd.to_datetime(df["DataCriacao"], errors="coerce")
        df.dropna(subset=["DataCriacao"], inplace=True)
        print("Coluna \"DataCriacao\" convertida para datetime e NaTs removidos.")

        df["Ano"] = df["DataCriacao"].dt.year
        df["MesNumero"] = df["DataCriacao"].dt.month
        df["MesNome"] = df["DataCriacao"].dt.strftime("%B")
        df["SemanaAno"] = df["DataCriacao"].dt.isocalendar().week
        print("Colunas \"Ano\", \"MesNumero\", \"MesNome\", \"SemanaAno\" extraídas de \"DataCriacao\".")

        df["QuantidadeKPI"] = pd.to_numeric(df["QuantidadeKPI"], errors="coerce").fillna(0).astype(int)
        print("Coluna \"QuantidadeKPI\" verificada/convertida para inteiro.")
        df["ValorFaturadoKPI"] = pd.to_numeric(df["ValorFaturadoKPI"], errors="coerce").fillna(0)
        print("Coluna \"ValorFaturadoKPI\" verificada/convertida para numérico.")

        for col in ["TipoClienteY", "CanalAA", "TresP_AH", "SalesOrgE", "StatusKPI",
                "BrandCode", "CollectionDesc", "BrandCategory", "OticoSport", "CanalBI"]:
            if col in df.columns:
                df[col] = df[col].astype(str).fillna("").astype("category")
                print(f"Coluna \"{col}\" convertida para category.")

        if "GrupoFranqueadoW" in df.columns:
            null_count = df["GrupoFranqueadoW"].isnull().sum()
            if null_count > 0:
                print(f"Preenchendo {null_count} valores nulos em \"GrupoFranqueadoW\" com \"Não Especificado\".")
                df["GrupoFranqueadoW"].fillna("Não Especificado", inplace=True)
            df["GrupoFranqueadoW"] = df["GrupoFranqueadoW"].astype("category")
            print("Coluna \"GrupoFranqueadoW\" convertida para category.")

        if "Franqueado" in df.columns:
            null_count_franqueado = df["Franqueado"].isnull().sum()
            if null_count_franqueado > 0:
                print(f"Preenchendo {null_count_franqueado} valores nulos em \"Franqueado\" com \"Não Especificado\".")
                df["Franqueado"].fillna("Não Especificado", inplace=True)
            df["Franqueado"] = df["Franqueado"].astype("category")
            print("Coluna \"Franqueado\" convertida para category.")

        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        present_months = df["MesNome"].unique()
        ordered_present_months = [m for m in month_order if m in present_months]
        df["MesNome"] = pd.Categorical(df["MesNome"], categories=ordered_present_months, ordered=True)
        print("Coluna \"MesNome\" convertida para category.")
        print("Coluna \"MesNome\" ordenada cronologicamente.")

        final_columns = [
            "DataCriacao", "Ano", "MesNumero", "MesNome", "SemanaAno",
            "NumPedido", "StatusKPI", "QuantidadeKPI", "ValorFaturadoKPI",
            "CanalAA", "TipoClienteY", "TresP_AH", "SalesOrgE",
            "GrupoFranqueadoW", "Franqueado", "NomeCompletoZ", "MotivoRejeicao",
        # NOVAS COLUNAS
            "BrandCode", "CollectionDesc", "BrandCategory", "OticoSport", "CanalBI"
        ]
        final_columns = [col for col in final_columns if col in df.columns]
        df_final = df[final_columns]

        print("Tratamento de dados (CSV) para Streamlit concluído.")
        print(f"Dados processados: {df_final.shape[0]} linhas, {df_final.shape[1]} colunas.")
        print(f"Colunas finais selecionadas: {list(df_final.columns)}")
        return df_final

    except KeyError as e:
        print(f"Erro de Chave (Coluna não encontrada no CSV após mapeamento): {e}. Verifique os nomes das colunas no arquivo CSV.")
        traceback.print_exc()
        return None
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a limpeza e preparação dos dados (CSV): {e}")
        traceback.print_exc()
        return None



