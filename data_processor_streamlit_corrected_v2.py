# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import time # Import time for retry delay
import shutil # Import shutil for file copy
import traceback # For detailed error logging
import locale # Import locale for month names

import streamlit as st

@st.cache_data(show_spinner=False)
def load_and_clean_data_streamlit_cached(file_path, is_csv=False):
    # Ensure locale setting is consistent across cached calls if needed,
    # or handle locale within the non-cached function only.
    # For simplicity, we call the main function which handles locale.
    return load_and_clean_data_streamlit(file_path, is_csv)

def load_and_clean_data_streamlit(file_path, is_csv=False, retries=3, delay=3):
    """Loads data from the specified file (Excel or CSV), cleans it, and prepares it for the Streamlit dashboard."""
    import traceback
    import sys
    print(f"File path to load: {file_path}")
    print(f"Is CSV: {is_csv}")

    # Set locale to Portuguese for month names
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        print("Locale set to pt_BR.UTF-8 for month names.")
    except locale.Error:
        print("Warning: Locale pt_BR.UTF-8 not available. Using default locale for month names.")
        # Consider fallback or error handling if pt_BR is critical

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
                df_base = pd.read_csv(file_path, low_memory=False)
                print(f"Dados carregados do CSV: {df_base.shape[0]} linhas, {df_base.shape[1]} colunas.")
            else: # Original Excel logic
                print(f"Attempt {attempt + 1}/{retries}: Reading Excel file directly using openpyxl engine...")
                xls = pd.ExcelFile(file_path, engine="openpyxl")
                print(f"Planilhas encontradas no arquivo Excel: {xls.sheet_names}")
                # Assuming the first sheet is the target
                if not xls.sheet_names:
                    print(f"Error: No sheets found in the Excel file: {file_path}")
                    return None
                df_base = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=0)
                print(f"Dados originais carregados da planilha '{xls.sheet_names[0]}': {df_base.shape[0]} linhas, {df_base.shape[1]} colunas.")
                # print("Preview das primeiras linhas:")
                # print(df_base.head())

            break # Exit loop if successful

        except FileNotFoundError:
            print(f"Error: File not found at {file_path} on attempt {attempt + 1}")
            return None # Stop if file not found
        except pd.errors.EmptyDataError:
             print(f"Error: The file at {file_path} is empty on attempt {attempt + 1}.")
             return None
        except Exception as e:
            print(f"An unexpected error occurred during file loading on attempt {attempt + 1}: {e}")
            traceback.print_exc(file=sys.stdout) # Print detailed traceback to stdout
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
    try:
        # Define expected column names based on previous analysis (adapt if needed)
        expected_columns = {
            "Order Creation Date: Date": "DataCriacao",
            "Tipo Cliente": "TipoClienteY",
            "Nome Completo": "NomeCompletoZ",
            "CANAL": "CanalAA", # This seems duplicated later as "CanalBI", check source file
            "3P": "TresP_AH",
            "Customer By SO: Buying Group Name": "GrupoFranqueadoW", # Also duplicated as "Franqueado"?
            "Sales Organization Code": "SalesOrgE",
            "STATUS": "StatusKPI",
            "Orders - TOTAL Orders Qty": "QuantidadeKPI",
            "Orders - TOTAL Gross Amount (Document Currency)": "ValorFaturadoKPI",
            "Orders Detail - Order Document Number": "NumPedido",
            "Reject Reason Code": "MotivoRejeicao",
            # "Customer By SO: Buying Group Name": "Franqueado", # Duplicated key, use distinct source if needed
            "Brand & Segment - Code": "BrandCode",
            "PLM Attributes - Collection Mix Desc": "CollectionDesc",
            "Brand & Segment - Category": "BrandCategory",
            "Otico/Sport": "OticoSport",
            "Canal": "CanalBI" # Potential duplicate key, check source Excel header exactly
        }

        # Resolve duplicate keys in expected_columns if they point to different source columns
        # Example: If "CANAL" and "Canal" are two distinct columns in Excel mapped to different desired names
        # For now, assume the last definition ("Canal": "CanalBI") overrides the first if keys are identical strings
        # And "Customer By SO: Buying Group Name" maps to "GrupoFranqueadoW"
        # Add mapping for "Franqueado" if it comes from a different source column
        # Assuming "Franqueado" should also map from "Customer By SO: Buying Group Name" for now
        expected_columns["Customer By SO: Buying Group Name_Franqueado"] = "Franqueado" # Create a unique key if needed

        actual_columns = df_base.columns.tolist()
        print(f"Colunas encontradas no arquivo: {actual_columns}")

        rename_map = {}
        missing_expected = []
        used_actual_columns = set()

        # Special handling for potentially duplicated source columns mapped to different names
        if "Customer By SO: Buying Group Name" in actual_columns:
            rename_map["Customer By SO: Buying Group Name"] = "GrupoFranqueadoW"
            used_actual_columns.add("Customer By SO: Buying Group Name")
            # If "Franqueado" is meant to be the *same* column, we don't need a separate mapping.
            # If it's a *different* column with a similar name, adjust `expected_columns` key.
            # For now, let's assume Franqueado is derived or handled later if it's the same source.
            if "Customer By SO: Buying Group Name_Franqueado" in expected_columns:
                 del expected_columns["Customer By SO: Buying Group Name_Franqueado"] # Remove the temp key

        if "CANAL" in actual_columns:
             rename_map["CANAL"] = "CanalAA" # First mapping
             used_actual_columns.add("CANAL")
        if "Canal" in actual_columns and "Canal" not in used_actual_columns:
             rename_map["Canal"] = "CanalBI" # Second mapping (case-sensitive)
             used_actual_columns.add("Canal")
        elif "Canal" not in actual_columns and "CANAL" in actual_columns: # If only "CANAL" exists
             # Decide which mapping takes precedence or if one column serves both purposes
             # Assuming CanalBI is the primary one if only "CANAL" exists:
             if "CANAL" in rename_map: del rename_map["CANAL"] # Remove first mapping
             rename_map["CANAL"] = "CanalBI"
             used_actual_columns.add("CANAL")

        # Map remaining columns
        for excel_name, desired_name in expected_columns.items():
            if excel_name in actual_columns and excel_name not in used_actual_columns:
                rename_map[excel_name] = desired_name
                used_actual_columns.add(excel_name)
            elif excel_name not in actual_columns:
                 # Check if it's one of the specially handled ones already mapped
                 if desired_name not in ["GrupoFranqueadoW", "CanalAA", "CanalBI", "Franqueado"]:
                     missing_expected.append(excel_name)
                     print(f"Warning: Expected column \"{excel_name}\" not found in source file header.")

        if not rename_map:
             print("Error: Could not map any expected columns to the source file headers. Check file structure.")
             print(f"Source Headers found: {actual_columns}")
             return None

        # Select and rename columns
        df = df_base[list(rename_map.keys())].copy()
        df.rename(columns=rename_map, inplace=True)
        print(f"Colunas renomeadas: {list(df.columns)}")

        # --- Continue with cleaning steps --- 
        df["DataCriacao"] = pd.to_datetime(df["DataCriacao"], errors="coerce")
        df.dropna(subset=["DataCriacao"], inplace=True)
        print("Coluna \"DataCriacao\" convertida para datetime e NaTs removidos.")

        df["Ano"] = df["DataCriacao"].dt.year
        df["MesNumero"] = df["DataCriacao"].dt.month
        # Generate month name using the locale set earlier (pt_BR)
        df["MesNome"] = df["DataCriacao"].dt.strftime("%B").str.capitalize()
        # Ensure SemanaAno is numeric, handle potential errors
        df["SemanaAno"] = pd.to_numeric(df["DataCriacao"].dt.strftime("%U"), errors='coerce').fillna(0).astype(int) # Using %U for week number starting Sunday
        # df["SemanaAno"] = df["DataCriacao"].dt.isocalendar().week # ISO week number if preferred

        print("Colunas \"Ano\", \"MesNumero\", \"MesNome\" (PT-BR), \"SemanaAno\" extraídas de \"DataCriacao\".")

        df["QuantidadeKPI"] = pd.to_numeric(df["QuantidadeKPI"], errors="coerce").fillna(0).astype(int)
        print("Coluna \"QuantidadeKPI\" verificada/convertida para inteiro.")
        df["ValorFaturadoKPI"] = pd.to_numeric(df["ValorFaturadoKPI"], errors="coerce").fillna(0)
        print("Coluna \"ValorFaturadoKPI\" verificada/convertida para numérico.")

        # Convert categorical columns
        categorical_cols = [
            "TipoClienteY", "CanalAA", "TresP_AH", "SalesOrgE", "StatusKPI",
            "BrandCode", "CollectionDesc", "BrandCategory", "OticoSport", "CanalBI",
            "GrupoFranqueadoW", "MotivoRejeicao" # Add Franqueado if it exists as a separate column
        ]
        if "Franqueado" not in df.columns and "GrupoFranqueadoW" in df.columns:
             df["Franqueado"] = df["GrupoFranqueadoW"] # If Franqueado should be same as GrupoFranqueadoW
             print("Coluna 'Franqueado' criada como cópia de 'GrupoFranqueadoW'.")
             categorical_cols.append("Franqueado")
        elif "Franqueado" in df.columns:
             categorical_cols.append("Franqueado")

        for col in categorical_cols:
            if col in df.columns:
                # Fill NaNs before converting to category
                fill_value = "Não Especificado"
                if df[col].isnull().any():
                     print(f"Preenchendo valores nulos em \"{col}\" com \"{fill_value}\".")
                     df[col].fillna(fill_value, inplace=True)
                df[col] = df[col].astype(str).astype("category")
                print(f"Coluna \"{col}\" convertida para category.")
            else:
                print(f"Warning: Coluna categórica esperada '{col}' não encontrada após renomeação.")

        # Order MesNome chronologically using Portuguese names
        month_order_pt = [
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
        ]
        present_months = df["MesNome"].unique().tolist()
        # Always use the full list of months for categories to ensure all appear in filters
        df["MesNome"] = pd.Categorical(df["MesNome"], categories=month_order_pt, ordered=True)
        print("Coluna \"MesNome\" convertida para category ordenada (PT-BR) com todos os 12 meses.")

        # Define final columns, ensuring they exist after processing
        final_columns_base = [
            "DataCriacao", "Ano", "MesNumero", "MesNome", "SemanaAno",
            "NumPedido", "StatusKPI", "QuantidadeKPI", "ValorFaturadoKPI",
            "CanalAA", "TipoClienteY", "TresP_AH", "SalesOrgE",
            "GrupoFranqueadoW", "Franqueado", "NomeCompletoZ", "MotivoRejeicao",
            "BrandCode", "CollectionDesc", "BrandCategory", "OticoSport", "CanalBI"
        ]
        final_columns = [col for col in final_columns_base if col in df.columns]
        df_final = df[final_columns]

        print("Tratamento de dados para Streamlit concluído.")
        print(f"Dados processados: {df_final.shape[0]} linhas, {df_final.shape[1]} colunas.")
        print(f"Colunas finais selecionadas: {list(df_final.columns)}")
        return df_final

    except KeyError as e:
        print(f"Erro de Chave (Coluna não encontrada durante processamento): {e}. Verifique os nomes das colunas no arquivo fonte e mapeamento.")
        traceback.print_exc()
        return None
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a limpeza e preparação dos dados: {e}")
        traceback.print_exc()
        return None

