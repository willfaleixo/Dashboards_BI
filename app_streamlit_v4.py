import locale

# Definir local para português para nome do mês
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252') # Windows
    except locale.Error:
        pass  # Não usar st.warning aqui para evitar erro de ordem

import sys
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
import os
import time
import io

# Importar a função de carregamento e limpeza de dados da versão corrigida
from data_processor_streamlit_corrected_v2 import load_and_clean_data_streamlit, load_and_clean_data_streamlit_cached

# --- Configuração da Página ---
st.set_page_config(
    page_title="Dashboard Doações EssilorLuxottica",
    page_icon="assets/logo.png",
    layout="wide"
)

# --- Carregamento e Preparação dos Dados ---
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running in a normal Python environment
    base_path = os.path.abspath(".")

DATA_FILE_PATH = os.path.join(base_path, "upload", "teste_d11.xlsx")
IS_CSV = False

@st.cache_data(ttl=600)
def load_data():
    print(f"Tentando carregar dados de: {os.path.abspath(DATA_FILE_PATH)}")
    if not os.path.exists(DATA_FILE_PATH):
        st.error(f"Arquivo de dados não encontrado em: {os.path.abspath(DATA_FILE_PATH)}")
        return None
    start_time = time.time()
    try:
        df_loaded = load_and_clean_data_streamlit_cached(DATA_FILE_PATH, IS_CSV)
        if df_loaded is None:
            st.error(f"Falha ao carregar ou processar dados de {DATA_FILE_PATH}.")
            return None
        if 'DataCriacao' not in df_loaded.columns or not pd.api.types.is_datetime64_any_dtype(df_loaded['DataCriacao']):
             st.error("Coluna 'DataCriacao' não encontrada ou não está no formato datetime.")
             return None
        # Garantir que MesNome seja categórico ordenado após o carregamento
        if 'MesNome' in df_loaded.columns and not isinstance(df_loaded['MesNome'].dtype, pd.CategoricalDtype):
            month_order_pt = [
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            ]
            present_months = df_loaded['MesNome'].unique().tolist()
            ordered_present_months = [m for m in month_order_pt if m in present_months]
            if ordered_present_months:
                df_loaded['MesNome'] = pd.Categorical(df_loaded['MesNome'], categories=month_order_pt, ordered=True)
                print("Coluna \"MesNome\" convertida para category ordenada (PT-BR) na app.")
            else:
                print("Warning: Nenhum mês encontrado para ordenação categórica de MesNome na app.")

        end_time = time.time()
        print(f"Dados carregados e verificados em {end_time - start_time:.2f} segundos.")
        return df_loaded
    except Exception as e:
        st.error(f"Erro crítico durante o carregamento dos dados: {e}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        return None

# Add a button to reload data and clear cache
if 'reload_data' not in st.session_state:
    st.session_state.reload_data = False

def reload_data_callback():
    st.cache_data.clear()
    st.session_state.reload_data = True

if st.sidebar.button("Atualizar Dados"):
    reload_data_callback()

if st.session_state.reload_data:
    df = load_data()
    st.session_state.reload_data = False
else:
    df = load_data()

# Obter timestamp da última modificação do arquivo de dados
last_update_timestamp = None
try:
    if df is not None and not df.empty and os.path.exists(DATA_FILE_PATH):
        last_update_timestamp = datetime.fromtimestamp(os.path.getmtime(DATA_FILE_PATH))
        last_update_str = last_update_timestamp.strftime("%d/%m/%Y %H:%M:%S")
    else:
        last_update_str = "N/A (Dados não carregados)"
except Exception as e:
    st.warning(f"Não foi possível obter o timestamp do arquivo de dados: {e}. Usando hora atual.")
    last_update_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# --- Layout Principal ---

# Cabeçalho
col1, col2, col3 = st.columns([1, 5, 1])
with col1:
    # Determine the base path for bundled files (PyInstaller or normal environment)
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        base_path_logo = sys._MEIPASS
    else:
        # Running in a normal Python environment
        base_path_logo = os.path.abspath(".")

    logo_path = os.path.join(base_path_logo, "assets", "logo.png")
    if os.path.exists(logo_path):
        try:
            st.image(logo_path, width=100)
        except Exception as e:
            st.warning(f"Erro ao carregar a imagem do logo: {e}")
with col2:
    st.title("Dashboard Doações EssilorLuxottica")
with col3:
    st.caption(f"Última Atualização: {last_update_str}")

st.markdown("---")

# --- Barra Lateral (Sidebar) para Filtros ---
st.sidebar.header("Filtros")

if df is None or df.empty:
    st.error("Erro ao carregar os dados ou dados vazios. Dashboard não pode ser exibido.")
    st.stop()

# Função para obter opções de filtro, garantindo a ordem para MesNome
@st.cache_data(show_spinner=False)
def get_options(df_input, column_name):
    if column_name not in df_input.columns:
        return []
    # Para MesNome, usar as categorias ordenadas definidas
    if column_name == 'MesNome' and isinstance(df_input[column_name].dtype, pd.CategoricalDtype):
        return df_input[column_name].cat.categories.tolist()
    # Para outras colunas categóricas ordenadas
    if isinstance(df_input[column_name].dtype, pd.CategoricalDtype) and df_input[column_name].cat.ordered:
        return df_input[column_name].cat.categories.tolist()
    # Para colunas numéricas ou outras
    options = df_input[column_name].dropna().unique()
    try:
        # Tenta ordenar como números (Ano, SemanaAno)
        options_sorted = sorted([int(opt) for opt in options])
        return [str(opt) for opt in options_sorted]
    except (ValueError, TypeError):
        # Ordena como string
        return sorted([str(opt) for opt in options])

query_params = st.query_params
def get_query_param_list(param_name):
    return query_params.get(param_name, [])

placeholder_text = "Escolha uma opção"

selected_anos = st.sidebar.multiselect("Ano", options=get_options(df, "Ano"), default=get_query_param_list("ano"), placeholder=placeholder_text)
# Usar a função get_options que retorna a lista ordenada completa para MesNome
selected_meses = st.sidebar.multiselect("Mês", options=[str(i) for i in range(1,13)], default=get_query_param_list("mes"), placeholder=placeholder_text)
selected_semanas = st.sidebar.multiselect("Semana", options=get_options(df, "SemanaAno"), default=get_query_param_list("semana"), placeholder=placeholder_text)
selected_canais = st.sidebar.multiselect("Canal", options=get_options(df, "CanalBI"), default=get_query_param_list("canal"), placeholder=placeholder_text)
selected_3p = st.sidebar.multiselect("3P/LUX", options=get_options(df, "TresP_AH"), default=get_query_param_list("3p"), placeholder=placeholder_text)
selected_sales_org = st.sidebar.multiselect("Organização de vendas", options=get_options(df, "SalesOrgE"), default=get_query_param_list("sales_org"), placeholder=placeholder_text)
franqueado_options = [opt for opt in get_options(df, "Franqueado") if opt and opt != "Não Especificado"]
selected_franqueado = st.sidebar.multiselect("Franqueado", options=franqueado_options, default=get_query_param_list("franqueado"), placeholder=placeholder_text)
selected_brand_code = st.sidebar.multiselect("Marca (Brand)", options=get_options(df, "BrandCode"), default=get_query_param_list("brand_code"), placeholder=placeholder_text)
selected_collection_desc = st.sidebar.multiselect("Tipo do Produto", options=get_options(df, "CollectionDesc"), default=get_query_param_list("collection_desc"), placeholder=placeholder_text)
selected_brand_category = st.sidebar.multiselect("Categorização da marca", options=get_options(df, "BrandCategory"), default=get_query_param_list("brand_category"), placeholder=placeholder_text)
selected_otico_sport = st.sidebar.multiselect("Otico / Sport", options=get_options(df, "OticoSport"), default=get_query_param_list("otico_sport"), placeholder=placeholder_text)

# --- Filtrar DataFrame com base nas seleções ---
@st.cache_data(show_spinner=False)
def apply_filters(df_input, anos, meses, semanas, canais, p3, sales_org, franqueado, brand, collection, category, otico):
    dff = df_input.copy()
    # Usar get_options para verificar se *todas* as opções estão selecionadas (não aplicar filtro nesse caso)
    if anos and len(anos) < len(get_options(df_input, "Ano")):
        dff = dff[dff["Ano"].astype(str).isin(anos)]
    if meses and len(meses) < len(get_options(df_input, "MesNumero")):
        dff = dff[dff["MesNumero"].isin(meses)]
    if semanas and len(semanas) < len(get_options(df_input, "SemanaAno")):
        dff = dff[dff["SemanaAno"].astype(str).isin(semanas)]
    if canais and len(canais) < len(get_options(df_input, "CanalBI")):
        dff = dff[dff["CanalBI"].astype(str).isin(canais)]
    if p3 and len(p3) < len(get_options(df_input, "TresP_AH")):
        dff = dff[dff["TresP_AH"].astype(str).isin(p3)]
    if sales_org and len(sales_org) < len(get_options(df_input, "SalesOrgE")):
        dff = dff[dff["SalesOrgE"].astype(str).isin(sales_org)]
    franqueado_all_options = [opt for opt in get_options(df_input, "Franqueado") if opt and opt != "Não Especificado"]
    if franqueado and len(franqueado) < len(franqueado_all_options):
        dff = dff[dff["Franqueado"].astype(str).isin(franqueado)]
    if brand and len(brand) < len(get_options(df_input, "BrandCode")):
        dff = dff[dff["BrandCode"].astype(str).isin(brand)]
    if collection and len(collection) < len(get_options(df_input, "CollectionDesc")):
        dff = dff[dff["CollectionDesc"].astype(str).isin(collection)]
    if category and len(category) < len(get_options(df_input, "BrandCategory")):
        dff = dff[dff["BrandCategory"].astype(str).isin(category)]
    if otico and len(otico) < len(get_options(df_input, "OticoSport")):
        dff = dff[dff["OticoSport"].astype(str).isin(otico)]
    return dff

# Convert selected_meses from string to int for filtering
selected_meses_int = [int(m) for m in selected_meses] if selected_meses else []

dff = apply_filters(df, selected_anos, selected_meses_int, selected_semanas, selected_canais, selected_3p, selected_sales_org, selected_franqueado, selected_brand_code, selected_collection_desc, selected_brand_category, selected_otico_sport)

# --- Função para converter DataFrame para CSV --- 
@st.cache_data
def convert_df_to_csv(df_to_convert):
    return df_to_convert.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

# --- Função para converter DataFrame para Excel --- 
@st.cache_data
def convert_df_to_excel(df_to_convert):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_to_convert.to_excel(writer, index=False, sheet_name='DadosFiltrados')
    processed_data = output.getvalue()
    return processed_data

# --- Botões de Download na Sidebar ---
st.sidebar.markdown("---")
st.sidebar.header("Downloads")

# Botão para baixar dados filtrados (Excel)
if not dff.empty:
    excel_filtered = convert_df_to_excel(dff)
    st.sidebar.download_button(
        label="Baixar Dados Filtrados (Excel)",
        data=excel_filtered,
        file_name=f'dados_filtrados_{datetime.now().strftime("%Y%m%d")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
else:
    st.sidebar.info("Nenhum dado filtrado para baixar.")

# Botão para baixar base original completa (Excel)
if os.path.exists(DATA_FILE_PATH):
    with open(DATA_FILE_PATH, "rb") as fp:
        st.sidebar.download_button(
            label="Baixar Base Original Completa",
            data=fp,
            file_name=os.path.basename(DATA_FILE_PATH), # Usa o nome original do arquivo
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.sidebar.warning("Arquivo original não encontrado para download.")


# --- Exibição Principal do Dashboard ---
if dff.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    # Adiciona a assinatura mesmo se não houver dados filtrados
    st.write("---")
    current_year_sig = datetime.now().year
    st.markdown(f"**_BI After Sales EssilorLuxottica | {current_year_sig}_**")
    st.markdown("_<small>Created by Willian Aleixo</small>_", unsafe_allow_html=True)
else:
    # --- KPIs Comparativos (Reestruturados) ---
    st.subheader("Indicadores Comparativos")

    def create_custom_kpi_card(title, value1, value2, label1, label2, delta_percentage):
        delta_str = "(N/D)"
        arrow = ""
        color = "gray"
        if delta_percentage is not None and pd.notna(delta_percentage) and delta_percentage != float('inf') and delta_percentage != float('-inf'):
            delta_formatted = f"{delta_percentage:.1f}".replace(".", ",") + "%"
            if delta_percentage > 0.1:
                arrow = "▲"
                color = "green"
            elif delta_percentage < -0.1:
                arrow = "▼"
                color = "red"
            else:
                 arrow = "▶"
                 color = "gray"
            delta_str = f"<span style='color:{color}; font-size: small;'>{delta_formatted} {arrow}</span>"
        elif value1 > 0 and value2 == 0:
             delta_str = "<span style='color:green; font-size: small;'>(Novo) ▲</span>"
        elif value1 == 0 and value2 > 0:
             delta_str = "<span style='color:red; font-size: small;'>(Zero) ▼</span>"
        elif value1 == value2:
             delta_str = "<span style='color:gray; font-size: small;'>0,0% ▶</span>"

        value1_formatted = f"{value1:,.0f}".replace(",", ".")
        value2_formatted = f"{value2:,.0f}".replace(",", ".")

        markdown_string = f'''
<div style="text-align: center; border: 1px solid #eee; padding: 10px; border-radius: 5px; height: 100%; display: flex; flex-direction: column; justify-content: space-between;">
<div style="font-size: small; color: gray;">{title}</div>
<div style="font-size: large; margin-top: 5px; margin-bottom: 5px;">
<span style="display: inline-block; min-width: 40%; text-align: right;">{value1_formatted}</span>
<span style="margin: 0 5px;">x</span>
<span style="display: inline-block; min-width: 40%; text-align: left;">{value2_formatted}</span>
</div>
<div style="font-size: x-small; color: gray; margin-bottom: 5px;">
<span style="display: inline-block; min-width: 40%; text-align: right;">{label1}</span>
<span style="margin: 0 5px;"> </span>
<span style="display: inline-block; min-width: 40%; text-align: left;">{label2}</span>
</div>
<div>{delta_str}</div>
</div>
'''
        return markdown_string

    # --- Cálculos para KPIs Comparativos --- 
    latest_data_date = df["DataCriacao"].max().date()
    today = latest_data_date

    current_year = today.year
    current_month = today.month
    current_day = today.day
    current_day_of_year = today.timetuple().tm_yday

    prev_year = current_year - 1
    prev_month_year = current_year if current_month > 1 else current_year - 1
    prev_month = current_month - 1 if current_month > 1 else 12
    prev_week_date = today - timedelta(days=7)

    # Períodos Atuais (YTD, MTD, WTD)
    start_current_year = date(current_year, 1, 1)
    start_current_month = date(current_year, current_month, 1)
    start_current_week = today - timedelta(days=today.weekday()) # Segunda-feira da semana atual

    df_ytd = df[(df['DataCriacao'].dt.date >= start_current_year) & (df['DataCriacao'].dt.date <= today)]
    df_mtd = df[(df['DataCriacao'].dt.date >= start_current_month) & (df['DataCriacao'].dt.date <= today)]
    df_wtd = df[(df['DataCriacao'].dt.date >= start_current_week) & (df['DataCriacao'].dt.date <= today)]

    # Períodos Anteriores Correspondentes (YTD, MTD, WTD)
    start_prev_year_ytd = date(prev_year, 1, 1)
    try:
        end_prev_year_ytd = date(prev_year, current_month, current_day)
    except ValueError: # Dia não existe no ano anterior (ex: 29 Fev)
        end_prev_year_ytd = date(prev_year, current_month, current_day -1)

    start_prev_month_mtd = date(prev_month_year, prev_month, 1)
    last_day_prev_month = (start_current_month - timedelta(days=1)).day
    end_prev_month_mtd = date(prev_month_year, prev_month, min(current_day, last_day_prev_month))

    start_prev_week_wtd = start_current_week - timedelta(days=7)
    end_prev_week_wtd = prev_week_date

    df_prev_ytd = df[(df['DataCriacao'].dt.date >= start_prev_year_ytd) & (df['DataCriacao'].dt.date <= end_prev_year_ytd)]
    df_prev_mtd = df[(df['DataCriacao'].dt.date >= start_prev_month_mtd) & (df['DataCriacao'].dt.date <= end_prev_month_mtd)]
    df_prev_wtd = df[(df['DataCriacao'].dt.date >= start_prev_week_wtd) & (df['DataCriacao'].dt.date <= end_prev_week_wtd)]

    # Calcular KPIs CRIADOS
    qtd_criada_ytd = df_ytd['QuantidadeKPI'].sum()
    qtd_criada_prev_ytd = df_prev_ytd['QuantidadeKPI'].sum()
    delta_criada_yoy = ((qtd_criada_ytd - qtd_criada_prev_ytd) / qtd_criada_prev_ytd * 100) if qtd_criada_prev_ytd != 0 else (float('inf') if qtd_criada_ytd > 0 else 0)

    qtd_criada_mtd = df_mtd['QuantidadeKPI'].sum()
    qtd_criada_prev_mtd = df_prev_mtd['QuantidadeKPI'].sum()
    delta_criada_mom = ((qtd_criada_mtd - qtd_criada_prev_mtd) / qtd_criada_prev_mtd * 100) if qtd_criada_prev_mtd != 0 else (float('inf') if qtd_criada_mtd > 0 else 0)

    qtd_criada_wtd = df_wtd['QuantidadeKPI'].sum()
    qtd_criada_prev_wtd = df_prev_wtd['QuantidadeKPI'].sum()
    delta_criada_wow = ((qtd_criada_wtd - qtd_criada_prev_wtd) / qtd_criada_prev_wtd * 100) if qtd_criada_prev_wtd != 0 else (float('inf') if qtd_criada_wtd > 0 else 0)

    # Calcular KPIs FATURADOS
    df_faturado_ytd = df_ytd[df_ytd['StatusKPI'] == "Faturado"]
    df_faturado_prev_ytd = df_prev_ytd[df_prev_ytd['StatusKPI'] == "Faturado"]
    qtd_faturada_ytd = df_faturado_ytd['QuantidadeKPI'].sum()
    qtd_faturada_prev_ytd = df_faturado_prev_ytd['QuantidadeKPI'].sum()
    delta_faturada_yoy = ((qtd_faturada_ytd - qtd_faturada_prev_ytd) / qtd_faturada_prev_ytd * 100) if qtd_faturada_prev_ytd != 0 else (float('inf') if qtd_faturada_ytd > 0 else 0)

    df_faturado_mtd = df_mtd[df_mtd['StatusKPI'] == "Faturado"]
    df_faturado_prev_mtd = df_prev_mtd[df_prev_mtd['StatusKPI'] == "Faturado"]
    qtd_faturada_mtd = df_faturado_mtd['QuantidadeKPI'].sum()
    qtd_faturada_prev_mtd = df_faturado_prev_mtd['QuantidadeKPI'].sum()
    delta_faturada_mom = ((qtd_faturada_mtd - qtd_faturada_prev_mtd) / qtd_faturada_prev_mtd * 100) if qtd_faturada_prev_mtd != 0 else (float('inf') if qtd_faturada_mtd > 0 else 0)

    df_faturado_wtd = df_wtd[df_wtd['StatusKPI'] == "Faturado"]
    df_faturado_prev_wtd = df_prev_wtd[df_prev_wtd['StatusKPI'] == "Faturado"]
    qtd_faturada_wtd = df_faturado_wtd['QuantidadeKPI'].sum()
    qtd_faturada_prev_wtd = df_faturado_prev_wtd['QuantidadeKPI'].sum()
    delta_faturada_wow = ((qtd_faturada_wtd - qtd_faturada_prev_wtd) / qtd_faturada_prev_wtd * 100) if qtd_faturada_prev_wtd != 0 else (float('inf') if qtd_faturada_wtd > 0 else 0)

    # --- Exibir KPIs Comparativos (2 Fileiras) --- 
    st.markdown("##### Volume Criado")
    comp_kpi_c1, comp_kpi_c2, comp_kpi_c3 = st.columns(3)
    with comp_kpi_c1:
        st.markdown(create_custom_kpi_card("YoY Criado", qtd_criada_ytd, qtd_criada_prev_ytd, f"YTD {current_year}", f"YTD {prev_year}", delta_criada_yoy), unsafe_allow_html=True)
    with comp_kpi_c2:
        st.markdown(create_custom_kpi_card("MoM Criado", qtd_criada_mtd, qtd_criada_prev_mtd, f"MTD {today.strftime('%b/%Y')}", f"MTD {end_prev_month_mtd.strftime('%b/%Y')}", delta_criada_mom), unsafe_allow_html=True)
    with comp_kpi_c3:
        st.markdown(create_custom_kpi_card("WoW Criado", qtd_criada_wtd, qtd_criada_prev_wtd, f"WTD {today.strftime('%d/%b')}", f"WTD {end_prev_week_wtd.strftime('%d/%b')}", delta_criada_wow), unsafe_allow_html=True)

    st.markdown("##### Volume Faturado")
    comp_kpi_f1, comp_kpi_f2, comp_kpi_f3 = st.columns(3)
    with comp_kpi_f1:
        st.markdown(create_custom_kpi_card("YoY Faturado", qtd_faturada_ytd, qtd_faturada_prev_ytd, f"YTD {current_year}", f"YTD {prev_year}", delta_faturada_yoy), unsafe_allow_html=True)
    with comp_kpi_f2:
        st.markdown(create_custom_kpi_card("MoM Faturado", qtd_faturada_mtd, qtd_faturada_prev_mtd, f"MTD {today.strftime('%b/%Y')}", f"MTD {end_prev_month_mtd.strftime('%b/%Y')}", delta_faturada_mom), unsafe_allow_html=True)
    with comp_kpi_f3:
        st.markdown(create_custom_kpi_card("WoW Faturado", qtd_faturada_wtd, qtd_faturada_prev_wtd, f"WTD {today.strftime('%d/%b')}", f"WTD {end_prev_week_wtd.strftime('%d/%b')}", delta_faturada_wow), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True) # Espaçamento

    # --- KPIs Filtrados (Status) ---
    st.subheader("Indicadores Chave (Filtro Aplicado)")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    qtd_criada_filtrada = dff["QuantidadeKPI"].sum()
    qtd_cancelada_filtrada = dff[dff["StatusKPI"] == "Cancelado"]["QuantidadeKPI"].sum()
    qtd_faturada_filtrada = dff[dff["StatusKPI"] == "Faturado"]["QuantidadeKPI"].sum()
    # Aberta = Criada - Cancelada - Faturada (considerando apenas esses 3 status principais)
    qtd_aberta_filtrada = qtd_criada_filtrada - qtd_cancelada_filtrada - qtd_faturada_filtrada
    # Ou, se houver outros status: qtd_aberta = dff[~dff["StatusKPI"].isin(["Cancelado", "Faturado"])]["QuantidadeKPI"].sum()

    kpi1.metric(label="Qtd. Criada (Filtro)", value=f"{qtd_criada_filtrada:,}".replace(",", "."))
    kpi2.metric(label="Qtd. Cancelada (Filtro)", value=f"{qtd_cancelada_filtrada:,}".replace(",", "."))
    kpi3.metric(label="Qtd. Faturada (Filtro)", value=f"{qtd_faturada_filtrada:,}".replace(",", "."))
    kpi4.metric(label="Qtd. em Aberto (Filtro)", value=f"{qtd_aberta_filtrada:,}".replace(",", "."))

    st.markdown("---")

    # --- Gráficos (Mantidos como na v3) ---
    st.subheader("Análise Temporal")
    col_tempo1, col_tempo2 = st.columns(2)

    with col_tempo1:
        criado_tempo_mes = dff.resample("MS", on="DataCriacao")["QuantidadeKPI"].sum().reset_index()
        fig_criado_tempo = px.line(criado_tempo_mes, x="DataCriacao", y="QuantidadeKPI", title="Volume Criado por Mês", markers=True, labels={"DataCriacao": "Mês", "QuantidadeKPI": "Quantidade"}, color_discrete_sequence=["black"])
        fig_criado_tempo.update_layout(hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_criado_tempo, use_container_width=True)

        criado_ano = dff.groupby("Ano")["QuantidadeKPI"].sum().reset_index()
        fig_criado_ano = px.bar(criado_ano, x="Ano", y="QuantidadeKPI", title="Volume Criado por Ano", labels={"Ano": "Ano", "QuantidadeKPI": "Quantidade"}, text_auto=True, color_discrete_sequence=["black"])
        fig_criado_ano.update_layout(margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_criado_ano, use_container_width=True)

    with col_tempo2:
        faturado_tempo_mes = dff[dff["StatusKPI"] == "Faturado"].resample("MS", on="DataCriacao")["QuantidadeKPI"].sum().reset_index()
        fig_faturado_tempo = px.line(faturado_tempo_mes, x="DataCriacao", y="QuantidadeKPI", title="Volume Faturado por Mês", markers=True, labels={"DataCriacao": "Mês", "QuantidadeKPI": "Quantidade Faturada"}, color_discrete_sequence=["black"])
        fig_faturado_tempo.update_layout(hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_faturado_tempo, use_container_width=True)

        faturado_ano = dff[dff["StatusKPI"] == "Faturado"].groupby("Ano")["QuantidadeKPI"].sum().reset_index()
        fig_faturado_ano = px.bar(faturado_ano, x="Ano", y="QuantidadeKPI", title="Volume Faturado por Ano", labels={"Ano": "Ano", "QuantidadeKPI": "Quantidade Faturada"}, text_auto=True, color_discrete_sequence=["black"])
        fig_faturado_ano.update_layout(margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_faturado_ano, use_container_width=True)

    st.markdown("---")
    st.subheader("Análise por Grupos")
    col_grupo1, col_grupo2 = st.columns(2)

    with col_grupo1:
        if "Franqueado" in dff.columns:
            franqueado_faturado = dff[dff["StatusKPI"] == "Faturado"]
            franqueado_faturado = franqueado_faturado[franqueado_faturado["Franqueado"] != "Não Especificado"]
            franqueado_faturado_vol = franqueado_faturado.groupby("Franqueado")["QuantidadeKPI"].sum().reset_index()
            top_franqueados_faturado = franqueado_faturado_vol.nlargest(15, "QuantidadeKPI")
            fig_franqueado_faturado = px.bar(top_franqueados_faturado, y="Franqueado", x="QuantidadeKPI", title="Top 15 Franqueados Faturados (Quantidade)", orientation="h", labels={"Franqueado": "Franqueado", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text="QuantidadeKPI")
            fig_franqueado_faturado.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
            fig_franqueado_faturado.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_franqueado_faturado, use_container_width=True)
        else:
            st.info("Coluna 'Franqueado' não encontrada.")

    with col_grupo2:
        if "NomeCompletoZ" in dff.columns:
            nome_faturado = dff[dff["StatusKPI"] == "Faturado"]
            nome_faturado = nome_faturado[nome_faturado["NomeCompletoZ"] != "-"]
            nome_faturado_vol = nome_faturado.groupby("NomeCompletoZ")["QuantidadeKPI"].sum().nlargest(10).reset_index()
            fig_nome_faturado = px.bar(nome_faturado_vol, y="NomeCompletoZ", x="QuantidadeKPI", title="Top 10 Colaboradores Faturados (Quantidade)", orientation="h", labels={"NomeCompletoZ": "Colaborador", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text="QuantidadeKPI")
            fig_nome_faturado.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
            fig_nome_faturado.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_nome_faturado, use_container_width=True)
        else:
             st.info("Coluna 'NomeCompletoZ' não encontrada.")

    st.markdown("---")
    st.subheader("Análise Adicional")
    col_add1, col_add2 = st.columns(2)

    with col_add1:
        if "BrandCategory" in dff.columns:
            pie_data = dff.groupby("BrandCategory").size().reset_index(name='count')
            fig_pie = px.pie(pie_data, names="BrandCategory", values="count", title="Distribuição por Categoria de Marca", color_discrete_sequence=px.colors.sequential.Darkmint)
            fig_pie.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Coluna 'BrandCategory' não encontrada.")

    with col_add2:
        if "BrandCode" in dff.columns:
            top10_brandcode = dff.groupby("BrandCode")["QuantidadeKPI"].sum().nlargest(10).reset_index()
            fig_top10 = px.bar(top10_brandcode, x="BrandCode", y="QuantidadeKPI", title="Top 10 por Marca (Quantidade)", labels={"BrandCode": "Marca", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text_auto=True)
            fig_top10.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_top10, use_container_width=True)
        else:
            st.info("Coluna 'BrandCode' não encontrada.")

    # --- Tabela de Dados Filtrados --- 
    st.markdown("---")
    st.subheader("Dados Filtrados")
    # Usar st.dataframe para melhor interatividade (rolagem, ordenação)
    # Selecionar colunas relevantes ou mostrar todas, se necessário
    # Exemplo: Mostrar um subconjunto de colunas
    colunas_tabela = [
        'DataCriacao', 'Ano', 'MesNome', 'SemanaAno', 'NumPedido', 'StatusKPI', 
        'QuantidadeKPI', 'CanalBI', 'Franqueado', 'BrandCode', 'CollectionDesc'
    ]
    colunas_existentes_tabela = [col for col in colunas_tabela if col in dff.columns]
    st.dataframe(dff[colunas_existentes_tabela], use_container_width=True)
    # Ou mostrar todas as colunas do dataframe filtrado:
    # st.dataframe(dff, use_container_width=True)

# --- Assinatura Final --- (Fora do else para sempre aparecer, exceto se erro inicial)
st.write("---")
current_year_sig = datetime.now().year
st.markdown(f"**_BI After Sales EssilorLuxottica | {current_year_sig}_**")
st.markdown("_<small>Created by Willian Aleixo</small>_", unsafe_allow_html=True)
