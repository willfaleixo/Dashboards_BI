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
    page_icon="📊",  # Usando emoji como fallback
    layout="wide"
)

# --- Carregamento e Preparação dos Dados ---
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running in a normal Python environment
    base_path = os.path.abspath(".")

# Lista de possíveis localizações do arquivo de dados
possible_data_paths = [
    os.path.join(base_path, "upload", "teste_d11.xlsx"),
    os.path.join(base_path, "teste_d11.xlsx"),
    "upload/teste_d11.xlsx",
    "teste_d11.xlsx",
    os.path.join(os.getcwd(), "upload", "teste_d11.xlsx"),
    os.path.join(os.getcwd(), "teste_d11.xlsx")
]

DATA_FILE_PATH = None
for path in possible_data_paths:
    if os.path.exists(path):
        DATA_FILE_PATH = path
        break

IS_CSV = False

@st.cache_data(ttl=600)
def load_data():
    if DATA_FILE_PATH is None:
        st.error("Arquivo de dados não encontrado. Verifique se o arquivo 'teste_d11.xlsx' está presente no repositório.")
        st.info("Locais verificados:")
        for path in possible_data_paths:
            st.write(f"- {path}")
        return None
    
    print(f"Tentando carregar dados de: {os.path.abspath(DATA_FILE_PATH)}")
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
    if df is not None and not df.empty and DATA_FILE_PATH and os.path.exists(DATA_FILE_PATH):
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
    # Lista de possíveis localizações do logo
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        base_path_logo = sys._MEIPASS
    else:
        # Running in a normal Python environment
        base_path_logo = os.path.abspath(".")

    possible_logo_paths = [
        os.path.join(base_path_logo, "assets", "logo.png"),
        os.path.join(base_path_logo, "logo.png"),
        "assets/logo.png",
        "logo.png",
        os.path.join(os.getcwd(), "assets", "logo.png"),
        os.path.join(os.getcwd(), "logo.png")
    ]
    
    logo_displayed = False
    for logo_path in possible_logo_paths:
        if os.path.exists(logo_path):
            try:
                # Verificar se o arquivo é uma imagem válida
                from PIL import Image
                img = Image.open(logo_path)
                st.image(logo_path, width=100)
                logo_displayed = True
                break
            except Exception as e:
                print(f"Erro ao carregar logo de {logo_path}: {e}")
                continue
    
    if not logo_displayed:
        # Fallback: usar emoji ou texto
        st.markdown("### 🏢")
        
with col2:
    st.title("Dashboard Doações EssilorLuxottica")
with col3:
    st.caption(f"Última Atualização: {last_update_str}")

st.markdown("---")

# --- Barra Lateral (Sidebar) para Filtros ---
st.sidebar.header("Filtros")

if df is None or df.empty:
    st.error("Erro ao carregar os dados ou dados vazios. Dashboard não pode ser exibido.")
    st.info("**Instruções para correção:**")
    st.write("1. Verifique se o arquivo 'teste_d11.xlsx' está presente no repositório GitHub")
    st.write("2. O arquivo deve estar na pasta raiz ou em uma pasta chamada 'upload'")
    st.write("3. Certifique-se de que o arquivo não está corrompido")
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
if DATA_FILE_PATH and os.path.exists(DATA_FILE_PATH):
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

    # Períodos Anteriores (Mesmo período do ano anterior)
    start_prev_year = date(prev_year, 1, 1)
    end_prev_year_ytd = date(prev_year, current_month, current_day) if current_month <= 12 else date(prev_year, 12, 31)
    
    start_prev_month = date(prev_month_year, prev_month, 1)
    try:
        end_prev_month_mtd = date(prev_month_year, prev_month, current_day)
    except ValueError:
        # Caso o dia atual não exista no mês anterior (ex: 31 de março vs 28/29 de fevereiro)
        import calendar
        last_day_prev_month = calendar.monthrange(prev_month_year, prev_month)[1]
        end_prev_month_mtd = date(prev_month_year, prev_month, min(current_day, last_day_prev_month))

    start_prev_week = prev_week_date - timedelta(days=prev_week_date.weekday())
    end_prev_week = prev_week_date

    df_ytd_prev = df[(df['DataCriacao'].dt.date >= start_prev_year) & (df['DataCriacao'].dt.date <= end_prev_year_ytd)]
    df_mtd_prev = df[(df['DataCriacao'].dt.date >= start_prev_month) & (df['DataCriacao'].dt.date <= end_prev_month_mtd)]
    df_wtd_prev = df[(df['DataCriacao'].dt.date >= start_prev_week) & (df['DataCriacao'].dt.date <= end_prev_week)]

    # Cálculo dos KPIs
    def calculate_kpis(df_period):
        if df_period.empty:
            return 0, 0
        total_qty = df_period['QuantidadeKPI'].sum()
        total_value = df_period['ValorFaturadoKPI'].sum()
        return total_qty, total_value

    qty_ytd, value_ytd = calculate_kpis(df_ytd)
    qty_ytd_prev, value_ytd_prev = calculate_kpis(df_ytd_prev)

    qty_mtd, value_mtd = calculate_kpis(df_mtd)
    qty_mtd_prev, value_mtd_prev = calculate_kpis(df_mtd_prev)

    qty_wtd, value_wtd = calculate_kpis(df_wtd)
    qty_wtd_prev, value_wtd_prev = calculate_kpis(df_wtd_prev)

    # Cálculo das variações percentuais
    def calculate_percentage_change(current, previous):
        if previous == 0:
            return None if current == 0 else float('inf')
        return ((current - previous) / previous) * 100

    qty_ytd_change = calculate_percentage_change(qty_ytd, qty_ytd_prev)
    value_ytd_change = calculate_percentage_change(value_ytd, value_ytd_prev)

    qty_mtd_change = calculate_percentage_change(qty_mtd, qty_mtd_prev)
    value_mtd_change = calculate_percentage_change(value_mtd, value_mtd_prev)

    qty_wtd_change = calculate_percentage_change(qty_wtd, qty_wtd_prev)
    value_wtd_change = calculate_percentage_change(value_wtd, value_wtd_prev)

    # Exibição dos KPIs
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(create_custom_kpi_card(
            "Quantidade YTD",
            qty_ytd, qty_ytd_prev,
            f"{current_year}", f"{prev_year}",
            qty_ytd_change
        ), unsafe_allow_html=True)

    with col2:
        st.markdown(create_custom_kpi_card(
            "Quantidade MTD",
            qty_mtd, qty_mtd_prev,
            f"{current_month:02d}/{current_year}", f"{prev_month:02d}/{prev_month_year}",
            qty_mtd_change
        ), unsafe_allow_html=True)

    with col3:
        st.markdown(create_custom_kpi_card(
            "Quantidade WTD",
            qty_wtd, qty_wtd_prev,
            "Semana Atual", "Semana Anterior",
            qty_wtd_change
        ), unsafe_allow_html=True)

    # Segunda linha de KPIs (Valor)
    col4, col5, col6 = st.columns(3)

    with col4:
        st.markdown(create_custom_kpi_card(
            "Valor YTD (R$)",
            value_ytd, value_ytd_prev,
            f"{current_year}", f"{prev_year}",
            value_ytd_change
        ), unsafe_allow_html=True)

    with col5:
        st.markdown(create_custom_kpi_card(
            "Valor MTD (R$)",
            value_mtd, value_mtd_prev,
            f"{current_month:02d}/{current_year}", f"{prev_month:02d}/{prev_month_year}",
            value_mtd_change
        ), unsafe_allow_html=True)

    with col6:
        st.markdown(create_custom_kpi_card(
            "Valor WTD (R$)",
            value_wtd, value_wtd_prev,
            "Semana Atual", "Semana Anterior",
            value_wtd_change
        ), unsafe_allow_html=True)

    st.markdown("---")

    # --- Gráficos ---
    st.subheader("Análises Visuais")

    # Gráfico 1: Evolução Mensal
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Evolução Mensal - Quantidade**")
        if 'MesNome' in dff.columns:
            monthly_qty = dff.groupby('MesNome')['QuantidadeKPI'].sum().reset_index()
            fig_monthly_qty = px.bar(monthly_qty, x='MesNome', y='QuantidadeKPI',
                                   title="Quantidade por Mês")
            fig_monthly_qty.update_layout(xaxis_title="Mês", yaxis_title="Quantidade")
            st.plotly_chart(fig_monthly_qty, use_container_width=True)
        else:
            st.warning("Coluna 'MesNome' não encontrada nos dados filtrados.")

    with col2:
        st.markdown("**Evolução Mensal - Valor**")
        if 'MesNome' in dff.columns:
            monthly_value = dff.groupby('MesNome')['ValorFaturadoKPI'].sum().reset_index()
            fig_monthly_value = px.bar(monthly_value, x='MesNome', y='ValorFaturadoKPI',
                                     title="Valor por Mês")
            fig_monthly_value.update_layout(xaxis_title="Mês", yaxis_title="Valor (R$)")
            st.plotly_chart(fig_monthly_value, use_container_width=True)
        else:
            st.warning("Coluna 'MesNome' não encontrada nos dados filtrados.")

    # Gráfico 2: Top 10 por Canal
    col3, col4 = st.columns(2)

    with col3:
        st.markdown("**Top 10 Canais - Quantidade**")
        if 'CanalBI' in dff.columns:
            top_canais_qty = dff.groupby('CanalBI')['QuantidadeKPI'].sum().sort_values(ascending=False).head(10).reset_index()
            fig_canais_qty = px.bar(top_canais_qty, x='QuantidadeKPI', y='CanalBI',
                                  orientation='h', title="Top 10 Canais por Quantidade")
            fig_canais_qty.update_layout(xaxis_title="Quantidade", yaxis_title="Canal")
            st.plotly_chart(fig_canais_qty, use_container_width=True)
        else:
            st.warning("Coluna 'CanalBI' não encontrada nos dados filtrados.")

    with col4:
        st.markdown("**Top 10 Canais - Valor**")
        if 'CanalBI' in dff.columns:
            top_canais_value = dff.groupby('CanalBI')['ValorFaturadoKPI'].sum().sort_values(ascending=False).head(10).reset_index()
            fig_canais_value = px.bar(top_canais_value, x='ValorFaturadoKPI', y='CanalBI',
                                    orientation='h', title="Top 10 Canais por Valor")
            fig_canais_value.update_layout(xaxis_title="Valor (R$)", yaxis_title="Canal")
            st.plotly_chart(fig_canais_value, use_container_width=True)
        else:
            st.warning("Coluna 'CanalBI' não encontrada nos dados filtrados.")

    # Gráfico 3: Distribuição por Status
    col5, col6 = st.columns(2)

    with col5:
        st.markdown("**Distribuição por Status - Quantidade**")
        if 'StatusKPI' in dff.columns:
            status_qty = dff.groupby('StatusKPI')['QuantidadeKPI'].sum().reset_index()
            fig_status_qty = px.pie(status_qty, values='QuantidadeKPI', names='StatusKPI',
                                  title="Distribuição por Status (Quantidade)")
            st.plotly_chart(fig_status_qty, use_container_width=True)
        else:
            st.warning("Coluna 'StatusKPI' não encontrada nos dados filtrados.")

    with col6:
        st.markdown("**Distribuição por Status - Valor**")
        if 'StatusKPI' in dff.columns:
            status_value = dff.groupby('StatusKPI')['ValorFaturadoKPI'].sum().reset_index()
            fig_status_value = px.pie(status_value, values='ValorFaturadoKPI', names='StatusKPI',
                                    title="Distribuição por Status (Valor)")
            st.plotly_chart(fig_status_value, use_container_width=True)
        else:
            st.warning("Coluna 'StatusKPI' não encontrada nos dados filtrados.")

    # --- Tabela de Dados Filtrados ---
    st.subheader("Dados Filtrados")
    
    # Mostrar apenas as primeiras 1000 linhas para performance
    display_df = dff.head(1000)
    st.dataframe(display_df, use_container_width=True)
    
    if len(dff) > 1000:
        st.info(f"Mostrando as primeiras 1000 linhas de {len(dff)} registros totais. Use os filtros para refinar a visualização ou baixe o arquivo completo.")

    # Adiciona a assinatura
    st.write("---")
    current_year_sig = datetime.now().year
    st.markdown(f"**_BI After Sales EssilorLuxottica | {current_year_sig}_**")
    st.markdown("_<small>Created by Willian Aleixo</small>_", unsafe_allow_html=True)

