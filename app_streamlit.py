import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import time
import locale

# Importar a função de carregamento e limpeza de dados
from data_processor_streamlit_corrected import load_and_clean_data_streamlit

# --- Configuração da Página (MOVIDO PARA CÁ) ---
st.set_page_config(
    page_title="Dashboard Doações EssilorLuxottica",
    page_icon="assets/logo.png",  # Use a logo como ícone
    layout="wide"
)

# Definir local para português para nome do mês
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    st.warning("Locale pt_BR não disponível. Usando locale padrão para nomes de mês.")


# --- Carregamento e Preparação dos Dados ---
import os
DATA_FILE_PATH = os.path.abspath("upload/teste_d11.xlsx")
IS_CSV = False

from data_processor_streamlit_corrected import load_and_clean_data_streamlit_cached

@st.cache_data
def load_data():
    start_time = time.time()
    # Removido o print do caminho absoluto do arquivo de dados
    try:
        df_loaded = load_and_clean_data_streamlit_cached(DATA_FILE_PATH, IS_CSV)
        if df_loaded is None:
            st.error(f"Falha ao carregar dados de {DATA_FILE_PATH}. Verifique o arquivo e as permissões.")
            return None
        # Certifica que a coluna DataCriacao existe e é datetime
        if 'DataCriacao' not in df_loaded.columns or not pd.api.types.is_datetime64_any_dtype(df_loaded['DataCriacao']):
             st.error("Coluna 'DataCriacao' não encontrada ou não está no formato datetime após o carregamento.")
             # Tenta converter novamente se possível
             if 'DataCriacao' in df_loaded.columns:
                 try:
                     df_loaded['DataCriacao'] = pd.to_datetime(df_loaded['DataCriacao'], errors='coerce')
                     df_loaded.dropna(subset=['DataCriacao'], inplace=True)
                     if df_loaded['DataCriacao'].isnull().all():
                         st.error("Conversão de 'DataCriacao' para datetime falhou ou resultou em todos os valores nulos.")
                         return None
                 except Exception as e:
                     st.error(f"Erro ao tentar converter 'DataCriacao': {e}")
                     return None
             else:
                 return None
        return df_loaded
    except Exception as e:
        st.error(f"Erro crítico durante o carregamento dos dados: {e}")
        return None

df = load_data()

# Obter timestamp da última modificação do arquivo de dados
last_update_timestamp = None
try:
    last_update_timestamp = datetime.now() # Usar hora atual como fallback
    if os.path.exists(DATA_FILE_PATH):
        last_update_timestamp = datetime.fromtimestamp(os.path.getmtime(DATA_FILE_PATH))
    last_update_str = last_update_timestamp.strftime("%d/%m/%Y %H:%M:%S")
except Exception as e:
    st.warning(f"Não foi possível obter o timestamp do arquivo de dados: {e}. Usando hora atual.")
    last_update_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# --- Layout Principal ---

# Cabeçalho
col1, col2, col3 = st.columns([1, 5, 1])
with col1:
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)
    else:
        st.warning("Logo não encontrado.")
with col2:
    st.title("Dashboard Doações EssilorLuxottica OnePage")
with col3:
    st.caption(f"Última Atualização: {last_update_str}")

st.markdown("---")

# --- Barra Lateral (Sidebar) para Filtros ---
st.sidebar.header("Filtros")

if df is None:
    st.error("Erro ao carregar os dados. Não é possível configurar os filtros ou KPIs.")
    st.stop() # Interrompe a execução se os dados não carregaram

# Obter opções para filtros (valores únicos e ordenados)
@st.cache_data(show_spinner=False)
def get_options(column_name):
    if column_name in df.columns:
        if isinstance(df[column_name].dtype, pd.CategoricalDtype) and df[column_name].cat.ordered:
            return df[column_name].cat.categories.tolist()
        options = df[column_name].dropna().unique()
        options = [str(opt) for opt in options]
        return sorted(options)
    return []

# Sincronizar filtros com parâmetros da URL
query_params = st.query_params

def get_query_param_list(param_name):
    if param_name in query_params:
        return query_params[param_name]
    return []

selected_anos = st.sidebar.multiselect("Ano", options=get_options("Ano"), default=get_query_param_list("ano"))
selected_meses = st.sidebar.multiselect("Mês", options=get_options("MesNome"), default=get_query_param_list("mes"))

@st.cache_data(show_spinner=False)
def get_sorted_week_options():
    options = get_options("SemanaAno")
    try:
        options_int = sorted([int(opt) for opt in options])
        return [str(opt) for opt in options_int]
    except Exception:
        return options

selected_semanas = st.sidebar.multiselect("Semana", options=get_sorted_week_options(), default=get_query_param_list("semana"))
selected_canais = st.sidebar.multiselect("Canal", options=get_options("CanalBI"), default=get_query_param_list("canal"))
selected_3p = st.sidebar.multiselect("3P/LUX", options=get_options("TresP_AH"), default=get_query_param_list("3p"))
selected_sales_org = st.sidebar.multiselect("Organização de vendas", options=get_options("SalesOrgE"), default=get_query_param_list("sales_org"))
selected_franqueado = st.sidebar.multiselect("Franqueado", options=[opt for opt in get_options("Franqueado") if opt], default=get_query_param_list("franqueado"))
selected_brand_code = st.sidebar.multiselect("Marca (Brand)", options=get_options("BrandCode"), default=get_query_param_list("brand_code"))
selected_collection_desc = st.sidebar.multiselect("Tipo do Produto", options=get_options("CollectionDesc"), default=get_query_param_list("collection_desc"))
selected_brand_category = st.sidebar.multiselect("Categorização da marca", options=get_options("BrandCategory"), default=get_query_param_list("brand_category"))
selected_otico_sport = st.sidebar.multiselect("Otico / Sport", options=get_options("OticoSport"), default=get_query_param_list("otico_sport"))

def update_query_params():
    params = {
        "ano": selected_anos,
        "mes": selected_meses,
        "semana": selected_semanas,
        "canal": selected_canais,
        "3p": selected_3p,
        "sales_org": selected_sales_org,
        "franqueado": selected_franqueado
    }
    params = {k: v for k, v in params.items() if v}
    st.query_params.update(params)

update_query_params()

# --- Filtrar DataFrame com base nas seleções ---
def apply_filters(df, selected_anos, selected_meses, selected_semanas, selected_canais, selected_3p, selected_sales_org, selected_franqueado, selected_brand_code, selected_collection_desc, selected_brand_category, selected_otico_sport):
    dff = df.copy()
    if selected_anos and len(selected_anos) < len(get_options("Ano")):
        dff = dff[dff["Ano"].astype(str).isin(selected_anos)]
    if selected_meses and len(selected_meses) < len(get_options("MesNome")):
        dff = dff[dff["MesNome"].astype(str).isin(selected_meses)]
    if selected_semanas and len(selected_semanas) < len(get_options("SemanaAno")):
        dff = dff[dff["SemanaAno"].astype(str).isin(selected_semanas)]
    if selected_canais and len(selected_canais) < len(get_options("CanalBI")):
        dff = dff[dff["CanalBI"].astype(str).isin(selected_canais)]
    if selected_3p and len(selected_3p) < len(get_options("TresP_AH")):
        dff = dff[dff["TresP_AH"].astype(str).isin(selected_3p)]
    if selected_sales_org and len(selected_sales_org) < len(get_options("SalesOrgE")):
        dff = dff[dff["SalesOrgE"].astype(str).isin(selected_sales_org)]
    if selected_franqueado and len(selected_franqueado) < len(get_options("Franqueado")):
        dff = dff[dff["Franqueado"].astype(str).isin(selected_franqueado)]
    if selected_brand_code and len(selected_brand_code) < len(get_options("BrandCode")):
        dff = dff[dff["BrandCode"].astype(str).isin(selected_brand_code)]
    if selected_collection_desc and len(selected_collection_desc) < len(get_options("CollectionDesc")):
        dff = dff[dff["CollectionDesc"].astype(str).isin(selected_collection_desc)]
    if selected_brand_category and len(selected_brand_category) < len(get_options("BrandCategory")):
        dff = dff[dff["BrandCategory"].astype(str).isin(selected_brand_category)]
    if selected_otico_sport and len(selected_otico_sport) < len(get_options("OticoSport")):
        dff = dff[dff["OticoSport"].astype(str).isin(selected_otico_sport)]
    return dff

dff = apply_filters(df, selected_anos, selected_meses, selected_semanas, selected_canais, selected_3p, selected_sales_org, selected_franqueado, selected_brand_code, selected_collection_desc, selected_brand_category, selected_otico_sport)

if dff.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.sidebar.markdown("---")
    st.sidebar.header("Download")
    st.sidebar.info("Base de dados original indisponível para download (sem dados nos filtros).")
else:
    # --- KPIs Principais ---
    st.subheader("Indicadores Chave (Período Filtrado)")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    qtd_criada = dff["QuantidadeKPI"].sum()
    qtd_cancelada = dff[dff["StatusKPI"] == "Cancelado"]["QuantidadeKPI"].sum()
    qtd_faturada = dff[dff["StatusKPI"] == "Faturado"]["QuantidadeKPI"].sum()
    qtd_aberta = dff[(dff["StatusKPI"] != "Cancelado") & (dff["StatusKPI"] != "Faturado")]["QuantidadeKPI"].sum()

    kpi1.metric(label="Qtd. Criada", value=f"{qtd_criada:,}".replace(",", "."))
    kpi2.metric(label="Qtd. Cancelada", value=f"{qtd_cancelada:,}".replace(",", "."))
    kpi3.metric(label="Qtd. Faturada", value=f"{qtd_faturada:,}".replace(",", "."))
    kpi4.metric(label="Qtd. em Aberta", value=f"{qtd_aberta:,}".replace(",", "."))

    # --- KPIs Comparativos (usando df original) ---
    st.markdown("---")
    st.subheader("Indicadores Comparativos")

    # Função para criar o cartão KPI customizado
    def create_custom_kpi_card(title, value1, value2, label1, label2, delta_percentage):
        delta_str = "(N/D)"
        arrow = ""
        color = "gray"
        if delta_percentage is not None:
            delta_formatted = f"{delta_percentage:.1f}".replace(".", ",") + "%"
            if delta_percentage > 0:
                arrow = "▲"
                color = "green"
            elif delta_percentage < 0:
                arrow = "▼"
                color = "red"
            delta_str = f"<span style='color:{color}; font-size: small;'>{delta_formatted} {arrow}</span>"
        elif value1 > 0 and value2 == 0:
             delta_str = "<span style='color:green; font-size: small;'>(Novo) ▲</span>"
        elif value1 == 0 and value2 > 0:
             delta_str = "<span style='color:red; font-size: small;'>(Zero) ▼</span>"


        # Formata os valores com separador de milhar
        value1_formatted = f"{value1:,}".replace(",", ".")
        value2_formatted = f"{value2:,}".replace(",", ".")

        # Usando HTML/Markdown para layout
        markdown_string = f"""
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
        """
        return markdown_string

    # Obter data mais recente no DataFrame original
    latest_date = df["DataCriacao"].max()
    current_year = latest_date.year
    current_month = latest_date.month
    current_week = latest_date.isocalendar().week
    current_month_name = latest_date.strftime("%B").capitalize()

    # Calcular períodos anteriores
    prev_year = current_year - 1
    prev_month_date = latest_date.replace(day=1) - timedelta(days=1)
    prev_month = prev_month_date.month
    prev_month_year = prev_month_date.year
    prev_month_name = prev_month_date.strftime("%B").capitalize()
    prev_week_date = latest_date - timedelta(weeks=1)
    prev_week = prev_week_date.isocalendar().week
    prev_week_year = prev_week_date.isocalendar().year

    # Calcular deltas
    def calculate_delta(current_val, previous_val):
        if previous_val > 0:
            return ((current_val - previous_val) / previous_val) * 100
        elif current_val > 0:
            return float('inf') # Representa crescimento infinito (ou 'Novo')
        else:
            return None # Nenhuma mudança ou N/D

    # --- Comparativo Ano (YoY) ---
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

    # YoY Criada
    qtd_criada_cy = df[df["Ano"] == current_year]["QuantidadeKPI"].sum()
    qtd_criada_py = df[df["Ano"] == prev_year]["QuantidadeKPI"].sum()
    delta_yoy_criada = calculate_delta(qtd_criada_cy, qtd_criada_py)
    md_yoy_criada = create_custom_kpi_card("Qtd. Criada (Ano)", qtd_criada_cy, qtd_criada_py, str(current_year), str(prev_year), delta_yoy_criada)
    kpi_col1.markdown(md_yoy_criada, unsafe_allow_html=True)

    # YoY Faturada
    qtd_faturada_cy = df[(df["Ano"] == current_year) & (df["StatusKPI"] == "Faturado")]["QuantidadeKPI"].sum()
    qtd_faturada_py = df[(df["Ano"] == prev_year) & (df["StatusKPI"] == "Faturado")]["QuantidadeKPI"].sum()
    delta_yoy_faturada = calculate_delta(qtd_faturada_cy, qtd_faturada_py)
    md_yoy_faturada = create_custom_kpi_card("Qtd. Faturada (Ano)", qtd_faturada_cy, qtd_faturada_py, str(current_year), str(prev_year), delta_yoy_faturada)
    kpi_col2.markdown(md_yoy_faturada, unsafe_allow_html=True)

    # MoM Criada
    qtd_criada_cm = df[(df["Ano"] == current_year) & (df["MesNumero"] == current_month)]["QuantidadeKPI"].sum()
    qtd_criada_pm = df[(df["Ano"] == prev_month_year) & (df["MesNumero"] == prev_month)]["QuantidadeKPI"].sum()
    delta_mom_criada = calculate_delta(qtd_criada_cm, qtd_criada_pm)
    md_mom_criada = create_custom_kpi_card("Qtd. Criada (Mês)", qtd_criada_cm, qtd_criada_pm, current_month_name, prev_month_name, delta_mom_criada)
    kpi_col3.markdown(md_mom_criada, unsafe_allow_html=True)

    # WoW Criada
    qtd_criada_cw = df[(df["Ano"] == current_year) & (df["SemanaAno"] == current_week)]["QuantidadeKPI"].sum()
    qtd_criada_pw = df[(df["Ano"] == prev_week_year) & (df["SemanaAno"] == prev_week)]["QuantidadeKPI"].sum()
    delta_wow_criada = calculate_delta(qtd_criada_cw, qtd_criada_pw)
    md_wow_criada = create_custom_kpi_card("Qtd. Criada (Semana)", qtd_criada_cw, qtd_criada_pw, f"Sem {current_week}", f"Sem {prev_week}", delta_wow_criada)
    kpi_col4.markdown(md_wow_criada, unsafe_allow_html=True)

    # --- Análise Temporal ---
    st.markdown("---")
    st.subheader("Análise Temporal (Período Filtrado)")
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

    # --- Análise por Grupos ---
    st.markdown("---")
    st.subheader("Análise por Grupos (Período Filtrado)")
    col_grupo1, col_grupo2 = st.columns(2)

    with col_grupo1:
        if "Franqueado" in dff.columns:
            franqueado_faturado = dff[dff["StatusKPI"] == "Faturado"]
            franqueado_faturado = franqueado_faturado[franqueado_faturado["Franqueado"] != "Não Especificado"]
            franqueado_faturado_vol = franqueado_faturado.groupby("Franqueado", observed=True)["QuantidadeKPI"].sum().reset_index()
            top_franqueados_faturado = franqueado_faturado_vol.nlargest(15, "QuantidadeKPI")
            fig_franqueado_faturado = px.bar(top_franqueados_faturado, y="Franqueado", x="QuantidadeKPI", title="Top 15 Franqueados Faturados (Quantidade)", orientation="h", labels={"Franqueado": "Franqueado", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text="QuantidadeKPI")
            fig_franqueado_faturado.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
            fig_franqueado_faturado.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_franqueado_faturado, use_container_width=True)
        else:
            st.info("Coluna 'Franqueado' não encontrada nos dados filtrados.")

    with col_grupo2:
        if "NomeCompletoZ" in dff.columns:
            nome_faturado = dff[dff["StatusKPI"] == "Faturado"]
            nome_faturado = nome_faturado[nome_faturado["NomeCompletoZ"] != "-"]
            nome_faturado_vol = nome_faturado.groupby("NomeCompletoZ", observed=True)["QuantidadeKPI"].sum().nlargest(10).reset_index()
            fig_nome_faturado = px.bar(nome_faturado_vol, y="NomeCompletoZ", x="QuantidadeKPI", title="Top 10 Colaboradores Faturados (Quantidade)", orientation="h", labels={"NomeCompletoZ": "Colaborador", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text="QuantidadeKPI")
            fig_nome_faturado.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
            fig_nome_faturado.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_nome_faturado, use_container_width=True)
        else:
            st.info("Coluna 'NomeCompletoZ' não encontrada nos dados filtrados.")

    # --- Análise Adicional (usando dff) ---
    st.markdown("---")
    st.subheader("Análise Adicional (Período Filtrado)")
    col_add1, col_add2 = st.columns(2)

    with col_add1:
        if "BrandCategory" in dff.columns and "OticoSport" in dff.columns:
            pie_data = dff.groupby("BrandCategory", observed=True).size().reset_index(name='counts')
            fig_pie = px.pie(pie_data, names="BrandCategory", values="counts", title="Distribuição por Categoria de Marca", color_discrete_sequence=px.colors.sequential.Darkmint)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Colunas 'BrandCategory' e/ou 'OticoSport' não encontradas para gráfico de pizza.")

    with col_add2:
        if "BrandCode" in dff.columns:
            top10_brandcode = dff.groupby("BrandCode", observed=True)["QuantidadeKPI"].sum().nlargest(10).reset_index()
            fig_top10 = px.bar(top10_brandcode, x="BrandCode", y="QuantidadeKPI", title="Top 10 Marcas (Quantidade)", labels={"BrandCode": "Marca", "QuantidadeKPI": "Quantidade Total"}, color_discrete_sequence=["black"], text_auto=True)
            st.plotly_chart(fig_top10, use_container_width=True)
        else:
            st.info("Coluna 'BrandCode' não encontrada para gráfico Top 10.")

# --- Rodapé e Download ---
st.markdown("---")
current_year_footer = datetime.now().year
st.markdown(f"**_BI After Sales EssilorLuxottica | {current_year_footer}_**")

st.markdown("<p style='font-size: small; color: gray; margin-top: -10px;'>designed by Willian Aleixo</p>", unsafe_allow_html=True)

st.markdown("---")
try:
    with open(DATA_FILE_PATH, "rb") as fp:
        btn = st.download_button(
            label="Download Base de Dados (Excel)",
            data=fp,
            file_name="teste_d11.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except FileNotFoundError:
    st.error(f"Arquivo {DATA_FILE_PATH} não encontrado para download.")
except Exception as e:
    st.error(f"Erro ao preparar arquivo para download: {e}")

