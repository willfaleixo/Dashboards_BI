import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import time

# Importar a função de carregamento e limpeza de dados
from data_processor_streamlit_corrected import load_and_clean_data_streamlit

# --- Configuração da Página ---
st.set_page_config(
    page_title="Dashboard Doações EssilorLuxottica",
    page_icon="assets/logo.png",  # Use a logo como ícone, fixed relative path
    layout="wide"
)


# --- Carregamento e Preparação dos Dados ---
DATA_FILE_PATH = r"c:/Users/alexiow/Desktop/Automações/dashboard_streamlit_local/upload/teste_d11.xlsx"
IS_CSV = False  # Indicador para usar CSV ao invés de Excel

from data_processor_streamlit_corrected import load_and_clean_data_streamlit_cached

@st.cache_data
def load_data():
    start_time = time.time()
    df_loaded = load_and_clean_data_streamlit_cached(DATA_FILE_PATH, IS_CSV)
    end_time = time.time()
    return df_loaded
df = load_data()

# Salvar os dados filtrados em um arquivo Excel para análise
if df is not None:
    output_path = "dados_dashboard_exportados.xlsx"
    df.to_excel(output_path, index=False)
    # st.info(f"Dados do dashboard exportados para o arquivo: {output_path}")

# Debug: mostrar as primeiras linhas das colunas CanalAA e GrupoFranqueadoW para verificar dados
# Debug: mostrar as primeiras linhas das colunas CanalAA e GrupoFranqueadoW para verificar dados
if df is not None:
    # st.write("Colunas disponíveis no DataFrame carregado:")
    # st.write(df.columns.tolist())

# Debug: mostrar tipo e valores únicos da coluna CanalAA
    if "CanalAA" in df.columns:
        # st.write(f"Tipo da coluna CanalAA: {df['CanalAA'].dtype}")
        # st.write(f"Valores únicos da coluna CanalAA: {df['CanalAA'].unique()[:20]}")
        pass  # Instrução válida para evitar erro de indentação

# Obter timestamp da última modificação do arquivo de dados
last_update_timestamp = None
try:
    last_update_timestamp = datetime.now()  # Usar hora atual como fallback
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
    st.title("Dashboard Doações EssilorLuxottica")
with col3:
    st.caption(f"Última Atualização: {last_update_str}")

st.markdown("---")

# --- Barra Lateral (Sidebar) para Filtros ---
st.sidebar.header("Filtros")

if df is None:
    st.error("Erro ao carregar os dados. Não é possível configurar os filtros.")
    st.stop()  # Interrompe a execução se os dados não carregaram

# Obter opções para filtros (valores únicos e ordenados)
@st.cache_data(show_spinner=False)
def get_options(column_name):
    if column_name in df.columns:
        if isinstance(df[column_name].dtype, pd.CategoricalDtype) and df[column_name].cat.ordered:
            return df[column_name].cat.categories.tolist()
        # Ensure consistent type for options by converting categories to strings
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

# Organizar opções da semana do menor para o maior
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
selected_brand_code = st.sidebar.multiselect("BrandCode", options=get_options("BrandCode"), default=get_query_param_list("brand_code"))
selected_collection_desc = st.sidebar.multiselect("CollectionDesc", options=get_options("CollectionDesc"), default=get_query_param_list("collection_desc"))
selected_brand_category = st.sidebar.multiselect("BrandCategory", options=get_options("BrandCategory"), default=get_query_param_list("brand_category"))
selected_otico_sport = st.sidebar.multiselect("OticoSport", options=get_options("OticoSport"), default=get_query_param_list("otico_sport"))

# Atualizar parâmetros da URL com as seleções atuais
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
    # Remover chaves com listas vazias para limpar a URL
    params = {k: v for k, v in params.items() if v}
    st.query_params.update(params)

update_query_params()

# --- Filtrar DataFrame com base nas seleções ---
# Aplicar filtros apenas se houver seleção e a seleção não for igual a todas as opções disponíveis
def apply_filters(df, selected_anos, selected_meses, selected_semanas, selected_canais, selected_3p, selected_sales_org, selected_franqueado, selected_brand_code, selected_collection_desc, selected_brand_category, selected_otico_sport):
    dff = df
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
else:
    # --- KPIs ---
    st.subheader("Indicadores Chave")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    # Calcular KPIs usando a coluna QuantidadeKPI (Coluna S) e dados filtrados (dff)
    qtd_criada = dff["QuantidadeKPI"].sum()
    qtd_cancelada = dff[dff["StatusKPI"] == "Cancelado"]["QuantidadeKPI"].sum()
    qtd_faturada = dff[dff["StatusKPI"] == "Faturado"]["QuantidadeKPI"].sum()
    valor_faturado_total = dff[dff["StatusKPI"] == "Faturado"]["ValorFaturadoKPI"].sum()

    kpi1.metric(label="Quantidade Criada", value=f"{qtd_criada:,}".replace(",", "."))
    kpi2.metric(label="Quantidade Cancelada", value=f"{qtd_cancelada:,}".replace(",", "."))
    kpi3.metric(label="Quantidade Faturada", value=f"{qtd_faturada:,}".replace(",", "."))
    # Adicionar Quantidade em Aberta (não cancelada e não faturada)
    qtd_aberta = dff[(dff["StatusKPI"] != "Cancelado") & (dff["StatusKPI"] != "Faturado")]["QuantidadeKPI"].sum()
    kpi4.metric(label="Quantidade em Aberta", value=f"{qtd_aberta:,}".replace(",", "."))

    st.markdown("---")
    st.subheader("Análise Temporal")

    # --- Gráficos Temporais ---
    col_tempo1, col_tempo2 = st.columns(2)

    # Inicializar filtros interativos
    filtro_ano = None
    filtro_mes = None

    with col_tempo1:
        criado_tempo_mes = dff.resample("MS", on="DataCriacao")["QuantidadeKPI"].sum().reset_index()
        fig_criado_tempo = px.line(criado_tempo_mes, x="DataCriacao", y="QuantidadeKPI", title="Volume Criado por Mês", markers=True, labels={"DataCriacao": "Mês", "QuantidadeKPI": "Quantidade"})
        fig_criado_tempo.update_layout(hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20))
        selected_points = st.plotly_chart(fig_criado_tempo, use_container_width=True)

        criado_ano = dff.groupby("Ano")["QuantidadeKPI"].sum().reset_index()
        fig_criado_ano = px.bar(criado_ano, x="Ano", y="QuantidadeKPI", title="Volume Criado por Ano", labels={"Ano": "Ano", "QuantidadeKPI": "Quantidade"}, text_auto=True)
        fig_criado_ano.update_layout(margin=dict(l=20, r=20, t=40, b=20))
        selected_points_ano = st.plotly_chart(fig_criado_ano, use_container_width=True)
        # Exemplo simples: se clicar em um ano, filtrar os KPIs (a implementar)

    with col_tempo2:
        faturado_tempo_mes = dff[dff["StatusKPI"] == "Faturado"].resample("MS", on="DataCriacao")["QuantidadeKPI"].sum().reset_index()
        fig_faturado_tempo = px.line(faturado_tempo_mes, x="DataCriacao", y="QuantidadeKPI", title="Volume Faturado por Mês", markers=True, labels={"DataCriacao": "Mês", "QuantidadeKPI": "Quantidade Faturada"})
        fig_faturado_tempo.update_layout(hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_faturado_tempo, use_container_width=True)

        faturado_ano = dff[dff["StatusKPI"] == "Faturado"].groupby("Ano")["QuantidadeKPI"].sum().reset_index()
        fig_faturado_ano = px.bar(faturado_ano, x="Ano", y="QuantidadeKPI", title="Volume Faturado por Ano", labels={"Ano": "Ano", "QuantidadeKPI": "Quantidade Faturada"}, text_auto=True)
        fig_faturado_ano.update_layout(margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_faturado_ano, use_container_width=True)

    st.markdown("---")
    st.subheader("Análise por Grupos")

    # --- Gráficos de Grupo ---
    col_grupo1, col_grupo2 = st.columns(2)

    with col_grupo1:
        if "Franqueado" in dff.columns:
            franqueado_volume = dff.groupby("Franqueado")["QuantidadeKPI"].sum().reset_index()
            if "Não Especificado" in franqueado_volume["Franqueado"].tolist() and len(selected_franqueado) > 0 and (len(selected_franqueado) != 1 or selected_franqueado[0] != "Não Especificado"):
                 franqueado_volume_filtrado = franqueado_volume[franqueado_volume["Franqueado"] != "Não Especificado"]
            else:
                 franqueado_volume_filtrado = franqueado_volume
            top_franqueados = franqueado_volume_filtrado.nlargest(15, "QuantidadeKPI")
            fig_grupo_franqueado_vol = px.bar(top_franqueados, y="Franqueado", x="QuantidadeKPI", title="Top 15 Franqueados por Volume (Quantidade)", orientation="h", labels={"Franqueado": "Franqueado", "QuantidadeKPI": "Quantidade Total"})
            fig_grupo_franqueado_vol.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
            fig_grupo_franqueado_vol.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_grupo_franqueado_vol, use_container_width=True)
        else:
            st.info("Coluna 'Franqueado' não encontrada nos dados filtrados. Gráfico de Franqueados não será exibido.")

    with col_grupo2:
        nome_volume = dff.groupby("NomeCompletoZ")["QuantidadeKPI"].sum().nlargest(10).reset_index()
        fig_nome_completo_vol = px.bar(nome_volume, y="NomeCompletoZ", x="QuantidadeKPI", title="Top 10 Colaboradores por Volume (Quantidade)", orientation="h", labels={"NomeCompletoZ": "Colaborador", "QuantidadeKPI": "Quantidade Total"})
        fig_nome_completo_vol.update_layout(xaxis_title="Quantidade Total", yaxis_title=None, margin=dict(l=20, r=20, t=40, b=20))
        fig_nome_completo_vol.update_yaxes(autorange="reversed")
        st.plotly_chart(fig_nome_completo_vol, use_container_width=True)

    # --- Novos Gráficos Solicitados ---

    # Carregar dados exportados para criar gráficos adicionais
    try:
        df_exported = pd.read_excel("dados_dashboard_exportados.xlsx")
    except Exception as e:
        st.error(f"Erro ao carregar dados exportados para gráficos adicionais: {e}")
        df_exported = None

    if df_exported is not None:
        st.markdown("---")
        st.subheader("Análise Adicional")

        # Gráfico de pizza com colunas BrandCategory (Q) e OticoSport (R)
        if "BrandCategory" in df_exported.columns and "OticoSport" in df_exported.columns:
            pie_data = df_exported.groupby("BrandCategory")["OticoSport"].count().reset_index()
            fig_pie = px.pie(pie_data, names="BrandCategory", values="OticoSport", title="Distribuição por BrandCategory e OticoSport")
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Colunas 'BrandCategory' e/ou 'OticoSport' não encontradas nos dados exportados para gráfico de pizza.")

        # Top 10 com coluna BrandCode (O)
        if "BrandCode" in df_exported.columns:
            top10_brandcode = df_exported.groupby("BrandCode")["QuantidadeKPI"].sum().nlargest(10).reset_index()
            fig_top10 = px.bar(top10_brandcode, x="BrandCode", y="QuantidadeKPI", title="Top 10 BrandCode por Quantidade", labels={"BrandCode": "BrandCode", "QuantidadeKPI": "Quantidade Total"})
            st.plotly_chart(fig_top10, use_container_width=True)
        else:
            st.info("Coluna 'BrandCode' não encontrada nos dados exportados para gráfico Top 10.")

import datetime

st.write("---")
current_year = datetime.datetime.now().year
st.markdown(f"**_BI After Sales EssilorLuxottica | {current_year}_**")

# Link para download da planilha completa
file_path = "upload/teste_d11.xlsx"
if os.path.exists(file_path):
    with open(file_path, "rb") as file:
        btn = st.download_button(
            label="Download da planilha completa teste_d11.xlsx",
            data=file,
            file_name="teste_d11.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Arquivo de planilha completa não encontrado para download.")
