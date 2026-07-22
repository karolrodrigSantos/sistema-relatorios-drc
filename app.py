import streamlit as st
import pandas as pd
import sqlite3
import numpy as np
import plotly.express as px
from datetime import datetime

# ==========================================
# CONSTANTES E CONFIGURAÇÕES DA INTERFACE
# ==========================================
st.set_page_config(page_title="Sistema de Inteligência de Ativos - DRC", layout="wide")

DB_NAME = "ativos_engenharia.db"

# ==========================================
# 1. ESTRUTURA DE DADOS (SQL & BANCO DE DADOS)
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Tabela de Unidades de Patrimônio (UP)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tabela_up (
        codigo_up TEXT PRIMARY KEY,
        nomenclatura_up TEXT NOT NULL,
        tipo_massa_individual TEXT
    )""")
    
    # Tabela de Unidades de Acréscimo e Recuperação (UAR)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tabela_uar (
        codigo_uar TEXT PRIMARY KEY,
        codigo_up TEXT,
        nomenclatura_uar TEXT NOT NULL,
        FOREIGN KEY (codigo_up) REFERENCES tabela_up(codigo_up)
    )""")
    
    # Tabela Histórica de Preços com chaves para i0 (Mês Ref) e Atributos Técnicos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS historico_precos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo_uar TEXT,
        banco TEXT NOT NULL, -- Sabesp, SINAPI, TCPO
        disciplina TEXT NOT NULL, -- Civil, Mecânica, Elétrica, Equipamentos
        i0_data TEXT NOT NULL, -- Data de Referência do Índice (YYYY-MM)
        preco_unitario REAL NOT NULL,
        atributo_acionamento TEXT, -- Tabela IV do manual
        FOREIGN KEY (codigo_uar) REFERENCES tabela_uar(codigo_uar)
    )""")
    
    # Carga inicial de Mock Data estruturado caso esteja vazio para garantir Forward Momentum
    cursor.execute("SELECT COUNT(*) FROM tabela_up")
    if cursor.fetchone()[0] == 0:
        # Carga UP
        ups = [
            ('02', 'Estruturas Saneamento', 'Individual'),
            ('05', 'Equipamentos de Bombeamentos', 'Individual'),
            ('06', 'Instalações e Equipamentos Elétricos', 'Individual')
        ]
        cursor.executemany("INSERT INTO tabela_up VALUES (?,?,?)", ups)
        
        # Carga UAR
        uars = [
            ('0200110', '02', 'Base de concreto'),
            ('0500101', '05', 'Bomba centrífuga horizontal'),
            ('0601001', '06', 'Inversor de frequência')
        ]
        cursor.executemany("INSERT INTO tabela_uar VALUES (?,?,?)", uars)
        
        # Histórico de Preços fake cobrindo variações temporais de i0
        precos_mock = [
            ('0500101', 'Sabesp', 'Equipamentos', '2026-01', 15000.00, 'Elétrico'),
            ('0500101', 'Sabesp', 'Equipamentos', '2026-06', 15800.00, 'Elétrico'),
            ('0500101', 'SINAPI', 'Equipamentos', '2026-01', 14200.00, 'Elétrico'),
            ('0500101', 'SINAPI', 'Equipamentos', '2026-06', 15100.00, 'Elétrico'),
            ('0500101', 'TCPO', 'Equipamentos', '2026-01', 14800.00, 'Elétrico'),
            ('0500101', 'TCPO', 'Equipamentos', '2026-06', 15400.00, 'Elétrico'),
            
            ('0200110', 'Sabesp', 'Civil', '2026-01', 250.00, 'Não se Aplica'),
            ('0200110', 'Sabesp', 'Civil', '2026-06', 270.00, 'Não se Aplica'),
            ('0200110', 'SINAPI', 'Civil', '2026-01', 240.00, 'Não se Aplica'),
            ('0200110', 'SINAPI', 'Civil', '2026-06', 268.00, 'Não se Aplica'),
        ]
        cursor.executemany("""
            INSERT INTO historico_precos (codigo_uar, banco, disciplina, i0_data, preco_unitario, atributo_acionamento) 
            VALUES (?,?,?,?,?,?)
        """, precos_mock)
        
    conn.commit()
    conn.close()

init_db()

# ==========================================
# HELPER FUNCTIONS (SQL FETCH & PANDAS)
# ==========================================
def get_connection():
    return sqlite3.connect(DB_NAME)

# ==========================================
# INTERFACE STREAMLIT
# ==========================================
st.title("⚡ Sistema de Inteligência de Ativos e Relatórios de Engenharia")
st.markdown("### Metodologia DRC & Análise de Séries Temporais ($i_0$)")
st.write("---")

# Sidebar - Filtros Dinâmicos e Upload
st.sidebar.header("⚙️ Painel de Controle e Carga")

# Funcionalidade 3.1: Botão para upload de novos arquivos de preços
uploaded_file = st.sidebar.file_uploader("Upload de Preços (CSV/Excel)", type=["csv", "xlsx"])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_upload = pd.read_csv(uploaded_file)
        else:
            df_upload = pd.read_excel(uploaded_file)
        st.sidebar.success("Arquivo processado com sucesso!")
        # Próximo passo em produção: mapear e dar `df_upload.to_sql('historico_precos', ...)`
    except Exception as e:
        st.sidebar.error(f"Erro ao ler arquivo: {e}")

# Funcionalidade 3.2: Filtros Dinâmicos baseados no BD
conn = get_connection()
df_ups = pd.read_sql_query("SELECT * FROM tabela_up", conn)
df_uars = pd.read_sql_query("SELECT * FROM tabela_uar", conn)

selected_up = st.sidebar.selectbox("Filtrar por Código UP", ["Todos"] + list(df_ups['codigo_up'].unique()))

if selected_up != "Todos":
    filtered_uars = df_uars[df_uars['codigo_up'] == selected_up]['codigo_uar'].unique()
else:
    filtered_uars = df_uars['codigo_uar'].unique()

selected_uar = st.sidebar.selectbox("Filtrar por Código UAR", ["Todos"] + list(filtered_uars))
conn.close()

# Abas de Relatórios Obrigatórios (Requisitos de Negócio)
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Evolução Temporal & Índices", 
    "📊 Variação por Segmento", 
    "🏭 Inflação Setorial", 
    "🔍 Módulo DRC (Comparativo)"
])

# ==========================================
# ABA 1: RELATÓRIO DE ÍNDICES PARA ATUALIZAÇÃO & GRÁFICOS
# ==========================================
with tab1:
    st.header("Relatório de Índices para Atualização de Ativos Históricos")
    
    query = """
        SELECT hp.*, u.codigo_up, u.nomenclatura_uar 
        FROM historico_precos hp
        JOIN tabela_uar u ON hp.codigo_uar = u.codigo_uar
    """
    conn = get_connection()
    df_precos = pd.read_sql_query(query, conn)
    conn.close()
    
    # Aplicação dos filtros dinâmicos
    if selected_up != "Todos":
        df_precos = df_precos[df_precos['codigo_up'] == selected_up]
    if selected_uar != "Todos":
        df_precos = df_precos[df_precos['codigo_uar'] == selected_uar]
        
    if not df_precos.empty:
        # Funcionalidade 3.3: Visualização de gráficos de linha para evolução dos preços
        st.subheader("Visualização Temporal de Evolução do Preço de Referência")
        fig = px.line(df_precos, x='i0_data', y='preco_unitario', color='banco', 
                      text='preco_unitario', markers=True, title="Evolução Temporal por Banco (i0)")
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
        
        # Tabela de Índices Calulados (i0 Base vs i0 Atual)
        st.subheader("Tabela Corrente de Dados Históricos ($i_0$)")
        st.dataframe(df_precos[['codigo_up', 'codigo_uar', 'nomenclatura_uar', 'banco', 'i0_data', 'preco_unitario', 'atributo_acionamento']], use_container_width=True)
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")

# ==========================================
# ABA 2: RELATÓRIO DE VARIAÇÃO DE PREÇO POR SEGMENTO
# ==========================================
with tab2:
    st.header("Relatório de Variação de Preço por Segmento / Disciplina")
    st.write("Comparativo direto entre as oscilações de mercado nos bancos: Sabesp, SINAPI e TCPO.")
    
    conn = get_connection()
    df_seg = pd.read_sql_query("SELECT banco, disciplina, AVG(preco_unitario) as preco_medio FROM historico_precos GROUP BY banco, disciplina", conn)
    conn.close()
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Preço Médio por Disciplina / Banco")
        st.dataframe(df_seg, use_container_width=True)
    with col2:
        fig_bar = px.bar(df_seg, x='disciplina', y='preco_medio', color='banco', barmode='group',
                         title="Análise Comparativa de Preço Médio por Disciplina")
        st.plotly_chart(fig_bar, use_container_width=True)

# ==========================================
# ABA 3: RELATÓRIO DE INFLAÇÃO SETORIAL
# ==========================================
with tab3:
    st.header("Relatório de Inflação Setorial (i0 Base vs i0 Fim)")
    st.write("Foco analítico segregado em **Inflação de Demanda** (Máquinas/Equipamentos) e **Inflação de Mão de Obra** (Civil).")
    
    # Cálculo dinâmico comparando 2026-01 com 2026-06
    conn = get_connection()
    df_inf = pd.read_sql_query("SELECT disciplina, i0_data, AVG(preco_unitario) as preco FROM historico_precos GROUP BY disciplina, i0_data", conn)
    conn.close()
    
    if not df_inf.empty and len(df_inf['i0_data'].unique()) >= 2:
        meses = sorted(df_inf['i0_data'].unique())
        mes_ini, mes_fim = meses[0], meses[-1]
        
        st.info(f"Análise inflacionária calculada automaticamente entre o período **{mes_ini}** e **{mes_fim}**.")
        
        linhas_inflacao = []
        for disc in df_inf['disciplina'].unique():
            p_ini = df_inf[(df_inf['disciplina'] == disc) & (df_inf['i0_data'] == mes_ini)]['preco'].values
            p_fim = df_inf[(df_inf['disciplina'] == disc) & (df_inf['i0_data'] == mes_fim)]['preco'].values
            
            if len(p_ini) > 0 and len(p_fim) > 0:
                variacao = ((p_fim[0] - p_ini[0]) / p_ini[0]) * 100
                tipo_inflacao = "Inflação de Demanda" if disc in ['Equipamentos', 'Mecânica'] else "Inflação de Mão de Obra / Insumos"
                linhas_inflacao.append({"Disciplina/Segmento": disc, "Tipo de Métrica": tipo_inflacao, f"Preço em {mes_ini}": f"R$ {p_ini[0]:,.2f}", f"Preço em {mes_fim}": f"R$ {p_fim[0]:,.2f}", "Variação Acumulada": f"{variacao:.2f}%"})
        
        st.table(pd.DataFrame(linhas_inflacao))
    else:
        st.warning("É necessário possuir pelo menos dois meses distintos de referência ($i_0$) para rodar o cálculo de inflação setorial.")

# ==========================================
# ABA 4: MÓDULO DRC (COMPARATIVO DE PROPOSTAS) - REGRA CRÍTICA
# ==========================================
with tab4:
    st.header("Módulo DRC - Comparativo Unificado de Propostas de Fornecedores")
    st.markdown("> **Regra Crítica Regulatória de Negócio:** Se o preço do fornecedor desviar mais que **± 5%** em relação ao referencial técnico do mês, a célula será destacada em **vermelho**.")
    
    # Inputs de teste para simulação do Fornecedor
    st.subheader("Simulador de Análise de Proposta")
    col_sim1, col_sim2, col_sim3 = st.columns(3)
    with col_sim1:
        ref_banco = st.selectbox("Banco Referencial Técnico", ["Sabesp", "SINAPI", "TCPO"])
    with col_sim2:
        item_analise = st.selectbox("Item UAR para Análise", ["0500101 - Bomba centrífuga horizontal", "0200110 - Base de concreto"])
    with col_sim3:
        preco_fornecedor = st.number_input("Preço Ofertado pelo Fornecedor (R$)", value=16000.0)
        
    cod_uar_sim = item_analise.split(" - ")[0]
    
    # Busca o referencial no BD para o mês mais recente
    conn = get_connection()
    ref_data = pd.read_sql_query(f"SELECT preco_unitario FROM historico_precos WHERE codigo_uar = '{cod_uar_sim}' AND banco = '{ref_banco}' ORDER BY i0_data DESC LIMIT 1", conn)
    conn.close()
    
    if not ref_data.empty:
        preco_ref = ref_data['preco_unitario'].values[0]
        desvio_percentual = ((preco_fornecedor - preco_ref) / preco_ref) * 100
        
        # Construção da tabela de exibição
        df_drc = pd.DataFrame([{
            "Item UAR": cod_uar_sim,
            "Banco Ref": ref_banco,
            "Preço de Referência (R$)": preco_ref,
            "Preço do Fornecedor (R$)": preco_fornecedor,
            "Desvio (%)": desvio_percentual
        }])
        
        # Função de Styler para aplicar a Regra Crítica de ± 5%
        def colorir_desvio(val):
            # Aplica a regra nas colunas de preço do fornecedor e desvio caso saia da margem de 5%
            color = 'background-color: #ffcccc; color: #cc0000; font-weight: bold;' if abs(desvio_percentual) > 5 else ''
            return [color if col in ['Preço do Fornecedor (R$)', 'Desvio (%)'] else '' for col in df_drc.columns]

        st.markdown("### Resultado do Laudo de Valoração (DRC vs Proposta)")
        
        # Exibe o dataframe aplicando o style por linha baseado na regra crítica
        st.dataframe(df_drc.style.apply(colorir_desvio, axis=1).format({
            "Preço de Referência (R$)": "R$ {:,.2f}",
            "Preço do Fornecedor (R$)": "R$ {:,.2f}",
            "Desvio (%)": "{:.2f}%"
        }), use_container_width=True)
        
        if abs(desvio_percentual) > 5:
            st.error(f"🚨 Atenção: A proposta do fornecedor apresenta um desvio de {desvio_percentual:.2f}%, estourando o limite regulatório de ±5%!")
        else:
            st.success("✅ Proposta validada e em conformidade técnica com as margens do DRC regulatório.")
    else:
        st.warning("Insira dados de referência válidos para realizar o cruzamento DRC.")