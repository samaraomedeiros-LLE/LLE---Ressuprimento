import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ressuprimento | Grupo LLE",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CORES LLE ────────────────────────────────────────────────────────────────
AZUL_ESCURO  = "#041747"
AZUL_MEDIO   = "#0071FE"
AMARELO      = "#FAC318"
VERDE        = "#0F8C3B"
VERMELHO     = "#C0392B"
LARANJA      = "#E67E22"
CINZA_CLARO  = "#F4F6FA"
BRANCO       = "#FFFFFF"

# ─── CSS CUSTOMIZADO ──────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');

  html, body, [class*="css"] {{
      font-family: 'Montserrat', sans-serif;
  }}

  /* Sidebar */
  [data-testid="stSidebar"] {{
      background-color: {AZUL_ESCURO};
  }}
  [data-testid="stSidebar"] * {{
      color: {BRANCO} !important;
  }}
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stFileUploader label {{
      color: {AMARELO} !important;
      font-weight: 700;
      font-size: 0.82rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
  }}

  /* Header principal */
  .header-banner {{
      background: linear-gradient(135deg, {AZUL_ESCURO} 0%, #0a2560 100%);
      padding: 18px 28px;
      border-radius: 12px;
      margin-bottom: 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
  }}
  .header-title {{
      color: {BRANCO};
      font-size: 1.55rem;
      font-weight: 800;
      margin: 0;
  }}
  .header-subtitle {{
      color: {AMARELO};
      font-size: 0.82rem;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      margin-top: 2px;
  }}
  .header-date {{
      color: {AMARELO};
      font-size: 0.9rem;
      font-weight: 700;
      text-align: right;
  }}

  /* KPI Cards */
  .kpi-card {{
      background: {BRANCO};
      border-radius: 12px;
      padding: 18px 22px;
      box-shadow: 0 2px 12px rgba(4,23,71,0.09);
      border-top: 4px solid {AZUL_MEDIO};
      text-align: center;
  }}
  .kpi-card.critico {{ border-top-color: {VERMELHO}; }}
  .kpi-card.super   {{ border-top-color: {AMARELO}; }}
  .kpi-card.verde   {{ border-top-color: {VERDE}; }}
  .kpi-label {{
      font-size: 0.7rem;
      font-weight: 700;
      color: #7a8499;
      text-transform: uppercase;
      letter-spacing: 0.07em;
      margin-bottom: 6px;
  }}
  .kpi-value {{
      font-size: 2.1rem;
      font-weight: 800;
      color: {AZUL_ESCURO};
      line-height: 1;
  }}
  .kpi-value.red   {{ color: {VERMELHO}; }}
  .kpi-value.amber {{ color: {AMARELO}; }}
  .kpi-value.green {{ color: {VERDE}; }}
  .kpi-sub {{
      font-size: 0.72rem;
      color: #9aa0b0;
      margin-top: 4px;
  }}

  /* Tabela de prioridades */
  .priority-badge {{
      display: inline-block;
      padding: 2px 10px;
      border-radius: 20px;
      font-size: 0.72rem;
      font-weight: 700;
  }}

  /* Section titles */
  .section-title {{
      font-size: 1rem;
      font-weight: 700;
      color: {AZUL_ESCURO};
      margin: 18px 0 10px 0;
      padding-left: 10px;
      border-left: 4px solid {AMARELO};
  }}

  /* Dataframe override */
  .stDataFrame {{
      border-radius: 10px;
      overflow: hidden;
  }}

  /* Remove padding padrão do main */
  .main .block-container {{
      padding-top: 1rem;
      padding-bottom: 1rem;
  }}

  /* Botão upload */
  [data-testid="stFileUploader"] section {{
      border: 2px dashed {AZUL_MEDIO};
      border-radius: 10px;
      background: #eaf2ff;
  }}

  /* Alerta */
  .alert-box {{
      background: #fff3cd;
      border: 1px solid {AMARELO};
      border-radius: 8px;
      padding: 10px 16px;
      font-size: 0.82rem;
      color: {AZUL_ESCURO};
      margin: 8px 0;
  }}
  .alert-critical {{
      background: #fdecea;
      border-color: {VERMELHO};
  }}
</style>
""", unsafe_allow_html=True)


# ─── FUNÇÕES ──────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def carregar_dados(file_bytes, filename):
    """Carrega e processa a planilha de ressuprimento."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="new sheet", header=5, engine="openpyxl")
    except Exception as e:
        return None, f"Erro ao ler 'new sheet': {e}"

    cols = [
        'pct_saldo', 'maximo', 'rua', 'curva', 'pulmao', 'score', 'estacao',
        'concatenar', 'cod_produto', 'referencia', 'desc_produto', 'end_reduzido',
        'desc_endereco', 'dt_validade', 'est_disponivel', 'un_estoque',
        'est_atual_un_padrao', 'un_padrao', 'ativo', 'picking', 'bloqueado',
        'multi_produto', 'expedicao', 'complemento', 'endereco'
    ]
    df.columns = cols[: len(df.columns)]

    # Filtrar mezanino (ruas 46 e 47)
    df = df[df['rua'].isin([46, 47])].copy()

    # Tipos
    df['score']     = pd.to_numeric(df['score'],     errors='coerce').fillna(0)
    df['pct_saldo'] = pd.to_numeric(df['pct_saldo'], errors='coerce')
    df['maximo']    = pd.to_numeric(df['maximo'],    errors='coerce')
    df['curva']     = df['curva'].astype(str).str.strip()
    df['estacao']   = df['estacao'].astype(str).str.strip()
    df['pulmao']    = df['pulmao'].astype(str).str.strip()

    # Classificação de prioridade
    def classi(row):
        if pd.isna(row['pct_saldo']):
            return 'Sem Pulmão'
        s, p = row['score'], row['pct_saldo']
        if row['curva'] in ['A', 'B'] and p <= 0.10:
            return '🔴 Supercrítico'
        elif s >= 500 or p < 0.10:
            return '🟠 Crítico'
        elif s > 0 and p <= 0.51:
            return '🟡 Atenção'
        elif s > 0:
            return '🟢 Normal'
        return 'Inativo'

    df['prioridade'] = df.apply(classi, axis=1)
    return df, None


def kpi_card(label, value, sub="", css_class="", value_class=""):
    return f"""
    <div class="kpi-card {css_class}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {value_class}">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


def fmt_pct(v):
    if pd.isna(v):
        return "—"
    return f"{v*100:.1f}%"


def cor_saldo(v):
    if pd.isna(v):
        return "background-color: #e9ecef; color: #6c757d;"
    if v < 0.10:
        return f"background-color: #fdecea; color: {VERMELHO}; font-weight:700;"
    if v <= 0.31:
        return f"background-color: #fff3cd; color: #856404; font-weight:700;"
    if v <= 0.51:
        return "background-color: #fff8e1; color: #a07800;"
    return f"background-color: #e8f5e9; color: {VERDE};"


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 10px 0 20px 0;">
        <div style="font-size:2.2rem;">📦</div>
        <div style="font-size:1.1rem; font-weight:800; color:{AMARELO};">GRUPO LLE</div>
        <div style="font-size:0.72rem; color:#ccc; letter-spacing:0.1em;">RESSUPRIMENTO DIÁRIO</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "📂 CARREGAR PLANILHA (.xlsx)",
        type=["xlsx"],
        help="Exporte do Sankhya e carregue aqui a planilha de ressuprimento.",
    )

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.75rem; color:{AMARELO}; font-weight:700; letter-spacing:0.08em;'>FILTROS</div>", unsafe_allow_html=True)

    filtro_estacao  = st.multiselect("🏭 ESTAÇÃO", [], key="est")
    filtro_curva    = st.multiselect("📊 CURVA ABC", ["A", "B", "C", "D", "E", "F"], key="curva")
    filtro_pulmao   = st.selectbox("💧 PULMÃO", ["Todos", "COM PULMÃO", "SEM PULMÃO"], key="pulmao")
    filtro_prio     = st.multiselect(
        "🚦 PRIORIDADE",
        ["🔴 Supercrítico", "🟠 Crítico", "🟡 Atenção", "🟢 Normal"],
        key="prio"
    )

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:0.68rem; color:#aab0c0; text-align:center; padding:4px;">
        Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}
    </div>""", unsafe_allow_html=True)


# ─── HEADER ───────────────────────────────────────────────────────────────────
hoje = datetime.now().strftime("%A, %d/%m/%Y").capitalize()
st.markdown(f"""
<div class="header-banner">
    <div>
        <div class="header-title">📦 Ressuprimento — Picking Mezanino</div>
        <div class="header-subtitle">Grupo LLE · Coordenação Logística</div>
    </div>
    <div class="header-date">📅 {hoje}</div>
</div>
""", unsafe_allow_html=True)


# ─── ESTADO: SEM ARQUIVO ──────────────────────────────────────────────────────
if uploaded is None:
    st.markdown(f"""
    <div style="text-align:center; padding: 60px 20px; background:{CINZA_CLARO}; border-radius:16px; margin-top:20px;">
        <div style="font-size:3.5rem; margin-bottom:16px;">📋</div>
        <div style="font-size:1.3rem; font-weight:700; color:{AZUL_ESCURO};">Carregue a planilha do Sankhya</div>
        <div style="font-size:0.88rem; color:#7a8499; margin-top:8px;">
            Use o painel lateral para fazer upload do arquivo <b>.xlsx</b> exportado pelo ERP.<br>
            Os dados serão processados automaticamente.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ─── CARREGAR DADOS ───────────────────────────────────────────────────────────
with st.spinner("⏳ Processando planilha..."):
    df_raw, erro = carregar_dados(uploaded.read(), uploaded.name)

if erro:
    st.error(f"❌ {erro}")
    st.stop()

# Atualizar filtro de estações disponíveis
estacoes_disp = sorted(df_raw['estacao'].dropna().unique().tolist())
# Recriar selectbox de estação com opções reais
with st.sidebar:
    st.empty()  # placeholder já foi renderizado acima

# Aplicar filtros
df = df_raw.copy()

# Refaz filtros com os valores reais no session_state
est_sel   = st.session_state.get("est",  [])
curva_sel = st.session_state.get("curva", [])
pulmao_sel= st.session_state.get("pulmao", "Todos")
prio_sel  = st.session_state.get("prio", [])

if est_sel:
    df = df[df['estacao'].isin(est_sel)]
if curva_sel:
    df = df[df['curva'].isin(curva_sel)]
if pulmao_sel != "Todos":
    df = df[df['pulmao'] == pulmao_sel]
if prio_sel:
    df = df[df['prioridade'].isin(prio_sel)]

# Reconstruir sidebar filtro estação com opções reais
with st.sidebar:
    # Limpa e re-renderiza apenas o multiselect de estação
    st.session_state['_est_opts'] = estacoes_disp

# ─── KPIs ─────────────────────────────────────────────────────────────────────
df_ativo      = df_raw[df_raw['score'] > 0]
total_tarefas = len(df_ativo)
criticos      = ((df_raw['score'] >= 500) | (df_raw['pct_saldo'] < 0.10)).sum()
supercrit     = ((df_raw['curva'].isin(['A', 'B'])) & (df_raw['pct_saldo'] <= 0.30)).sum()
pct_cobertura = criticos / total_tarefas if total_tarefas else 0

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(kpi_card("Total de Tarefas", f"{total_tarefas:,}", "Score > 0 no mezanino"), unsafe_allow_html=True)
with col2:
    st.markdown(kpi_card("Críticos", f"{criticos:,}", "Score≥500 ou saldo<10%", "critico", "red"), unsafe_allow_html=True)
with col3:
    st.markdown(kpi_card("Supercríticos", f"{supercrit:,}", "Curva A/B com saldo ≤30%", "super", "amber"), unsafe_allow_html=True)
with col4:
    st.markdown(kpi_card("% Cobertura Crítica", f"{pct_cobertura*100:.1f}%", "Críticos / Total tarefas", "verde", "green"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── TABS PRINCIPAIS ──────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🎯 Prioridades do Dia", "📊 Análise por Estação", "📋 Base Completa"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — PRIORIDADES DO DIA
# ══════════════════════════════════════════════════════════════════════════════
with tab1:

    # Alerta crítico
    n_sc = ((df['curva'].isin(['A', 'B'])) & (df['pct_saldo'] <= 0.10)).sum()
    if n_sc > 0:
        st.markdown(f"""
        <div class="alert-box alert-critical">
            🚨 <b>{n_sc} itens Curva A/B com saldo ≤ 10%!</b> — Ação imediata necessária.
        </div>""", unsafe_allow_html=True)

    # Montar tabela de prioridades (score > 0, ordenado por score desc)
    df_prio = df[df['score'] > 0].copy()
    df_prio = df_prio.sort_values('score', ascending=False).reset_index(drop=True)
    df_prio.index = df_prio.index + 1  # linha 1-based

    col_a, col_b = st.columns([2, 1])

    with col_a:
        st.markdown('<div class="section-title">🏆 Ranking de Prioridades</div>', unsafe_allow_html=True)

        # Selecionar top N
        top_n = st.slider("Exibir top:", min_value=10, max_value=min(200, len(df_prio)), value=50, step=10)
        df_view = df_prio.head(top_n)[
            ['prioridade', 'curva', 'estacao', 'cod_produto', 'desc_produto',
             'desc_endereco', 'pct_saldo', 'pulmao', 'score']
        ].rename(columns={
            'prioridade':    'Prioridade',
            'curva':         'Curva',
            'estacao':       'Estação',
            'cod_produto':   'Cód.',
            'desc_produto':  'Descrição',
            'desc_endereco': 'Endereço',
            'pct_saldo':     '% Saldo',
            'pulmao':        'Pulmão',
            'score':         'Score',
        })

        # Formatar % Saldo
        df_view['% Saldo'] = df_view['% Saldo'].apply(fmt_pct)
        df_view['Score']   = df_view['Score'].apply(lambda x: f"{x:,.0f}")

        st.dataframe(
            df_view,
            use_container_width=True,
            height=480,
            hide_index=False,
        )

        # Download
        csv = df_prio[['prioridade','curva','estacao','cod_produto','desc_produto',
                        'desc_endereco','pct_saldo','pulmao','score']].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "⬇️ Exportar lista completa (.csv)",
            data=csv,
            file_name=f"prioridades_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )

    with col_b:
        st.markdown('<div class="section-title">📌 Distribuição de Prioridades</div>', unsafe_allow_html=True)

        dist = df_prio['prioridade'].value_counts().reset_index()
        dist.columns = ['Prioridade', 'Qtd']
        color_map = {
            '🔴 Supercrítico': VERMELHO,
            '🟠 Crítico':      LARANJA,
            '🟡 Atenção':      AMARELO,
            '🟢 Normal':       VERDE,
        }
        fig_pizza = px.pie(
            dist, names='Prioridade', values='Qtd',
            color='Prioridade', color_discrete_map=color_map,
            hole=0.45,
        )
        fig_pizza.update_traces(textposition='outside', textinfo='percent+label')
        fig_pizza.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=False,
            height=260,
            paper_bgcolor='rgba(0,0,0,0)',
        )
        st.plotly_chart(fig_pizza, use_container_width=True)

        st.markdown('<div class="section-title">📦 Curva ABC — Críticos</div>', unsafe_allow_html=True)
        crit_curva = df_prio[df_prio['prioridade'].isin(['🔴 Supercrítico', '🟠 Crítico'])]
        crit_curva_agg = crit_curva['curva'].value_counts().reset_index()
        crit_curva_agg.columns = ['Curva', 'Críticos']
        fig_bar = px.bar(
            crit_curva_agg.sort_values('Curva'),
            x='Curva', y='Críticos',
            color='Curva',
            color_discrete_map={'A': VERMELHO, 'B': LARANJA, 'C': AMARELO,
                                 'D': '#3498db', 'E': '#9b59b6', 'F': '#95a5a6'},
            text='Críticos',
        )
        fig_bar.update_traces(textposition='outside')
        fig_bar.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=False, height=220,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(showgrid=False, showticklabels=False),
            xaxis=dict(showgrid=False),
        )
        st.plotly_chart(fig_bar, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ANÁLISE POR ESTAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
with tab2:

    df_est = df_raw.copy()

    # Agrupar por estação
    resumo = (
        df_est.groupby('estacao', as_index=False)
        .agg(
            Total=('score', lambda x: (x > 0).sum()),
            Criticos=('score', lambda x: (
                ((x >= 500) | (df_est.loc[x.index, 'pct_saldo'] < 0.10)) & (x >= 0)
            ).sum()),
            Supercriticos=('curva', lambda x: (
                (x.isin(['A', 'B'])) & (df_est.loc[x.index, 'pct_saldo'] <= 0.30)
            ).sum()),
        )
    )
    resumo['% Crítica'] = (resumo['Criticos'] / resumo['Total'].replace(0, np.nan) * 100).round(1)
    resumo = resumo.sort_values('Criticos', ascending=False).reset_index(drop=True)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown('<div class="section-title">🏭 Resumo por Estação</div>', unsafe_allow_html=True)

        st.dataframe(
            resumo.rename(columns={
                'estacao': 'Estação',
                'Total': 'Total Tarefas',
                'Criticos': 'Críticos',
                'Supercriticos': 'Supercríticos',
                '% Crítica': '% Crítica',
            }),
            use_container_width=True,
            height=420,
            hide_index=True,
        )

    with col2:
        st.markdown('<div class="section-title">📈 % Crítica por Estação</div>', unsafe_allow_html=True)

        fig_h = px.bar(
            resumo.sort_values('% Crítica'),
            y='estacao', x='% Crítica',
            orientation='h',
            text='% Crítica',
            color='% Crítica',
            color_continuous_scale=[[0, VERDE], [0.3, AMARELO], [0.6, LARANJA], [1, VERMELHO]],
        )
        fig_h.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_h.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            height=420,
            coloraxis_showscale=False,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showgrid=False, showticklabels=False, title=''),
            yaxis=dict(title=''),
        )
        st.plotly_chart(fig_h, use_container_width=True)

    # Linha 2: Itens críticos A/B por estação
    st.markdown('<div class="section-title">🔴 Itens Supercríticos (A/B · saldo ≤ 30%) por Estação</div>', unsafe_allow_html=True)

    fig_sc = px.bar(
        resumo.sort_values('Supercriticos', ascending=False),
        x='estacao', y='Supercriticos',
        text='Supercriticos',
        color_discrete_sequence=[AMARELO],
    )
    fig_sc.update_traces(textposition='outside')
    fig_sc.update_layout(
        margin=dict(t=10, b=40, l=10, r=10),
        height=280,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(tickangle=-30, title=''),
        yaxis=dict(showgrid=True, gridcolor='#e8ecf0', title='Qtd Itens'),
    )
    st.plotly_chart(fig_sc, use_container_width=True)

    # Detalhe por estação
    st.markdown('<div class="section-title">🔍 Top SKUs Críticos por Estação</div>', unsafe_allow_html=True)
    est_selecionada = st.selectbox("Selecione a estação:", options=estacoes_disp)

    df_est_det = df_raw[
        (df_raw['estacao'] == est_selecionada) &
        ((df_raw['score'] >= 500) | (df_raw['pct_saldo'] < 0.10))
    ].sort_values('score', ascending=False).head(20)

    if df_est_det.empty:
        st.info("Nenhum item crítico nesta estação com os filtros selecionados.")
    else:
        df_show = df_est_det[['curva', 'cod_produto', 'desc_produto', 'desc_endereco', 'pct_saldo', 'pulmao', 'score']].copy()
        df_show['pct_saldo'] = df_show['pct_saldo'].apply(fmt_pct)
        df_show['score']     = df_show['score'].apply(lambda x: f"{x:,.0f}")
        df_show.columns      = ['Curva', 'Cód.', 'Descrição', 'Endereço', '% Saldo', 'Pulmão', 'Score']
        st.dataframe(df_show, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — BASE COMPLETA
# ══════════════════════════════════════════════════════════════════════════════
with tab3:

    st.markdown('<div class="section-title">📋 Base Completa do Mezanino (filtros ativos)</div>', unsafe_allow_html=True)

    # Busca textual
    busca = st.text_input("🔍 Buscar por código, descrição ou endereço:", placeholder="Ex: BUCHA, 10314, 5.46.01...")

    df_base = df.copy()
    if busca:
        mask = (
            df_base['desc_produto'].str.contains(busca, case=False, na=False) |
            df_base['cod_produto'].astype(str).str.contains(busca, case=False, na=False) |
            df_base['desc_endereco'].str.contains(busca, case=False, na=False)
        )
        df_base = df_base[mask]

    st.caption(f"{len(df_base):,} registros exibidos")

    colunas = ['prioridade', 'curva', 'estacao', 'cod_produto', 'desc_produto',
               'desc_endereco', 'pct_saldo', 'pulmao', 'score', 'rua']
    df_exp = df_base[colunas].copy()
    df_exp['pct_saldo'] = df_exp['pct_saldo'].apply(fmt_pct)
    df_exp['score']     = df_exp['score'].apply(lambda x: f"{x:,.1f}" if pd.notna(x) else "—")
    df_exp.columns      = ['Prioridade', 'Curva', 'Estação', 'Cód.', 'Descrição',
                            'Endereço', '% Saldo', 'Pulmão', 'Score', 'Rua']

    st.dataframe(df_exp, use_container_width=True, height=500, hide_index=True)

    csv_base = df_base[colunas].to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "⬇️ Exportar base filtrada (.csv)",
        data=csv_base,
        file_name=f"base_ressuprimento_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center; padding:20px 0 6px 0; font-size:0.72rem; color:#aab0c0;">
    Grupo LLE · Coordenação Logística · {datetime.now().strftime('%Y')}
</div>
""", unsafe_allow_html=True)
