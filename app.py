"""
Automatização Slides Semanais — Banco Bari
Interface web para geração dos slides de diretoria.

Deploy: Streamlit Cloud (https://streamlit.io/cloud)
"""

import streamlit as st
import os, io, calendar, tempfile
from collections import defaultdict
from datetime import datetime, date, timedelta
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.path import Path
import matplotlib.patches as mpatches
import openpyxl
from pptx import Presentation
from pptx.util import Inches

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG & CSS
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Slides Semanais — Banco Bari",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&display=swap');

    /* ═══ BARI BRAND TOKENS ═══ */
    :root {
        --bari-blue: #4A90E2;
        --bari-blue-dark: #2563EB;
        --bari-blue-light: #EBF2FC;
        --bari-navy: #0A1628;
        --bari-navy-2: #0D1B2A;
        --bari-navy-3: #142236;
        --bari-gray-50: #F5F7FA;
        --bari-gray-100: #E8ECF2;
        --bari-gray-200: #D0D6E0;
        --bari-gray-400: #8E99A8;
        --bari-gray-600: #556070;
        --bari-gray-800: #2D3142;
        --bari-white: #FFFFFF;
        --bari-orange: #F59E0B;
        --bari-red: #E53E3E;
        --bari-green-ok: #10B981;
    }

    /* ═══ GLOBAL ═══ */
    .stApp {
        font-family: 'Outfit', 'Segoe UI', system-ui, sans-serif !important;
        background: var(--bari-gray-50) !important;
    }
    .block-container { max-width: 940px !important; padding-top: 0 !important; }
    header[data-testid="stHeader"] { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    div[data-testid="stDecoration"] { display: none !important; }

    /* ═══ HEADER ═══ */
    .bari-header {
        background: var(--bari-navy-2);
        padding: 0;
        margin: -1rem -1rem 32px -1rem;
        color: white;
        position: relative;
        overflow: hidden;
    }
    .bari-header-inner {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 22px 36px;
        position: relative;
        z-index: 2;
    }
    .bari-header::after {
        content: '';
        position: absolute;
        bottom: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, var(--bari-blue) 0%, var(--bari-blue-dark) 50%, transparent 100%);
    }
    /* Decorative dots pattern */
    .bari-header::before {
        content: '';
        position: absolute;
        top: 0; right: 0;
        width: 300px; height: 100%;
        background: radial-gradient(circle at 2px 2px, rgba(74,144,226,0.08) 1px, transparent 0);
        background-size: 20px 20px;
    }
    .bari-logo {
        font-size: 32px;
        font-weight: 900;
        letter-spacing: -1.5px;
        color: white;
        line-height: 1;
    }
    .bari-logo span {
        color: var(--bari-blue);
    }
    .bari-header-title {
        font-size: 14px;
        font-weight: 500;
        color: rgba(255,255,255,0.5);
        margin-top: 2px;
        letter-spacing: 0.3px;
    }
    .bari-header-right {
        display: flex;
        align-items: center;
        gap: 8px;
        background: rgba(255,255,255,0.06);
        padding: 8px 16px;
        border-radius: 8px;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .bari-header-right span {
        font-size: 12px;
        color: rgba(255,255,255,0.45);
        font-weight: 500;
    }
    .bari-header-dot {
        width: 7px; height: 7px;
        border-radius: 50%;
        background: var(--bari-blue);
        animation: pulse-dot 2s ease-in-out infinite;
    }
    @keyframes pulse-dot {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.4; }
    }

    /* ═══ SECTION HEADERS ═══ */
    .step-header {
        display: flex; align-items: center; gap: 12px;
        margin-bottom: 16px; margin-top: 8px;
    }
    .step-num {
        width: 30px; height: 30px; border-radius: 8px;
        background: var(--bari-blue);
        color: white;
        display: flex; align-items: center; justify-content: center;
        font-size: 14px; font-weight: 700;
        flex-shrink: 0;
    }
    .step-num-inactive {
        background: var(--bari-gray-200) !important;
        color: var(--bari-gray-400) !important;
    }
    .step-title {
        font-weight: 700; font-size: 16px; color: var(--bari-navy);
    }
    .step-sub {
        font-size: 12px; color: var(--bari-gray-400); margin-left: 4px;
    }

    /* ═══ CARDS BASE ═══ */
    .bari-card {
        background: var(--bari-white);
        border-radius: 12px;
        border: 1px solid var(--bari-gray-100);
        padding: 20px 22px;
        transition: all 0.2s ease;
    }
    .bari-card:hover {
        border-color: var(--bari-gray-200);
        box-shadow: 0 2px 12px rgba(0,0,0,0.04);
    }

    /* ═══ FILE UPLOAD AREA ═══ */
    .stFileUploader > div > div { padding: 6px !important; }
    div[data-testid="stFileUploaderDropzone"] {
        padding: 10px !important;
        border-radius: 10px !important;
        border-color: var(--bari-gray-200) !important;
        background: var(--bari-gray-50) !important;
    }
    div[data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--bari-blue) !important;
        background: var(--bari-blue-light) !important;
    }
    div[data-testid="stFileUploaderDropzone"] > div > span {
        font-size: 12px !important;
        color: var(--bari-gray-400) !important;
    }
    div[data-testid="stFileUploaderDropzone"] button {
        font-size: 12px !important;
        padding: 5px 14px !important;
        border-radius: 8px !important;
        background: var(--bari-blue) !important;
        color: white !important;
        border: none !important;
    }

    /* ═══ BADGES ═══ */
    .badge-ok {
        background: var(--bari-blue-light); color: var(--bari-blue-dark);
        padding: 2px 10px; border-radius: 6px;
        font-size: 10px; font-weight: 700;
        letter-spacing: 0.5px;
    }
    .badge-req {
        background: #FFF3E0; color: var(--bari-orange);
        padding: 2px 10px; border-radius: 6px;
        font-size: 10px; font-weight: 700;
        letter-spacing: 0.5px;
    }

    /* ═══ DATE CARDS ═══ */
    .date-card {
        background: var(--bari-white);
        border-radius: 10px;
        padding: 14px 16px;
        border: 1px solid var(--bari-gray-100);
        border-top: 3px solid;
    }
    .date-card-green  { border-top-color: var(--bari-blue); }
    .date-card-blue   { border-top-color: #3B82F6; }
    .date-card-purple { border-top-color: #8B5CF6; }
    .date-label {
        font-size: 10px; color: var(--bari-gray-400); font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.8px;
        margin-bottom: 6px;
    }
    .date-value { font-size: 20px; font-weight: 800; }
    .date-green  .date-value { color: var(--bari-blue-dark); }
    .date-blue   .date-value { color: #2563EB; }
    .date-purple .date-value { color: #7C3AED; }

    /* ═══ SUMMARY ═══ */
    .summary-box {
        background: var(--bari-white); border-radius: 10px;
        padding: 14px 18px; border: 1px solid var(--bari-gray-100);
        font-size: 13px; color: var(--bari-gray-600);
        line-height: 1.5; margin-bottom: 14px;
        border-left: 3px solid var(--bari-blue);
    }

    /* ═══ LOG PANEL ═══ */
    .log-panel {
        background: var(--bari-navy-2);
        border-radius: 10px;
        padding: 16px 18px;
        max-height: 340px;
        overflow-y: auto;
        font-family: 'JetBrains Mono', 'Fira Code', 'Consolas', monospace;
        font-size: 11.5px;
        line-height: 1.7;
        border: 1px solid rgba(74,144,226,0.15);
    }
    .log-normal  { color: #cbd5e1; }
    .log-success { color: #4ade80; }
    .log-warning { color: #fbbf24; }
    .log-error   { color: #f87171; }
    .log-info    { color: #60a5fa; }

    /* ═══ SUCCESS BANNER ═══ */
    .success-banner {
        background: linear-gradient(135deg, var(--bari-blue-light) 0%, #DBEAFE 100%);
        border-radius: 12px;
        padding: 28px 32px;
        border: 2px solid var(--bari-blue);
        text-align: center;
    }

    /* ═══ NOTE BOX ═══ */
    .note-box {
        background: #FFF8F0;
        border-radius: 10px;
        padding: 14px 18px;
        border: 1px solid #FFDDB5;
        border-left: 3px solid var(--bari-orange);
        font-size: 13px;
        color: #7C4A1E;
        line-height: 1.5;
    }

    /* ═══ BUTTONS ═══ */
    .stButton > button[kind="primary"],
    .stDownloadButton > button {
        background: var(--bari-blue) !important;
        color: white !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        letter-spacing: 0.3px !important;
        padding: 14px 24px !important;
        font-size: 15px !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button:hover {
        background: var(--bari-blue-dark) !important;
        box-shadow: 0 4px 16px rgba(74,144,226,0.3) !important;
    }

    /* ═══ TOGGLE ═══ */
    .stToggle label span { font-size: 13px !important; }

    /* ═══ EXPANDER ═══ */
    .streamlit-expanderHeader {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: var(--bari-gray-600) !important;
    }

    /* ═══ DIVIDER ═══ */
    .soft-divider {
        height: 1px;
        background: linear-gradient(90deg, var(--bari-gray-100) 0%, transparent 100%);
        margin: 28px 0;
    }

    /* ═══ LABEL STYLING ═══ */
    .upload-label {
        font-weight: 700;
        font-size: 13px;
        color: var(--bari-navy);
        margin-bottom: 2px;
    }
    .upload-label-req {
        display: inline-block;
        background: #FFF3E0;
        color: var(--bari-orange);
        font-size: 9px;
        font-weight: 700;
        padding: 1px 6px;
        border-radius: 4px;
        margin-left: 6px;
        vertical-align: middle;
        letter-spacing: 0.5px;
    }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════

FASES_ORDEM = [
    ("Novo",                                   "Data de criação"),
    ("Tentativa de contato",                   "Data Etapa Tentativa de contato"),
    ("Trabalhando/negociação",                 "Data Etapa Trabalhando/negociação"),
    ("Aguardando documentação",                "Data Etapa Aguardando documentação"),
    ("Pré-análise",                            "Data Etapa Pré-análise"),
    ("Análise de crédito",                     "Data Etapa Análise de Crédito"),
    ("Crédito aprovado",                       "Data Etapa Crédito aprovado"),
    ("Análise jurídica / Avaliação do imóvel", "Data Etapa Análise Jurídica"),
    ("Emissão do contrato",                    "Data Etapa Emissão de contrato"),
    ("Assinatura",                             "Data Etapa Assinatura"),
]

FASES_NOMES = [f for f, _ in FASES_ORDEM]

CORES = {
    "Novo":                                   "#4472C4",
    "Tentativa de contato":                   "#1F3864",
    "Trabalhando/negociação":                 "#70D4CE",
    "Aguardando documentação":                "#2E8B6A",
    "Pré-análise":                            "#D4BE6A",
    "Análise de crédito":                     "#E8922A",
    "Crédito aprovado":                       "#C0392B",
    "Análise jurídica / Avaliação do imóvel": "#C97A1A",
    "Emissão do contrato":                    "#27AE60",
    "Assinatura":                             "#1A5276",
}

CANAIS = {
    "B2C":           ["B2C"],
    "GP":            ["GP"],
    "PC":            ["PC"],
    "Relacionamento":["Relacionamento"],
    "Comercial":     ["B2C", "GP", "PC"],
    "Todos":         ["B2C", "GP", "PC", "Relacionamento"],
}

FASES_PC = [f for f in FASES_NOMES if f not in (
    "Novo", "Tentativa de contato", "Trabalhando/negociação", "Aguardando documentação"
)]

SLIDES_FUNIL = [
    (9,  "B2C",       "volume",    "mensal"),
    (10, "B2C",       "volume",    "semanal"),
    (11, "B2C",       "propostas", "mensal"),
    (12, "B2C",       "propostas", "semanal"),
    (14, "GP",        "volume",    "mensal"),
    (15, "GP",        "volume",    "semanal"),
    (16, "GP",        "propostas", "mensal"),
    (17, "GP",        "propostas", "semanal"),
    (19, "PC",        "volume",    "mensal"),
    (20, "PC",        "volume",    "semanal"),
    (21, "PC",        "propostas", "mensal"),
    (22, "PC",        "propostas", "semanal"),
    (24, "Comercial", "volume",    "mensal"),
    (25, "Comercial", "volume",    "semanal"),
    (26, "Comercial", "propostas", "mensal"),
    (27, "Comercial", "propostas", "semanal"),
]

POS_ESQ     = (1.26, 1.21, 3.24, 4.39)
POS_DIR     = (4.70, 1.21, 3.24, 4.39)
POS_LEGENDA = (7.40, 3.40, 2.50, 2.20)

SLIDES_DASH = [
    (13, "B2C"),
    (18, "GP"),
    (23, "PC"),
    (28, "Relacionamento"),
    (29, "Todos"),
]

POS_DASH = (1.20, 0.74, 8.03, 4.85)

EXCLUIR_TIMES = {"GP": ["FRANQ"], "Todos": ["FRANQ"]}

DIST_MTD = {m: {1: 0.20, 2: 0.45, 3: 0.70, 4: 0.90, 5: 1.00} for m in range(1, 13)}

PLAN_STAGE_MAP = {
    "lead": "Lead", "wl": "Workable Lead", "workable lead": "Workable Lead",
    "novo": "Novo consultor", "trabalhando": "Trabalhando",
    "aguardando documentação": "Documentação", "aguardando documentacao": "Documentação",
    "pré-análise": "Pré-análise", "pre-analise": "Pré-análise",
    "análise de crédito": "Análise crédito", "analise de credito": "Análise crédito",
    "crédito aprovado": "Crédito aprovado", "credito aprovado": "Crédito aprovado",
    "jurídico/imóvel": "Jurídica", "jurídico/ imóvel": "Jurídica",
    "juridico/imovel": "Jurídica", "juridico/ imovel": "Jurídica",
    "emissão": "Emissão", "emissao": "Emissão",
    "assinatura": "Assinatura", "efetivado": "Novos contratos",
}

PLAN_TABS = {"B2C": "B2C", "GP": "GP", "PC": "PC", "Rel": "Relacionamento", "Total CGI": "Todos"}

PLAN_MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3,
    "abril": 4, "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

FASES_DASH = {
    "B2C": ["Lead","Workable Lead","Novo consultor","Trabalhando","Documentação","Pré-análise","Análise crédito","Crédito aprovado","Jurídica","Emissão","Assinatura","Novos contratos"],
    "GP": ["Workable Lead","Novo consultor","Trabalhando","Documentação","Pré-análise","Análise crédito","Crédito aprovado","Jurídica","Emissão","Assinatura","Novos contratos"],
    "PC": ["Novo consultor","Pré-análise","Análise crédito","Crédito aprovado","Jurídica","Emissão","Assinatura","Novos contratos"],
    "Relacionamento": ["Novo consultor","Pré-análise","Análise crédito","Crédito aprovado","Jurídica","Emissão","Assinatura","Novos contratos"],
    "Todos": ["Lead","Workable Lead","Novo consultor","Trabalhando","Documentação","Pré-análise","Análise crédito","Crédito aprovado","Jurídica","Emissão","Assinatura","Novos contratos"],
}

CONVERSOES_DASH = {
    "B2C": [("Trabalhando → Efetivado","Trabalhando","Novos contratos"),("Novo → Pré-Análise","Novo consultor","Pré-análise"),("Crédito Aprov. → Efetivado","Crédito aprovado","Novos contratos")],
    "GP": [("Trabalhando → Efetivado","Trabalhando","Novos contratos"),("Novo → Pré-Análise","Novo consultor","Pré-análise"),("Crédito Aprov. → Efetivado","Crédito aprovado","Novos contratos")],
    "PC": [("Pré-análise → Efetivado","Pré-análise","Novos contratos"),("Novo → Pré-Análise","Novo consultor","Pré-análise"),("Crédito Aprov. → Efetivado","Crédito aprovado","Novos contratos")],
    "Relacionamento": [("Pré-análise → Efetivado","Pré-análise","Novos contratos"),("Novo → Pré-Análise","Novo consultor","Pré-análise"),("Crédito Aprov. → Efetivado","Crédito aprovado","Novos contratos")],
    "Todos": [("Novo consultor → Efetivado","Novo consultor","Novos contratos"),("Novo → Pré-Análise","Novo consultor","Pré-análise"),("Crédito Aprov. → Efetivado","Crédito aprovado","Novos contratos")],
}

DASH_FASE_TO_OPP = {
    "Novo consultor": "Novo", "Trabalhando": "Trabalhando/negociação",
    "Documentação": "Aguardando documentação", "Pré-análise": "Pré-análise",
    "Análise crédito": "Análise de crédito", "Crédito aprovado": "Crédito aprovado",
    "Jurídica": "Análise jurídica / Avaliação do imóvel",
    "Emissão": "Emissão do contrato", "Assinatura": "Assinatura",
}

DASH_DOT_CORES = {
    "Lead": "#3b82f6", "Workable Lead": "#2563eb", "Novo consultor": "#7c3aed",
    "Trabalhando": "#0891b2", "Documentação": "#0e7490", "Pré-análise": "#059669",
    "Análise crédito": "#047857", "Crédito aprovado": "#065f46", "Jurídica": "#d97706",
    "Emissão": "#b45309", "Assinatura": "#92400e", "Novos contratos": "#166534",
}

STAGE_COL_DF = {
    'Novo consultor': 'Data de criação', 'Trabalhando': 'Data Etapa Trabalhando/negociação',
    'Documentação': 'Data Etapa Aguardando documentação', 'Pré-análise': 'Data Etapa Pré-análise',
    'Análise crédito': 'Data Etapa Análise de Crédito', 'Crédito aprovado': 'Data Etapa Crédito aprovado',
    'Jurídica': 'Data Etapa Análise Jurídica', 'Emissão': 'Data Etapa Emissão de contrato',
    'Assinatura': 'Data Etapa Assinatura',
}

DATE_COLS_OPPS = [
    'Data de criação', 'Data Etapa Trabalhando/negociação',
    'Data Etapa Aguardando documentação', 'Data Etapa Pré-análise',
    'Data Etapa Análise de Crédito', 'Data Etapa Crédito aprovado',
    'Data Etapa Análise Jurídica', 'Data Etapa Emissão de contrato',
    'Data Etapa Assinatura', 'Data de fechamento',
    'Data da última mudança de fase',
]

# ══════════════════════════════════════════════════════════════
# FUNÇÕES DE PROCESSAMENTO
# ══════════════════════════════════════════════════════════════

def parse_data(val):
    if val is None or val == '':
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ('%d/%m/%Y %H:%M', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def carregar_base(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    hdrs = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i for i, h in enumerate(hdrs) if h}
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        fase = r[col['Fase']]
        time_ = r[col['Time']]
        if not fase or not time_:
            continue
        valor = float(r[col['Valor do Derivado']] or 0)
        data_fech = parse_data(r[col['Data de fechamento']])
        datas = {}
        for fase_nome, col_nome in FASES_ORDEM:
            if col_nome in col:
                datas[fase_nome] = parse_data(r[col[col_nome]])
        rows.append({'time': time_, 'fase': fase, 'valor': valor, 'datas': datas, 'data_fechamento': data_fech})
    wb.close()
    return rows


def carregar_planejamento(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    metas = {}
    for tab_nome, canal in PLAN_TABS.items():
        if tab_nome not in wb.sheetnames:
            continue
        ws = wb[tab_nome]
        mes_cell = None
        for row in ws.iter_rows():
            for cell in row:
                if str(cell.value or "").strip().lower() in ("mês", "mes"):
                    mes_cell = cell
                    break
            if mes_cell:
                break
        if not mes_cell:
            continue
        stage_col = mes_cell.column
        month_row = mes_cell.row
        col_mes_map = {}
        for cell in ws[month_row]:
            if cell.column <= stage_col:
                continue
            val = str(cell.value or "").strip().lower()
            if val in PLAN_MESES:
                col_mes_map[cell.column] = PLAN_MESES[val]
        metas[canal] = {}
        for row in ws.iter_rows(min_row=month_row + 1):
            nome_cell = ws.cell(row=row[0].row, column=stage_col)
            raw = str(nome_cell.value or "").strip().lower()
            fase_nome = PLAN_STAGE_MAP.get(raw)
            if not fase_nome:
                continue
            for col_idx, mes_num in col_mes_map.items():
                cell = ws.cell(row=nome_cell.row, column=col_idx)
                if cell.value is None:
                    continue
                try:
                    val_int = int(float(cell.value))
                except (ValueError, TypeError):
                    continue
                metas[canal].setdefault(mes_num, {})[fase_nome] = val_int
    wb.close()
    return metas


def retrato_funil(rows, canal, data_ref):
    times = CANAIS[canal]
    count = defaultdict(int)
    volume = defaultdict(float)
    for row in rows:
        if row['time'] not in times:
            continue
        if row['fase'] in ('Fechado ganho', 'Fechado perdido'):
            if row['data_fechamento'] and row['data_fechamento'] <= data_ref:
                continue
        fase_na_data = None
        for fase_nome, _ in FASES_ORDEM:
            dt = row['datas'].get(fase_nome)
            if dt is not None and dt <= data_ref:
                fase_na_data = fase_nome
        if fase_na_data:
            count[fase_na_data] += 1
            volume[fase_na_data] += row['valor']
    return count, volume


def fmt_valor(v):
    if v >= 1_000_000: return f"R${v/1_000_000:.0f}m"
    if v >= 1_000: return f"R${v/1_000:.0f}K"
    return f"R${v:,.0f}"


def gerar_funil_png(count, volume, canal, tipo, data_label):
    # Usa todas as fases, mas filtra as que têm valor zero
    fases_base = FASES_NOMES
    if tipo == "volume":
        fases = [f for f in fases_base if volume.get(f, 0) > 0]
        valores = [volume.get(f, 0) for f in fases]
    else:
        fases = [f for f in fases_base if count.get(f, 0) > 0]
        valores = [count.get(f, 0) for f in fases]
    total = sum(valores)
    n = len(fases)
    fig, ax = plt.subplots(figsize=(3.2, 4.2))
    fig.patch.set_facecolor('white')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    ax.text(0.50, 0.99, data_label, ha='center', va='top', fontsize=9, fontweight='bold', color='#222222')
    titulo = f"Soma de Valor do Derivado: R${total/1e6:.0f}m" if tipo == "volume" else f"Contagem de registros: {int(total)}"
    ax.text(0.50, 0.91, titulo, ha='center', va='top', fontsize=7.0, color='#555555')
    if n == 0 or total == 0:
        ax.text(0.5, 0.5, 'Sem dados', ha='center', va='center', fontsize=10, color='#999')
        buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white'); plt.close(); buf.seek(0); return buf.read()
    ft, fb = 0.84, 0.04; fh = ft - fb; top_hw, bot_hw = 0.45, 0.17; cx = 0.50
    props = [v / total for v in valores]; cum = 0.0
    for fase, val, prop in zip(fases, valores, props):
        yt = ft - fh * cum; yb = ft - fh * (cum + prop)
        hw_top = top_hw * (1 - cum) + bot_hw * cum
        hw_bot = top_hw * (1 - (cum+prop)) + bot_hw * (cum+prop)
        verts = [(cx-hw_top,yt),(cx+hw_top,yt),(cx+hw_bot,yb),(cx-hw_bot,yb),(cx-hw_top,yt)]
        codes = [Path.MOVETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
        patch = mpatches.PathPatch(Path(verts, codes), facecolor=CORES[fase], edgecolor='none')
        ax.add_patch(patch)
        pct = prop * 100; ym = (yt + yb) / 2; wm = hw_top + hw_bot; sh = yt - yb
        txt = f"{fmt_valor(val)} ({pct:.2f}%)" if tipo == "volume" else f"{int(val)} ({pct:.2f}%)"
        fs = 7.5 if wm > 0.50 else 6.5 if wm > 0.38 else 5.5
        if val > 0 and sh > 0.03:
            ax.text(cx, ym, txt, ha='center', va='center', fontsize=fs, fontweight='bold', color='black')
        cum += prop
    plt.tight_layout(pad=0.2)
    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white'); plt.close(); buf.seek(0); return buf.read()


def gerar_legenda_png(fases_list=None):
    fases = fases_list if fases_list else FASES_NOMES
    fig, ax = plt.subplots(figsize=(2.4, 2.6)); fig.patch.set_facecolor('white')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    ax.text(0.95, 0.97, "Fase", ha='right', va='top', fontsize=9, color='#444444', fontweight='bold')
    n = len(fases); spacing = min(0.087, 0.85 / max(n, 1))
    for i, fase in enumerate(fases):
        y = 0.89 - i * spacing
        ax.add_patch(plt.Circle((0.90, y), 0.022, color=CORES[fase], transform=ax.transData, clip_on=False))
        nome = fase.replace("Análise jurídica / Avaliação do imóvel", "Análise jurídica / Av. imóvel")
        ax.text(0.83, y, nome, ha='right', va='center', fontsize=8, color='#444444')
    plt.tight_layout(pad=0.2)
    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white'); plt.close(); buf.seek(0); return buf.read()


def carregar_opps_df(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    for col in DATE_COLS_OPPS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
    return df


def carregar_leads_df(file_bytes):
    if file_bytes is None: return None
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    if 'Data de criação' in df.columns:
        df['Data de criação'] = pd.to_datetime(df['Data de criação'], dayfirst=True, errors='coerce')
    return df


def calcular_ref_periodos(df):
    ref_cols = [c for c in ['Data de criação','Data Etapa Trabalhando/negociação','Data da última mudança de fase'] if c in df.columns]
    REF = df[ref_cols].max(numeric_only=False).max()
    ano, mes, dia = REF.year, REF.month, REF.day
    re_start = pd.Timestamp(ano, mes, 1); re_end = REF.normalize()
    sw_start = pd.Timestamp(ano, mes, 1); sw_end = (REF - pd.Timedelta(days=7)).normalize()
    prev_mes = mes - 1 if mes > 1 else 12; prev_ano = ano if mes > 1 else ano - 1
    ultimo_dia_prev = calendar.monthrange(prev_ano, prev_mes)[1]
    mp_start = pd.Timestamp(prev_ano, prev_mes, 1); mp_end = pd.Timestamp(prev_ano, prev_mes, min(dia, ultimo_dia_prev))
    return REF, re_start, re_end, mp_start, mp_end, sw_start, sw_end


def filter_opps_df(df, canal):
    if 'Time' not in df.columns: return df
    if canal == 'Todos': return df
    if canal == 'Relacionamento': return df[df['Time'].isin(['Relacionamento', 'Rel'])]
    if canal in ('B2C', 'PC', 'GP'): return df[df['Time'] == canal]
    return df


def count_stage_df(df, stage, start, end):
    id_col = 'ID da oportunidade' if 'ID da oportunidade' in df.columns else None
    if stage == 'Novos contratos':
        if 'Data de fechamento' not in df.columns: return 0
        mask = (df['Data de fechamento'] >= start) & (df['Data de fechamento'] <= end) & (df['Fase'] == 'Fechado ganho')
        return int(df[mask][id_col].nunique()) if id_col else int(mask.sum())
    col = STAGE_COL_DF.get(stage)
    if not col or col not in df.columns: return 0
    mask = (df[col] >= start) & (df[col] <= end)
    return int(df[mask][id_col].nunique()) if id_col else int(mask.sum())


def count_leads_df(df_leads, canal, start, end):
    if df_leads is None or len(df_leads) == 0: return {'Lead': 0, 'Workable Lead': 0}
    df_c = df_leads
    if canal not in ('Todos',) and 'Canal' in df_leads.columns:
        if canal in ('B2C', 'GP'): df_c = df_leads[df_leads['Canal'] == canal]
        else: return {'Lead': 0, 'Workable Lead': 0}
    if 'Data de criação' not in df_c.columns: return {'Lead': 0, 'Workable Lead': 0}
    mask = (df_c['Data de criação'] >= start) & (df_c['Data de criação'] <= end)
    df_p = df_c[mask]; total = len(df_p)
    wl_col = next((c for c in df_c.columns if 'workable' in c.lower()), None)
    workable = int(df_p[wl_col].apply(lambda v: str(v).strip().upper() in ('1','TRUE','VERDADEIRO')).sum()) if wl_col else 0
    return {'Lead': total, 'Workable Lead': workable}


def calcular_metricas_dash_df(df_opps, df_leads, canal, re_start, re_end, mp_start, mp_end, sw_start, sw_end):
    df_c = filter_opps_df(df_opps, canal); fases = FASES_DASH[canal]
    tem_leads = any(f in fases for f in ('Lead', 'Workable Lead'))
    resultado = {}
    for periodo, start, end in [('realizado',re_start,re_end),('mes_passado',mp_start,mp_end),('semana',sw_start,sw_end)]:
        counts = {}
        if tem_leads:
            lc = count_leads_df(df_leads, canal, start, end)
            if 'Lead' in fases: counts['Lead'] = lc['Lead']
            if 'Workable Lead' in fases: counts['Workable Lead'] = lc['Workable Lead']
        for fase in fases:
            if fase in ('Lead', 'Workable Lead'): continue
            counts[fase] = count_stage_df(df_c, fase, start, end)
        resultado[periodo] = counts
    return resultado


def perc_mtd_ref(ref_date):
    sem = (ref_date.day - 1) // 7 + 1
    return DIST_MTD.get(ref_date.month, {}).get(sem, 1.0)


def badge_semaforo(conv_real, conv_plan=None):
    if conv_plan is None: return '#f3f4f6', '#4b5563'
    if conv_plan and conv_plan > 0:
        ratio = conv_real / conv_plan
        if ratio >= 1.00: return '#d1fae5', '#065f46'
        if ratio >= 0.95: return '#fef3c7', '#92400e'
        return '#fee2e2', '#991b1b'
    if conv_real >= 75: return '#d1fae5', '#065f46'
    if conv_real >= 50: return '#fef3c7', '#92400e'
    return '#fee2e2', '#991b1b'


def cor_numero_vs_meta(valor, meta):
    if not meta or meta == 0: return '#1a1a2e'
    r = valor / meta * 100
    if r >= 90: return '#065f46'
    if r >= 70: return '#92400e'
    return '#991b1b'


def kpi_cor(pct):
    if pct >= 20: return '#065f46'
    if pct >= 10: return '#92400e'
    return '#991b1b'


def fmt_num(v):
    if v is None or v == 0: return '—'
    if v >= 1000: return f"{int(v):,}".replace(',', '.')
    return str(int(v))


def gerar_dashboard_png(canal, metricas_plan, metricas_mes, metricas_sem, metricas_real,
                         ref=None, re_start=None, re_end=None, mp_start=None, mp_end=None, sw_start=None, sw_end=None):
    fases = FASES_DASH[canal]; convs = CONVERSOES_DASH[canal]; n_f = len(fases); n_c = len(convs)
    RH = 0.52; HDR_H = RH*1.30; SUB_H = RH*0.60; SEP_H = RH*0.40; CONV_H = RH*0.90
    fig_w = 10.0; fig_h = HDR_H + SUB_H + n_f*RH + SEP_H + n_c*CONV_H + 0.20
    fig, ax = plt.subplots(figsize=(fig_w, fig_h)); fig.patch.set_facecolor('white')
    ax.set_xlim(0, fig_w); ax.set_ylim(0, fig_h); ax.axis('off')
    cx = [0.0, 2.85, 4.72, 6.58, 8.38, fig_w]
    def mid(i): return (cx[i] + cx[i+1]) / 2
    HDR_TXT = ['#9ca3af','#0369a1','#1d4ed8','#6d28d9','#065f46']
    HDR_BG = ['white','#f0f9ff','#eff6ff','#faf5ff','#f0fdf4']
    CEL_BG = ['white','#f8fcff','#f8fbff','#fbf8ff','#f8fffe']
    HDR_LABELS = ['','PLANEJADO','MÊS PASSADO','SEMANA PASSADA','REALIZADO']
    _ref = ref or pd.Timestamp.now(); _re_start = re_start or pd.Timestamp.now().replace(day=1)
    _re_end = re_end or pd.Timestamp.now(); _mp_start = mp_start or pd.Timestamp.now().replace(day=1)
    _mp_end = mp_end or pd.Timestamp.now(); _sw_start = sw_start or pd.Timestamp.now().replace(day=1)
    _sw_end = sw_end or pd.Timestamp.now(); _pct_mtd = perc_mtd_ref(_ref)
    HDR_DATES = ['', f"{MESES_PT[_ref.month]} {_ref.year} · MTD {_pct_mtd*100:.0f}%",
                 f"{_mp_start.strftime('%d/%m')} → {_mp_end.strftime('%d/%m')}",
                 f"{_sw_start.strftime('%d/%m')} → {_sw_end.strftime('%d/%m')}",
                 f"{_re_start.strftime('%d/%m')} → {_re_end.strftime('%d/%m')}"]
    y = fig_h
    y -= HDR_H
    for i in range(5):
        ax.add_patch(plt.Rectangle((cx[i],y),cx[i+1]-cx[i],HDR_H,facecolor=HDR_BG[i],edgecolor='none',zorder=1))
        ax.text(mid(i),y+HDR_H/2,HDR_LABELS[i],ha='center',va='center',fontsize=12,fontweight='bold',color=HDR_TXT[i],zorder=3)
    y -= SUB_H
    for i in range(5):
        ax.add_patch(plt.Rectangle((cx[i],y),cx[i+1]-cx[i],SUB_H,facecolor=HDR_BG[i],edgecolor='none',zorder=1))
        ax.text(mid(i),y+SUB_H/2,HDR_DATES[i],ha='center',va='center',fontsize=9,color='#9ca3af',zorder=3)
    ax.axhline(y,color='#e8eaf0',linewidth=1.2,zorder=2)
    all_metricas = [metricas_plan, metricas_mes, metricas_sem, metricas_real]
    def conv_plan_fase(ri):
        if ri == 0: return None
        p = metricas_plan.get(fases[ri]); pp = metricas_plan.get(fases[ri-1])
        return p/pp*100 if pp and pp > 0 and p else None
    for ri, fase in enumerate(fases):
        y -= RH; row_bg = '#f9fafb' if ri%2==1 else 'white'
        ax.add_patch(plt.Rectangle((cx[0],y),cx[1]-cx[0],RH,facecolor=row_bg,edgecolor='none',zorder=1))
        for ci in range(1,5):
            bg = CEL_BG[ci] if row_bg=='white' else '#f5f8fb' if ci<=2 else '#f7f4fc' if ci==3 else '#f4fcf7'
            ax.add_patch(plt.Rectangle((cx[ci],y),cx[ci+1]-cx[ci],RH,facecolor=bg,edgecolor='none',zorder=1))
        dot_c = DASH_DOT_CORES.get(fase,'#888')
        ax.add_patch(plt.Circle((0.22,y+RH/2),0.090,color=dot_c,zorder=3))
        is_nc = fase=="Novos contratos"
        ax.text(0.38,y+RH/2,fase,ha='left',va='center',fontsize=11,fontweight='bold' if is_nc else 'normal',color='#1a1a2e',zorder=3)
        pct_plan = conv_plan_fase(ri)
        for ci, metricas in enumerate(all_metricas):
            col_i = ci+1; val = metricas.get(fase); meta = metricas_plan.get(fase); is_plan_col = (col_i==1)
            pct = None
            if ri > 0:
                prev_val = metricas.get(fases[ri-1])
                if prev_val and prev_val > 0 and val: pct = val/prev_val*100
            if not val:
                ax.text(mid(col_i),y+RH/2,'—',ha='center',va='center',fontsize=10,color='#9ca3af',zorder=3); continue
            val_str = fmt_num(val); fw = 'bold'
            num_c = ('#0369a1' if is_nc else '#1a1a2e') if is_plan_col else cor_numero_vs_meta(val, meta)
            if pct is not None:
                col_w = cx[col_i+1]-cx[col_i]; val_x = cx[col_i]+col_w*0.36
                bw, bh = 0.46, RH*0.52; bx = val_x+0.07; by = y+(RH-bh)/2
                ax.text(val_x,y+RH/2,val_str,ha='right',va='center',fontsize=11,fontweight=fw,color=num_c,zorder=3)
                bg_b, fg_b = badge_semaforo(pct, None if is_plan_col else pct_plan)
                ax.add_patch(mpatches.FancyBboxPatch((bx,by),bw,bh,boxstyle="round,pad=0.03",facecolor=bg_b,edgecolor='none',zorder=3))
                ax.text(bx+bw/2,y+RH/2,f"{pct:.0f}%",ha='center',va='center',fontsize=9,fontweight='bold',color=fg_b,zorder=4)
            else:
                ax.text(mid(col_i),y+RH/2,val_str,ha='center',va='center',fontsize=11,fontweight=fw,color=num_c,zorder=3)
        ax.axhline(y,color='#f0f2f5',linewidth=0.5,zorder=2)
    y -= SEP_H; ax.axhline(y+SEP_H*0.5,color='#e8eaf0',linewidth=1.0,zorder=2)
    def calc_kpi(metricas, fase_num, fase_den):
        n = metricas.get(fase_num); d = metricas.get(fase_den)
        return d/n*100 if n and n > 0 and d else None
    for label, fase_num, fase_den in convs:
        y -= CONV_H
        ax.add_patch(plt.Rectangle((0,y),fig_w,CONV_H,facecolor='#f9fafb',edgecolor='none',zorder=1))
        ax.text(0.38,y+CONV_H/2,label,ha='left',va='center',fontsize=9.5,style='italic',color='#9ca3af',zorder=3)
        for ci, metricas in enumerate(all_metricas):
            col_i = ci+1; v = calc_kpi(metricas, fase_num, fase_den)
            if v is None: txt_c = '#9ca3af'; txt = '—'
            else: txt = f"{v:.1f}%"; txt_c = '#4b5563' if col_i==1 else kpi_cor(v)
            ax.text(mid(col_i),y+CONV_H/2,txt,ha='center',va='center',fontsize=9.5,style='italic',fontweight='bold',color=txt_c,zorder=3)
    for i in range(1,5): ax.axvline(cx[i],color='#e8eaf0',linewidth=0.6,zorder=2)
    ax.add_patch(mpatches.FancyBboxPatch((0.03,0.03),fig_w-0.06,fig_h-0.06,boxstyle="round,pad=0.03",facecolor='none',edgecolor='#e8eaf0',linewidth=1.5,zorder=5))
    plt.subplots_adjust(left=0,right=1,top=1,bottom=0)
    buf = io.BytesIO(); plt.savefig(buf,format='png',dpi=300,bbox_inches='tight',facecolor='white'); plt.close(); buf.seek(0); return buf.read()


# ── Slide manipulation ──

def remover_funis_existentes(slide):
    shapes_rem = [s for s in slide.shapes if s.shape_type == 13 and s.width > 1_500_000]
    for s in shapes_rem:
        sp = s._element; sp.getparent().remove(sp)

def add_img(slide, blob, pos):
    left, top, w, h = pos
    slide.shapes.add_picture(io.BytesIO(blob), Inches(left), Inches(top), Inches(w), Inches(h))

def fix_dates(slide, subs):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for old, new in subs.items():
                    if old in run.text:
                        run.text = run.text.replace(old, new)

# ══════════════════════════════════════════════════════════════
# HELPERS DE DATA
# ══════════════════════════════════════════════════════════════

def sexta_mais_recente(ref=None):
    d = ref or date.today()
    dias = (d.weekday() - 4) % 7
    return d - timedelta(days=dias)

def calcular_datas_auto():
    atual = sexta_mais_recente()
    sem_pass = atual - timedelta(days=7)
    mes_ant_mes = atual.month - 1 if atual.month > 1 else 12
    mes_ant_ano = atual.year if atual.month > 1 else atual.year - 1
    ultimo_dia = calendar.monthrange(mes_ant_ano, mes_ant_mes)[1]
    dia_ref = min(atual.day, ultimo_dia)
    mes_ant = date(mes_ant_ano, mes_ant_mes, dia_ref)
    return atual, sem_pass, mes_ant

# ══════════════════════════════════════════════════════════════
# PROCESSAMENTO PRINCIPAL
# ══════════════════════════════════════════════════════════════

def processar_tudo(pptx_bytes, base_funil_bytes, base_dash_bytes, base_leads_bytes,
                   plan_bytes, data_atual, data_sem_pass, data_mes_ant, progress_bar, status_text):
    """Processa tudo e retorna (bytes_pptx, lista_de_logs)."""
    logs = []
    def log(msg):
        logs.append(msg)

    total_steps = 50
    step = 0
    def advance(n=1, msg=None):
        nonlocal step
        step += n
        progress_bar.progress(min(step / total_steps, 1.0))
        if msg:
            status_text.markdown(f"<span style='font-size:13px;color:#64748b'>{msg}</span>", unsafe_allow_html=True)

    log("╔══════════════════════════════════════════════╗")
    log("║   Automatização Slides Semanais — Bari       ║")
    log("╚══════════════════════════════════════════════╝")
    log("")

    STR_ATUAL = data_atual.strftime("%d/%m"); MES_ATUAL = MESES_PT[data_atual.month]
    STR_SEM_PASS = data_sem_pass.strftime("%d/%m"); MES_SEM_PASS = MESES_PT[data_sem_pass.month]
    STR_MES_ANT = data_mes_ant.strftime("%d/%m"); MES_ANT = MESES_PT[data_mes_ant.month]
    log(f"📅 Datas: atual={STR_ATUAL} | sem_pass={STR_SEM_PASS} | mes_ant={STR_MES_ANT}")

    # Planejamento
    METAS_2026 = {}
    if plan_bytes:
        advance(2, "📅 Carregando planejamento...")
        METAS_2026 = carregar_planejamento(plan_bytes)
        total_p = sum(len(f) for cd in METAS_2026.values() for f in cd.values())
        log(f"\n📅 Planejamento: {total_p} valores carregados")

    # Base funil
    advance(3, "📊 Carregando base do funil...")
    rows = carregar_base(base_funil_bytes)
    log(f"\n📊 Base do funil: {len(rows)} oportunidades")

    # Gerar funis
    advance(2, "🎨 Gerando funis...")
    log("\n🎨 Gerando funis...")
    cache = {}
    combos = list(dict.fromkeys((c, t) for _, c, t, _ in SLIDES_FUNIL))
    periodos = [("atual", data_atual), ("sem_pass", data_sem_pass), ("mes_ant", data_mes_ant)]
    # Coletar fases ativas por canal (para legendas dinâmicas)
    fases_ativas = {}
    for canal, tipo in combos:
        for chave, data_ref in periodos:
            count, volume = retrato_funil(rows, canal, data_ref)
            mes = MESES_PT[data_ref.month]
            data_label = f"{mes} ({data_ref.strftime('%d/%m')})"
            cache[(canal, tipo, chave)] = gerar_funil_png(count, volume, canal, tipo, data_label)
            # Coleta fases com valor > 0 para a legenda
            for f in FASES_NOMES:
                if count.get(f, 0) > 0 or volume.get(f, 0) > 0:
                    fases_ativas.setdefault(canal, set()).add(f)
        c_atual, v_atual = retrato_funil(rows, canal, data_atual)
        if tipo == "volume":
            lbl = f"R${sum(v_atual.values())/1e6:.1f}M"
        else:
            lbl = f"{int(sum(c_atual.values()))} propostas"
        log(f"  ✅ {canal:12s} | {tipo:10s} | atual: {lbl}")
        advance(2)

    # Gerar legendas dinâmicas por canal (mantém a ordem original, só exclui zerados)
    legendas = {}
    for canal in set(c for c, _ in combos):
        fases_canal = [f for f in FASES_NOMES if f in fases_ativas.get(canal, set())]
        legendas[canal] = gerar_legenda_png(fases_canal if fases_canal else FASES_NOMES)
    log("  ✅ Legendas")
    advance(1)

    # Apresentação
    advance(2, "📂 Carregando apresentação...")
    prs = Presentation(io.BytesIO(pptx_bytes))
    log(f"\n📂 Apresentação: {len(prs.slides)} slides")

    # Atualizar slides funil
    advance(2, "🔄 Atualizando slides dos funis...")
    log("\n🔄 Atualizando slides dos funis...")
    nomes = {"B2C": "B2C", "GP": "Grandes Parcerias", "PC": "Correspondentes (PC)", "Comercial": "Comercial s/ Rel"}
    for num, canal, tipo, comp in SLIDES_FUNIL:
        idx = num - 1
        if idx >= len(prs.slides):
            log(f"  ⚠️  Slide {num} não existe"); continue
        slide = prs.slides[idx]; nome = nomes.get(canal, canal)
        periodo_esq = "mes_ant" if comp == "mensal" else "sem_pass"
        img_esq = cache[(canal, tipo, periodo_esq)]; img_dir = cache[(canal, tipo, "atual")]
        leg = legendas.get(canal, legendas.get("B2C", gerar_legenda_png()))
        remover_funis_existentes(slide)
        add_img(slide, img_esq, POS_ESQ); add_img(slide, img_dir, POS_DIR); add_img(slide, leg, POS_LEGENDA)
        if comp == "mensal":
            fix_dates(slide, {STR_MES_ANT: STR_MES_ANT, MES_ANT: MES_ANT, STR_SEM_PASS: STR_ATUAL, MES_SEM_PASS: MES_ATUAL})
        else:
            fix_dates(slide, {STR_SEM_PASS: STR_SEM_PASS, STR_ATUAL: STR_ATUAL})
        log(f"  Slide {num:2d} — {nome:20s} | {tipo:10s} | {comp}")
        advance(1)

    # Dashboards
    if base_dash_bytes:
        advance(2, "📋 Gerando dashboards...")
        log("\n📋 Gerando dashboards...")
        df_opps = carregar_opps_df(base_dash_bytes)
        log(f"   {len(df_opps)} oportunidades (dash)")
        df_leads = carregar_leads_df(base_leads_bytes)
        if df_leads is not None:
            log(f"   {len(df_leads)} leads")
        else:
            log("   ⚠️ Sem base de leads — Lead/WL zerados")
        REF, re_start, re_end, mp_start, mp_end, sw_start, sw_end = calcular_ref_periodos(df_opps)
        log(f"   REF={REF.strftime('%d/%m/%Y')} | realizado={re_start.strftime('%d/%m')}→{re_end.strftime('%d/%m')}")
        for num, canal in SLIDES_DASH:
            idx = num - 1
            if idx >= len(prs.slides):
                log(f"  ⚠️  Slide {num} não existe"); continue
            mes = REF.month; pct_mtd = perc_mtd_ref(REF)
            plan_raw = METAS_2026.get(canal, {}).get(mes, {})
            metricas_plan = {f: (round(plan_raw[f]*pct_mtd) if plan_raw.get(f) else None) for f in FASES_DASH[canal]}
            resultado = calcular_metricas_dash_df(df_opps, df_leads, canal, re_start, re_end, mp_start, mp_end, sw_start, sw_end)
            png = gerar_dashboard_png(canal, metricas_plan, resultado['mes_passado'], resultado['semana'], resultado['realizado'],
                                       ref=REF, re_start=re_start, re_end=re_end, mp_start=mp_start, mp_end=mp_end, sw_start=sw_start, sw_end=sw_end)
            slide = prs.slides[idx]; remover_funis_existentes(slide); add_img(slide, png, POS_DASH)
            log(f"  Slide {num:2d} — {canal} ✅")
            advance(2)
    else:
        log("\n⚠️ Sem base do dashboard — pulando dashboards")

    # Salvar
    advance(2, "💾 Salvando apresentação...")
    output = io.BytesIO(); prs.save(output); output.seek(0)
    log("\n✅ Apresentação gerada com sucesso!")
    log("🎉 Pronto! Baixe o arquivo abaixo.")
    progress_bar.progress(1.0)
    status_text.markdown("<span style='font-size:13px;color:#2563EB;font-weight:700'>✅ Concluído!</span>", unsafe_allow_html=True)

    return output.getvalue(), logs


# ══════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════

def render_file_card(title, subtitle, icon, accent, file_obj, required=False):
    """Renders info about a file card (purely visual, uploader is separate)."""
    has_file = file_obj is not None
    badge = ""
    if required:
        badge = f'<span class="badge-{"ok" if has_file else "req"}">{"OK" if has_file else "OBRIGATÓRIO"}</span>'

    icon_class = f"{accent}" if has_file else "empty"
    file_info = ""
    if has_file and file_obj is not None:
        size = f"({file_obj.size / 1024:.0f} KB)" if hasattr(file_obj, 'size') else ""
        name = file_obj.name if hasattr(file_obj, 'name') else "arquivo"
        file_info = f'<span class="upload-filename">{name}</span> <span style="color:#94a3b8;font-size:11px">{size}</span>'
    else:
        file_info = subtitle

    st.markdown(f"""
    <div class="upload-card {'upload-card-'+accent+' has-file file-'+accent if has_file else ''}">
        <div style="display:flex;align-items:center;gap:12px">
            <div class="upload-icon {icon_class}" style="color:white;font-size:16px">{icon if not has_file else '✓'}</div>
            <div style="flex:1">
                <div style="display:flex;align-items:center;gap:8px">
                    <span class="upload-title">{title}</span>
                    {badge}
                </div>
                <div class="upload-sub">{file_info}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def main():
    # ── Header ──
    st.markdown("""
    <div class="bari-header">
        <div class="bari-header-inner">
            <div>
                <div class="bari-logo">bari<span>.</span></div>
                <div class="bari-header-title">Slides Semanais — Diretoria Comercial</div>
            </div>
            <div class="bari-header-right">
                <div class="bari-header-dot"></div>
                <span>Gerador de Apresentações</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Step 1: Bases ──
    st.markdown("""
    <div class="step-header">
        <div class="step-num">1</div>
        <span class="step-title">Carregar bases</span>
        <span class="step-sub">Arraste ou clique para selecionar</span>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📈 Base do Funil** `OBRIGATÓRIO`", help="Atualizar Entrada nas Fases.xlsx")
        st.caption("Exportação do Salesforce — Atualizar Entrada nas Fases.xlsx")
        f_funil = st.file_uploader("Base do Funil", type=["xlsx"], key="f_funil", label_visibility="collapsed")
    with col2:
        st.markdown("**📑 Apresentação Modelo** `OBRIGATÓRIO`", help="Arquivo .pptx da diretoria")
        st.caption("Arquivo .pptx atual da diretoria")
        f_pptx = st.file_uploader("Apresentação", type=["pptx"], key="f_pptx", label_visibility="collapsed")

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**📊 Base Dashboard (Opps)**")
        st.caption("Entrada nas Fases Dash.xlsx — período mais amplo")
        f_dash = st.file_uploader("Base Dash", type=["xlsx"], key="f_dash", label_visibility="collapsed")
    with col4:
        st.markdown("**👥 Base Dashboard (Leads)**")
        st.caption("Entradas nas Fases Leads Dash.xlsx")
        f_leads = st.file_uploader("Base Leads", type=["xlsx"], key="f_leads", label_visibility="collapsed")

    # Planejamento (colapsável)
    with st.expander("🎯 Planejamento — muda pouco, carregue apenas se necessário"):
        st.caption("Planejamento.xlsx — metas mensais por canal")
        f_plan = st.file_uploader("Planejamento", type=["xlsx"], key="f_plan", label_visibility="collapsed")

    st.markdown('<div class="soft-divider"></div>', unsafe_allow_html=True)

    # ── Step 2: Datas ──
    st.markdown("""
    <div class="step-header">
        <div class="step-num">2</div>
        <span class="step-title">Conferir datas</span>
    </div>
    """, unsafe_allow_html=True)

    auto_atual, auto_sem, auto_mes = calcular_datas_auto()
    usar_auto = st.toggle("Calcular automaticamente", value=True, key="auto_datas")

    if usar_auto:
        data_atual, data_sem_pass, data_mes_ant = auto_atual, auto_sem, auto_mes
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""<div class="date-card date-card-green date-green">
                <div class="date-label">DATA ATUAL (SEXTA)</div>
                <div class="date-value">{data_atual.strftime('%d/%m/%Y')}</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class="date-card date-card-blue date-blue">
                <div class="date-label">SEMANA PASSADA</div>
                <div class="date-value">{data_sem_pass.strftime('%d/%m/%Y')}</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class="date-card date-card-purple date-purple">
                <div class="date-label">MÊS ANTERIOR</div>
                <div class="date-value">{data_mes_ant.strftime('%d/%m/%Y')}</div>
            </div>""", unsafe_allow_html=True)
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            data_atual = st.date_input("Data atual (sexta)", value=auto_atual, key="d_atual")
        with c2:
            data_sem_pass = st.date_input("Semana passada", value=auto_sem, key="d_sem")
        with c3:
            data_mes_ant = st.date_input("Mês anterior", value=auto_mes, key="d_mes")

    st.markdown('<div class="soft-divider"></div>', unsafe_allow_html=True)

    # ── Step 3: Gerar ──
    can_generate = f_funil is not None and f_pptx is not None

    step_bg = "var(--bari-blue)" if can_generate else "var(--bari-gray-200)"
    step_color = "white" if can_generate else "var(--bari-gray-400)"
    title_color = "var(--bari-navy)" if can_generate else "var(--bari-gray-400)"

    st.markdown(f"""
    <div class="step-header">
        <div class="step-num" style="background:{step_bg};color:{step_color}">3</div>
        <span class="step-title" style="color:{title_color}">Gerar apresentação</span>
    </div>
    """, unsafe_allow_html=True)

    if can_generate:
        # Summary
        parts = [f"Funis ({f_funil.name})"]
        if f_dash: parts.append(f"Dashboards ({f_dash.name})")
        if f_leads: parts.append("Leads")
        if f_plan: parts.append("Planejamento")
        st.markdown(f"""<div class="summary-box">
            <strong style="color:#0A1628">Resumo:</strong> {' + '.join(parts)} → <strong style="color:#2563EB">{f_pptx.name}</strong>
        </div>""", unsafe_allow_html=True)

    if not can_generate:
        st.markdown("""<div class="note-box">
            ⏳ Carregue pelo menos a <strong>Base do Funil</strong> e a <strong>Apresentação Modelo</strong> para continuar.
        </div>""", unsafe_allow_html=True)
        return

    if st.button("🚀  Gerar Apresentação", type="primary", use_container_width=True, disabled=not can_generate):
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            result_bytes, log_lines = processar_tudo(
                pptx_bytes=f_pptx.read(),
                base_funil_bytes=f_funil.read(),
                base_dash_bytes=f_dash.read() if f_dash else None,
                base_leads_bytes=f_leads.read() if f_leads else None,
                plan_bytes=f_plan.read() if f_plan else None,
                data_atual=data_atual,
                data_sem_pass=data_sem_pass,
                data_mes_ant=data_mes_ant,
                progress_bar=progress_bar,
                status_text=status_text,
            )

            # Success
            nome_saida = f"Apresentacao_{MESES_PT[data_atual.month]}_{data_atual.strftime('%d%m%Y')}.pptx"

            st.markdown("""<div class="success-banner">
                <div style="font-size:40px;margin-bottom:8px">🎉</div>
                <div style="font-size:18px;font-weight:800;color:#2563EB;margin-bottom:4px">Apresentação pronta!</div>
            </div>""", unsafe_allow_html=True)

            st.download_button(
                label=f"📥  Baixar {nome_saida}",
                data=result_bytes,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

            # Log (colapsável)
            with st.expander("📋 Ver log completo"):
                log_html = ""
                for l in log_lines:
                    css = "log-normal"
                    if "✅" in l: css = "log-success"
                    elif "⚠" in l: css = "log-warning"
                    elif "❌" in l: css = "log-error"
                    elif "🎉" in l or "📅" in l or "📊" in l or "📂" in l or "📋" in l or "🔄" in l or "💾" in l: css = "log-info"
                    log_html += f'<div class="{css}">{l}</div>'
                st.markdown(f'<div class="log-panel">{log_html}</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Erro durante o processamento: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

    # Nota quando faltam arquivos opcionais
    if not f_dash and not st.session_state.get('_generated'):
        st.markdown("""<div class="note-box" style="margin-top:14px">
            <strong>Nota:</strong> Sem a base do Dashboard, apenas os funis serão gerados (slides 9-27).
            Os dashboards (slides 13, 18, 23, 28, 29) precisam da base de Oportunidades Dash.
        </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
