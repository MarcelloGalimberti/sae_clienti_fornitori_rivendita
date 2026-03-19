# =============================================================================
# Analisi Mark-Up Rivendita — App Unificata Clienti & Fornitori
# Filtro automatico: Rivendita standard | Rivendita WR
# Requisiti: pip install fpdf2 kaleido streamlit plotly pandas openpyxl xlsxwriter
# =============================================================================

import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import warnings
import tempfile
import os
from datetime import datetime, date

import plotly.express as px
import plotly.graph_objects as go
import plotly.figure_factory as ff
import plotly.io as pio

warnings.filterwarnings('ignore')

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="Analisi Mark-Up Rivendita")

# ══════════════════════════════════════════════════════════════════════════════
# PALETTE COLORI — coerente con logo SAE Scientifica
# ══════════════════════════════════════════════════════════════════════════════
C = {
    'verde':         '#5ab030',   # SAE green (primary)
    'verde_chiaro':  '#a3d47a',   # SAE green light
    'verde_scuro':   '#3a7a1e',   # SAE green dark
    'navy':          '#1a3a5c',   # dark navy (linee, testi forti)
    'arancio':       '#f5a623',   # warning orange
    'rosso':         '#d63031',   # error red
    'grigio':        '#636e72',   # neutral gray
    'grigio_chiaro': '#dfe6e9',   # light gray background
    'giallo_adi':    '#ffc107',   # ADI yellow (accent)
    'bianco':        '#ffffff',
    'sfondo_kpi':    '#f4f9f1',   # very light green background for cards
}

TIPI_RIVENDITA = ['Rivendita standard', 'Rivendita WR']

# Articoli di servizio esclusi dall'analisi marginalità Fornitori
ARTICOLI_ESCLUSI_FORN = {'SP TRASP', 'F00001', 'T00001'}

# Classificazioni articolo da escludere
CLASSI_ESCLUSE = {'Interventi e servizi connessi', 'Servizi', 'Spese e acconti'}

COLONNE = [
    'Codice anagrafica', 'Cliente/Fornitore', 'Codice articolo','Classificazione articolo',
    'Descrizione Articolo', 'Quantità', 'Prezzo finale',
    'Cliente - Fornitore', 'Data consegna', 'MACROPROGETTO',
    'Anno data consegna', 'Mese data consegna', 'Tipo Sottoprogetto', 'Progetto'
]

COLONNE_VIS_ART = [
    'Codice articolo', 'Descrizione Articolo',
    'Quantità_Cliente',
    'Prezzo medio Cliente',    # prezzo unitario di vendita al cliente
    'Costo medio Fornitore',   # prezzo unitario di acquisto dal fornitore
    'Mark-Up',
    'Utile/Perdita'
]

# Colonne visualizzate per la tab Fornitori (logica minimo_quantità)
COLONNE_VIS_ART_FORN = [
    'Codice articolo', 'Descrizione Articolo',
    'Quantità_Cliente',
    'Quantità_Fornitore',
    'minimo_quantità',
    'Prezzo medio Cliente',    # prezzo unitario di vendita
    'Costo medio Fornitore',   # prezzo unitario di acquisto
    'Mark-Up',
    'Utile/Perdita'
]

# ══════════════════════════════════════════════════════════════════════════════
# FORMATTAZIONE NUMERI — formato italiano
# ══════════════════════════════════════════════════════════════════════════════

def _it(val, decimali=0):
    """Numero in formato italiano (punto migliaia, virgola decimali)."""
    if pd.isna(val):
        return "N/D"
    try:
        s = f"{abs(float(val)):,.{decimali}f}"
        s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
        return ('-' if float(val) < 0 else '') + s
    except Exception:
        return str(val)

def fmt_eur(val, decimali=0):
    return f"€ {_it(val, decimali)}" if not pd.isna(val) else "N/D"

def fmt_pct(val, decimali=1):
    return f"{_it(val, decimali)}%" if not pd.isna(val) else "N/D"

def fmt_num(val, decimali=2):
    return _it(val, decimali)


def formatta_df(df):
    """
    Formatta automaticamente le colonne numeriche di un dataframe
    in formato italiano, riconoscendo il tipo dalla colonna.
    Restituisce un dataframe con colonne stringa.
    """
    df_out = df.copy()
    for col in df_out.columns:
        if not pd.api.types.is_numeric_dtype(df_out[col]):
            continue
        col_l = col.lower()
        if '(€)' in col or col_l in ['utile/perdita', 'fatturato', 'costo acquisto',
                                       'prezzo finale_cliente', 'prezzo finale_fornitore integrato',
                                       'costo medio fornitore', 'prezzo medio cliente',
                                       'prezzo medio acquisto', 'fatturato_vendita', 'costo_acquisto']:
            df_out[col] = df_out[col].apply(lambda x: fmt_eur(x, 2))
        elif '(%)' in col or col_l in ['mark-up', 'marginalità', 'marginalita']:
            df_out[col] = df_out[col].apply(lambda x: fmt_pct(x, 1))
        else:
            df_out[col] = df_out[col].apply(lambda x: fmt_num(x, 2))
    return df_out


def tabella_semaforo(df_num, col_marg, soglia_bassa, soglia_alta):
    """
    Mostra un dataframe con:
    - numeri formattati in italiano
    - colonna marginalità colorata con semaforo
    """
    colori_marg = df_num[col_marg].apply(
        lambda x: C.get(semaforo(x, soglia_bassa, soglia_alta), C['grigio'])
    ).tolist()

    df_fmt = formatta_df(df_num)

    def _highlight(s):
        return [f'color:{colori_marg[i]};font-weight:bold' for i in range(len(s))]

    return df_fmt.style.apply(_highlight, subset=[col_marg], axis=0)


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════

def semaforo(val, soglia_bassa, soglia_alta):
    if pd.isna(val):
        return 'grigio'
    if val >= soglia_alta:
        return 'verde'
    if val >= soglia_bassa:
        return 'arancio'
    return 'rosso'


def flatten_cols(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.copy()
        df.columns = [
            '_'.join(str(p) for p in col if p != '' and not (isinstance(p, float) and np.isnan(p)))
            for col in df.columns
        ]
    return df


def kpi_card(label, valore, colore='grigio'):
    hex_col = C.get(colore, C['grigio'])
    st.markdown(
        f"""
        <div style="background:{C['sfondo_kpi']};border-radius:8px;padding:16px 12px;
                    text-align:center;border-left:6px solid {hex_col};margin-bottom:6px">
            <div style="font-size:12px;color:{C['grigio']};margin-bottom:4px;
                        font-family:Arial,sans-serif">{label}</div>
            <div style="font-size:22px;font-weight:bold;color:{hex_col};
                        font-family:Arial,sans-serif">{valore}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ══════════════════════════════════════════════════════════════════════════════
# CARICAMENTO & PRE-PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data
def carica_e_preproces(file):
    try:
        df = pd.read_excel(file, usecols=COLONNE, engine='openpyxl')
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return None
    df['Data consegna'] = pd.to_datetime(df['Data consegna'], errors='coerce')
    df = df.dropna(subset=['Data consegna'])
    df['Codice articolo'] = df['Codice articolo'].astype(str).str.strip()
    df['Anno data consegna'] = df['Data consegna'].dt.year
    df['Mese data consegna'] = df['Data consegna'].dt.month
    df['Anno-Mese'] = df['Data consegna'].dt.to_period('M')
    df = df[df['Tipo Sottoprogetto'].isin(TIPI_RIVENDITA)].copy()
    df = df[~df['Classificazione articolo'].isin(CLASSI_ESCLUSE)].copy()
    return df.reset_index(drop=True)


def filtra_periodo(df, da: date, a: date):
    return df[
        (df['Data consegna'].dt.date >= da) &
        (df['Data consegna'].dt.date <= a)
    ].copy()


# ══════════════════════════════════════════════════════════════════════════════
# CALCOLO MARGINALITÀ
# ══════════════════════════════════════════════════════════════════════════════

def calcola_marginalita_articoli(df_calc, df_storico):
    if df_calc is None or df_calc.empty:
        return pd.DataFrame()
    pivot = df_calc.pivot_table(
        index=['Codice articolo', 'Descrizione Articolo'],
        columns='Cliente - Fornitore',
        values=['Prezzo finale', 'Quantità'],
        aggfunc='sum'
    ).reset_index()
    pivot = flatten_cols(pivot)
    for col in ['Prezzo finale_Cliente', 'Quantità_Cliente',
                'Prezzo finale_Fornitore', 'Quantità_Fornitore']:
        if col not in pivot.columns:
            pivot[col] = np.nan
    pivot['Prezzo medio Cliente'] = np.where(
        pivot['Quantità_Cliente'].fillna(0) > 0,
        pivot['Prezzo finale_Cliente'] / pivot['Quantità_Cliente'], np.nan
    )
    pivot['Costo medio Fornitore'] = np.where(
        pivot['Quantità_Fornitore'].fillna(0) > 0,
        pivot['Prezzo finale_Fornitore'] / pivot['Quantità_Fornitore'], np.nan
    )
    # Fallback costo storico (vettorizzato)
    if df_storico is not None and not df_storico.empty:
        st_piv = df_storico.pivot_table(
            index=['Codice articolo'], columns='Cliente - Fornitore',
            values=['Prezzo finale', 'Quantità'], aggfunc='sum'
        ).reset_index()
        st_piv = flatten_cols(st_piv)
        for col in ['Prezzo finale_Fornitore', 'Quantità_Fornitore']:
            if col not in st_piv.columns:
                st_piv[col] = np.nan
        st_piv['_costo_storico'] = np.where(
            st_piv['Quantità_Fornitore'].fillna(0) > 0,
            st_piv['Prezzo finale_Fornitore'] / st_piv['Quantità_Fornitore'], np.nan
        )
        storico_map = st_piv.set_index('Codice articolo')['_costo_storico'].to_dict()
        mask_nan = pivot['Costo medio Fornitore'].isna()
        pivot.loc[mask_nan, 'Costo medio Fornitore'] = (
            pivot.loc[mask_nan, 'Codice articolo'].map(storico_map)
        )
    pivot['Prezzo finale_Fornitore integrato'] = (
        pivot['Costo medio Fornitore'] * pivot['Quantità_Cliente']
    )
    pivot['Mark-Up'] = np.where(
        (pd.notna(pivot['Costo medio Fornitore'])) & (pivot['Costo medio Fornitore'] > 0),
        (pivot['Prezzo medio Cliente'] / pivot['Costo medio Fornitore'] - 1) * 100,
        np.nan
    )
    pivot['Utile/Perdita'] = (
        pivot['Prezzo finale_Cliente'].fillna(0)
        - pivot['Prezzo finale_Fornitore integrato'].fillna(0)
    )
    return pivot


def marginalita_complessiva(df_art):
    if df_art is None or df_art.empty:
        return np.nan
    tot_cli  = df_art['Prezzo finale_Cliente'].sum()
    tot_forn = df_art['Prezzo finale_Fornitore integrato'].sum()
    return (tot_cli / tot_forn - 1) * 100 if tot_forn > 0 else np.nan


# ── Funzioni specifiche per la logica Fornitori (Analisi_Fornitori 2.py) ──────

def calcola_marginalita_forn(df_calc):
    """
    Calcolo marginalità per tab Fornitori — logica da Analisi_Fornitori 2.py:
    - Esclude articoli di servizio (SP TRASP, F00001, T00001)
    - Mark-Up per articolo: (prezzo_medio_cli / costo_medio_forn - 1) × 100
    - fatturato_vendita / costo_acquisto calcolati su minimo_quantità
      per evitare distorsioni quando qty acquistata ≠ qty venduta
    - Nessun fallback storico: se manca il costo fornitore nel periodo → NaN
    """
    if df_calc is None or df_calc.empty:
        return pd.DataFrame()
    df_calc = df_calc[~df_calc['Codice articolo'].isin(ARTICOLI_ESCLUSI_FORN)].copy()
    if df_calc.empty:
        return pd.DataFrame()
    pivot = df_calc.pivot_table(
        index=['Codice articolo', 'Descrizione Articolo'],
        columns='Cliente - Fornitore',
        values=['Prezzo finale', 'Quantità'],
        aggfunc='sum'
    ).reset_index()
    pivot = flatten_cols(pivot)
    for col in ['Prezzo finale_Cliente', 'Quantità_Cliente',
                'Prezzo finale_Fornitore', 'Quantità_Fornitore']:
        if col not in pivot.columns:
            pivot[col] = np.nan
    pivot['Prezzo medio Cliente'] = np.where(
        pivot['Quantità_Cliente'].fillna(0) > 0,
        pivot['Prezzo finale_Cliente'] / pivot['Quantità_Cliente'], np.nan
    )
    pivot['Costo medio Fornitore'] = np.where(
        pivot['Quantità_Fornitore'].fillna(0) > 0,
        pivot['Prezzo finale_Fornitore'] / pivot['Quantità_Fornitore'], np.nan
    )
    # Mark-Up per articolo
    pivot['Mark-Up'] = np.where(
        (pd.notna(pivot['Costo medio Fornitore'])) & (pivot['Costo medio Fornitore'] > 0) &
        (pd.notna(pivot['Prezzo medio Cliente'])),
        (pivot['Prezzo medio Cliente'] / pivot['Costo medio Fornitore'] - 1) * 100,
        np.nan
    )
    # Quantità minima tra venduto e acquistato (base di confronto).
    # fillna(0) OBBLIGATORIO: pandas .min(axis=1) con skipna=True (default) restituisce
    # il valore non-NaN quando uno dei due è NaN — gonfiando il costo su articoli
    # acquistati ma non venduti nel periodo (Qty_Cliente assente → NaN → ignorato).
    pivot['minimo_quantità'] = (
        pivot[['Quantità_Cliente', 'Quantità_Fornitore']].fillna(0).min(axis=1)
    )
    # Fatturato e costo su quantità minima
    pivot['fatturato_vendita'] = pivot['Prezzo medio Cliente'] * pivot['minimo_quantità']
    pivot['costo_acquisto']    = pivot['Costo medio Fornitore'] * pivot['minimo_quantità']
    # Margine su venduto (utile/perdita su quantità minima)
    pivot['Utile/Perdita'] = (
        pivot['fatturato_vendita'].fillna(0) - pivot['costo_acquisto'].fillna(0)
    )
    # Acquistato ma non ancora venduto
    pivot['eccesso_quantità'] = (
        pivot['Quantità_Fornitore'].fillna(0) - pivot['minimo_quantità']
    ).clip(lower=0)
    pivot['costo_non_venduto'] = pivot['Costo medio Fornitore'] * pivot['eccesso_quantità']
    return pivot


def marginalita_complessiva_forn(df_art):
    """
    Mark-Up aggregato (logica Fornitori 2): fatturato_vendita / costo_acquisto - 1
    Considera solo articoli con entrambi i valori > 0.
    """
    if df_art is None or df_art.empty:
        return np.nan
    valid = df_art[
        (df_art['fatturato_vendita'].fillna(0) > 0) &
        (df_art['costo_acquisto'].fillna(0) > 0)
    ]
    if valid.empty:
        return np.nan
    fat  = valid['fatturato_vendita'].sum()
    cost = valid['costo_acquisto'].sum()
    return (fat / cost - 1) * 100 if cost > 0 else np.nan


def kpi_da_df_art_forn(df_art):
    fat         = df_art['fatturato_vendita'].sum()  if not df_art.empty else 0
    cost        = df_art['costo_acquisto'].sum()     if not df_art.empty else 0
    utile       = fat - cost
    marg        = marginalita_complessiva_forn(df_art)
    non_venduto = df_art['costo_non_venduto'].sum()  if not df_art.empty else 0
    return fat, cost, utile, marg, non_venduto


def calcola_markup_tutti_clienti(df_periodo, df_full):
    """
    Calcola il Mark-Up complessivo per ogni cliente nel periodo.
    Logica identica al drill-down singolo cliente (tab Clienti):
    usa calcola_marginalita_articoli con fallback storico df_full.
    """
    clienti = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Cliente']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])['Prezzo finale']
        .sum().reset_index().sort_values('Prezzo finale', ascending=False)
        .rename(columns={'Cliente/Fornitore': 'Nome'})
    )
    rows = []
    for _, r in clienti.iterrows():
        cod, nome = r['Codice anagrafica'], r['Nome']
        art_cli = df_periodo[
            (df_periodo['Cliente - Fornitore'] == 'Cliente') &
            (df_periodo['Codice anagrafica'] == cod)
        ]['Codice articolo'].unique().tolist()
        mask = (
            ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
             (df_periodo['Codice anagrafica'] == cod)) |
            ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
             (df_periodo['Codice articolo'].isin(art_cli)))
        )
        df_art = calcola_marginalita_articoli(df_periodo[mask], df_full)
        if df_art.empty:
            continue
        fat, cost, utile, marg = kpi_da_df_art(df_art)
        rows.append({
            'Cliente':           nome,
            'Fatturato (€)':     round(fat,   0),
            'Costo (€)':         round(cost,  0),
            'Utile/Perdita (€)': round(utile, 0),
            'Mark-Up (%)':       round(marg,  1) if not np.isnan(marg) else np.nan,
        })
    return (
        pd.DataFrame(rows)
        .sort_values('Mark-Up (%)', ascending=False)
        .reset_index(drop=True)
    )


def calcola_markup_tutti_fornitori(df_periodo):
    """
    Calcola il Mark-Up complessivo per ogni fornitore nel periodo.
    Logica identica al drill-down singolo fornitore (tab Fornitori):
    usa calcola_marginalita_forn (minimo_quantità, esclusi articoli di servizio).
    """
    fornitori = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Fornitore']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])
        .size().reset_index()[['Codice anagrafica', 'Cliente/Fornitore']]
        .rename(columns={'Cliente/Fornitore': 'Nome'})
    )
    rows = []
    for _, r in fornitori.iterrows():
        cod, nome = r['Codice anagrafica'], r['Nome']
        art_forn = [
            a for a in df_periodo[
                (df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                (df_periodo['Codice anagrafica'] == cod)
            ]['Codice articolo'].unique().tolist()
            if a not in ARTICOLI_ESCLUSI_FORN
        ]
        if not art_forn:
            continue
        mask = (
            ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
             (df_periodo['Codice anagrafica'] == cod) &
             (df_periodo['Codice articolo'].isin(art_forn))) |
            ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
             (df_periodo['Codice articolo'].isin(art_forn)))
        )
        df_art = calcola_marginalita_forn(df_periodo[mask])
        if df_art.empty:
            continue
        fat, cost, utile, marg, nv = kpi_da_df_art_forn(df_art)
        rows.append({
            'Fornitore':                nome,
            'Fatturato su venduto (€)': round(fat,   0),
            'Costo su venduto (€)':     round(cost,  0),
            'Utile/Perdita (€)':        round(utile, 0),
            'Mark-Up (%)':              round(marg,  1) if not np.isnan(marg) else np.nan,
            'Non ancora fatturato (€)': round(nv,    0),
        })
    return (
        pd.DataFrame(rows)
        .sort_values('Mark-Up (%)', ascending=False)
        .reset_index(drop=True)
    )


def trend_mensile_forn(df_periodo, mask=None):
    """Trend mensile per tab Fornitori — usa logica minimo_quantità."""
    df = df_periodo[mask].copy() if mask is not None else df_periodo.copy()
    rows = []
    for periodo in sorted(df['Anno-Mese'].unique()):
        df_m = df[df['Anno-Mese'] == periodo]
        art  = calcola_marginalita_forn(df_m)
        marg = marginalita_complessiva_forn(art)
        fat  = art['fatturato_vendita'].sum()   if not art.empty else 0
        cost = art['costo_acquisto'].sum()       if not art.empty else 0
        nv   = art['costo_non_venduto'].sum()    if not art.empty else 0
        rows.append({
            'Periodo':             periodo.strftime('%b %Y'),
            'Periodo_ord':         periodo,
            'Mark-Up (%)':     round(marg, 2) if not np.isnan(marg) else np.nan,
            'Fatturato su venduto': fat,
            'Costo su venduto':    cost,
            'Non ancora fatturato': nv,
        })
    return pd.DataFrame(rows).sort_values('Periodo_ord').reset_index(drop=True)


def trend_mensile(df_periodo, df_storico, mask=None):
    df = df_periodo[mask].copy() if mask is not None else df_periodo.copy()
    rows = []
    for periodo in sorted(df['Anno-Mese'].unique()):
        df_m = df[df['Anno-Mese'] == periodo]
        art  = calcola_marginalita_articoli(df_m, df_storico)
        marg = marginalita_complessiva(art)
        fat  = art['Prezzo finale_Cliente'].sum() if not art.empty else 0
        cost = art['Prezzo finale_Fornitore integrato'].sum() if not art.empty else 0
        rows.append({
            'Periodo':         periodo.strftime('%b %Y'),
            'Periodo_ord':     periodo,
            'Mark-Up (%)': round(marg, 2) if not np.isnan(marg) else np.nan,
            'Fatturato':       fat,
            'Costo Acquisto':  cost,
        })
    return pd.DataFrame(rows).sort_values('Periodo_ord').reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# GRAFICI
# ══════════════════════════════════════════════════════════════════════════════

FONT_BASE = dict(family='Arial, sans-serif', size=14, color=C['navy'])

def fig_trend(df_t, titolo, soglia_bassa, soglia_alta):
    """Grafico barre (fatturato/costo) + linea marginalità con label sui marker."""
    fig = go.Figure()

    # Barre fatturato
    fig.add_trace(go.Bar(
        x=df_t['Periodo'], y=df_t['Fatturato'],
        name='Fatturato Cliente',
        marker_color=C['verde'],
        opacity=0.9,
        hovertemplate='<b>%{x}</b><br>Fatturato: € %{y:,.0f}<extra></extra>'
    ))
    # Barre costo
    fig.add_trace(go.Bar(
        x=df_t['Periodo'], y=df_t['Costo Acquisto'],
        name='Costo Acquisto',
        marker_color=C['grigio_chiaro'],
        marker_line=dict(color=C['grigio'], width=1),
        opacity=0.95,
        hovertemplate='<b>%{x}</b><br>Costo: € %{y:,.0f}<extra></extra>'
    ))

    # Linea marginalità con marker grandi e label
    marg_vals = df_t['Mark-Up (%)'].tolist()
    labels = [f"{v:.1f}%" if not np.isnan(v) else "" for v in marg_vals]

    fig.add_trace(go.Scatter(
        x=df_t['Periodo'],
        y=df_t['Mark-Up (%)'],
        name='Mark-Up (%)',
        mode='lines+markers+text',
        text=labels,
        textposition='top center',
        textfont=dict(size=13, color=C['navy'], family='Arial, sans-serif'),
        line=dict(color=C['navy'], width=3),
        marker=dict(size=14, color=C['navy'],
                    line=dict(color=C['bianco'], width=2.5)),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Mark-Up: %{y:.1f}%<extra></extra>'
    ))

    # Linee soglia
    fig.add_hline(y=soglia_bassa, line_dash='dot', line_color=C['arancio'],
                  line_width=1.5, yref='y2',
                  annotation_text=f'  Soglia bassa {soglia_bassa:.1f}%',
                  annotation_font=dict(color=C['arancio'], size=12),
                  annotation_position='bottom right')
    fig.add_hline(y=soglia_alta, line_dash='dot', line_color=C['verde_scuro'],
                  line_width=1.5, yref='y2',
                  annotation_text=f'  Soglia budget {soglia_alta:.1f}%',
                  annotation_font=dict(color=C['verde_scuro'], size=12),
                  annotation_position='top right')

    fig.update_layout(
        title=dict(text=titolo, font=dict(size=17, color=C['navy'], family='Arial')),
        barmode='group',
        yaxis=dict(title='€', tickformat=',.0f',
                   title_font=dict(size=14), tickfont=dict(size=13)),
        yaxis2=dict(title='Mark-Up (%)', overlaying='y', side='right',
                    tickformat='.1f',
                    title_font=dict(size=14), tickfont=dict(size=13)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    xanchor='right', x=1, font=dict(size=13)),
        height=580,
        hovermode='x unified',
        plot_bgcolor=C['bianco'],
        paper_bgcolor=C['bianco'],
        font=FONT_BASE,
        margin=dict(t=80, b=40),
    )
    fig.update_xaxes(tickfont=dict(size=13), gridcolor=C['grigio_chiaro'])
    fig.update_yaxes(gridcolor=C['grigio_chiaro'])
    return fig


def fig_trend_forn(df_t, titolo, soglia_bassa, soglia_alta):
    """
    Grafico andamento mensile per tab Fornitori.
    2 barre raggruppate per mese:
      1. Fatturato su venduto  (Prezzo_medio_CLI × min_qty)
      2. Costo su venduto      (Costo_medio_FORN × min_qty)
    Linea secondaria: Mark-Up (%).
    Il valore "Acquistato non ancora venduto" è esposto solo nei KPI card.
    """
    fig = go.Figure()

    # Barra 1 — Fatturato su venduto
    fig.add_trace(go.Bar(
        x=df_t['Periodo'], y=df_t['Fatturato su venduto'],
        name='Fatturato su venduto',
        marker_color=C['verde'],
        opacity=0.9,
        hovertemplate='<b>%{x}</b><br>Fatturato su venduto: € %{y:,.0f}<extra></extra>'
    ))
    # Barra 2 — Costo su venduto
    fig.add_trace(go.Bar(
        x=df_t['Periodo'], y=df_t['Costo su venduto'],
        name='Costo su venduto',
        marker_color=C['grigio_chiaro'],
        marker_line=dict(color=C['grigio'], width=1),
        opacity=0.95,
        hovertemplate='<b>%{x}</b><br>Costo su venduto: € %{y:,.0f}<extra></extra>'
    ))

    # Linea marginalità con marker grandi e label
    marg_vals = df_t['Mark-Up (%)'].tolist()
    labels = [f"{v:.1f}%" if not np.isnan(v) else "" for v in marg_vals]

    fig.add_trace(go.Scatter(
        x=df_t['Periodo'],
        y=df_t['Mark-Up (%)'],
        name='Mark-Up (%)',
        mode='lines+markers+text',
        text=labels,
        textposition='top center',
        textfont=dict(size=13, color=C['navy'], family='Arial, sans-serif'),
        line=dict(color=C['navy'], width=3),
        marker=dict(size=14, color=C['navy'],
                    line=dict(color=C['bianco'], width=2.5)),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Mark-Up: %{y:.1f}%<extra></extra>'
    ))

    # Linee soglia
    fig.add_hline(y=soglia_bassa, line_dash='dot', line_color=C['arancio'],
                  line_width=1.5, yref='y2',
                  annotation_text=f'  Soglia bassa {soglia_bassa:.1f}%',
                  annotation_font=dict(color=C['arancio'], size=12),
                  annotation_position='bottom right')
    fig.add_hline(y=soglia_alta, line_dash='dot', line_color=C['verde_scuro'],
                  line_width=1.5, yref='y2',
                  annotation_text=f'  Soglia budget {soglia_alta:.1f}%',
                  annotation_font=dict(color=C['verde_scuro'], size=12),
                  annotation_position='top right')

    fig.update_layout(
        title=dict(text=titolo, font=dict(size=17, color=C['navy'], family='Arial')),
        barmode='group',
        yaxis=dict(title='€', tickformat=',.0f',
                   title_font=dict(size=14), tickfont=dict(size=13)),
        yaxis2=dict(title='Mark-Up (%)', overlaying='y', side='right',
                    tickformat='.1f',
                    title_font=dict(size=14), tickfont=dict(size=13)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    xanchor='right', x=1, font=dict(size=13)),
        height=580,
        hovermode='x unified',
        plot_bgcolor=C['bianco'],
        paper_bgcolor=C['bianco'],
        font=FONT_BASE,
        margin=dict(t=80, b=40),
    )
    fig.update_xaxes(tickfont=dict(size=13), gridcolor=C['grigio_chiaro'])
    fig.update_yaxes(gridcolor=C['grigio_chiaro'])
    return fig


def fig_distribuzione(df_art, marg_comp):
    vals = df_art['Mark-Up'].dropna().values
    if len(vals) < 2:
        return None
    fig = ff.create_distplot([vals], ['Mark-Up'], bin_size=25,
                             colors=[C['verde']], show_hist=False)
    fig.add_vline(x=marg_comp, line_dash='dash', line_color=C['rosso'], line_width=2,
                  annotation_text=f'  Mark-Up complessivo: {marg_comp:.1f}%',
                  annotation_font=dict(color=C['rosso'], size=13))
    fig.update_xaxes(range=[-100, 300], tickfont=dict(size=13))
    fig.update_yaxes(tickfont=dict(size=13))
    fig.update_layout(
        title=dict(text='Distribuzione Mark-Up per articolo',
                   font=dict(size=15, color=C['navy'])),
        height=460, xaxis_title='Mark-Up (%)',
        font=FONT_BASE, plot_bgcolor=C['bianco'],
        paper_bgcolor=C['bianco']
    )
    return fig


def fig_scatter_marg(df_art):
    df_p = df_art.dropna(subset=['Mark-Up', 'Prezzo finale_Cliente']).copy()
    if df_p.empty:
        return None
    df_p['Qty_scaled'] = np.sqrt(df_p['Quantità_Cliente'].fillna(1))
    fig = px.scatter(
        df_p, x='Mark-Up', y='Prezzo finale_Cliente',
        size='Qty_scaled', size_max=45,
        color_discrete_sequence=[C['verde']],
        hover_data=['Codice articolo', 'Descrizione Articolo', 'Quantità_Cliente'],
        title='Mark-Up vs Fatturato (asse Y logaritmico)',
        labels={'Mark-Up': 'Mark-Up (%)',
                'Prezzo finale_Cliente': 'Fatturato [€, log]'},
        log_y=True, height=460
    )
    fig.update_xaxes(range=[-100, 300], tickfont=dict(size=13))
    fig.update_yaxes(tickfont=dict(size=13))
    fig.update_layout(font=FONT_BASE, plot_bgcolor=C['bianco'],
                      paper_bgcolor=C['bianco'],
                      title_font=dict(size=15, color=C['navy']))
    return fig


def fig_pareto_clienti(df_periodo):
    """
    Pareto del fatturato (Prezzo finale Cliente) per cliente, periodo corrente.
    Barre ordinate per valore decrescente + linea cumulata % su asse secondario.
    """
    df = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Cliente']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])['Prezzo finale']
        .sum().reset_index()
        .rename(columns={'Prezzo finale': 'Fatturato', 'Cliente/Fornitore': 'Nome'})
        .sort_values('Fatturato', ascending=False)
        .reset_index(drop=True)
    )
    if df.empty:
        return None
    totale = df['Fatturato'].sum()
    df['Cumulato %'] = df['Fatturato'].cumsum() / totale * 100

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df['Nome'], y=df['Fatturato'],
        name='Fatturato',
        marker_color=C['verde'], opacity=0.85,
        hovertemplate='<b>%{x}</b><br>Fatturato: € %{y:,.0f}<extra></extra>'
    ))
    fig.add_trace(go.Scatter(
        x=df['Nome'], y=df['Cumulato %'],
        name='Cumulato %',
        mode='lines+markers',
        line=dict(color=C['navy'], width=2.5),
        marker=dict(size=9, color=C['navy'], line=dict(color=C['bianco'], width=2)),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Cumulato: %{y:.1f}%<extra></extra>'
    ))
    fig.add_hline(y=80, line_dash='dot', line_color=C['arancio'], line_width=1.5,
                  yref='y2',
                  annotation_text='  80 %',
                  annotation_font=dict(color=C['arancio'], size=12),
                  annotation_position='top left')
    fig.update_layout(
        title=dict(text='Pareto Fatturato per Cliente',
                   font=dict(size=16, color=C['navy'], family='Arial')),
        yaxis=dict(title='Fatturato (€)', tickformat=',.0f',
                   title_font=dict(size=13), tickfont=dict(size=12)),
        yaxis2=dict(title='Cumulato (%)', overlaying='y', side='right',
                    range=[0, 105], tickformat='.0f',
                    title_font=dict(size=13), tickfont=dict(size=12)),
        xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    xanchor='right', x=1, font=dict(size=12)),
        height=860, barmode='group',
        plot_bgcolor=C['bianco'], paper_bgcolor=C['bianco'],
        font=FONT_BASE, margin=dict(t=70, b=110, l=60, r=60),
    )
    fig.update_xaxes(gridcolor=C['grigio_chiaro'])
    fig.update_yaxes(gridcolor=C['grigio_chiaro'])
    return fig


def _pareto_data_fornitori(df_periodo):
    """
    Calcola per ogni fornitore (esclusi articoli di servizio):
      - fatturato_su_venduto  = sum(Prezzo_medio_CLI × min_qty)
      - acquistato_complessivo = sum(Costo_medio_FORN × Qty_Fornitore)
                               = costo_acquisto + costo_non_venduto
    Restituisce DataFrame pronto per i due grafici Pareto.
    """
    forn_ids = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Fornitore']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])
        .size().reset_index()[['Codice anagrafica', 'Cliente/Fornitore']]
        .rename(columns={'Cliente/Fornitore': 'Nome'})
    )
    rows = []
    for _, r in forn_ids.iterrows():
        cod, nome = r['Codice anagrafica'], r['Nome']
        art_f = [
            a for a in df_periodo[
                (df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                (df_periodo['Codice anagrafica'] == cod)
            ]['Codice articolo'].unique()
            if a not in ARTICOLI_ESCLUSI_FORN
        ]
        if not art_f:
            continue
        mask = (
            ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
             (df_periodo['Codice anagrafica'] == cod) &
             (df_periodo['Codice articolo'].isin(art_f))) |
            ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
             (df_periodo['Codice articolo'].isin(art_f)))
        )
        art = calcola_marginalita_forn(df_periodo[mask])
        if art.empty:
            continue
        fat = art['fatturato_vendita'].fillna(0).sum()
        acq = (art['costo_acquisto'].fillna(0) + art['costo_non_venduto'].fillna(0)).sum()
        rows.append({'Nome': nome, 'Fatturato su venduto': fat, 'Acquistato complessivo': acq})
    return pd.DataFrame(rows)


def fig_pareto_forn(df_pareto, col_val, titolo, bar_color):
    """
    Pareto generico per tab Fornitori.
    col_val: colonna da usare come valore (es. 'Fatturato su venduto').
    """
    df = (
        df_pareto[['Nome', col_val]].copy()
        .sort_values(col_val, ascending=False)
        .reset_index(drop=True)
    )
    if df.empty or df[col_val].sum() == 0:
        return None
    df['Cumulato %'] = df[col_val].cumsum() / df[col_val].sum() * 100

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df['Nome'], y=df[col_val],
        name=col_val,
        marker_color=bar_color, opacity=0.85,
        hovertemplate='<b>%{x}</b><br>' + col_val + ': € %{y:,.0f}<extra></extra>'
    ))
    fig.add_trace(go.Scatter(
        x=df['Nome'], y=df['Cumulato %'],
        name='Cumulato %',
        mode='lines+markers',
        line=dict(color=C['navy'], width=2.5),
        marker=dict(size=9, color=C['navy'], line=dict(color=C['bianco'], width=2)),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Cumulato: %{y:.1f}%<extra></extra>'
    ))
    fig.add_hline(y=80, line_dash='dot', line_color=C['arancio'], line_width=1.5,
                  yref='y2',
                  annotation_text='  80 %',
                  annotation_font=dict(color=C['arancio'], size=12),
                  annotation_position='top left')
    fig.update_layout(
        title=dict(text=titolo, font=dict(size=16, color=C['navy'], family='Arial')),
        yaxis=dict(title='€', tickformat=',.0f',
                   title_font=dict(size=13), tickfont=dict(size=12)),
        yaxis2=dict(title='Cumulato (%)', overlaying='y', side='right',
                    range=[0, 105], tickformat='.0f',
                    title_font=dict(size=13), tickfont=dict(size=12)),
        xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    xanchor='right', x=1, font=dict(size=12)),
        height=860, barmode='group',
        plot_bgcolor=C['bianco'], paper_bgcolor=C['bianco'],
        font=FONT_BASE, margin=dict(t=70, b=110, l=60, r=60),
    )
    fig.update_xaxes(gridcolor=C['grigio_chiaro'])
    fig.update_yaxes(gridcolor=C['grigio_chiaro'])
    return fig


def fig_treemap_fornitori(df_dati, cod_forn, soglia_bassa, soglia_alta):
    """
    Treemap Cliente > Macroprogetto per il fornitore selezionato.
    Colore = marginalità %, dimensione = fatturato.
    Usa logica Fornitori 2: minimo_quantità, esclusi SP TRASP/F00001/T00001.
    """
    df_forn_art = [
        a for a in df_dati[
            (df_dati['Cliente - Fornitore'] == 'Fornitore') &
            (df_dati['Codice anagrafica'] == cod_forn)
        ]['Codice articolo'].unique().tolist()
        if a not in ARTICOLI_ESCLUSI_FORN
    ]

    df_cli = df_dati[
        (df_dati['Cliente - Fornitore'] == 'Cliente') &
        (df_dati['Codice articolo'].isin(df_forn_art))
    ]

    if df_cli.empty:
        return None, pd.DataFrame()

    rows = []
    for (cod_c, nome_c, macro), grp in df_cli.groupby(
            ['Codice anagrafica', 'Cliente/Fornitore', 'MACROPROGETTO']):
        art_grp = grp['Codice articolo'].unique().tolist()
        mask = (
            ((df_dati['Cliente - Fornitore'] == 'Cliente') &
             (df_dati['Codice anagrafica'] == cod_c) &
             (df_dati['MACROPROGETTO'] == macro) &
             (df_dati['Codice articolo'].isin(df_forn_art))) |
            ((df_dati['Cliente - Fornitore'] == 'Fornitore') &
             (df_dati['Codice anagrafica'] == cod_forn) &
             (df_dati['Codice articolo'].isin(art_grp)))
        )
        art_res = calcola_marginalita_forn(df_dati[mask])
        if art_res.empty:
            continue
        mg = marginalita_complessiva_forn(art_res)
        ft = art_res['fatturato_vendita'].sum()
        co = art_res['costo_acquisto'].sum()
        rows.append({
            'Cliente':          nome_c,
            'Macroprogetto':    macro,
            'Fatturato (€)':    round(ft, 0),
            'Costo (€)':        round(co, 0),
            'Utile/Perdita (€)':round(ft - co, 0),
            'Mark-Up (%)':  round(mg, 1) if not np.isnan(mg) else np.nan,
        })

    if not rows:
        return None, pd.DataFrame()

    df_tree = pd.DataFrame(rows)
    df_tree_clean = df_tree.dropna(subset=['Mark-Up (%)']).copy()

    if df_tree_clean.empty:
        return None, df_tree

    # Scala colori ancorata alle soglie configurate dall'utente:
    #   rosso  = marginalità < soglia_bassa
    #   arancio = tra soglia_bassa e soglia_alta
    #   verde  = sopra soglia_alta
    # Range fisso [-50%, +60%] → i colori sono assoluti, non relativi ai dati.
    _RMIN, _RMAX = -50.0, 60.0
    _span = _RMAX - _RMIN
    _p_bassa = max(0.01, min(0.97, (soglia_bassa - _RMIN) / _span))
    _p_alta  = max(_p_bassa + 0.02, min(0.99, (soglia_alta - _RMIN) / _span))

    _color_scale = [
        [0.0,       C['rosso']],
        [_p_bassa,  C['arancio']],
        [_p_alta,   C['verde']],
        [1.0,       C['verde_scuro']],
    ]

    fig = px.treemap(
        df_tree_clean,
        path=['Cliente', 'Macroprogetto'],
        values='Fatturato (€)',
        color='Mark-Up (%)',
        color_continuous_scale=_color_scale,
        range_color=[_RMIN, _RMAX],
        title='Panoramica Cliente > Macroprogetto  |  colore = Mark-Up %  |  dimensione = Fatturato',
        hover_data={'Fatturato (€)': True, 'Costo (€)': True,
                    'Utile/Perdita (€)': True, 'Mark-Up (%)': True},
        height=600,
    )
    fig.update_traces(
        texttemplate='<b>%{label}</b><br>%{customdata[3]:.1f}%',
        textfont=dict(size=14),
        hovertemplate=(
            '<b>%{label}</b><br>'
            'Fatturato: € %{customdata[0]:,.0f}<br>'
            'Costo: € %{customdata[1]:,.0f}<br>'
            'Utile/Perdita: € %{customdata[2]:,.0f}<br>'
            'Mark-Up: %{customdata[3]:.1f}%<extra></extra>'
        )
    )
    fig.update_layout(font=FONT_BASE,
                      title_font=dict(size=15, color=C['navy']),
                      paper_bgcolor=C['bianco'])
    return fig, df_tree


def fig_bar_clienti_marginalita(df_tree, soglia_bassa, soglia_alta):
    """Bar chart orizzontale clienti ordinati per marginalità, colorati semaforo."""
    df_cli = (
        df_tree.dropna(subset=['Mark-Up (%)'])
        .groupby('Cliente')
        .apply(lambda g: pd.Series({
            'Mark-Up (%)': (
                (g['Fatturato (€)'].sum() / g['Costo (€)'].sum() - 1) * 100
                if g['Costo (€)'].sum() > 0 else np.nan
            ),
            'Fatturato (€)': g['Fatturato (€)'].sum()
        }))
        .reset_index()
        .sort_values('Mark-Up (%)', ascending=True)
    )
    colors = [C.get(semaforo(v, soglia_bassa, soglia_alta), C['grigio'])
              for v in df_cli['Mark-Up (%)']]
    fig = go.Figure(go.Bar(
        y=df_cli['Cliente'],
        x=df_cli['Mark-Up (%)'],
        orientation='h',
        marker_color=colors,
        text=[f"{v:.1f}%" for v in df_cli['Mark-Up (%)']],
        textposition='outside',
        textfont=dict(size=13),
        hovertemplate='<b>%{y}</b><br>Mark-Up: %{x:.1f}%<extra></extra>'
    ))
    fig.add_vline(x=soglia_bassa, line_dash='dot', line_color=C['arancio'],
                  line_width=1.5,
                  annotation_text=f'  {soglia_bassa:.1f}%',
                  annotation_font=dict(color=C['arancio']))
    fig.add_vline(x=soglia_alta, line_dash='dot', line_color=C['verde_scuro'],
                  line_width=1.5,
                  annotation_text=f'  {soglia_alta:.1f}%',
                  annotation_font=dict(color=C['verde_scuro']))
    fig.update_layout(
        title=dict(text='Mark-Up per Cliente (dal fornitore selezionato)',
                   font=dict(size=15, color=C['navy'])),
        xaxis_title='Mark-Up (%)',
        height=max(350, len(df_cli) * 42),
        font=FONT_BASE,
        plot_bgcolor=C['bianco'], paper_bgcolor=C['bianco'],
        margin=dict(l=20, r=80),
    )
    fig.update_xaxes(gridcolor=C['grigio_chiaro'])
    return fig, df_cli


def fig_bar_markup_riepilogo(df_mup, col_nome, col_mup, titolo, soglia_bassa, soglia_alta):
    """
    Bar chart orizzontale ordinato per Mark-Up (%), colorato con semaforo.
    Riutilizzabile sia per il riepilogo Clienti sia per il riepilogo Fornitori.
    """
    df_plot = df_mup.dropna(subset=[col_mup]).sort_values(col_mup, ascending=True).copy()
    if df_plot.empty:
        return None
    colors = [
        C.get(semaforo(v, soglia_bassa, soglia_alta), C['grigio'])
        for v in df_plot[col_mup]
    ]
    fig = go.Figure(go.Bar(
        y=df_plot[col_nome],
        x=df_plot[col_mup],
        orientation='h',
        marker_color=colors,
        text=[f"{v:.1f}%" for v in df_plot[col_mup]],
        textposition='outside',
        textfont=dict(size=12),
        hovertemplate='<b>%{y}</b><br>Mark-Up: %{x:.1f}%<extra></extra>'
    ))
    fig.add_vline(x=soglia_bassa, line_dash='dot', line_color=C['arancio'],
                  line_width=1.5,
                  annotation_text=f'  {soglia_bassa:.1f}%',
                  annotation_font=dict(color=C['arancio'], size=11))
    fig.add_vline(x=soglia_alta, line_dash='dot', line_color=C['verde_scuro'],
                  line_width=1.5,
                  annotation_text=f'  {soglia_alta:.1f}%',
                  annotation_font=dict(color=C['verde_scuro'], size=11))
    fig.update_layout(
        title=dict(text=titolo, font=dict(size=15, color=C['navy'])),
        xaxis_title='Mark-Up (%)',
        height=max(360, len(df_plot) * 38),
        font=FONT_BASE,
        plot_bgcolor=C['bianco'],
        paper_bgcolor=C['bianco'],
        margin=dict(l=20, r=90, t=60, b=40),
    )
    fig.update_xaxes(gridcolor=C['grigio_chiaro'])
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def to_excel(sheets: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt_hdr = wb.add_format({
            'bold': True, 'bg_color': '#1a3a5c',
            'font_color': 'white', 'border': 1, 'font_size': 11
        })
        fmt_cell = wb.add_format({'border': 1, 'font_size': 10})
        fmt_num  = wb.add_format({
            'border': 1, 'num_format': '#.##0,00', 'font_size': 10
        })
        for sheet_name, df in sheets.items():
            if df is None or (hasattr(df, 'empty') and df.empty):
                continue
            name = sheet_name[:31]
            df_out = df.copy().reset_index(drop=True)
            df_out.to_excel(writer, sheet_name=name, index=False,
                            startrow=1, header=False)
            ws = writer.sheets[name]
            for col_idx, col_name in enumerate(df_out.columns):
                ws.write(0, col_idx, col_name, fmt_hdr)
                col_width = max(len(str(col_name)) + 4, 16)
                ws.set_column(col_idx, col_idx, col_width)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT PDF
# ══════════════════════════════════════════════════════════════════════════════

# Mappa esplicita per caratteri fuori Latin-1 comuni nei testi finanziari
_UNICODE_MAP = str.maketrans({
    '\u2014': ' - ',   # em dash —
    '\u2013': ' - ',   # en dash –
    '\u2192': ' -> ',  # freccia →
    '\u2190': ' <- ',  # freccia ←
    '\u2019': "'",     # right single quotation mark '
    '\u2018': "'",     # left single quotation mark '
    '\u201c': '"',     # left double quotation mark "
    '\u201d': '"',     # right double quotation mark "
    '\u2022': '-',     # bullet •
    '\u2026': '...',   # ellipsis …
    '\u00b7': '.',     # middle dot ·
    '\u00d7': 'x',     # multiplication sign ×
    '\u20ac': 'EUR',   # euro sign €  ← causa dell'errore
    '\u00a3': 'GBP',   # pound sign £
    '\u00a5': 'JPY',   # yen sign ¥
    '\u00a9': '(c)',   # copyright ©
    '\u00ae': '(R)',   # registered ®
    '\u2122': 'TM',    # trademark ™
})

def _safe(text: str) -> str:
    """
    Rende una stringa compatibile con Helvetica (Latin-1) di fpdf2.
    1. Applica le mappings esplicite (€, —, →, …)
    2. Usa NFKD per i caratteri non-Latin-1 rimanenti (es. ü → u, ñ → n)
    3. Scarta definitivamente tutto ciò che non è encodabile Latin-1
    """
    import unicodedata
    # Step 1: sostituzioni esplicite
    text = str(text).translate(_UNICODE_MAP)
    # Step 2-3: normalizzazione NFKD + filtro Latin-1
    result = []
    for char in text:
        try:
            char.encode('latin-1')
            result.append(char)
        except UnicodeEncodeError:
            # prova a ottenere l'equivalente ASCII tramite decomposizione
            ascii_eq = unicodedata.normalize('NFKD', char).encode('ascii', 'ignore').decode('ascii')
            result.append(ascii_eq)   # stringa vuota se nessun equivalente → char ignorato
    return ''.join(result)


def genera_pdf(titolo, periodo_str, kpi_dict: dict,
               df_riepilogo, df_articoli,
               fig_trend_plotly=None,
               logo_sx='logo.png', logo_dx='logo_adi.png') -> bytes:
    try:
        from fpdf import FPDF
    except ImportError:
        st.error("Installa fpdf2: pip install fpdf2")
        return b""

    W, H   = 297, 210
    MARGIN = 12

    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    # ── Header ────────────────────────────────────────────────────────────────
    header_h = 22
    if os.path.exists(logo_sx):
        pdf.image(logo_sx, x=MARGIN, y=6, h=14)
    if os.path.exists(logo_dx):
        pdf.image(logo_dx, x=W - MARGIN - 22, y=4, h=18)
    pdf.set_xy(MARGIN + 35, 7)
    pdf.set_font('Helvetica', 'B', 15)
    pdf.set_text_color(26, 58, 92)
    pdf.cell(W - 2 * MARGIN - 60, 8, _safe(titolo), align='C', ln=False)
    pdf.set_xy(MARGIN + 35, 16)
    pdf.set_font('Helvetica', '', 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(W - 2 * MARGIN - 60, 6,
             _safe(f"Periodo: {periodo_str}   |   Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}"),
             align='C', ln=True)

    # Linea divisoria
    pdf.set_draw_color(90, 176, 48)
    pdf.set_line_width(0.8)
    pdf.line(MARGIN, header_h + 3, W - MARGIN, header_h + 3)
    pdf.ln(4)

    # ── KPI cards ─────────────────────────────────────────────────────────────
    pdf.set_font('Helvetica', 'B', 11)
    pdf.set_text_color(26, 58, 92)
    pdf.cell(0, 7, 'KPI Principali', ln=True)

    n     = len(kpi_dict)
    cw    = (W - 2 * MARGIN - (n - 1) * 3) / n
    ch    = 18
    sx    = MARGIN
    sy    = pdf.get_y()

    for label, (valore, colore) in kpi_dict.items():
        rgb = {'verde': (90, 176, 48), 'arancio': (245, 166, 35),
               'rosso': (214, 48, 49)}.get(colore, (99, 110, 114))
        pdf.set_fill_color(244, 249, 241)
        pdf.set_draw_color(*rgb)
        pdf.set_line_width(0.4)
        pdf.rect(sx, sy, cw, ch, 'FD')
        pdf.set_fill_color(*rgb)
        pdf.rect(sx, sy, 2, ch, 'F')
        pdf.set_font('Helvetica', '', 8)
        pdf.set_text_color(90, 90, 90)
        pdf.set_xy(sx + 4, sy + 2.5)
        pdf.cell(cw - 5, 5, _safe(label))
        pdf.set_font('Helvetica', 'B', 11)
        pdf.set_text_color(*rgb)
        pdf.set_xy(sx + 4, sy + 9)
        pdf.cell(cw - 5, 6, _safe(str(valore)))
        sx += cw + 3

    pdf.set_text_color(0, 0, 0)
    pdf.set_y(sy + ch + 4)

    # ── Grafico trend ─────────────────────────────────────────────────────────
    if fig_trend_plotly is not None:
        try:
            img_bytes = pio.to_image(fig_trend_plotly, format='png',
                                     width=1200, height=420, scale=1.5)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                tmp.write(img_bytes)
                tmp_path = tmp.name
            avail = H - pdf.get_y() - 16
            img_h = min(avail, 64)
            pdf.image(tmp_path, x=MARGIN, w=W - 2 * MARGIN, h=img_h)
            os.unlink(tmp_path)
            pdf.ln(3)
        except Exception:
            pdf.set_font('Helvetica', 'I', 9)
            pdf.set_text_color(160, 160, 160)
            pdf.cell(0, 6, '[Grafico non disponibile - installa kaleido]', ln=True)
            pdf.set_text_color(0, 0, 0)

    # ── Tabella riepilogativa ─────────────────────────────────────────────────
    if df_riepilogo is not None and not df_riepilogo.empty:
        if pdf.get_y() > H - 55:
            pdf.add_page()
        pdf.ln(2)
        pdf.set_font('Helvetica', 'B', 11)
        pdf.set_text_color(26, 58, 92)
        pdf.cell(0, 7, 'Riepilogo', ln=True)
        pdf.set_text_color(0, 0, 0)
        _pdf_table(pdf, df_riepilogo, W, MARGIN)

    # ── Tabella articoli ──────────────────────────────────────────────────────
    if df_articoli is not None and not df_articoli.empty:
        if pdf.get_y() > H - 55:
            pdf.add_page()
        pdf.ln(2)
        pdf.set_font('Helvetica', 'B', 11)
        pdf.set_text_color(26, 58, 92)
        pdf.cell(0, 7, 'Dettaglio articoli (top 20)', ln=True)
        pdf.set_text_color(0, 0, 0)
        _pdf_table(pdf, df_articoli.head(20), W, MARGIN)

    # ── Footer ────────────────────────────────────────────────────────────────
    pdf.set_y(-11)
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(160, 160, 160)
    pdf.cell(0, 5, 'SAE Scientifica | ADI Business Consulting - Analisi Mark-Up Rivendita',
             align='C')

    return bytes(pdf.output())


def _pdf_table(pdf, df, W, MARGIN):
    cols   = df.columns.tolist()
    col_w  = (W - 2 * MARGIN) / len(cols)
    # Header
    pdf.set_font('Helvetica', 'B', 8)
    pdf.set_fill_color(26, 58, 92)
    pdf.set_text_color(255, 255, 255)
    pdf.set_line_width(0.2)
    for col in cols:
        pdf.cell(col_w, 6, _safe(str(col))[:25], border=1, fill=True, align='C')
    pdf.ln()
    # Righe
    pdf.set_font('Helvetica', '', 7.5)
    pdf.set_text_color(0, 0, 0)
    for i, (_, row) in enumerate(df.iterrows()):
        pdf.set_fill_color(240, 248, 235) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
        for col in cols:
            val = row[col]
            if isinstance(val, float) and not pd.isna(val):
                txt = _it(val, 1)
            else:
                txt = _safe(str(val))[:25]
            pdf.cell(col_w, 5.5, txt, border=1, fill=True,
                     align='R' if isinstance(val, (int, float)) else 'L')
        pdf.ln()
        if pdf.get_y() > 196:
            break


# ══════════════════════════════════════════════════════════════════════════════
# BLOCCHI UI RIUTILIZZABILI
# ══════════════════════════════════════════════════════════════════════════════

def mostra_kpi_row(fat, cost, utile, marg, soglia_bassa, soglia_alta, prefisso=''):
    col1, col2, col3, col4 = st.columns(4)
    with col1:  kpi_card(f'{prefisso}Fatturato', fmt_eur(fat))
    with col2:  kpi_card(f'{prefisso}Costo acquisto', fmt_eur(cost))
    with col3:  kpi_card(f'{prefisso}Utile / Perdita', fmt_eur(utile),
                         semaforo(marg, soglia_bassa, soglia_alta))
    with col4:  kpi_card(f'{prefisso}Mark-Up complessivo', fmt_pct(marg),
                         semaforo(marg, soglia_bassa, soglia_alta))


def kpi_da_df_art(df_art):
    fat   = df_art['Prezzo finale_Cliente'].sum() if not df_art.empty else 0
    cost  = df_art['Prezzo finale_Fornitore integrato'].sum() if not df_art.empty else 0
    utile = fat - cost
    marg  = marginalita_complessiva(df_art)
    return fat, cost, utile, marg


def mostra_analisi_articoli(df_art, marg_comp, soglia_bassa, soglia_alta,
                             key_prefix='', label='', col_list=None):
    if col_list is None:
        col_list = COLONNE_VIS_ART
    cols_vis = [c for c in col_list if c in df_art.columns]

    df_plot = df_art.dropna(subset=['Mark-Up']).copy()
    if len(df_plot) >= 2:
        col_d, col_s = st.columns(2)
        with col_d:
            fd = fig_distribuzione(df_plot, marg_comp)
            if fd:
                st.plotly_chart(fd, use_container_width=True)
        with col_s:
            fs = fig_scatter_marg(df_plot)
            if fs:
                st.plotly_chart(fs, use_container_width=True)


    # ── TOP 10 / BOTTOM 10 per impatto economico ──────────────────────────────
    st.subheader(
        f'TOP 10 · BOTTOM 10 per impatto economico{(" — " + label) if label else ""}',
        divider='green'
    )
    st.caption(
        'Impatto = (Prezzo medio vendita − Costo medio acquisto) × Quantità venduta  '
        '→  ordinati per contributo assoluto in € (campo Utile/Perdita)'
    )
    df_peso = df_art.dropna(subset=['Mark-Up', 'Utile/Perdita']).copy()
    if df_peso.empty:
        st.info('Dati insufficienti per il calcolo TOP/BOTTOM 10.')
    else:
        n_top = min(10, len(df_peso))
        col_t10, col_b10 = st.columns(2)
        with col_t10:
            st.markdown('**🏆 TOP 10 — contributo positivo maggiore**')
            df_top10 = df_peso.nlargest(n_top, 'Utile/Perdita')[cols_vis].copy()
            df_top10['Mark-Up'] = df_top10['Mark-Up'].round(1)
            st.dataframe(
                tabella_semaforo(df_top10, 'Mark-Up', soglia_bassa, soglia_alta),
                hide_index=True, use_container_width=True
            )
        with col_b10:
            st.markdown('**⚠️ BOTTOM 10 — impatto negativo maggiore**')
            df_bot10 = df_peso.nsmallest(n_top, 'Utile/Perdita')[cols_vis].copy()
            df_bot10['Mark-Up'] = df_bot10['Mark-Up'].round(1)
            st.dataframe(
                tabella_semaforo(df_bot10, 'Mark-Up', soglia_bassa, soglia_alta),
                hide_index=True, use_container_width=True
            )

    num = st.slider('N articoli top / bottom Mark-Up:', 1, 20, 5, key=f'{key_prefix}_slider')
    col_t, col_b = st.columns(2)
    with col_t:
        st.markdown(f'**Top {num} per Mark-Up**')
        df_top = df_art.dropna(subset=['Mark-Up']).nlargest(num, 'Mark-Up')[cols_vis].copy()
        df_top['Mark-Up'] = df_top['Mark-Up'].round(1)
        st.dataframe(tabella_semaforo(df_top, 'Mark-Up', soglia_bassa, soglia_alta),
                     hide_index=True, use_container_width=True)
    with col_b:
        st.markdown(f'**Bottom {num} per Mark-Up**')
        df_bot = df_art.dropna(subset=['Mark-Up']).nsmallest(num, 'Mark-Up')[cols_vis].copy()
        df_bot['Mark-Up'] = df_bot['Mark-Up'].round(1)
        st.dataframe(tabella_semaforo(df_bot, 'Mark-Up', soglia_bassa, soglia_alta),
                     hide_index=True, use_container_width=True)

    st.subheader(f'Utile / Perdita per articolo{(" — " + label) if label else ""}',
                 divider='green')
    df_up = df_art[cols_vis].copy().sort_values('Utile/Perdita', ascending=False).reset_index(drop=True)
    df_up['Mark-Up'] = df_up['Mark-Up'].round(1)
    st.dataframe(tabella_semaforo(df_up, 'Mark-Up', soglia_bassa, soglia_alta),
                 use_container_width=True, hide_index=True)
    return df_up


# ══════════════════════════════════════════════════════════════════════════════
# HEADER DELL'APP
# ══════════════════════════════════════════════════════════════════════════════

col_logo_sx, col_titolo, col_logo_dx = st.columns([1.2, 4, 1.2])
with col_logo_sx:
    if os.path.exists('logo.png'):
        st.image('logo.png', width=160)
with col_titolo:
    st.markdown(
        f"<h2 style='text-align:center;color:{C['navy']};font-family:Arial,sans-serif;"
        f"margin-top:10px'>Analisi Mark-Up Rivendita</h2>",
        unsafe_allow_html=True
    )
with col_logo_dx:
    if os.path.exists('logo_adi.png'):
        st.image('logo_adi.png', width=100)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — SOGLIE
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header('Configurazione')
    st.subheader('Soglie Mark-Up')
    soglia_bassa = st.number_input('Soglia bassa (%)', min_value=0.0,
                                   max_value=100.0, value=10.0, step=0.1, format='%.1f')
    soglia_alta  = st.number_input('Soglia budget (%)', min_value=0.0,
                                   max_value=100.0, value=23.3, step=0.1, format='%.1f')
    st.markdown(
        f"<small><span style='color:{C['verde']}'>Verde</span> ≥ {soglia_alta}%  "
        f"&nbsp;|&nbsp;  <span style='color:{C['arancio']}'>Arancio</span> ≥ {soglia_bassa}%  "
        f"&nbsp;|&nbsp;  <span style='color:{C['rosso']}'>Rosso</span> &lt; {soglia_bassa}% (Mark-Up)</small>",
        unsafe_allow_html=True
    )
    st.divider()
    st.caption('v4.0 — SAE Scientifica')

# ══════════════════════════════════════════════════════════════════════════════
# CARICAMENTO FILE
# ══════════════════════════════════════════════════════════════════════════════

st.subheader('Caricamento dati', divider='green')
uploaded = st.file_uploader(
    'Carica il file: Analisi marginalità fornitori 2022-Today(aaaammgg).xlsx',
    type=['xlsx']
)
if not uploaded:
    st.info('Carica il file Excel per iniziare.')
    st.stop()

df_full = carica_e_preproces(uploaded)
if df_full is None:
    st.stop()

st.success(f'File caricato — {len(df_full):,} righe (filtrate su Rivendita)')

# ══════════════════════════════════════════════════════════════════════════════
# SELEZIONE PERIODO
# ══════════════════════════════════════════════════════════════════════════════

min_date = df_full['Data consegna'].min().date()
max_date = df_full['Data consegna'].max().date()

col_da, col_a, _ = st.columns([1, 1, 3])
with col_da:
    data_da = st.date_input('📅 Da:', value=min_date,
                             min_value=min_date, max_value=max_date,
                             key='date_da')
with col_a:
    data_a  = st.date_input('📅 A:', value=max_date,
                             min_value=min_date, max_value=max_date,
                             key='date_a')

if data_da > data_a:
    st.error('Il periodo "da" deve essere precedente o uguale al periodo "a".')
    st.stop()

df_periodo = filtra_periodo(df_full, data_da, data_a)
periodo_label = (
    f"{data_da.strftime('%d/%m/%Y')} — {data_a.strftime('%d/%m/%Y')}"
)
st.caption(f'Righe nel periodo selezionato: **{len(df_periodo):,}**')

if df_periodo.empty:
    st.warning('Nessun dato nel periodo selezionato.')
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# KPI AGGREGATI E TREND — prima delle tab (validi per entrambe)
# ══════════════════════════════════════════════════════════════════════════════

st.markdown(f'### KPI aggregati — {periodo_label}')
st.caption('Tutti i clienti · tutti i fornitori · Rivendita standard + Rivendita WR')

df_art_tutti = calcola_marginalita_articoli(df_periodo, df_full)
fat_all, cost_all, utile_all, marg_all = kpi_da_df_art(df_art_tutti)
mostra_kpi_row(fat_all, cost_all, utile_all, marg_all, soglia_bassa, soglia_alta)

with st.expander('Andamento mensile aggregato', expanded=True):
    with st.spinner('Calcolo trend...'):
        df_trend_all = trend_mensile(df_periodo, df_full)
    st.plotly_chart(
        fig_trend(df_trend_all,
                  f'Mark-Up mensile — tutti i clienti/fornitori | {periodo_label}',
                  soglia_bassa, soglia_alta),
        use_container_width=True
    )


with st.expander('TOP 10 · BOTTOM 10 articoli — periodo aggregato', expanded=False):
    st.caption(
        'Impatto = (Prezzo medio vendita − Costo medio acquisto) × Quantità venduta  '
        '→  tutti i clienti, tutti i fornitori nel periodo selezionato'
    )
    df_peso_all = df_art_tutti.dropna(subset=['Mark-Up', 'Utile/Perdita']).copy()
    if df_peso_all.empty:
        st.info('Dati insufficienti per il calcolo TOP/BOTTOM 10.')
    else:
        cols_vis_all = [c for c in COLONNE_VIS_ART if c in df_peso_all.columns]
        n_top_all = min(10, len(df_peso_all))
        col_ta, col_ba = st.columns(2)
        with col_ta:
            st.markdown('**🏆 TOP 10 — contributo positivo maggiore**')
            df_top_all = df_peso_all.nlargest(n_top_all, 'Utile/Perdita')[cols_vis_all].copy()
            df_top_all['Mark-Up'] = df_top_all['Mark-Up'].round(1)
            st.dataframe(
                tabella_semaforo(df_top_all, 'Mark-Up', soglia_bassa, soglia_alta),
                hide_index=True, use_container_width=True
            )
        with col_ba:
            st.markdown('**⚠️ BOTTOM 10 — impatto negativo maggiore**')
            df_bot_all = df_peso_all.nsmallest(n_top_all, 'Utile/Perdita')[cols_vis_all].copy()
            df_bot_all['Mark-Up'] = df_bot_all['Mark-Up'].round(1)
            st.dataframe(
                tabella_semaforo(df_bot_all, 'Mark-Up', soglia_bassa, soglia_alta),
                hide_index=True, use_container_width=True
            )

st.markdown('---')

# ══════════════════════════════════════════════════════════════════════════════
# TAB PRINCIPALI
# ══════════════════════════════════════════════════════════════════════════════

tab_cli, tab_forn = st.tabs(['  Clienti  ', '  Fornitori  '])


# ════════════════════════════════════════════════════════════════════════════════
# TAB CLIENTI
# ════════════════════════════════════════════════════════════════════════════════

with tab_cli:
    st.header('Clienti — Rivendita', divider='green')

    # ── Pareto fatturato per cliente ──────────────────────────────────────────
    st.subheader('Panoramica Fatturato per Cliente', divider='green')
    with st.spinner('Calcolo Pareto clienti...'):
        fig_par_cli = fig_pareto_clienti(df_periodo)
    if fig_par_cli is not None:
        st.plotly_chart(fig_par_cli, use_container_width=True)
    else:
        st.info('Nessun dato cliente disponibile per il periodo selezionato.')


    # ── Mark-Up riepilogativo per Cliente ─────────────────────────────────────
    st.subheader('Mark-Up per Cliente — riepilogo periodo', divider='green')
    st.caption(
        'Mark-Up complessivo per ogni cliente nel periodo selezionato. '
        'Stessa logica del drill-down: include il costo fornitore degli articoli acquistati per quel cliente.'
    )
    with st.spinner('Calcolo Mark-Up per cliente...'):
        df_mup_cli = calcola_markup_tutti_clienti(df_periodo, df_full)
    if df_mup_cli.empty:
        st.info('Nessun dato cliente disponibile per il periodo selezionato.')
    else:
        col_tbl_cli, col_bar_cli = st.columns([1, 1])
        with col_tbl_cli:
            st.dataframe(
                tabella_semaforo(df_mup_cli, 'Mark-Up (%)', soglia_bassa, soglia_alta),
                use_container_width=True, hide_index=True
            )
        with col_bar_cli:
            fig_mup_cli = fig_bar_markup_riepilogo(
                df_mup_cli, 'Cliente', 'Mark-Up (%)',
                'Mark-Up per Cliente', soglia_bassa, soglia_alta
            )
            if fig_mup_cli:
                st.plotly_chart(fig_mup_cli, use_container_width=True)

    st.markdown('---')

    # ── Selezione cliente ─────────────────────────────────────────────────────

    st.subheader('Drill-down per Cliente', divider='green')
    df_cli_list = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Cliente']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])['Prezzo finale']
        .sum().reset_index().sort_values('Prezzo finale', ascending=False)
    )
    df_cli_list['Label'] = (
        df_cli_list['Codice anagrafica'].astype(str) + '  —  '
        + df_cli_list['Cliente/Fornitore'].astype(str)
    )

    cliente_sel = st.selectbox(
        'Seleziona Cliente:',
        ['— Seleziona Cliente —'] + df_cli_list['Label'].tolist(),
        key='cli_sel'
    )

    if cliente_sel == '— Seleziona Cliente —':
        st.info('Seleziona un cliente per visualizzare i KPI e il drill-down.')
    else:
        cod_cli  = df_cli_list[df_cli_list['Label'] == cliente_sel]['Codice anagrafica'].values[0]
        nome_cli = df_cli_list[df_cli_list['Label'] == cliente_sel]['Cliente/Fornitore'].values[0]

        df_p_cli = df_periodo[
            (df_periodo['Cliente - Fornitore'] == 'Cliente') &
            (df_periodo['Codice anagrafica'] == cod_cli)
        ]
        art_cli = df_p_cli['Codice articolo'].unique().tolist()
        mask_cli = (
            ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
             (df_periodo['Codice anagrafica'] == cod_cli)) |
            ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
             (df_periodo['Codice articolo'].isin(art_cli)))
        )

        df_art_cli = calcola_marginalita_articoli(df_periodo[mask_cli], df_full)
        fat_cli, cost_cli, utile_cli, marg_cli = kpi_da_df_art(df_art_cli)

        # KPI cliente
        st.markdown(f'#### KPI — {nome_cli}')
        mostra_kpi_row(fat_cli, cost_cli, utile_cli, marg_cli,
                       soglia_bassa, soglia_alta)

        # Trend mensile cliente
        with st.expander(f'Andamento mensile — {nome_cli}', expanded=True):
            with st.spinner('Calcolo trend...'):
                df_trend_cli = trend_mensile(df_periodo, df_full, mask=mask_cli)
            fig_tc = fig_trend(df_trend_cli,
                               f'Mark-Up mensile — {nome_cli}',
                               soglia_bassa, soglia_alta)
            st.plotly_chart(fig_tc, use_container_width=True)

        st.markdown('---')

        # ── Tabella Macroprogetti ─────────────────────────────────────────────
        st.subheader('Mark-Up per Macroprogetto', divider='green')
        macroprog_list = sorted(df_p_cli['MACROPROGETTO'].dropna().unique().tolist())
        rows_macro = []
        for macro in macroprog_list:
            art_m = df_p_cli[df_p_cli['MACROPROGETTO'] == macro]['Codice articolo'].unique()
            mask_m = (
                ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
                 (df_periodo['Codice anagrafica'] == cod_cli) &
                 (df_periodo['MACROPROGETTO'] == macro)) |
                ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                 (df_periodo['MACROPROGETTO'] == macro) &
                 (df_periodo['Codice articolo'].isin(art_m)))
            )
            ar = calcola_marginalita_articoli(df_periodo[mask_m], df_full)
            if ar.empty:
                continue
            mg = marginalita_complessiva(ar)
            ft = ar['Prezzo finale_Cliente'].sum()
            co = ar['Prezzo finale_Fornitore integrato'].sum()
            rows_macro.append({
                'Macroprogetto':    macro,
                'Fatturato (€)':    round(ft, 0),
                'Costo (€)':        round(co, 0),
                'Utile/Perdita (€)':round(ft - co, 0),
                'Mark-Up (%)':  round(mg, 1) if not np.isnan(mg) else np.nan,
            })

        df_macro_tab = pd.DataFrame(rows_macro)
        if not df_macro_tab.empty:
            st.dataframe(
                tabella_semaforo(df_macro_tab, 'Mark-Up (%)', soglia_bassa, soglia_alta),
                use_container_width=True, hide_index=True
            )

        macro_sel = st.multiselect(
            'Filtra Macroprogetto:', macroprog_list,
            default=macroprog_list, key='cli_macro'
        )

        if macro_sel:
            # ── Tabella Progetti ──────────────────────────────────────────────
            st.subheader('Mark-Up per Progetto', divider='green')
            proj_list = sorted(
                df_p_cli[df_p_cli['MACROPROGETTO'].isin(macro_sel)]
                ['Progetto'].dropna().unique().tolist()
            )
            rows_proj = []
            for proj in proj_list:
                art_p = df_p_cli[
                    (df_p_cli['MACROPROGETTO'].isin(macro_sel)) &
                    (df_p_cli['Progetto'] == proj)
                ]['Codice articolo'].unique()
                mask_p = (
                    ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
                     (df_periodo['Codice anagrafica'] == cod_cli) &
                     (df_periodo['MACROPROGETTO'].isin(macro_sel)) &
                     (df_periodo['Progetto'] == proj)) |
                    ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                     (df_periodo['MACROPROGETTO'].isin(macro_sel)) &
                     (df_periodo['Progetto'] == proj) &
                     (df_periodo['Codice articolo'].isin(art_p)))
                )
                ar_p = calcola_marginalita_articoli(df_periodo[mask_p], df_full)
                if ar_p.empty:
                    continue
                mg_p = marginalita_complessiva(ar_p)
                ft_p = ar_p['Prezzo finale_Cliente'].sum()
                co_p = ar_p['Prezzo finale_Fornitore integrato'].sum()
                rows_proj.append({
                    'Progetto':          proj,
                    'Fatturato (€)':     round(ft_p, 0),
                    'Costo (€)':         round(co_p, 0),
                    'Utile/Perdita (€)': round(ft_p - co_p, 0),
                    'Mark-Up (%)':   round(mg_p, 1) if not np.isnan(mg_p) else np.nan,
                })

            df_proj_tab = pd.DataFrame(rows_proj)
            if not df_proj_tab.empty:
                st.dataframe(
                    tabella_semaforo(df_proj_tab, 'Mark-Up (%)', soglia_bassa, soglia_alta),
                    use_container_width=True, hide_index=True
                )

            proj_sel = st.multiselect(
                'Filtra Progetto:', proj_list,
                default=proj_list, key='cli_proj'
            )

            if proj_sel:
                # ── Analisi articoli ──────────────────────────────────────────
                st.subheader('Analisi articoli', divider='green')
                art_sel = df_p_cli[
                    (df_p_cli['MACROPROGETTO'].isin(macro_sel)) &
                    (df_p_cli['Progetto'].isin(proj_sel))
                ]['Codice articolo'].unique().tolist()

                mask_a = (
                    ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
                     (df_periodo['Codice anagrafica'] == cod_cli) &
                     (df_periodo['MACROPROGETTO'].isin(macro_sel)) &
                     (df_periodo['Progetto'].isin(proj_sel))) |
                    ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                     (df_periodo['MACROPROGETTO'].isin(macro_sel)) &
                     (df_periodo['Progetto'].isin(proj_sel)) &
                     (df_periodo['Codice articolo'].isin(art_sel)))
                )
                df_art_sel = calcola_marginalita_articoli(df_periodo[mask_a], df_full)
                fat_s, cost_s, utile_s, marg_s = kpi_da_df_art(df_art_sel)

                st.markdown(
                    f"**Selezione corrente** — Fatturato: {fmt_eur(fat_s)}  |  "
                    f"Costo: {fmt_eur(cost_s)}  |  Mark-Up: **{fmt_pct(marg_s)}**"
                )
                cols_vis = [c for c in COLONNE_VIS_ART if c in df_art_sel.columns]
                df_up = mostra_analisi_articoli(
                    df_art_sel, marg_s, soglia_bassa, soglia_alta,
                    key_prefix='cli', label=nome_cli
                )

                # ── Export ────────────────────────────────────────────────────
                st.subheader('Export', divider='green')
                col_x1, col_x2 = st.columns(2)

                kpi_df = pd.DataFrame([{
                    'Cliente':           nome_cli,
                    'Periodo':           periodo_label,
                    'Fatturato (€)':     round(fat_cli, 2),
                    'Costo (€)':         round(cost_cli, 2),
                    'Utile/Perdita (€)': round(utile_cli, 2),
                    'Mark-Up (%)':   round(marg_cli, 1) if not np.isnan(marg_cli) else None,
                }])
                with col_x1:
                    xl = to_excel({
                        'KPI':           kpi_df,
                        'Macroprogetti': df_macro_tab,
                        'Progetti':      df_proj_tab,
                        'Articoli':      df_art_sel[cols_vis].round(2),
                        'Trend mensile': df_trend_cli[['Periodo', 'Mark-Up (%)',
                                                        'Fatturato', 'Costo Acquisto']],
                    })
                    st.download_button(
                        'Download Excel',
                        data=xl,
                        file_name=f'MarkUp_{nome_cli.replace(" ","_")}'
                                  f'_{data_da}_{data_a}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='dl_xl_cli'
                    )
                with col_x2:
                    if st.button('Genera PDF', key='pdf_cli_btn'):
                        with st.spinner('Generazione PDF...'):
                            kpi_pdf = {
                                'Fatturato':     (fmt_eur(fat_cli),  'grigio'),
                                'Costo acquisto':(fmt_eur(cost_cli), 'grigio'),
                                'Utile/Perdita': (fmt_eur(utile_cli),
                                                  semaforo(marg_cli, soglia_bassa, soglia_alta)),
                                'Mark-Up':   (fmt_pct(marg_cli),
                                                  semaforo(marg_cli, soglia_bassa, soglia_alta)),
                            }
                            pdf_b = genera_pdf(
                                titolo=f'Analisi Mark-Up — {nome_cli}',
                                periodo_str=periodo_label,
                                kpi_dict=kpi_pdf,
                                df_riepilogo=df_macro_tab,
                                df_articoli=df_art_sel[cols_vis]
                                    .sort_values('Mark-Up', ascending=False)
                                    .head(20).round(2),
                                fig_trend_plotly=fig_tc,
                            )
                        st.download_button(
                            'Scarica PDF',
                            data=pdf_b,
                            file_name=f'MarkUp_{nome_cli.replace(" ","_")}'
                                      f'_{data_da}_{data_a}.pdf',
                            mime='application/pdf',
                            key='dl_pdf_cli'
                        )


# ════════════════════════════════════════════════════════════════════════════════
# TAB FORNITORI
# ════════════════════════════════════════════════════════════════════════════════

with tab_forn:
    st.header('Fornitori — Rivendita', divider='green')

    # ── Pareto fornitori ──────────────────────────────────────────────────────
    st.subheader('Panoramica Fornitori', divider='green')
    st.caption(
        '**Fatturato su venduto**: ricavo sugli articoli acquistati e venduti nel periodo '
        '(Prezzo_medio_CLI × min_qty).  '
        '**Acquistato complessivo**: costo totale acquistato dal fornitore, '
        'incluso quanto non ancora fatturato ai clienti '
        '(Costo_medio_FORN × Qty_Fornitore).'
    )
    with st.spinner('Calcolo Pareto fornitori...'):
        df_par_forn = _pareto_data_fornitori(df_periodo)

    if df_par_forn.empty:
        st.info('Nessun dato fornitore disponibile per il periodo selezionato.')
    else:
        fig_pf1 = fig_pareto_forn(
            df_par_forn,
            col_val='Fatturato su venduto',
            titolo='Pareto Fatturato su Venduto per Fornitore',
            bar_color=C['verde']
        )
        if fig_pf1:
            st.plotly_chart(fig_pf1, use_container_width=True)

        fig_pf2 = fig_pareto_forn(
            df_par_forn,
            col_val='Acquistato complessivo',
            titolo='Pareto Valore Acquistato Complessivo per Fornitore',
            bar_color=C['navy']
        )
        if fig_pf2:
            st.plotly_chart(fig_pf2, use_container_width=True)


    # ── Mark-Up riepilogativo per Fornitore ───────────────────────────────────
    st.subheader('Mark-Up per Fornitore — riepilogo periodo', divider='green')
    st.caption(
        'Mark-Up complessivo per ogni fornitore nel periodo selezionato. '
        'Calcolato sulla quantità effettivamente venduta (minimo tra acquistato e venduto). '
        'Articoli di servizio esclusi (SP TRASP, F00001, T00001).'
    )
    with st.spinner('Calcolo Mark-Up per fornitore...'):
        df_mup_forn = calcola_markup_tutti_fornitori(df_periodo)
    if df_mup_forn.empty:
        st.info('Nessun dato fornitore disponibile per il periodo selezionato.')
    else:
        col_tbl_forn, col_bar_forn = st.columns([1, 1])
        with col_tbl_forn:
            st.dataframe(
                tabella_semaforo(df_mup_forn, 'Mark-Up (%)', soglia_bassa, soglia_alta),
                use_container_width=True, hide_index=True
            )
        with col_bar_forn:
            fig_mup_forn = fig_bar_markup_riepilogo(
                df_mup_forn, 'Fornitore', 'Mark-Up (%)',
                'Mark-Up per Fornitore', soglia_bassa, soglia_alta
            )
            if fig_mup_forn:
                st.plotly_chart(fig_mup_forn, use_container_width=True)

    st.markdown('---')

    # ── Selezione fornitore ───────────────────────────────────────────────────

    st.subheader('Drill-down per Fornitore', divider='green')
    df_forn_list = (
        df_periodo[df_periodo['Cliente - Fornitore'] == 'Fornitore']
        .groupby(['Codice anagrafica', 'Cliente/Fornitore'])['Prezzo finale']
        .sum().reset_index().sort_values('Prezzo finale', ascending=False)
    )
    df_forn_list['Label'] = (
        df_forn_list['Codice anagrafica'].astype(str) + '  —  '
        + df_forn_list['Cliente/Fornitore'].astype(str)
    )

    fornitore_sel = st.selectbox(
        'Seleziona Fornitore:',
        ['— Seleziona Fornitore —'] + df_forn_list['Label'].tolist(),
        key='forn_sel'
    )

    if fornitore_sel == '— Seleziona Fornitore —':
        st.info('Seleziona un fornitore per visualizzare i KPI e il drill-down.')
    else:
        cod_forn  = df_forn_list[df_forn_list['Label'] == fornitore_sel]['Codice anagrafica'].values[0]
        nome_forn = df_forn_list[df_forn_list['Label'] == fornitore_sel]['Cliente/Fornitore'].values[0]

        # Articoli del fornitore, esclusi articoli di servizio
        art_forn = [
            a for a in df_periodo[
                (df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                (df_periodo['Codice anagrafica'] == cod_forn)
            ]['Codice articolo'].unique().tolist()
            if a not in ARTICOLI_ESCLUSI_FORN
        ]

        mask_forn = (
            ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
             (df_periodo['Codice anagrafica'] == cod_forn) &
             (df_periodo['Codice articolo'].isin(art_forn))) |
            ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
             (df_periodo['Codice articolo'].isin(art_forn)))
        )
        df_art_forn = calcola_marginalita_forn(df_periodo[mask_forn])
        fat_f, cost_f, utile_f, marg_f, nv_f = kpi_da_df_art_forn(df_art_forn)

        # ── KPI fornitore ─────────────────────────────────────────────────────
        st.markdown(f'#### KPI — {nome_forn}')
        st.caption('I valori sono calcolati sulla quantità minima tra venduto e acquistato nel periodo.')
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: kpi_card('Fatturato su venduto', fmt_eur(fat_f))
        with col2: kpi_card('Costo su venduto',     fmt_eur(cost_f))
        with col3: kpi_card('Margine su venduto',   fmt_eur(utile_f),
                            semaforo(marg_f, soglia_bassa, soglia_alta))
        with col4: kpi_card('Mark-Up',           fmt_pct(marg_f),
                            semaforo(marg_f, soglia_bassa, soglia_alta))
        with col5: kpi_card('Non ancora fatturato',  fmt_eur(nv_f), 'grigio')

        # ── Trend mensile fornitore ───────────────────────────────────────────
        with st.expander(f'Andamento mensile — {nome_forn}', expanded=True):
            with st.spinner('Calcolo trend...'):
                df_trend_forn = trend_mensile_forn(df_periodo, mask=mask_forn)
            fig_tf = fig_trend_forn(df_trend_forn,
                                    f'Mark-Up mensile — {nome_forn}',
                                    soglia_bassa, soglia_alta)
            st.plotly_chart(fig_tf, use_container_width=True)

        st.markdown('---')

        # ── PANORAMICA CLIENTE x MACROPROGETTO ────────────────────────────────
        st.subheader('Panoramica Cliente × Macroprogetto', divider='green')
        st.caption(
            'Il treemap mostra tutti i clienti e macroprogetti associati al fornitore. '
            'Colore = marginalità %  (rosso → arancio → verde)  |  Dimensione = fatturato. '
            'Usa il bar chart per identificare rapidamente i clienti critici, '
            'poi seleziona un cliente per il drill-down dettagliato.'
        )

        with st.spinner('Calcolo panoramica...'):
            fig_tree, df_tree = fig_treemap_fornitori(
                df_periodo, cod_forn, soglia_bassa, soglia_alta
            )

        if fig_tree is not None:
            st.plotly_chart(fig_tree, use_container_width=True)

            # Bar chart clienti per marginalità
            if not df_tree.empty:
                fig_bar_c, df_cli_marg = fig_bar_clienti_marginalita(
                    df_tree, soglia_bassa, soglia_alta
                )
                st.plotly_chart(fig_bar_c, use_container_width=True)

                # Tabella riepilogativa Cliente x Macroprogetto
                with st.expander('Tabella riepilogativa Cliente x Macroprogetto'):
                    st.dataframe(
                        tabella_semaforo(df_tree.sort_values('Mark-Up (%)', ascending=True),
                                         'Mark-Up (%)', soglia_bassa, soglia_alta),
                        use_container_width=True, hide_index=True
                    )
        else:
            st.warning('Nessun dato disponibile per il fornitore selezionato.')

        st.markdown('---')

        # ── DRILL-DOWN: SELEZIONA UN CLIENTE ─────────────────────────────────
        st.subheader('Drill-down per Cliente', divider='green')
        clienti_del_forn = sorted(
            df_periodo[
                (df_periodo['Cliente - Fornitore'] == 'Cliente') &
                (df_periodo['Codice articolo'].isin(art_forn))
            ]['Cliente/Fornitore'].dropna().unique().tolist()
        )

        if not clienti_del_forn:
            st.warning('Nessun cliente trovato per questo fornitore nel periodo selezionato.')
        else:
            cli_drill = st.selectbox(
                'Seleziona un Cliente per il drill-down:',
                ['— Seleziona Cliente —'] + clienti_del_forn,
                key='forn_cli_drill'
            )

            if cli_drill != '— Seleziona Cliente —':
                cod_c_drill = df_periodo[
                    (df_periodo['Cliente - Fornitore'] == 'Cliente') &
                    (df_periodo['Cliente/Fornitore'] == cli_drill)
                ]['Codice anagrafica'].values[0]

                mask_drill = (
                    ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                     (df_periodo['Codice anagrafica'] == cod_forn) &
                     (df_periodo['Codice articolo'].isin(art_forn))) |
                    ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
                     (df_periodo['Codice anagrafica'] == cod_c_drill) &
                     (df_periodo['Codice articolo'].isin(art_forn)))
                )
                df_art_drill = calcola_marginalita_forn(df_periodo[mask_drill])
                fat_d, cost_d, utile_d, marg_d, nv_d = kpi_da_df_art_forn(df_art_drill)

                # KPI coppia fornitore-cliente
                st.markdown(f'##### {nome_forn}  →  {cli_drill}')
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1: kpi_card('Fatturato su venduto', fmt_eur(fat_d))
                with col2: kpi_card('Costo su venduto',     fmt_eur(cost_d))
                with col3: kpi_card('Margine su venduto',   fmt_eur(utile_d),
                                    semaforo(marg_d, soglia_bassa, soglia_alta))
                with col4: kpi_card('Mark-Up',           fmt_pct(marg_d),
                                    semaforo(marg_d, soglia_bassa, soglia_alta))
                with col5: kpi_card('Non ancora fatturato',  fmt_eur(nv_d), 'grigio')

                # Tabella macroprogetti (informativa, senza selezione)
                st.subheader('Mark-Up per Macroprogetto', divider='green')
                macroprog_drill = sorted(
                    df_periodo[
                        (df_periodo['Cliente - Fornitore'] == 'Cliente') &
                        (df_periodo['Codice anagrafica'] == cod_c_drill) &
                        (df_periodo['Codice articolo'].isin(art_forn))
                    ]['MACROPROGETTO'].dropna().unique().tolist()
                )
                rows_md = []
                for macro in macroprog_drill:
                    art_md = df_periodo[
                        (df_periodo['Cliente - Fornitore'] == 'Cliente') &
                        (df_periodo['Codice anagrafica'] == cod_c_drill) &
                        (df_periodo['Codice articolo'].isin(art_forn)) &
                        (df_periodo['MACROPROGETTO'] == macro)
                    ]['Codice articolo'].unique()
                    mask_md = (
                        ((df_periodo['Cliente - Fornitore'] == 'Cliente') &
                         (df_periodo['Codice anagrafica'] == cod_c_drill) &
                         (df_periodo['MACROPROGETTO'] == macro) &
                         (df_periodo['Codice articolo'].isin(art_forn))) |
                        ((df_periodo['Cliente - Fornitore'] == 'Fornitore') &
                         (df_periodo['Codice anagrafica'] == cod_forn) &
                         (df_periodo['Codice articolo'].isin(art_md)))
                    )
                    ar_md = calcola_marginalita_forn(df_periodo[mask_md])
                    if ar_md.empty:
                        continue
                    mg_md = marginalita_complessiva_forn(ar_md)
                    ft_md = ar_md['fatturato_vendita'].sum()
                    co_md = ar_md['costo_acquisto'].sum()
                    rows_md.append({
                        'Macroprogetto':         macro,
                        'Fatturato su venduto (€)': round(ft_md, 0),
                        'Costo su venduto (€)':  round(co_md, 0),
                        'Margine su venduto (€)': round(ft_md - co_md, 0),
                        'Mark-Up (%)':       round(mg_md, 1) if not np.isnan(mg_md) else np.nan,
                    })

                df_macro_drill = pd.DataFrame(rows_md)
                if not df_macro_drill.empty:
                    st.dataframe(
                        tabella_semaforo(df_macro_drill, 'Mark-Up (%)',
                                         soglia_bassa, soglia_alta),
                        use_container_width=True, hide_index=True
                    )

                # Analisi articoli diretta
                st.subheader('Analisi articoli', divider='green')
                st.markdown(
                    f"**Selezione** — Fatturato su venduto: {fmt_eur(fat_d)}  |  "
                    f"Costo su venduto: {fmt_eur(cost_d)}  |  Mark-Up: **{fmt_pct(marg_d)}**"
                )
                cols_vis_f = [c for c in COLONNE_VIS_ART_FORN if c in df_art_drill.columns]
                df_up_f = mostra_analisi_articoli(
                    df_art_drill, marg_d, soglia_bassa, soglia_alta,
                    key_prefix='forn', label=f'{nome_forn} → {cli_drill}',
                    col_list=COLONNE_VIS_ART_FORN
                )

                # Export fornitore
                st.subheader('Export', divider='green')
                col_ef1, col_ef2 = st.columns(2)
                kpi_df_f = pd.DataFrame([{
                    'Fornitore':                  nome_forn,
                    'Cliente':                    cli_drill,
                    'Periodo':                    periodo_label,
                    'Fatturato su venduto (€)':   round(fat_d, 2),
                    'Costo su venduto (€)':       round(cost_d, 2),
                    'Margine su venduto (€)':     round(utile_d, 2),
                    'Mark-Up (%)':            round(marg_d, 1) if not np.isnan(marg_d) else None,
                    'Non ancora fatturato (€)':   round(nv_d, 2),
                }])
                with col_ef1:
                    xl_f = to_excel({
                        'KPI':                   kpi_df_f,
                        'Clienti_Macroprogetti': df_tree if not df_tree.empty else pd.DataFrame(),
                        'Macroprogetti_Cliente':  df_macro_drill,
                        'Articoli':              df_art_drill[cols_vis_f].round(2),
                        'Trend mensile':         df_trend_forn[['Periodo', 'Mark-Up (%)',
                                                                  'Fatturato su venduto',
                                                                  'Costo su venduto',
                                                                  'Non ancora fatturato']],
                    })
                    st.download_button(
                        'Download Excel',
                        data=xl_f,
                        file_name=f'MarkUp_Forn_{nome_forn.replace(" ","_")}'
                                  f'_{data_da}_{data_a}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='dl_xl_forn'
                    )
                with col_ef2:
                    if st.button('Genera PDF', key='pdf_forn_btn'):
                        with st.spinner('Generazione PDF...'):
                            kpi_pdf_f = {
                                'Fatturato su venduto': (fmt_eur(fat_d),   'grigio'),
                                'Costo su venduto':     (fmt_eur(cost_d),  'grigio'),
                                'Margine su venduto':   (fmt_eur(utile_d),
                                                         semaforo(marg_d, soglia_bassa, soglia_alta)),
                                'Mark-Up':          (fmt_pct(marg_d),
                                                         semaforo(marg_d, soglia_bassa, soglia_alta)),
                                'Non ancora fatturato': (fmt_eur(nv_d),   'grigio'),
                            }
                            pdf_bf = genera_pdf(
                                titolo=f'Analisi Mark-Up — {nome_forn} → {cli_drill}',
                                periodo_str=periodo_label,
                                kpi_dict=kpi_pdf_f,
                                df_riepilogo=df_macro_drill,
                                df_articoli=df_art_drill[cols_vis_f]
                                    .sort_values('Mark-Up', ascending=False)
                                    .head(20).round(2),
                                fig_trend_plotly=fig_tf,
                            )
                        st.download_button(
                            'Scarica PDF',
                            data=pdf_bf,
                            file_name=f'MarkUp_Forn_{nome_forn.replace(" ","_")}'
                                      f'_{data_da}_{data_a}.pdf',
                            mime='application/pdf',
                            key='dl_pdf_forn'
                        )
            else:
                st.info('Seleziona un cliente per il drill-down sugli articoli.')
