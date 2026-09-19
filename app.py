import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io
from google import genai
from google.genai import types

# Impostazioni della pagina
st.set_page_config(page_title="EcoSite Tracker | LCA Dashboard", layout="wide")

# --- STILE CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1f2937;
        background-color: #faf9f6;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Card minimaliste con bordo verde pastello */
    div.minimal-card {
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        border-radius: 8px !important;
        padding: 32px !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.01) !important;
        margin-bottom: 24px !important;
    }

    h1 {
        font-weight: 600;
        letter-spacing: -0.025em;
        color: #111827;
        font-size: 2.25rem;
    }
    
    h2, h3 {
        font-weight: 500;
        letter-spacing: -0.01em;
        color: #1f2937;
    }

    .stRadio label, .stNumberInput label, .stDateInput label, .stFileUploader label, .stSelectbox label {
        font-weight: 500;
        font-size: 0.9rem;
        color: #4b5563;
    }

    /* Stile per i Bottoni Nativi (Download e Link) in Verde Pastello */
    [data-testid="stDownloadButton"] button, 
    [data-testid="stLinkButton"] button,
    [data-testid="baseButton-secondary"] {
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        color: #374151 !important;
        border-radius: 8px !important;
        transition: all 0.25s ease-in-out !important;
        font-weight: 500 !important;
        width: 100% !important;
    }
    
    [data-testid="stDownloadButton"] button:hover, 
    [data-testid="stLinkButton"] button:hover,
    [data-testid="baseButton-secondary"]:hover {
        border-color: #a7b89f !important;
        background-color: #f4f7f3 !important;
        color: #111827 !important;
    }

    [data-testid="stDownloadButton"] button:focus:not(:active), 
    [data-testid="stLinkButton"] button:focus:not(:active),
    [data-testid="baseButton-secondary"]:focus:not(:active) {
        border-color: #a7b89f !important;
        box-shadow: 0 0 0 1px #a7b89f !important;
        color: #374151 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- DIZIONARIO COLORI GLOBALE ---
colori_parametri = {
    'Materiali': "#B80D0D", 
    'Rifiuti': "#078303", 
    'Trasporti': "#732BB7", 
    'Energia': "#f6de03",
    'Acqua': "#53DCFE", 
    'Macchinari': "#8a929e"
}

# --- GUIDA METODOLOGICA ---
with st.expander("Note metodologiche e specifiche di utilizzo"):
    st.markdown("""
    <p style='color: #4b5563; font-size: 0.95rem; line-height: 1.6; margin-bottom: 0;'>
    Questo strumento calcola l'impronta di carbonio (kg di CO₂e) in fase di progetto incrociando i dati di consumo giornaliero con il database LCI di riferimento. 
    </p>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

# --- BLOCCO CONFIGURAZIONE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Configurazione dell'Unità Funzionale")
st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Definisci i parametri di normalizzazione globale per l'opera.</p>", unsafe_allow_html=True)

tipo_cantiere = st.radio(
    "Tipologia di modellazione spaziale",
    options=[
        "Cantiere Lineare (normalizzazione per metro lineare - m)", 
        "Cantiere Standard (normalizzazione per metro quadro - m²)"
    ],
    label_visibility="collapsed"
)

st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
col_cfg1, col_cfg2 = st.columns([2, 1])
with col_cfg1:
    if "Lineare" in tipo_cantiere:
        unita = "m"
        dimensione_cantiere = st.number_input("Estensione longitudinale complessiva (metri)", min_value=0.1, value=100.0, step=1.0)
    else:
        unita = "m²"
        dimensione_cantiere = st.number_input("Superficie coperta complessiva (metri quadri)", min_value=0.1, value=100.0, step=1.0)
st.markdown("</div>", unsafe_allow_html=True)

# --- LETTURA DATABASE LCI ---
@st.cache_data
def carica_database_lci(percorso_file):
    if not os.path.exists(percorso_file):
        return None
    try:
        xls = pd.ExcelFile(percorso_file)
    except Exception:
        return None
        
    df_inv = pd.DataFrame()
    mappatura = {
        'Materiali': {'nome_elemento': 'nome_materiale', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Rifiuti': {'nome_elemento': 'tipo_rifiuto', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Energia': {'nome_elemento': 'Tecnologia di generazione elettrica', 'nome_fattore': 'Fattori di emissione (kg CO2eq/kWh)'},
        'Acqua': {'nome_elemento': 'Elemento', 'nome_fattore': 'Fattore_Emissione'},
        'Trasporti': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'},
        'Macchinari': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'}
    }
    
    for foglio in xls.sheet_names:
        df_temp = pd.read_excel(xls, sheet_name=foglio)
        col_el = None
        col_fat = None
        
        if foglio in mappatura:
            col_el = mappatura[foglio]['nome_elemento']
            col_fat = mappatura[foglio]['nome_fattore']
            if foglio in ['Trasporti', 'Macchinari']:
                df_temp = df_temp.iloc[0:9].copy()
        else:
            possibili_nomi_elemento = ['Elemento', 'Macchinario', 'Nome', 'Acqua', 'Tipo', 'Fuel']
            possibili_nomi_fattore = ['Fattore_Emissione', 'kg CO2eq', 'Emissione', 'emissioni_kg_co2eq_kg', 'kg co2 per kg of fuel']
            
            for c in df_temp.columns:
                if any(x.lower() in str(c).lower() for x in possibili_nomi_elemento) and col_el is None:
                    col_el = c
                if any(x.lower() in str(c).lower() for x in possibili_nomi_fattore) and col_fat is None:
                    col_fat = c
        
        if col_el in df_temp.columns and col_fat in df_temp.columns:
            df_temp = df_temp[[col_el, col_fat]].copy()
            df_temp.rename(columns={col_el: 'Elemento', col_fat: 'Fattore_Emissione'}, inplace=True)
            df_temp['Parametro'] = foglio 
            df_temp['Fattore_Emissione'] = pd.to_numeric(df_temp['Fattore_Emissione'], errors='coerce')
            df_inv = pd.concat([df_inv, df_temp], ignore_index=True)
                
    return df_inv.dropna(subset=['Elemento', 'Fattore_Emissione'])

percorso_lci = "LCI.xlsx"
df_inventario = carica_database_lci(percorso_lci)

if df_inventario is None or df_inventario.empty:
    st.error(f"Errore critico: Il database '{percorso_lci}' non è reperibile o non è valido sul server.")
    st.stop()

# --- BLOCCO INPUT DATI (IA vs MANUALE) ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Caricamento Dataset di Progetto")

tab1, tab2 = st.tabs(["✨ Elaborazione Intelligente (IA)", "📂 Caricamento CSV Manuale"])

with tab1:
    st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Carica i tuoi elaborati (PDF, TXT, Excel). L'IA estrarrà i dati, mapperà i materiali e calcolerà i trasporti automaticamente.</p>", unsafe_allow_html=True)
    
    file_computo = st.file_uploader("Computo Metrico (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'], key="ia_comp")
    file_cronoprogramma = st.file_uploader("Cronoprogramma / Gantt (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'], key="ia_crono")
    file_trasporti = st.file_uploader("Note distanze trasporti (TXT o PDF)", type=['txt', 'pdf', 'xlsx', 'csv'], key="ia_trasp")

    if st.button("Elabora e Normalizza con IA", key="btn_ia"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("Chiave API mancante nei Secrets. Inserisci GEMINI_API_KEY per utilizzare questa funzione.")
        elif not file_computo:
            st.warning("Carica almeno il file del computo metrico per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("L'intelligenza artificiale sta analizzando gli elaborati..."):
                    contents = []
                    
                    for file_obj in [file_computo, file_cronoprogramma, file_trasporti]:
                        if file_obj is not None:
                            estensione = file_obj.name.split('.')[-1].lower()
                            if estensione == 'pdf':
                                mime = 'application/pdf'
                            elif estensione == 'xlsx':
                                mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            elif estensione == 'csv':
                                mime = 'text/csv'
                            else:
                                mime = 'text/plain'
                            
                            contents.append(
                                types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime)
                            )
                    
                    prompt_sistema = """
                    Sei un esperto ingegnere edile e analista LCA. 
                    Il tuo compito è analizzare i documenti di progetto forniti in input (computi, cronoprogrammi, note trasporti) e generarne un'unica tabella CSV pulita.
                    
                    Il CSV finale DEVE avere esattamente queste 4 intestazioni di colonna:
                    Data,Parametro,Elemento,Quantita
                    
                    Regole ferree:
                    1. 'Data': Formato AAAA-MM-GG. Se c'è un cronoprogramma, distribuisci le quantità nelle date corrette.
                    2. 'Parametro': Scegli tra: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
                    3. 'Elemento': Il nome dell'elemento o del combustibile coerente con i database LCA.
                    4. 'Quantita': Valore numerico (per trasporti: massa in tonnellate x km).
                    
                    Restituisci ESCLUSIVAMENTE il codice CSV grezzo, pronto per pd.read_csv().
                    """
                    contents.append(prompt_sistema)
                    
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=contents
                    )
                    
                    csv_testo = response.text.strip()
                    if csv_testo.startswith("```"):
                        csv_testo = csv_testo.split("
