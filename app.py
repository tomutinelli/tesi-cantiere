import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io
import base64
from difflib import get_close_matches
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

    /* Pulsanti di download stilizzati (Design moderno) */
    .custom-dl-btn {
        text-decoration: none !important;
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        color: #374151 !important;
        padding: 0.6rem 1rem !important;
        border-radius: 8px !important;
        font-size: 0.9rem !important;
        font-weight: 500 !important;
        display: block !important;
        text-align: center !important;
        width: 100% !important;
        margin-top: 8px !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02) !important;
        transition: all 0.25s ease-in-out !important;
        letter-spacing: 0.3px !important;
    }
    
    .custom-dl-btn:hover {
        border-color: #a7b89f !important;
        background-color: #f4f7f3 !important;
        color: #111827 !important;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.06) !important;
        transform: translateY(-2px) !important;
    }
</style>
""", unsafe_allow_html=True)

# --- FUNZIONE DOWNLOAD BASE64 CON TARGET TOP ---
def genera_link_download(data_bytes, filename, button_text):
    b64 = base64.b64encode(data_bytes).decode()
    mime = "text/csv" if filename.endswith('.csv') else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return f'<a class="custom-dl-btn" href="data:{mime};base64,{b64}" download="{filename}" target="_top">{button_text}</a>'

st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

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
        dimensione_cantiere = st.number_input("Area di cantiere complessiva (metri quadri)", min_value=0.1, value=100.0, step=1.0)
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

# --- BLOCCO INPUT DATI (IA vs CSV MANUALE) ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Caricamento Dataset di Progetto")

tab1, tab2 = st.tabs(["Elaborazione Intelligente (IA)", "Caricamento CSV Manuale"])

with tab1:
    st.markdown("""
    <div style='background-color: #f4f7f3; border: 1.5px solid #d5ddd1; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #111827; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>Guida Operativa: Procedura per l'Elaborazione con IA</h4>
        <ol style='color: #4b5563; font-size: 0.9rem; line-height: 1.6; margin-bottom: 0; padding-left: 20px;'>
            <li><b>Prepara la documentazione di cantiere:</b> Raccogli i file di progetto (es. computo metrico, cronoprogramma e log dei trasporti).</li>
            <li><b>Carica i file nella barra unica:</b> Trascina contemporaneamente tutti i documenti nel riquadro sottostante. L'IA provvederà anche a convertire automaticamente le unità di misura (es. $m^3$ di cls in kg tramite densità) per adeguarle agli standard LCI.</li>
            <li><b>Avvia l'analisi semantica:</b> Clicca sul pulsante <i>"Elabora e Normalizza con IA"</i> per estrarre e unificare i dati.</li>
            <li><b>Visualizza i risultati:</b> Controlla l'anteprima della tabella normalizzata, scarica il CSV pulito e analizza i grafici aggiornati.</li>
            <li><i>Se si possiede già il file CSV impostato correttamente, lo si inserisca direttamente nella scheda apposita "Caricamento CSV Manuale".</i></li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    files_unificati = st.file_uploader(
        "Carica tutti i documenti di cantiere insieme (Computo, Cronoprogramma, Trasporti)", 
        type=['pdf', 'txt', 'xlsx', 'csv'], 
        accept_multiple_files=True,
        key="ia_unified"
    )

    if st.button("Elabora e Normalizza con IA", key="btn_ia"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("Chiave API mancante nei Secrets. Inserisci GEMINI_API_KEY per utilizzare questa funzione.")
        elif not files_unificati:
            st.warning("Carica almeno un file per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("L'intelligenza artificiale sta analizzando e convertendo le unità di misura..."):
                    contents = []
                    for file_obj in files_unificati:
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
                    Sei un ingegnere edile e analista LCA. Analizza congiuntamente TUTTI i documenti forniti (computo metrico, cronoprogramma, trasporti).
                    Restituisci un'unica tabella CSV pulita con queste esatte 4 intestazioni di colonna:
                    Data,Parametro,Elemento,Quantita
                    
                    - Data: Formato AAAA-MM-GG. Distribuisci le quantità lungo le date del cantiere.
                    - Parametro: Scegli tassativamente tra: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
                    - Elemento: Descrizione standardizzata dell'elemento o materiale.
                    - Quantita: Valore numerico convertito nell'unità di misura coerente con il database LCI (es. converti i metri cubi di calcestruzzo in kg moltiplicando per la densità di 2400 kg/m3; converti tonnellate di acciaio in kg moltiplicando per 1000). Per i trasporti, massa in tonnellate moltiplicata per i km (t*km).
                    
                    Restituisci ESCLUSIVAMENTE il codice CSV grezzo, senza blocchi Markdown, pronto per pd.read_csv().
                    """
                    contents.append(prompt_sistema)
                    
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=contents
                    )
                    
                    csv_testo = response.text.strip()
                    if csv_testo.startswith("```"):
                        csv_testo = csv_testo.split("
