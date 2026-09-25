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
            In alternativa, Se si possiede già il file CSV impostato correttamente, lo si inserisca direttamente nella scheda apposita "Caricamento CSV Manuale".
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    files_unificati = st.file_uploader(
        "Carica tutti i documenti di cantiere insieme (Computo, Cronoprogramma, Trasporti)", 
        type=['pdf', 'txt', 'xlsx', 'csv', 'xml'], 
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
                with st.spinner("L'intelligenza artificiale sta analizzando il cronoprogramma e il computo..."):
                    contents = []
                    for file_obj in files_unificati:
                        estensione = file_obj.name.split('.')[-1].lower()
                        if estensione == 'pdf':
                            mime = 'application/pdf'
                        elif estensione == 'xlsx':
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        elif estensione in ['csv', 'xml', 'txt']:
                            mime = 'text/plain'
                        else:
                            mime = 'application/octet-stream'
                        
                        contents.append(
                            types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime)
                        )
                    
                    # PROMPT AGGIORNATO PER MATCHING FASI
                    prompt_sistema = """
                    Sei un ingegnere edile e analista LCA. Analizza congiuntamente TUTTI i documenti forniti, in particolare il computo metrico e il cronoprogramma (che può essere in formato XML o testo).
                    
                    CRITERIO FONDAMENTALE DI MATCHING:
                    Nota bene che le fasi del cronoprogramma e le fasi del computo sono strutturate e nominate in modo identico (ad esempio: Fase RI 1-3, Fase CO 1-3, Fase GR 1-3, ecc.).
                    Devi sfruttare questa identica suddivisione per associare in modo esatto le lavorazioni e le quantità (presenti nel Computo) al periodo temporale corretto (presente nel Cronoprogramma).
                    
                    Restituisci un'unica tabella CSV pulita con queste esatte 4 intestazioni di colonna:
                    Data,Parametro,Elemento,Quantita
                    
                    - Data: Formato AAAA-MM-GG. Spalma e distribuisci la quantità totale di una voce di computo lungo tutti i giorni lavorativi previsti nel cronoprogramma per quella specifica fase.
                    - Parametro: Scegli tassativamente tra: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
                    - Elemento: Descrizione standardizzata dell'elemento o materiale.
                    - Quantita: Valore numerico della singola riga convertito nell'unità di misura coerente con il database LCI (es. converti i metri cubi di calcestruzzo in kg moltiplicando per la densità di 2400 kg/m3; converti tonnellate di acciaio in kg moltiplicando per 1000). 
                    
                    Non inserire commenti o spiegazioni. Restituisci ESCLUSIVAMENTE il codice CSV grezzo, senza blocchi Markdown, pronto per pd.read_csv().
                    """
                    contents.append(prompt_sistema)
                    
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=contents
                    )
                    
                    csv_testo = response.text.strip()
                    if csv_testo.startswith("```"):
                        csv_testo = csv_testo.split("```")[1]
                        if csv_testo.startswith("csv"):
                            csv_testo = csv_testo[3:].strip()
                        elif csv_testo.startswith("\n"):
                            csv_testo = csv_testo.strip()
                    
                    df_cantiere_grezzo = pd.read_csv(io.StringIO(csv_testo))
                    
                    colonne_mappa = {}
                    for c in df_cantiere_grezzo.columns:
                        c_low = str(c).strip().lower()
                        if 'data' in c_low or 'date' in c_low or 'giorno' in c_low:
                            colonne_mappa[c] = 'Data'
                        elif 'param' in c_low or 'categ' in c_low:
                            colonne_mappa[c] = 'Parametro'
                        elif 'elem' in c_low or 'material' in c_low or 'voce' in c_low:
                            colonne_mappa[c] = 'Elemento'
                        elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low or 'amount' in c_low:
                            colonne_mappa[c] = 'Quantita'
                    
                    df_cantiere_grezzo.rename(columns=colonne_mappa, inplace=True)
                    
                    if 'Parametro' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Parametro'] = 'Materiali'
                    if 'Data' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Data'] = str(date.today())
                    if 'Elemento' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Elemento'] = 'Calcestruzzo'
                    if 'Quantita' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Quantita'] = 1.0

                    def mappa_voce_a_lci(parametro, elemento_grezzo):
                        p_str = str(parametro).strip()
                        e_str = str(elemento_grezzo).strip().lower()
                        
                        voci_disponibili = df_inventario[df_inventario['Parametro'].str.lower() == p_str.lower()]['Elemento'].tolist()
                        if not voci_disponibili:
                            voci_disponibili = df_inventario['Elemento'].tolist()
                            
                        for v in voci_disponibili:
                            if v.lower() in e_str or e_str in v.lower():
                                return v
                                
                        if p_str.lower() == 'materiali':
                            if 'calcestruzzo' in e_str or 'cls' in e_str: return 'Calcestruzzo'
                            if 'acciaio' in e_str or 'ferro' in e_str: return 'Acciaio'
                            if 'laterizio' in e_str or 'mattone' in e_str: return 'Laterizio'
                            if 'inerti' in e_str or 'sabbia' in e_str or 'ghiaia' in e_str: return 'Inerti'
                            if 'asfalto' in e_str or 'bitume' in e_str: return 'Asfalto/Bitume'
                            if 'legno' in e_str: return 'Legno'
                            if 'vetro' in e_str: return 'Vetro'
                            if 'isolante' in e_str or 'lana' in e_str or 'eps' in e_str: return 'Isolante EPS'
                        elif p_str.lower() == 'rifiuti':
                            if 'scavo' in e_str or 'terra' in e_str: return 'Inerti / Macerie di demolizione'
                            if 'calcestruzzo' in e_str: return 'Calcestruzzo di risulta'
                            if 'acciaio' in e_str or 'ferro' in e_str: return 'Metallo / Acciaio di scarto'
                            if 'legno' in e_str: return 'Legno da cantiere'
                        elif p_str.lower() in ['trasporti', 'macchinari']:
                            if 'diesel' in e_str or 'gasolio' in e_str: return 'Diesel'
                            if 'benzina' in e_str or 'petrol' in e_str: return 'Petrol'
                            
                        match = get_close_matches(str(elemento_grezzo), voci_disponibili, n=1, cutoff=0.1)
                        if match:
                            return match[0]
                            
                        return elemento_grezzo

                    df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(
                        lambda row: mappa_voce_a_lci(row['Parametro'], row['Elemento']), axis=1
                    )
                    
                    st.session_state['df_cantiere'] = df_cantiere_grezzo
                    st.success("Tutti i documenti sono stati analizzati e convertiti con successo sfruttando il matching delle fasi!")
            except Exception as e:
                st.error(f"Errore durante l'elaborazione con l'IA: {e}")

    # Anteprima e Download CSV IA
    if 'df_cantiere' in st.session_state:
        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 👁️ Anteprima Dati Elaborati dall'IA")
        st.dataframe(st.session_state['df_cantiere'], use_container_width=True)
        
        csv_esportato = st.session_state['df_cantiere'].to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Scarica CSV Normalizzato dall'IA", 
            data=csv_esportato, 
            file_name="dataset_cantiere_normalizzato.csv", 
            mime="text/csv",
            key="download_csv_ia"
        )

with tab2:
    st.markdown("""
    <p style='color: #4b5563; font-size: 0.95rem; line-height: 1.6; margin-bottom: 15px;'>
    Questo strumento calcola l'impronta di carbonio (kg di CO₂e) in fase di progetto incrociando i dati con il database LCI di riferimento. 
    Il dataset deve essere strutturato in 4 colonne: <code>Data</code> (AAAA-MM-GG), <code>Parametro</code>, <code>Elemento</code> e <code>Quantita</code> (con unità già convertite secondo gli standard LCI).
    </p>
    """, unsafe_allow_html=True)
    
    file_cantiere = st.file_uploader("Seleziona file CSV", type=['csv'], label_visibility="collapsed", key="csv_manuale")
    if file_cantiere:
        try:
            st.session_state['df_cantiere'] = pd.read_csv(file_cantiere)
            st.success("File CSV caricato correttamente!")
        except Exception as e:
            st.error(f"Errore nella lettura del file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# =====================================================================
# ELABORAZIONE E ANALISI LCA
# =====================================================================
if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere']
    
    mappa_colonne_finali = {}
    for col in df_cantiere.columns:
        c_low = str(col).strip().lower()
        if 'data' in c_low or 'date' in c_low:
            mappa_colonne_finali[col] = 'Data'
        elif 'param' in c_low or 'categ' in c_low:
            mappa_colonne_finali[col] = 'Parametro'
        elif 'elem' in c_low or 'material' in c_low:
            mappa_colonne_finali[col] = 'Elemento'
        elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low:
            mappa_colonne_finali[col] = 'Quantita'
            
    df_cantiere.rename(columns=mappa_colonne_finali, inplace=True)
    
    try:
        for col in ['Data', 'Parametro', 'Elemento', 'Quantita']:
            if col not in df_cantiere.columns:
                st.error(f"Errore di struttura: La colonna '{col}' risulta assente.")
                st.stop()
        
        df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
        
        df_cantiere['Parametro_match'] = df_cantiere['Parametro'].astype(str).str.strip().str.lower()
        df_cantiere['Elemento_match'] = df_cantiere['Elemento'].astype(str).str.strip().str.lower()
        
        df_inventario['Parametro_match'] = df_inventario['Parametro'].astype(str).str.strip().str.lower()
        df_inventario['Elemento_match'] = df_inventario['Elemento'].astype(str).str.strip().str.lower()
        df_inventario_clean = df_inventario.drop_duplicates(subset=['Parametro_match', 'Elemento_match']).copy()
        
        df_completo = pd.merge(
            df_cantiere, 
            df_inventario_clean, 
            on=['Parametro_match', 'Elemento_match'], 
            how='left',
            suffixes=('', '_lci')
        )
        
        df_completo['Parametro'] = df_completo['Parametro'].fillna(df_completo['Parametro_match'])
        df_completo['Elemento'] = df_completo['Elemento'].fillna(df_completo['Elemento_match'])
        df_completo.drop(columns=['Parametro_match', 'Elemento_match'], inplace=True, errors='ignore')
        if 'Parametro_lci' in df_completo.columns: df_completo.drop(columns=['Parametro_lci'], inplace=True)
        if 'Elemento_lci' in df_completo.columns: df_completo.drop(columns=['Elemento_lci'], inplace=True)
        
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.warning(f"Elementi o Parametri non riconosciuti nel database LCI (calcolati a zero): {mancanti['Elemento'].unique().tolist()}")
            df_completo['Fattore_Emissione'] = df_completo['Fattore_Emissione'].fillna(0)
            
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        df_completo['CO2_Normalizzata'] = df_completo['CO2_Totale_kg'] / dimensione_cantiere
        
        # --- FILTRO TEMPORALE ---
        st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
        st.subheader("Filtro Temporale")
        st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 15px;'>Seleziona l'arco temporale di interesse per l'analisi.</p>", unsafe_allow_html=True)
        
        min_date = df_completo['Data_dt'].min().date()
        max_date = df_completo['Data_dt'].max().date()

        date_range = st.date_input(
            "Intervallo temporale",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            label_visibility="collapsed"
        )

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df_completo['Data_dt'].dt.date >= start_date) & (df_completo['Data_dt'].dt.date <= end_date)
            df_filtrato = df_completo.loc[mask].copy()
        elif isinstance(date_range, date):
            mask = (df_completo['Data_dt'].dt.date == date_range)
            df_filtrato = df_completo.loc[mask].copy()
        else:
            df_filtrato = df_completo.copy()
            
        st.markdown("</div>", unsafe_allow_html=True)

        # --- VISUALIZZAZIONE GRAFICI ---
        st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
        st.subheader("Risultati Analitici")
        st.markdown(f"<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 25px;'>I valori visualizzati esprimono l'incidenza normalizzata rispetto all'unità funzionale complessiva (<b>{dimensione_cantiere} {unita}</b>).</p>", unsafe_allow_html=True)
        
        if df_filtrato.empty:
            st.warning("Nessuna evidenza registrata nell'intervallo temporale selezionato.")
        else:
            modo_visualizzazione = st.selectbox(
                "Modalità di visualizzazione grafica",
                options=[
                    "Standard (Barre temporali)", 
                    "Distribuzione (KDE / Frequenza)", 
                    "Linea Media (Andamento con media mobile)",
                    "Linea Media Fissa (Orizzontale)"
                ],
                index=0
            )

            def genera_figura(df_dat, x_col, y_col, titolo, colore_base):
                df_dat = df_dat.copy()
                if df_dat.empty:
                    return go.Figure()
                    
                if "Distribuzione" in modo_visualizzazione:
                    fig = px.histogram(
                        df_dat, x=y_col, nbins=20, marginal="violin",
                        title=f"{titolo} - Distribuzione Frazionale",
                        labels={y_col: f'kg CO₂e / {unita}', 'count': 'Frequenza'},
                        color_discrete_sequence=[colore_base]
                    )
                elif "Media mobile" in modo_visualizzazione or "Media (Andamento" in modo_visualizzazione:
                    df_dat = df_dat.sort_values(by=x_col)
                    df_dat['Media_Mobile'] = df_dat[y_col].rolling(window=7, min_periods=1).mean()
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Valore Giornaliero', marker_color=colore_base, opacity=0.4))
                    fig.add_trace(go.Scatter(x=df_dat[x_col], y=df_dat['Media_Mobile'], mode='lines', name='Linea Media (7gg)', line=dict(color='#111827', width=3)))
                    fig.update_layout(title=f"{titolo} - Andamento con Media Mobile")
                elif "Fissa" in modo_visualizzazione:
                    media_val = df_dat[y_col].mean()
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Valore Giornaliero', marker_color=colore_base, opacity=0.5))
                    fig.add_hline(
                        y=media_val, 
                        line_dash="dash", 
                        line_color="#27ae60", 
                        annotation_text=f"Media Fissa: {media_val:.2f}", 
                        annotation_position="top right"
                    )
                    fig.update_layout(title=f"{titolo} - Con Linea Media Fissa nell'Intervallo")
                else:
                    fig = px.bar(
                        df_dat, x=x_col, y=y_col,
                        title=titolo,
                        labels={y_col: f'kg CO₂e / {unita}', x_col: ''},
                        text_auto='.2f'
                    )
                    fig.update_traces(marker_color=colore_base)
                
                fig.update_layout(
                    height=380, 
                    font_family="Inter", 
                    plot_bgcolor='rgba(0,0,0,0)', 
                    paper_bgcolor='rgba(0,0,0,0)',
                    title_font_size=14,
                    title_font_color="#374151"
                )
                return fig

            # 1. Grafico Totale
            df_totale = df_filtrato.groupby('Data')['CO2_Normalizzata'].sum().reset_index()
            fig_tot = genera_figura(df_totale, 'Data', 'CO2_Normalizzata', f"Andamento Complessivo (kg CO₂e / {unita})", "#0B0752")
            st.plotly_chart(fig_tot, use_container_width=True)
            
            # 2. Diagramma a Torta (Incidenza Percentuale)
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            df_pie = df_filtrato.groupby('Parametro')['CO2_Normalizzata'].sum().reset_index()
            
            colori_parametri = {
                'Materiali': "#B80D0D", 
                'Rifiuti': "#078303", 
                'Trasporti': "#732BB7", 
                'Energia': "#f6de03",
                'Acqua': "#53DCFE", 
                'Macchinari': "#8a929e"
            }
            
            fig_pie = px.pie(
                df_pie, 
                values='CO2_Normalizzata', 
                names='Parametro',
                title="Incidenza Percentuale delle Categorie sulle Emissioni Totali",
                color='Parametro',
                color_discrete_map=colori_parametri,
                hole=0.4
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
            fig_pie.update_layout(
                height=450, 
                font_family="Inter", 
                showlegend=True, 
                title_font_size=14, 
                title_font_color="#374151",
                plot_bgcolor='rgba(0,0,0,0)', 
                paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_pie, use_container_width=True)

            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                csv_totale = df_totale.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Scarica dati totali (CSV)", data=csv_totale, file_name="emissioni_totali_giornaliere.csv", mime="text/csv")
            with col_exp2:
                output_xlsx = io.BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                    df_totale.to_excel(writer, index=False, sheet_name='Totale Giornaliero')
                st.download_button("📥 Scarica dati totali (XLSX)", data=output_xlsx.getvalue(), file_name="emissioni_totali_giornaliere.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
            st.subheader("Disaggregazione per Categoria")
            
            parametri_presenti = df_filtrato['Parametro'].dropna().unique()
            col_grafici = st.columns(2)
            
            for idx, param in enumerate(parametri_presenti):
                df_p = df_filtrato[df_filtrato['Parametro'] == param].groupby('Data')['CO2_Normalizzata'].sum().reset_index()
                colore_cat = colori_parametri.get(param, '#4b5563')
                
                fig_cat = genera_figura(df_p, 'Data', 'CO2_Normalizzata', f"{param}", colore_cat)
                
                with col_grafici[idx % 2]:
                    st.plotly_chart(fig_cat, use_container_width=True)
                    csv_cat = df_p.to_csv(index=False).encode('utf-8')
                    st.download_button(f"📥 Scarica {param} (CSV)", data=csv_cat, file_name=f"dati_{param.lower()}.csv", mime="text/csv", key=f"dl_{param}")

            # Tabella Dati Globale
            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
            with st.expander("Esporta / Visualizza matrice dati completa"):
                st.dataframe(df_filtrato.drop(columns=['Data_dt']), use_container_width=True)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_filtrato.drop(columns=['Data_dt']).to_excel(writer, index=False, sheet_name='Dettaglio Completo')
                st.download_button("📥 Scarica intero dataset filtrato (XLSX)", data=excel_buffer.getvalue(), file_name="dataset_completo_lca.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("</div>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Si è verificato un errore durante l'elaborazione: {e}")
else:
    st.info("Carica i documenti di progetto tramite IA o seleziona un file CSV per avviare l'elaborazione analitica.")
