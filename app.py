import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io
import base64
import time
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

    div.minimal-card {
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        border-radius: 8px !important;
        padding: 32px !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.01) !important;
        margin-bottom: 24px !important;
    }

    h1 { font-weight: 600; letter-spacing: -0.025em; color: #111827; font-size: 2.25rem; }
    h2, h3 { font-weight: 500; letter-spacing: -0.01em; color: #1f2937; }

    .stRadio label, .stNumberInput label, .stDateInput label, .stFileUploader label, .stSelectbox label {
        font-weight: 500; font-size: 0.9rem; color: #4b5563;
    }

    .custom-dl-btn {
        text-decoration: none !important; background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important; color: #374151 !important;
        padding: 0.6rem 1rem !important; border-radius: 8px !important;
        font-size: 0.9rem !important; font-weight: 500 !important;
        display: block !important; text-align: center !important;
        width: 100% !important; margin-top: 8px !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02) !important;
        transition: all 0.25s ease-in-out !important; letter-spacing: 0.3px !important;
    }
    
    .custom-dl-btn:hover {
        border-color: #a7b89f !important; background-color: #f4f7f3 !important;
        color: #111827 !important; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.06) !important;
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

# --- LETTURA DATABASE LCI ---
@st.cache_data
def carica_database_lci(percorso_file):
    if not os.path.exists(percorso_file): return None
    try: xls = pd.ExcelFile(percorso_file)
    except Exception: return None
        
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
        col_el, col_fat = None, None
        
        if foglio in mappatura:
            col_el = mappatura[foglio]['nome_elemento']
            col_fat = mappatura[foglio]['nome_fattore']
            if foglio in ['Trasporti', 'Macchinari']: df_temp = df_temp.iloc[0:9].copy()
        else:
            possibili_nomi_elemento = ['Elemento', 'Macchinario', 'Nome', 'Acqua', 'Tipo', 'Fuel']
            possibili_nomi_fattore = ['Fattore_Emissione', 'kg CO2eq', 'Emissione', 'emissioni_kg_co2eq_kg', 'kg co2 per kg of fuel']
            
            for c in df_temp.columns:
                if any(x.lower() in str(c).lower() for x in possibili_nomi_elemento) and col_el is None: col_el = c
                if any(x.lower() in str(c).lower() for x in possibili_nomi_fattore) and col_fat is None: col_fat = c
        
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
    st.error(f"Errore critico: Il database '{percorso_lci}' non è reperibile.")
    st.stop()

# --- FUNZIONE GLOBALE DI MAPPING ---
def mappa_voce_a_lci(parametro, elemento_grezzo, df_inv=df_inventario):
    p_str = str(parametro).strip()
    e_str = str(elemento_grezzo).strip().lower()
    
    voci_disponibili = df_inv[df_inv['Parametro'].str.lower() == p_str.lower()]['Elemento'].tolist()
    if not voci_disponibili: voci_disponibili = df_inv['Elemento'].tolist()
        
    for v in voci_disponibili:
        if v.lower() in e_str or e_str in v.lower(): return v
            
    if p_str.lower() == 'materiali':
        if 'calcestruzzo' in e_str or 'cls' in e_str: return 'Calcestruzzo'
        if 'acciaio' in e_str or 'ferro' in e_str: return 'Acciaio'
        if 'laterizio' in e_str or 'mattone' in e_str: return 'Laterizio'
        if 'inerti' in e_str or 'sabbia' in e_str or 'ghiaia' in e_str: return 'Inerti'
        if 'asfalto' in e_str or 'bitume' in e_str or 'strada' in e_str: return 'Asfalto/Bitume'
        if 'legno' in e_str: return 'Legno'
        if 'vetro' in e_str: return 'Vetro'
        if 'isolante' in e_str or 'lana' in e_str or 'eps' in e_str: return 'Isolante EPS'
    elif p_str.lower() == 'rifiuti':
        if 'scavo' in e_str or 'terra' in e_str: return 'Inerti / Macerie di demolizione'
        if 'calcestruzzo' in e_str: return 'Calcestruzzo di risulta'
        if 'acciaio' in e_str or 'ferro' in e_str: return 'Metallo / Acciaio di scarto'
    elif p_str.lower() in ['trasporti', 'macchinari']:
        if 'diesel' in e_str or 'gasolio' in e_str: return 'Diesel'
        if 'benzina' in e_str or 'petrol' in e_str: return 'Petrol'
        
    match = get_close_matches(str(elemento_grezzo), voci_disponibili, n=1, cutoff=0.1)
    if match: return match[0]
        
    return elemento_grezzo


# --- BLOCCO INPUT DATI ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Caricamento Dataset di Progetto")

def reset_dati():
    if 'df_cantiere' in st.session_state: del st.session_state['df_cantiere']

tab1, tab2 = st.tabs(["Elaborazione Intelligente (IA)", "Caricamento CSV Manuale"])

with tab1:
    st.markdown("""
    <div style='background-color: #f4f7f3; border: 1.5px solid #d5ddd1; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #111827; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>Guida Operativa: Procedura per l'Elaborazione con IA</h4>
        <ol style='color: #4b5563; font-size: 0.9rem; line-height: 1.6; margin-bottom: 0; padding-left: 20px;'>
            <li><b>Prepara la documentazione di cantiere:</b> Raccogli i file (es. esportazioni IFC, abachi da Revit, cronoprogramma).</li>
            <li><b>Carica i file nella barra unica:</b> Trascina contemporaneamente tutti i documenti nel riquadro.</li>
            <li><b>Avvia l'analisi semantica:</b> Clicca sul pulsante <i>"Elabora e Normalizza con IA"</i> per estrarre e unificare i dati temporali, dei materiali e dei macchinari.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    files_unificati = st.file_uploader(
        "Carica tutti i documenti di cantiere (IFC, Cronoprogramma, Abachi Materiali)", 
        type=['pdf', 'txt', 'xlsx', 'csv', 'xml', 'ifc'], 
        accept_multiple_files=True,
        key="ia_unified",
        on_change=reset_dati
    )

    if st.button("Elabora e Normalizza con IA", key="btn_ia"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key: st.error("Chiave API mancante nei Secrets.")
        elif not files_unificati: st.warning("Carica almeno un file per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("L'intelligenza artificiale sta analizzando la struttura logica dei documenti..."):
                    contents = []
                    for file_obj in files_unificati:
                        file_obj.seek(0)
                        estensione = file_obj.name.split('.')[-1].lower()
                        mime = 'application/pdf' if estensione == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if estensione == 'xlsx' else 'text/plain'
                        contents.append(types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime))
                    
                    prompt_sistema = """
                    Sei un data analyst e BIM manager. Il tuo compito è estrarre i dati dagli abachi e accoppiarli rigorosamente alle date del cronoprogramma di Gantt sfruttando i codici WBS univoci.
                    
                    REGOLE OPERATIVE STRETTE:
                    1. Analisi Abachi: Leggi i CSV e trova la colonna WBS (es. Ss_20_20, Ss_30_12_20, ecc.), il Nome del materiale e il totale in Massa (solo numero, es. 13775.92).
                    2. Analisi Cronoprogramma: Leggi il documento Gantt. Troverai le attività del progetto; le attività chiave iniziano proprio con il codice WBS (es. "Ss_20_20_Strutture acciaio e varo..."). 
                    3. Match Deterministico: Abbina il materiale dell'abaco ESATTAMENTE all'attività del cronoprogramma che riporta lo STESSO IDENTICO codice WBS all'inizio del nome.
                    4. Output Date: Per quell'attività abbinata, estrai la Data d'inizio e la Data di fine esatte come segnate nel documento.
                    
                    NON ESEGUIRE CALCOLI SUI GIORNI. Estrai solo i totali assoluti da distribuire e le date di inizio/fine della macro-fase.

                    Restituisci ESCLUSIVAMENTE un blocco di testo in formato CSV puro con le seguenti 5 colonne esatte, separate da virgola:
                    Parametro,Elemento,Quantita_Totale,Data_Inizio,Data_Fine

                    - Parametro: Scrivi "Materiali" (o "Macchinari").
                    - Elemento: Il nome del materiale.
                    - Quantita_Totale: Valore in cifre assolute con punto decimale.
                    - Data_Inizio: Formato AAAA-MM-GG dell'attività abbinata.
                    - Data_Fine: Formato AAAA-MM-GG dell'attività abbinata.

                    Niente markdown, niente chiacchiere. Solo CSV grezzo.
                    """
                    contents.append(prompt_sistema)
                    
                    max_tentativi = 3
                    csv_testo = ""

                    for tentativo in range(max_tentativi):
                        try:
                            response = client.models.generate_content(model='gemini-3.6-flash', contents=contents)
                            csv_testo = response.text.strip()
                            break
                        except Exception as e:
                            errore_str = str(e)
                            if "429" in errore_str or "503" in errore_str or "UNAVAILABLE" in errore_str:
                                st.warning(f"Server IA temporaneamente occupati. Attesa di 20 secondi (Tentativo {tentativo + 1}/{max_tentativi})...")
                                time.sleep(20)
                            else:
                                st.error(f"Errore imprevisto: {e}")
                                break
                    
                    if csv_testo:
                        if csv_testo.startswith("```"):
                            csv_testo = csv_testo.split("```")[1].strip()
                            if csv_testo.startswith("csv"): csv_testo = csv_testo[3:].strip()
                        elif "Parametro,Elemento" in csv_testo and "\n" in csv_testo:
                             csv_testo = csv_testo[csv_testo.find("Parametro,Elemento"):]
                        
                        df_estratti = pd.read_csv(io.StringIO(csv_testo))
                        df_estratti.columns = df_estratti.columns.str.strip()
                        
                        df_estratti['Data_Inizio'] = pd.to_datetime(df_estratti['Data_Inizio'], errors='coerce')
                        df_estratti['Data_Fine'] = pd.to_datetime(df_estratti['Data_Fine'], errors='coerce')
                        df_estratti['Quantita_Totale'] = pd.to_numeric(df_estratti['Quantita_Totale'], errors='coerce').fillna(0)
                        
                        df_estratti = df_estratti.dropna(subset=['Data_Inizio', 'Data_Fine'])
                        
                        righe_distribuite = []
                        for index, row in df_estratti.iterrows():
                            giorni_lavorativi = pd.bdate_range(start=row['Data_Inizio'], end=row['Data_Fine'])
                            num_giorni = len(giorni_lavorativi)
                            
                            if num_giorni < 1: 
                                num_giorni = 1
                                giorni_lavorativi = [row['Data_Inizio']]
                                
                            quantita_giorn = row['Quantita_Totale'] / num_giorni
                            
                            for data_curr in giorni_lavorativi:
                                righe_distribuite.append({
                                    'Data': data_curr.strftime('%Y-%m-%d'),
                                    'Parametro': row['Parametro'],
                                    'Elemento': row['Elemento'],
                                    'Quantita': quantita_giorn
                                })

                        df_cantiere_grezzo = pd.DataFrame(righe_distribuite)

                        df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(
                            lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1
                        )
                        
                        st.session_state['df_cantiere'] = df_cantiere_grezzo
                        st.success("Documenti analizzati, materiali estratti con successo!")
            except Exception as e:
                st.error(f"Errore di configurazione dell'elaborazione: {e}")

    if 'df_cantiere' in st.session_state:
        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 👁️ Anteprima Dati Elaborati dall'IA")
        st.dataframe(st.session_state['df_cantiere'], use_container_width=True)
        
        csv_esportato = st.session_state['df_cantiere'].to_csv(index=False).encode('utf-8')
        st.download_button("📥 Scarica CSV Elaborato", data=csv_esportato, file_name="dataset_cantiere_estratti.csv", mime="text/csv")

with tab2:
    st.markdown("""
    <div style='background-color: #f4f7f3; border: 1.5px solid #d5ddd1; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #111827; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>Guida Operativa: Procedura per l'Elaborazione con file .CSV</h4>
        <ol style='color: #4b5563; font-size: 0.9rem; line-height: 1.6; margin-bottom: 0; padding-left: 20px;'>
        Questo strumento calcola l'impronta di carbonio (espresso in kg di CO₂ equivalente). 
        Il file CSV deve essere strutturato in 4 colonne denominate esattamente:
        <ul style='margin-top: 8px; margin-bottom: 10px; padding-left: 20px;'>
            <li><code>Data</code>: Giorno della lavorazione (AAAA-MM-GG).</li>
            <li><code>Parametro</code>: Macro-categoria tra: <i>Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari</i>.</li>
            <li><code>Elemento</code>: La descrizione specifica della voce.</li>
            <li><code>Quantita</code>: Valore numerico del consumo (es. massa in kg per i materiali).</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    file_cantiere = st.file_uploader("Seleziona file CSV", type=['csv'], label_visibility="collapsed", key="csv_manuale", on_change=reset_dati)
    if file_cantiere:
        if 'df_cantiere' not in st.session_state:
            try:
                file_cantiere.seek(0)
                st.session_state['df_cantiere'] = pd.read_csv(file_cantiere)
                st.success("File CSV caricato correttamente!")
            except Exception as e:
                st.error(f"Errore nella lettura del file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# =====================================================================
# ELABORAZIONE E ANALISI LCA
# =====================================================================
if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere'].copy()
    
    mappa_colonne_finali = {}
    for col in df_cantiere.columns:
        c_low = str(col).strip().lower()
        if 'data' in c_low or 'date' in c_low: mappa_colonne_finali[col] = 'Data'
        elif 'param' in c_low or 'categ' in c_low: mappa_colonne_finali[col] = 'Parametro'
        elif 'elem' in c_low or 'material' in c_low: mappa_colonne_finali[col] = 'Elemento'
        elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low or 'massa' in c_low: mappa_colonne_finali[col] = 'Quantita'
            
    df_cantiere.rename(columns=mappa_colonne_finali, inplace=True)
    
    # -------------------------------------------------------------
    # CALCOLO AUTOMATICO DELLO SFRIDO (15%) SUI MATERIALI
    # -------------------------------------------------------------
    mat_mask = df_cantiere['Parametro'].astype(str).str.strip().str.lower() == 'materiali'
    if mat_mask.any():
        df_sfrido = df_cantiere[mat_mask].copy()
        df_sfrido['Parametro'] = 'Rifiuti'
        df_sfrido['Quantita'] = df_sfrido['Quantita'] * 0.15 
        df_sfrido['Elemento'] = df_sfrido.apply(lambda row: mappa_voce_a_lci(row['Parametro'], row['Elemento']), axis=1)
        df_cantiere = pd.concat([df_cantiere, df_sfrido], ignore_index=True)
    
    df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
    
    df_cantiere['Parametro_match'] = df_cantiere['Parametro'].astype(str).str.strip().str.lower()
    df_cantiere['Elemento_match'] = df_cantiere['Elemento'].astype(str).str.strip().str.lower()
    df_inventario['Parametro_match'] = df_inventario['Parametro'].astype(str).str.strip().str.lower()
    df_inventario['Elemento_match'] = df_inventario['Elemento'].astype(str).str.strip().str.lower()
    
    df_inventario_clean = df_inventario.drop_duplicates(subset=['Parametro_match', 'Elemento_match']).copy()
    
    df_completo = pd.merge(df_cantiere, df_inventario_clean, on=['Parametro_match', 'Elemento_match'], how='left', suffixes=('', '_lci'))
    df_completo['Parametro'] = df_completo['Parametro'].fillna(df_completo['Parametro_match'])
    df_completo['Elemento'] = df_completo['Elemento'].fillna(df_completo['Elemento_match'])
    df_completo.drop(columns=['Parametro_match', 'Elemento_match'], inplace=True, errors='ignore')
    if 'Parametro_lci' in df_completo.columns: df_completo.drop(columns=['Parametro_lci'], inplace=True)
    if 'Elemento_lci' in df_completo.columns: df_completo.drop(columns=['Elemento_lci'], inplace=True)
    
    mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
    if not mancanti.empty:
        st.warning(f"Elementi non riconosciuti nel LCI (calcolati a zero): {mancanti['Elemento'].unique().tolist()}")
        df_completo['Fattore_Emissione'] = df_completo['Fattore_Emissione'].fillna(0)
        
    df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
    
    # --- FILTRO TEMPORALE ---
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Filtro Temporale")
    min_date = df_completo['Data_dt'].min().date() if not df_completo.empty else date.today()
    max_date = df_completo['Data_dt'].max().date() if not df_completo.empty else date.today()

    date_range = st.date_input("Intervallo temporale", value=(min_date, max_date), min_value=min_date, max_value=max_date, label_visibility="collapsed")

    if isinstance(date_range, tuple) and len(date_range) == 2:
        df_filtrato = df_completo.loc[(df_completo['Data_dt'].dt.date >= date_range[0]) & (df_completo['Data_dt'].dt.date <= date_range[1])].copy()
    else: df_filtrato = df_completo.copy()
    st.markdown("</div>", unsafe_allow_html=True)

    # --- VISUALIZZAZIONE GRAFICI ---
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Risultati Analitici (Emissioni Assolute)")
    if df_filtrato.empty: 
        st.warning("Nessuna evidenza registrata.")
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
                    labels={y_col: 'kg CO₂e', 'count': 'Frequenza'},
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
                    y=media_val, line_dash="dash", line_color="#27ae60", 
                    annotation_text=f"Media Fissa: {media_val:.2f}", annotation_position="top right"
                )
                fig.update_layout(title=f"{titolo} - Con Linea Media Fissa nell'Intervallo")
            else:
                fig = px.bar(
                    df_dat, x=x_col, y=y_col, title=titolo,
                    labels={y_col: 'kg CO₂e', x_col: ''}, text_auto='.2f'
                )
                fig.update_traces(marker_color=colore_base)
            
            fig.update_layout(
                height=380, font_family="Inter", 
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                title_font_size=14, title_font_color="#374151"
            )
            return fig

        # 1. Grafico Totale
        df_totale = df_filtrato.groupby('Data')['CO2_Totale_kg'].sum().reset_index()
        fig_tot = genera_figura(df_totale, 'Data', 'CO2_Totale_kg', "Andamento Complessivo Emissioni (kg CO₂e)", "#0B0752")
        st.plotly_chart(fig_tot, use_container_width=True)
        
        # 2. Diagramma a Torta (Incidenza Percentuale)
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        
        # Normalizzazione testuale del Parametro (Prima lettera maiuscola) per il match colori esatto
        df_filtrato['Parametro_Title'] = df_filtrato['Parametro'].astype(str).str.title()
        df_pie = df_filtrato.groupby('Parametro_Title')['CO2_Totale_kg'].sum().reset_index()
        
        colori_parametri = {
            'Materiali': "#B80D0D", 
            'Rifiuti': "#078303", 
            'Trasporti': "#732BB7", 
            'Energia': "#f6de03",
            'Acqua': "#53DCFE", 
            'Macchinari': "#8a929e"
        }
        
        fig_pie = px.pie(
            df_pie, values='CO2_Totale_kg', names='Parametro_Title',
            title="Incidenza Percentuale delle Categorie sulle Emissioni Totali",
            color='Parametro_Title', color_discrete_map=colori_parametri, hole=0.4
        )
        fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
        fig_pie.update_layout(
            height=450, font_family="Inter", showlegend=True, 
            title_font_size=14, title_font_color="#374151",
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
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

        # 3. GRAFICI PARZIALI (Disaggregazione per categoria)
        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        st.subheader("Disaggregazione per Categoria")
        
        parametri_presenti = df_filtrato['Parametro_Title'].dropna().unique()
        col_grafici = st.columns(2)
        
        for idx, param in enumerate(parametri_presenti):
            df_p = df_filtrato[df_filtrato['Parametro_Title'] == param].groupby('Data')['CO2_Totale_kg'].sum().reset_index()
            colore_cat = colori_parametri.get(param, '#4b5563')
            
            fig_cat = genera_figura(df_p, 'Data', 'CO2_Totale_kg', f"{param}", colore_cat)
            
            with col_grafici[idx % 2]:
                st.plotly_chart(fig_cat, use_container_width=True)
                csv_cat = df_p.to_csv(index=False).encode('utf-8')
                st.download_button(f"📥 Scarica {param} (CSV)", data=csv_cat, file_name=f"dati_{param.lower()}.csv", mime="text/csv", key=f"dl_{param}")

        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        with st.expander("Esporta / Visualizza matrice dati completa"):
            st.dataframe(df_filtrato.drop(columns=['Data_dt', 'Parametro_Title', 'Parametro_match', 'Elemento_match'], errors='ignore'), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
