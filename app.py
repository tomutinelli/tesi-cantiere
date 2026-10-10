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
                            # Spalma calcolo SOLO nei giorni feriali (Business Days: Lunedì - Venerdì)
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
                        st.success("Documenti analizzati, materiali (Massa) e macchinari estratti con successo!")
            except Exception as e:
                st.error(f"Errore di configurazione dell'elaborazione: {e}")

    # Anteprima e Download
    if 'df_cantiere' in st.session_state:
        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 👁️ Anteprima Dati Elaborati dall'IA")
        st.dataframe(st.session_state['df_cantiere'], use_container_width=True)
        
        csv_esportato = st.session_state['df_cantiere'].to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Scarica CSV Elaborato dall'IA", 
            data=csv_esportato, 
            file_name="dataset_cantiere_estratti.csv", 
            mime="text/csv",
            key="download_csv_ia"
        )

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
    
    # SFRIDO 15% AUTOMATICO
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
    
    df_completo = pd.merge(df_cantiere, df_inventario_clean, on=['Parametro_match', 'Elemento_match'], how='left')
    df_completo['Parametro'] = df_completo['Parametro'].fillna(df_completo['Parametro_match'])
    df_completo['Elemento'] = df_completo['Elemento'].fillna(df_completo['Elemento_match'])
    
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

    # --- GRAFICI ---
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Risultati Analitici (Emissioni Assolute)")
    if df_filtrato.empty: st.warning("Nessuna evidenza registrata.")
    else:
        df_totale = df_filtrato.groupby('Data')['CO2_Totale_kg'].sum().reset_index()
        fig_tot = px.bar(df_totale, x='Data', y='CO2_Totale_kg', title="Andamento Complessivo (kg CO₂e)")
        fig_tot.update_traces(marker_color="#0B0752")
        fig_tot.update_layout(height=380, font_family="Inter", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_tot, use_container_width=True)
        
        df_pie = df_filtrato.groupby('Parametro')['CO2_Totale_kg'].sum().reset_index()
        colori = {'Materiali': "#B80D0D", 'Rifiuti': "#078303", 'Trasporti': "#732BB7", 'Energia': "#f6de03", 'Acqua': "#53DCFE", 'Macchinari': "#8a929e"}
        fig_pie = px.pie(df_pie, values='CO2_Totale_kg', names='Parametro', title="Incidenza Percentuale", color='Parametro', color_discrete_map=colori, hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
        with st.expander("Esporta / Visualizza matrice dati completa"):
            st.dataframe(df_filtrato.drop(columns=['Data_dt', 'Parametro_match', 'Elemento_match'], errors='ignore'), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
