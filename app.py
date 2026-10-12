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
    
    .stRadio div[role="radiogroup"] {
        background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1.5px solid #d5ddd1;
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


# --- BLOCCO INPUT DATI E SCELTA FASE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Configurazione Cantiere e Inserimento Dati")

def reset_dati():
    if 'df_cantiere' in st.session_state: del st.session_state['df_cantiere']

fase_cantiere = st.radio(
    "Seleziona la fase del ciclo di vita del progetto:",
    ["🏗️ Fase di Progettazione (Stima Predittiva IA)", "📡 Fase di Esecuzione (Monitoraggio Dinamico IoT)"],
    on_change=reset_dati
)

if "Progettazione" in fase_cantiere:
    st.markdown("""
    <div style='background-color: #f4f7f3; border: 1.5px solid #d5ddd1; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #111827; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>Stima Predittiva</h4>
        <p style='color: #4b5563; font-size: 0.9rem; margin-bottom: 10px;'>In assenza di sensori, imposta i consumi medi giornalieri previsti. L'IA aggancerà questi dati ai giorni lavorativi estratti da abachi e cronoprogramma.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        stima_macchinari = st.number_input("Macchinari (kg Diesel/giorno)", value=45.0, min_value=0.0, step=5.0)
    with col2:
        stima_energia = st.number_input("Energia Elettrica (kWh/giorno)", value=120.0, min_value=0.0, step=10.0)
    with col3:
        stima_acqua = st.number_input("Acqua (Litri/giorno)", value=500.0, min_value=0.0, step=50.0)
    
    files_unificati = st.file_uploader("Carica abachi materiali e cronoprogramma (PDF/CSV)", type=['pdf', 'txt', 'xlsx', 'csv', 'ifc'], accept_multiple_files=True, key="ia_unified")

    if st.button("Avvia Analisi Predittiva con IA"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key: st.error("Chiave API mancante nei Secrets.")
        elif not files_unificati: st.warning("Carica almeno un file per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("L'IA sta estraendo i dati e applicando le tue stime..."):
                    contents = []
                    for file_obj in files_unificati:
                        file_obj.seek(0)
                        estensione = file_obj.name.split('.')[-1].lower()
                        mime = 'application/pdf' if estensione == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if estensione == 'xlsx' else 'text/plain'
                        contents.append(types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime))
                    
                    prompt_sistema = """
                    Sei un BIM manager. Leggi gli abachi e individua WBS, materiale e massa.
                    Leggi il Gantt per trovare data inizio e fine di quella WBS.
                    Restituisci SOLO un CSV con colonne: Parametro,Elemento,Quantita_Totale,Data_Inizio,Data_Fine
                    Parametro deve essere sempre "Materiali".
                    """
                    contents.append(prompt_sistema)
                    
                    csv_testo = ""
                    for tentativo in range(5):
                        try:
                            response = client.models.generate_content(model='gemini-3.6-flash', contents=contents)
                            csv_testo = response.text.strip()
                            break
                        except Exception as e:
                            if any(err in str(e) for err in ["429", "503", "UNAVAILABLE"]): time.sleep(10)
                            else: break
                    
                    if not csv_testo: st.error("Errore IA. Riprova.")
                    else:
                        if csv_testo.startswith("```"): csv_testo = csv_testo.split("```")[1].replace("csv\n", "").strip()
                        elif "Parametro,Elemento" in csv_testo: csv_testo = csv_testo[csv_testo.find("Parametro,Elemento"):]
                        
                        df_estratti = pd.read_csv(io.StringIO(csv_testo))
                        df_estratti.columns = df_estratti.columns.str.strip()
                        df_estratti['Data_Inizio'] = pd.to_datetime(df_estratti['Data_Inizio'], errors='coerce')
                        df_estratti['Data_Fine'] = pd.to_datetime(df_estratti['Data_Fine'], errors='coerce')
                        df_estratti['Quantita_Totale'] = pd.to_numeric(df_estratti['Quantita_Totale'], errors='coerce').fillna(0)
                        df_estratti = df_estratti.dropna(subset=['Data_Inizio', 'Data_Fine'])
                        
                        righe_arricchite = []
                        giorni_unici = set()
                        
                        for index, row in df_estratti.iterrows():
                            giorni_lavorativi = pd.bdate_range(start=row['Data_Inizio'], end=row['Data_Fine'])
                            if len(giorni_lavorativi) < 1: giorni_lavorativi = [row['Data_Inizio']]
                            quantita_giorn = row['Quantita_Totale'] / len(giorni_lavorativi)
                            
                            for data_curr in giorni_lavorativi:
                                data_str = data_curr.strftime('%Y-%m-%d')
                                giorni_unici.add(data_str)
                                righe_arricchite.append({'Data': data_str, 'Parametro': row['Parametro'], 'Elemento': row['Elemento'], 'Quantita': quantita_giorn})
                                
                                q_rif = quantita_giorn * 0.15
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Rifiuti', 'Elemento': 'Inerti / Macerie di demolizione', 'Quantita': q_rif})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (quantita_giorn/28000)*30*0.25})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_rif/28000)*30*0.25})

                        for d in giorni_unici:
                            righe_arricchite.append({'Data': d, 'Parametro': 'Macchinari', 'Elemento': 'Diesel', 'Quantita': stima_macchinari})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': stima_energia})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': stima_acqua})

                        df_cantiere_grezzo = pd.DataFrame(righe_arricchite)
                        df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
                        st.session_state['df_cantiere'] = df_cantiere_grezzo
                        st.success("Modello LCA predittivo generato con le tue stime!")
            except Exception as e: st.error(f"Errore: {e}")

elif "Esecuzione" in fase_cantiere:
    st.markdown("""
    <div style='background-color: #eef2ff; border: 1.5px solid #c7d2fe; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #1e3a8a; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>Integrazione Telemetria IoT</h4>
        <p style='color: #3730a3; font-size: 0.9rem; margin-bottom: 0;'>Carica i log generati dai sensori (GPS, Smart Meter, Flussimetri) e il CSV base dei materiali. Il sistema unificherà i dati reali in un unico Digital Twin LCA.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        file_mat = st.file_uploader("1. Dataset Materiali Base (CSV)", type=['csv'], help="Formato standard: Data, Parametro, Elemento, Quantita")
        file_gps = st.file_uploader("2. Log Trasporti GPS (CSV)", type=['csv'], help="Riconosce colonna 'Distance_km' o 'Distanza'")
    with col_file2:
        file_en = st.file_uploader("3. Log Smart Meter Energia (CSV)", type=['csv'], help="Riconosce colonna 'ActiveEnergy_kWh' o 'kWh'")
        file_acq = st.file_uploader("4. Log Flussimetro Acqua (CSV)", type=['csv'], help="Riconosce colonna 'TotalVolume_Liters' o 'Litri'")
        
    if st.button("Elabora e Unifica Dati IoT"):
        righe_reali = []
        
        # Elaborazione Materiali
        if file_mat:
            file_mat.seek(0)
            df_m = pd.read_csv(file_mat)
            for col in df_m.columns:
                c_low = str(col).strip().lower()
                if 'data' in c_low or 'date' in c_low: col_data = col
                elif 'param' in c_low or 'categ' in c_low: col_param = col
                elif 'elem' in c_low or 'material' in c_low: col_elem = col
                elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low or 'massa' in c_low: col_quant = col
            
            for _, r in df_m.iterrows():
                righe_reali.append({'Data': str(r[col_data])[:10], 'Parametro': r[col_param], 'Elemento': r[col_elem], 'Quantita': r[col_quant]})
        
        # Elaborazione Trasporti GPS
        if file_gps:
            file_gps.seek(0)
            df_g = pd.read_csv(file_gps)
            dist_col = next((c for c in df_g.columns if 'dist' in c.lower() or 'km' in c.lower()), None)
            date_col = next((c for c in df_g.columns if 'dat' in c.lower()), df_g.columns[0])
            if dist_col:
                df_g_agg = df_g.groupby(df_g[date_col].astype(str).str[:10])[dist_col].sum().reset_index()
                for _, r in df_g_agg.iterrows():
                    righe_reali.append({'Data': r[date_col], 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': r[dist_col] * 0.25}) # 0.25 kg/km
                    
        # Elaborazione Energia
        if file_en:
            file_en.seek(0)
            df_e = pd.read_csv(file_en)
            kwh_col = next((c for c in df_e.columns if 'kwh' in c.lower() or 'energ' in c.lower()), None)
            date_col = next((c for c in df_e.columns if 'time' in c.lower() or 'dat' in c.lower()), df_e.columns[0])
            if kwh_col:
                df_e_agg = df_e.groupby(df_e[date_col].astype(str).str[:10])[kwh_col].sum().reset_index()
                for _, r in df_e_agg.iterrows():
                    righe_reali.append({'Data': r[date_col], 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': r[kwh_col]})

        # Elaborazione Acqua
        if file_acq:
            file_acq.seek(0)
            df_a = pd.read_csv(file_acq)
            lit_col = next((c for c in df_a.columns if 'lit' in c.lower() or 'vol' in c.lower()), None)
            date_col = next((c for c in df_a.columns if 'time' in c.lower() or 'dat' in c.lower()), df_a.columns[0])
            if lit_col:
                df_a_agg = df_a.groupby(df_a[date_col].astype(str).str[:10])[lit_col].sum().reset_index()
                for _, r in df_a_agg.iterrows():
                    righe_reali.append({'Data': r[date_col], 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': r[lit_col]})

        if righe_reali:
            df_cantiere_reale = pd.DataFrame(righe_reali)
            # Raggruppa eventuali duplicati dello stesso giorno
            df_cantiere_reale = df_cantiere_reale.groupby(['Data', 'Parametro', 'Elemento'], as_index=False)['Quantita'].sum()
            df_cantiere_reale['Elemento'] = df_cantiere_reale.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
            st.session_state['df_cantiere'] = df_cantiere_reale
            st.success("Dati dei sensori integrati con successo!")
        else:
            st.warning("Carica almeno un file dati per procedere.")

st.markdown("</div>", unsafe_allow_html=True)

# =====================================================================
# ELABORAZIONE E ANALISI LCA (COMUNE A ENTRAMBE LE FASI)
# =====================================================================
if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere'].copy()
    
    rifiuti_presenti = (df_cantiere['Parametro'].astype(str).str.strip().str.lower() == 'rifiuti').any()
    if not rifiuti_presenti:
        mat_mask = df_cantiere['Parametro'].astype(str).str.strip().str.lower() == 'materiali'
        if mat_mask.any():
            df_sfrido = df_cantiere[mat_mask].copy()
            df_sfrido['Parametro'] = 'Rifiuti'
            df_sfrido['Quantita'] = df_sfrido['Quantita'] * 0.15 
            df_sfrido['Elemento'] = 'Inerti / Macerie di demolizione'
            df_cantiere = pd.concat([df_cantiere, df_sfrido], ignore_index=True)
    
    df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='mixed', errors='coerce')
    
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
    
    df_completo['Fattore_Emissione'] = df_completo['Fattore_Emissione'].fillna(0)
    df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
    
    # -------------------------------------------------------------
    # FILTRO TEMPORALE
    # -------------------------------------------------------------
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Filtro Temporale")
    
    min_date = df_completo['Data_dt'].min().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].min()) else date.today()
    max_date = df_completo['Data_dt'].max().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].max()) else date.today()

    date_range = st.date_input("Seleziona l'intervallo temporale:", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")

    if isinstance(date_range, tuple) and len(date_range) == 2:
        df_filtrato = df_completo.loc[(df_completo['Data_dt'].dt.date >= date_range[0]) & (df_completo['Data_dt'].dt.date <= date_range[1])].copy()
    else: 
        df_filtrato = df_completo.copy()
    st.markdown("</div>", unsafe_allow_html=True)

    # --- VISUALIZZAZIONE GRAFICI ---
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Risultati Analitici (Emissioni Assolute)")
    if df_filtrato.empty: 
        st.warning("Nessuna evidenza registrata per il periodo selezionato.")
    else:
        modo_visualizzazione = st.selectbox("Modalità di visualizzazione grafica", ["Standard (Barre temporali)", "Distribuzione (KDE / Frequenza)", "Linea Media (Andamento)", "Linea Media Fissa"], index=0)

        def genera_figura(df_dat, x_col, y_col, titolo, colore_base):
            df_dat = df_dat.copy()
            if df_dat.empty: return go.Figure()
            if "Distribuzione" in modo_visualizzazione:
                fig = px.histogram(df_dat, x=y_col, nbins=20, marginal="violin", title=f"{titolo} - Distribuzione Frazionale", labels={y_col: 'kg CO₂e'}, color_discrete_sequence=[colore_base])
            elif "Andamento" in modo_visualizzazione:
                df_dat = df_dat.sort_values(by=x_col)
                df_dat['Media_Mobile'] = df_dat[y_col].rolling(window=7, min_periods=1).mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Giornaliero', marker_color=colore_base, opacity=0.4))
                fig.add_trace(go.Scatter(x=df_dat[x_col], y=df_dat['Media_Mobile'], mode='lines', name='Media (7gg)', line=dict(color='#111827', width=3)))
                fig.update_layout(title=f"{titolo} - Andamento Media Mobile")
            elif "Fissa" in modo_visualizzazione:
                media_val = df_dat[y_col].mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Giornaliero', marker_color=colore_base, opacity=0.5))
                fig.add_hline(y=media_val, line_dash="dash", line_color="#27ae60", annotation_text=f"Media: {media_val:.2f}")
                fig.update_layout(title=f"{titolo} - Linea Media Fissa")
            else:
                fig = px.bar(df_dat, x=x_col, y=y_col, title=titolo, labels={y_col: 'kg CO₂e', x_col: ''})
                fig.update_traces(marker_color=colore_base)
            
            fig.update_layout(height=380, font_family="Inter", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', title_font_size=14)
            return fig

        df_totale = df_filtrato.groupby('Data')['CO2_Totale_kg'].sum().reset_index()
        st.plotly_chart(genera_figura(df_totale, 'Data', 'CO2_Totale_kg', "Andamento Complessivo Emissioni", "#0B0752"), use_container_width=True)
        
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        df_filtrato['Parametro_Title'] = df_filtrato['Parametro'].astype(str).str.title()
        df_pie = df_filtrato.groupby('Parametro_Title')['CO2_Totale_kg'].sum().reset_index()
        df_pie = df_pie[df_pie['CO2_Totale_kg'] > 0]
        
        colori_parametri = {'Materiali': "#B80D0D", 'Rifiuti': "#078303", 'Trasporti': "#732BB7", 'Energia': "#f6de03", 'Acqua': "#53DCFE", 'Macchinari': "#8a929e"}
        if not df_pie.empty:
            fig_pie = px.pie(df_pie, values='CO2_Totale_kg', names='Parametro_Title', title="Incidenza Percentuale Categorie", color='Parametro_Title', color_discrete_map=colori_parametri, hole=0.4)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
            fig_pie.update_layout(height=450, font_family="Inter", showlegend=True, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_pie, use_container_width=True)

        col_exp1, col_exp2 = st.columns(2)
        with col_exp1: st.download_button("📥 Scarica dati totali (CSV)", data=df_totale.to_csv(index=False).encode('utf-8'), file_name="emissioni_totali.csv", mime="text/csv")
        with col_exp2:
            output_xlsx = io.BytesIO()
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer: df_totale.to_excel(writer, index=False, sheet_name='Totale')
            st.download_button("📥 Scarica dati totali (XLSX)", data=output_xlsx.getvalue(), file_name="emissioni_totali.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        st.subheader("Disaggregazione per Categoria")
        col_grafici = st.columns(2)
        for idx, param in enumerate(df_filtrato['Parametro_Title'].dropna().unique()):
            df_p = df_filtrato[df_filtrato['Parametro_Title'] == param].groupby('Data')['CO2_Totale_kg'].sum().reset_index()
            with col_grafici[idx % 2]:
                st.plotly_chart(genera_figura(df_p, 'Data', 'CO2_Totale_kg', param, colori_parametri.get(param, '#4b5563')), use_container_width=True)

        with st.expander("Esporta Matrice Dati Completa"):
            st.dataframe(df_filtrato.drop(columns=['Data_dt', 'Parametro_Title', 'Parametro_match', 'Elemento_match'], errors='ignore'), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
