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
st.set_page_config(page_title="EcoSite Tracker | Analisi LCA", layout="wide")

# --- STILE CSS MINIMAL ---
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
        border: 1px solid #e5e7eb !important;
        border-radius: 4px !important;
        padding: 32px !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.02) !important;
        margin-bottom: 24px !important;
    }

    h1 { font-weight: 600; letter-spacing: -0.02em; color: #111827; font-size: 2rem; }
    h2, h3 { font-weight: 500; letter-spacing: -0.01em; color: #374151; }
    h4 { font-weight: 500; color: #4b5563; font-size: 1.1rem; margin-top: 0; }

    .stRadio label, .stNumberInput label, .stDateInput label, .stFileUploader label, .stSelectbox label {
        font-weight: 500; font-size: 0.85rem; color: #4b5563; text-transform: uppercase; letter-spacing: 0.05em;
    }
    
    .stRadio div[role="radiogroup"] {
        background-color: #ffffff; padding: 16px; border-radius: 4px; border: 1px solid #e5e7eb;
    }

    .custom-dl-btn {
        text-decoration: none !important; background-color: #ffffff !important;
        border: 1px solid #d1d5db !important; color: #374151 !important;
        padding: 0.5rem 1rem !important; border-radius: 4px !important;
        font-size: 0.85rem !important; font-weight: 500 !important;
        display: block !important; text-align: center !important;
        width: 100% !important; margin-top: 8px !important;
        transition: all 0.2s ease-in-out !important;
    }
    
    .custom-dl-btn:hover {
        border-color: #9ca3af !important; background-color: #f9fafb !important;
        color: #111827 !important;
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
    st.error(f"Errore di sistema: Il database LCI '{percorso_lci}' non è reperibile o è corrotto.")
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
    "Seleziona la fase operativa dell'analisi:",
    ["Fase di Progettazione (Stima Predittiva)", "Fase di Esecuzione (Monitoraggio Dinamico)"],
    on_change=reset_dati
)

if "Progettazione" in fase_cantiere:
    st.markdown("""
    <div style='background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 4px; padding: 20px; margin-bottom: 20px;'>
        <h4>Analisi Predittiva</h4>
        <p style='color: #4b5563; font-size: 0.9rem; margin-bottom: 10px; line-height: 1.5;'>
        In assenza di telemetria, il sistema modella l'impronta carbonica distribuendo i materiali estratti dal modello BIM sulle finestre temporali del cronoprogramma. Compila i consumi operativi medi previsti per l'accantieramento. I rifiuti verranno calcolati automaticamente al 15% di sfrido, e i trasporti stimati su 30 km.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        stima_macchinari = st.number_input("Macchinari (kg Diesel/giorno)", value=45.0, min_value=0.0, step=5.0)
    with col2:
        stima_energia = st.number_input("Energia Elettrica (kWh/giorno)", value=120.0, min_value=0.0, step=10.0)
    with col3:
        stima_acqua = st.number_input("Acqua (Litri/giorno)", value=500.0, min_value=0.0, step=50.0)
    
    files_unificati = st.file_uploader("Carica abachi materiali e cronoprogramma (PDF, CSV, XLSX)", type=['pdf', 'txt', 'xlsx', 'csv', 'ifc'], accept_multiple_files=True, key="ia_unified")

    if st.button("Elabora Modello Predittivo"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key: st.error("Errore di configurazione: Chiave API mancante.")
        elif not files_unificati: st.warning("È necessario caricare la documentazione di cantiere per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("Estrazione semantica in corso..."):
                    contents = []
                    for file_obj in files_unificati:
                        file_obj.seek(0)
                        estensione = file_obj.name.split('.')[-1].lower()
                        mime = 'application/pdf' if estensione == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if estensione == 'xlsx' else 'text/plain'
                        contents.append(types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime))
                    
                    prompt_sistema = """
                    Analizza gli abachi e il cronoprogramma. Identifica per ogni voce WBS il materiale, la massa totale (in kg) e le date di inizio e fine dell'attività associata.
                    Restituisci unicamente un file CSV strutturato esattamente in questo modo:
                    Parametro,Elemento,Quantita_Totale,Data_Inizio,Data_Fine
                    
                    Regole: 
                    - 'Parametro' deve sempre riportare la dicitura 'Materiali'.
                    - Nessun testo aggiuntivo, unicamente il blocco CSV.
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
                    
                    if not csv_testo: st.error("I server di analisi sono sovraccarichi. Riprova più tardi.")
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
                                
                                # Inserimento Materiali
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Materiali', 'Elemento': row['Elemento'], 'Quantita': quantita_giorn})
                                
                                # Calcolo Rifiuti e Trasporti
                                q_rif = quantita_giorn * 0.15
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Rifiuti', 'Elemento': 'Inerti / Macerie di demolizione', 'Quantita': q_rif})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (quantita_giorn/28000)*30*0.25})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_rif/28000)*30*0.25})

                        # Inserimento Consumi Fissi
                        for d in giorni_unici:
                            righe_arricchite.append({'Data': d, 'Parametro': 'Macchinari', 'Elemento': 'Diesel', 'Quantita': stima_macchinari})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': stima_energia})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': stima_acqua})

                        df_cantiere_grezzo = pd.DataFrame(righe_arricchite)
                        df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
                        st.session_state['df_cantiere'] = df_cantiere_grezzo
                        st.success("Modello LCA predittivo generato con successo.")
            except Exception as e: st.error(f"Errore di esecuzione: {e}")

elif "Esecuzione" in fase_cantiere:
    st.markdown("""
    <div style='background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 4px; padding: 20px; margin-bottom: 20px;'>
        <h4>Integrazione Telemetria e Monitoraggio Dinamico</h4>
        <p style='color: #4b5563; font-size: 0.9rem; margin-bottom: 0; line-height: 1.5;'>
        Carica nei rispettivi moduli l'output strutturato dei materiali estratti dal modello BIM e i file di log generati dalla sensoristica di cantiere. Il sistema mapperà le grandezze fisiche (kWh, Litri, Km) normalizzando il tracciato per la conformità con il database LCI, creando un digital twin ambientale accurato.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        file_mat = st.file_uploader("Dataset Materiali BIM (CSV singolo)", type=['csv'], help="Formato richiesto: Data, Parametro, Elemento, Quantita")
    with col_file2:
        files_sensori = st.file_uploader("Log Sensori IoT (CSV multipli)", type=['csv'], accept_multiple_files=True, help="Supporta log da Smart Meter, Flussimetri e Moduli GPS veicolari")
        
    if st.button("Elabora e Unifica Modello Dinamico"):
        righe_reali = []
        
        # Elaborazione Modulo Materiali
        if file_mat:
            try:
                file_mat.seek(0)
                df_m = pd.read_csv(file_mat)
                mappa_colonne = {}
                for col in df_m.columns:
                    c_low = str(col).strip().lower()
                    if 'data' in c_low or 'date' in c_low: mappa_colonne['Data'] = col
                    elif 'param' in c_low or 'categ' in c_low: mappa_colonne['Parametro'] = col
                    elif 'elem' in c_low or 'material' in c_low: mappa_colonne['Elemento'] = col
                    elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low or 'massa' in c_low: mappa_colonne['Quantita'] = col
                
                if 'Data' in mappa_colonne and 'Elemento' in mappa_colonne and 'Quantita' in mappa_colonne:
                    param_col_name = mappa_colonne.get('Parametro', None)
                    for _, r in df_m.iterrows():
                        param_val = r[param_col_name] if param_col_name else 'Materiali'
                        righe_reali.append({
                            'Data': str(r[mappa_colonne['Data']])[:10],
                            'Parametro': param_val,
                            'Elemento': r[mappa_colonne['Elemento']],
                            'Quantita': r[mappa_colonne['Quantita']]
                        })
            except Exception as e:
                st.error(f"Errore nella lettura del Dataset Materiali: {e}")

        # Elaborazione Modulo Sensori
        if files_sensori:
            for f in files_sensori:
                try:
                    f.seek(0)
                    df_s = pd.read_csv(f)
                    col_lower = [c.lower() for c in df_s.columns]
                    
                    date_col = next((c for c in df_s.columns if 'date' in c.lower() or 'time' in c.lower() or 'data' in c.lower()), df_s.columns[0])
                    
                    # Riconoscimento log Energia Elettrica (kWh)
                    if any('kwh' in c for c in col_lower) or any('energy' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'kwh' in c.lower() or 'energy' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': r[val_col]})
                            
                    # Riconoscimento log Acqua (Litri)
                    elif any('lit' in c for c in col_lower) or any('vol' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'lit' in c.lower() or 'vol' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': r[val_col]})
                            
                    # Riconoscimento log Trasporti (Km)
                    elif any('dist' in c for c in col_lower) or any('km' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'dist' in c.lower() or 'km' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            # Conversione dinamica distanza in consumo carburante (0.25 kg diesel al km)
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': r[val_col] * 0.25})
                            
                except Exception as e:
                    st.error(f"Errore nella lettura del file {f.name}: {e}")

        if righe_reali:
            df_cantiere_reale = pd.DataFrame(righe_reali)
            # Normalizzazione formati per compatibilità LCI
            df_cantiere_reale['Quantita'] = pd.to_numeric(df_cantiere_reale['Quantita'], errors='coerce').fillna(0)
            df_cantiere_reale = df_cantiere_reale.groupby(['Data', 'Parametro', 'Elemento'], as_index=False)['Quantita'].sum()
            df_cantiere_reale['Elemento'] = df_cantiere_reale.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
            
            st.session_state['df_cantiere'] = df_cantiere_reale
            st.success("Analisi telemetrica e normalizzazione dei dati completata con successo.")
        else:
            st.warning("Eseguire il caricamento di almeno un modulo (Materiali o Sensori) per generare il calcolo.")

st.markdown("</div>", unsafe_allow_html=True)

# =====================================================================
# ELABORAZIONE E ANALISI LCA (MODULO COMUNE)
# =====================================================================
if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere'].copy()
    
    # Generazione automatica rifiuti se assenti nel caricamento dinamico
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
    
    # Processo di incrocio dati (Matching) con il database LCI
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
    st.subheader("Configurazione Temporale")
    
    min_date = df_completo['Data_dt'].min().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].min()) else date.today()
    max_date = df_completo['Data_dt'].max().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].max()) else date.today()

    date_range = st.date_input("Selezione intervallo di calcolo:", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")

    if isinstance(date_range, tuple) and len(date_range) == 2:
        df_filtrato = df_completo.loc[(df_completo['Data_dt'].dt.date >= date_range[0]) & (df_completo['Data_dt'].dt.date <= date_range[1])].copy()
    else: 
        df_filtrato = df_completo.copy()
    st.markdown("</div>", unsafe_allow_html=True)

    # --- VISUALIZZAZIONE GRAFICI ---
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Risultati Analitici")
    if df_filtrato.empty: 
        st.info("La selezione temporale non presenta dati elaborabili.")
    else:
        modo_visualizzazione = st.selectbox("Modalità di visualizzazione dati", ["Standard (Distribuzione Temporale)", "Analisi di Frequenza (KDE)", "Andamento (Media Mobile)", "Andamento (Media Fissa)"], index=0)

        def genera_figura(df_dat, x_col, y_col, titolo, colore_base):
            df_dat = df_dat.copy()
            if df_dat.empty: return go.Figure()
            if "Frequenza" in modo_visualizzazione:
                fig = px.histogram(df_dat, x=y_col, nbins=20, marginal="violin", title=f"{titolo} - Analisi di Frequenza", labels={y_col: 'kg CO₂e'}, color_discrete_sequence=[colore_base])
            elif "Media Mobile" in modo_visualizzazione:
                df_dat = df_dat.sort_values(by=x_col)
                df_dat['Media_Mobile'] = df_dat[y_col].rolling(window=7, min_periods=1).mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Dato Giornaliero', marker_color=colore_base, opacity=0.4))
                fig.add_trace(go.Scatter(x=df_dat[x_col], y=df_dat['Media_Mobile'], mode='lines', name='Media Mobile', line=dict(color='#111827', width=2)))
                fig.update_layout(title=f"{titolo} - Variazione su Media Mobile")
            elif "Media Fissa" in modo_visualizzazione:
                media_val = df_dat[y_col].mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Dato Giornaliero', marker_color=colore_base, opacity=0.5))
                fig.add_hline(y=media_val, line_dash="dash", line_color="#4b5563", annotation_text=f"Media: {media_val:.2f}")
                fig.update_layout(title=f"{titolo} - Riferimento su Media Fissa")
            else:
                fig = px.bar(df_dat, x=x_col, y=y_col, title=titolo, labels={y_col: 'kg CO₂e', x_col: ''})
                fig.update_traces(marker_color=colore_base)
            
            fig.update_layout(height=380, font_family="Inter", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', title_font_size=13, title_font_color="#374151")
            return fig

        df_totale = df_filtrato.groupby('Data')['CO2_Totale_kg'].sum().reset_index()
        st.plotly_chart(genera_figura(df_totale, 'Data', 'CO2_Totale_kg', "Andamento Complessivo Impronta Carbonica", "#111827"), use_container_width=True)
        
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        df_filtrato['Parametro_Title'] = df_filtrato['Parametro'].astype(str).str.title()
        df_pie = df_filtrato.groupby('Parametro_Title')['CO2_Totale_kg'].sum().reset_index()
        df_pie = df_pie[df_pie['CO2_Totale_kg'] > 0]
        
        colori_parametri = {'Materiali': "#991b1b", 'Rifiuti': "#166534", 'Trasporti': "#4338ca", 'Energia': "#ca8a04", 'Acqua': "#0369a1", 'Macchinari': "#4b5563"}
        if not df_pie.empty:
            fig_pie = px.pie(df_pie, values='CO2_Totale_kg', names='Parametro_Title', title="Incidenza Strutturale per Categorie Operative", color='Parametro_Title', color_discrete_map=colori_parametri, hole=0.4)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
            fig_pie.update_layout(height=450, font_family="Inter", showlegend=True, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', title_font_size=13, title_font_color="#374151")
            st.plotly_chart(fig_pie, use_container_width=True)

        col_exp1, col_exp2 = st.columns(2)
        with col_exp1: st.download_button("Esporta Dati Totali (Formato CSV)", data=df_totale.to_csv(index=False).encode('utf-8'), file_name="emissioni_totali.csv", mime="text/csv")
        with col_exp2:
            output_xlsx = io.BytesIO()
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer: df_totale.to_excel(writer, index=False, sheet_name='Totale')
            st.download_button("Esporta Dati Totali (Formato XLSX)", data=output_xlsx.getvalue(), file_name="emissioni_totali.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        st.subheader("Analisi Disaggregata per Voci")
        col_grafici = st.columns(2)
        for idx, param in enumerate(df_filtrato['Parametro_Title'].dropna().unique()):
            df_p = df_filtrato[df_filtrato['Parametro_Title'] == param].groupby('Data')['CO2_Totale_kg'].sum().reset_index()
            with col_grafici[idx % 2]:
                st.plotly_chart(genera_figura(df_p, 'Data', 'CO2_Totale_kg', param, colori_parametri.get(param, '#4b5563')), use_container_width=True)

        with st.expander("Ispezione Matrice Dati"):
            st.dataframe(df_filtrato.drop(columns=['Data_dt', 'Parametro_Title', 'Parametro_match', 'Elemento_match'], errors='ignore'), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
