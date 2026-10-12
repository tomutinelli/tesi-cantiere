import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io

st.set_page_config(page_title="EcoSite Tracker | Analisi LCA", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1f2937;
        background-color: #f9fafb;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    div.minimal-card {
        background-color: #ffffff !important;
        border: 1px solid #e5e7eb !important;
        border-radius: 6px !important;
        padding: 32px !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.03) !important;
        margin-bottom: 24px !important;
    }

    h1 { font-weight: 600; letter-spacing: -0.02em; color: #111827; font-size: 1.75rem; }
    h2, h3 { font-weight: 500; letter-spacing: -0.01em; color: #374151; }
    h4 { font-weight: 500; color: #4b5563; font-size: 1.05rem; margin-top: 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

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

def mappa_voce_a_lci(parametro, elemento_grezzo, df_inv=df_inventario):
    p_str = str(parametro).strip()
    e_str = str(elemento_grezzo).strip().lower()
    
    if df_inv is None: return elemento_grezzo
        
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
        
    return elemento_grezzo

# --- NUOVO ALGORITMO DETERMINISTICO WBS ---
def elabora_materiali_wbs_pandas(lista_file):
    df_crono = pd.DataFrame()
    df_abachi = []
    
    for f in lista_file:
        f.seek(0)
        ext = f.name.split('.')[-1].lower()
        try:
            if ext == 'csv': df = pd.read_csv(f, sep=None, engine='python')
            elif ext in ['xlsx', 'xls']: df = pd.read_excel(f)
            else: continue # PDF saltati
            
            cols = [str(c).lower() for c in df.columns]
            
            # Se ha date di inizio/fine è il cronoprogramma
            if any('iniz' in c or 'start' in c for c in cols) and any('fin' in c or 'end' in c for c in cols):
                df_crono = df
            else:
                df_abachi.append(df)
        except Exception:
            pass
            
    if df_crono.empty or not df_abachi:
        return pd.DataFrame() 
        
    # Standardizza il Cronoprogramma
    crono_wbs_col = next((c for c in df_crono.columns if any(k in str(c).lower() for k in ['wbs', 'codice', 'code', 'task'])), None)
    crono_start_col = next((c for c in df_crono.columns if 'iniz' in str(c).lower() or 'start' in str(c).lower()), None)
    crono_end_col = next((c for c in df_crono.columns if 'fin' in str(c).lower() or 'end' in str(c).lower()), None)
    
    if not all([crono_wbs_col, crono_start_col, crono_end_col]):
        return pd.DataFrame()
        
    df_crono = df_crono[[crono_wbs_col, crono_start_col, crono_end_col]].copy()
    df_crono.columns = ['WBS', 'Data_Inizio', 'Data_Fine']
    df_crono['WBS'] = df_crono['WBS'].astype(str).str.strip()
    df_crono['Data_Inizio'] = pd.to_datetime(df_crono['Data_Inizio'], errors='coerce')
    df_crono['Data_Fine'] = pd.to_datetime(df_crono['Data_Fine'], errors='coerce')
    df_crono.dropna(subset=['Data_Inizio', 'Data_Fine'], inplace=True)

    # Standardizza gli Abachi e fa il merge
    risultati = []
    for df_a in df_abachi:
        wbs_col = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['wbs', 'codice', 'code'])), None)
        mat_col = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['material', 'famiglia', 'nome', 'descrizione'])), None)
        q_col = next((c for c in df_a.columns if any(k in str(c).lower() for k in ['quantit', 'volum', 'massa', 'peso', 'kg'])), None)
        
        if not all([wbs_col, mat_col, q_col]): continue
            
        df_a['WBS'] = df_a[wbs_col].astype(str).str.strip()
        merged = pd.merge(df_a, df_crono, on='WBS', how='inner')
        
        for _, row in merged.iterrows():
            val = pd.to_numeric(row[q_col], errors='coerce')
            if pd.notna(val) and val > 0:
                risultati.append({
                    'Parametro': 'Materiali',
                    'Elemento': str(row[mat_col]),
                    'Quantita_Totale': val,
                    'Data_Inizio': row['Data_Inizio'],
                    'Data_Fine': row['Data_Fine']
                })
                
    return pd.DataFrame(risultati)

# --- NUOVO ALGORITMO DETERMINISTICO SENSORI ---
def elabora_sensori_pandas(files_sensori):
    righe_reali = []
    for f in files_sensori:
        f.seek(0)
        try:
            ext = f.name.split('.')[-1].lower()
            if ext == 'csv': df_s = pd.read_csv(f, sep=None, engine='python')
            elif ext in ['xlsx', 'xls']: df_s = pd.read_excel(f)
            else: continue
                
            col_lower = [str(c).lower() for c in df_s.columns]
            date_col = next((c for c in df_s.columns if any(k in str(c).lower() for k in ['date', 'time', 'dat', 'timestamp'])), df_s.columns[0])
            
            # Match consumi
            if any('kwh' in c for c in col_lower) or any('energy' in c for c in col_lower) or any('energia' in c for c in col_lower):
                val_col = next(c for c in df_s.columns if 'kwh' in str(c).lower() or 'energy' in str(c).lower() or 'energia' in str(c).lower())
                df_s[val_col] = pd.to_numeric(df_s[val_col], errors='coerce')
                df_agg = df_s.groupby(pd.to_datetime(df_s[date_col], errors='coerce').dt.date)[val_col].sum().reset_index()
                for _, r in df_agg.dropna().iterrows():
                    righe_reali.append({'Data': r[date_col].strftime('%Y-%m-%d'), 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': r[val_col]})
                    
            elif any('lit' in c for c in col_lower) or any('vol' in c for c in col_lower) or any('water' in c for c in col_lower) or any('acqua' in c for c in col_lower):
                val_col = next(c for c in df_s.columns if 'lit' in str(c).lower() or 'vol' in str(c).lower() or 'water' in str(c).lower() or 'acqua' in str(c).lower())
                df_s[val_col] = pd.to_numeric(df_s[val_col], errors='coerce')
                df_agg = df_s.groupby(pd.to_datetime(df_s[date_col], errors='coerce').dt.date)[val_col].sum().reset_index()
                for _, r in df_agg.dropna().iterrows():
                    righe_reali.append({'Data': r[date_col].strftime('%Y-%m-%d'), 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': r[val_col]})
                    
            elif any('dist' in c for c in col_lower) or any('km' in c for c in col_lower):
                val_col = next(c for c in df_s.columns if 'dist' in str(c).lower() or 'km' in str(c).lower())
                df_s[val_col] = pd.to_numeric(df_s[val_col], errors='coerce')
                df_agg = df_s.groupby(pd.to_datetime(df_s[date_col], errors='coerce').dt.date)[val_col].sum().reset_index()
                for _, r in df_agg.dropna().iterrows():
                    righe_reali.append({'Data': r[date_col].strftime('%Y-%m-%d'), 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': r[val_col] * 0.25})
            
            elif any('ore' in c for c in col_lower) or any('hour' in c for c in col_lower) or any('tempo' in c for c in col_lower):
                val_col = next(c for c in df_s.columns if 'ore' in str(c).lower() or 'hour' in str(c).lower() or 'tempo' in str(c).lower())
                df_s[val_col] = pd.to_numeric(df_s[val_col], errors='coerce')
                df_agg = df_s.groupby(pd.to_datetime(df_s[date_col], errors='coerce').dt.date)[val_col].sum().reset_index()
                for _, r in df_agg.dropna().iterrows():
                    righe_reali.append({'Data': r[date_col].strftime('%Y-%m-%d'), 'Parametro': 'Macchinari', 'Elemento': 'Diesel', 'Quantita': r[val_col] * 15.0})
        except Exception: pass
    return righe_reali


st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Impostazioni architettura dati")

def reset_dati():
    if 'df_cantiere' in st.session_state: del st.session_state['df_cantiere']

fase_cantiere = st.radio(
    "Seleziona la tipologia di analisi LCA:",
    ["**Fase di progettazione**", "**Fase di esecuzione**"],
    on_change=reset_dati
)

if "**Fase di progettazione**" in fase_cantiere:
    col1, col2, col3 = st.columns(3)
    with col1: stima_macchinari = st.number_input("**Macchinari** (kg diesel/giorno)", value=45.0, min_value=0.0, step=5.0)
    with col2: stima_energia = st.number_input("**Energia elettrica** (kWh/giorno)", value=120.0, min_value=0.0, step=10.0)
    with col3: stima_acqua = st.number_input("**Acqua** (litri/giorno)", value=500.0, min_value=0.0, step=50.0)
    
    files_unificati = st.file_uploader("Documentazione (Abachi e Cronoprogramma in CSV/XLSX)", type=['xlsx', 'csv'], accept_multiple_files=True)

    if st.button("Elabora modello predittivo (Matching WBS)"):
        if not files_unificati: st.warning("Carica almeno un abaco e un cronoprogramma per avviare il matching.")
        else:
            with st.spinner("Estrazione ed elaborazione matrici WBS in corso..."):
                df_estratti = elabora_materiali_wbs_pandas(files_unificati)
                
                if df_estratti.empty:
                    st.error("Matching WBS fallito. Assicurati che i file contengano le colonne WBS/Codice, e le date di inizio/fine.")
                else:
                    righe_arricchite = []
                    giorni_unici = set()
                    
                    for _, row in df_estratti.iterrows():
                        giorni_lavorativi = pd.bdate_range(start=row['Data_Inizio'], end=row['Data_Fine'])
                        if len(giorni_lavorativi) < 1: giorni_lavorativi = [row['Data_Inizio']]
                        quantita_giorn = row['Quantita_Totale'] / len(giorni_lavorativi)
                        
                        for d in giorni_lavorativi:
                            data_str = d.strftime('%Y-%m-%d')
                            giorni_unici.add(data_str)
                            righe_arricchite.append({'Data': data_str, 'Parametro': 'Materiali', 'Elemento': row['Elemento'], 'Quantita': quantita_giorn})
                            
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
                    st.success("Modello LCA predittivo generato correttamente via WBS.")

elif "**Fase di esecuzione**" in fase_cantiere:
    col_file1, col_file2 = st.columns(2)
    with col_file1: files_bim = st.file_uploader("1. Abachi BIM e Cronoprogramma (CSV/XLSX)", type=['csv', 'xlsx'], accept_multiple_files=True)
    with col_file2: files_sensori = st.file_uploader("2. Log sensoristica IoT (CSV/XLSX)", type=['csv', 'xlsx'], accept_multiple_files=True)
        
    if st.button("Elabora monitoraggio dinamico (WBS + Sensori)"):
        if not files_bim and not files_sensori:
            st.warning("Eseguire il caricamento di almeno un pacchetto dati per generare il calcolo.")
        else:
            with st.spinner("Allineamento matrici WBS e parsing temporale dei sensori in corso..."):
                righe_reali = []
                
                # Parte 1: Materiali tramite WBS
                if files_bim:
                    df_estratti = elabora_materiali_wbs_pandas(files_bim)
                    if not df_estratti.empty:
                        for _, row in df_estratti.iterrows():
                            giorni_lavorativi = pd.bdate_range(start=row['Data_Inizio'], end=row['Data_Fine'])
                            if len(giorni_lavorativi) < 1: giorni_lavorativi = [row['Data_Inizio']]
                            quantita_giorn = row['Quantita_Totale'] / len(giorni_lavorativi)
                            
                            for d in giorni_lavorativi:
                                d_str = d.strftime('%Y-%m-%d')
                                righe_reali.append({'Data': d_str, 'Parametro': 'Materiali', 'Elemento': row['Elemento'], 'Quantita': quantita_giorn})
                                
                                q_rif = quantita_giorn * 0.15
                                righe_reali.append({'Data': d_str, 'Parametro': 'Rifiuti', 'Elemento': 'Inerti / Macerie di demolizione', 'Quantita': q_rif})
                                righe_reali.append({'Data': d_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (quantita_giorn/28000)*30*0.25})
                                righe_reali.append({'Data': d_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_rif/28000)*30*0.25})
                    else:
                        st.warning("Materiali non trovati: Nessun match WBS effettuato tra abachi e cronoprogramma.")

                # Parte 2: Consumi tramite parsing CSV
                if files_sensori:
                    righe_sensori = elabora_sensori_pandas(files_sensori)
                    righe_reali.extend(righe_sensori)

                if righe_reali:
                    df_cantiere_reale = pd.DataFrame(righe_reali)
                    df_cantiere_reale['Quantita'] = pd.to_numeric(df_cantiere_reale['Quantita'], errors='coerce').fillna(0)
                    df_cantiere_reale = df_cantiere_reale.groupby(['Data', 'Parametro', 'Elemento'], as_index=False)['Quantita'].sum()
                    df_cantiere_reale['Elemento'] = df_cantiere_reale.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
                    
                    st.session_state['df_cantiere'] = df_cantiere_reale
                    st.success("Modello dinamico unificato calcolato deterministicamente con successo.")
                else:
                    st.error("Nessun dato valido estratto dai file caricati.")

st.markdown("</div>", unsafe_allow_html=True)


if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere'].copy()
    
    df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='mixed', errors='coerce')
    
    df_cantiere['Parametro_match'] = df_cantiere['Parametro'].astype(str).str.strip().str.lower()
    df_cantiere['Elemento_match'] = df_cantiere['Elemento'].astype(str).str.strip().str.lower()
    if df_inventario is not None:
        df_inventario['Parametro_match'] = df_inventario['Parametro'].astype(str).str.strip().str.lower()
        df_inventario['Elemento_match'] = df_inventario['Elemento'].astype(str).str.strip().str.lower()
        df_inventario_clean = df_inventario.drop_duplicates(subset=['Parametro_match', 'Elemento_match']).copy()
        
        df_completo = pd.merge(df_cantiere, df_inventario_clean, on=['Parametro_match', 'Elemento_match'], how='left', suffixes=('', '_lci'))
        df_completo['Parametro'] = df_completo['Parametro'].fillna(df_completo['Parametro_match'])
        df_completo['Elemento'] = df_completo['Elemento'].fillna(df_completo['Elemento_match'])
        df_completo.drop(columns=['Parametro_match', 'Elemento_match', 'Parametro_lci', 'Elemento_lci'], errors='ignore', inplace=True)
        df_completo['Fattore_Emissione'] = df_completo['Fattore_Emissione'].fillna(0)
    else:
        df_completo = df_cantiere
        df_completo['Fattore_Emissione'] = 0.0

    df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
    
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Configurazione temporale dell'analisi")
    
    min_date = df_completo['Data_dt'].min().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].min()) else date.today()
    max_date = df_completo['Data_dt'].max().date() if not df_completo.empty and pd.notna(df_completo['Data_dt'].max()) else date.today()

    date_range = st.date_input("Selezione intervallo di calcolo:", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")

    if isinstance(date_range, tuple) and len(date_range) == 2:
        df_filtrato = df_completo.loc[(df_completo['Data_dt'].dt.date >= date_range[0]) & (df_completo['Data_dt'].dt.date <= date_range[1])].copy()
    else: 
        df_filtrato = df_completo.copy()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("Risultati analitici")
    if df_filtrato.empty: 
        st.info("La selezione temporale non presenta dati elaborabili.")
    else:
        modo_visualizzazione = st.selectbox("Modalità di visualizzazione grafica", ["Distribuzione temporale", "Analisi di frequenza (KDE)", "Andamento (media mobile)", "Riferimento media fissa"], index=0)

        def genera_figura(df_dat, x_col, y_col, titolo, colore_base):
            df_dat = df_dat.copy()
            if df_dat.empty: return go.Figure()
            if "frequenza" in modo_visualizzazione:
                fig = px.histogram(df_dat, x=y_col, nbins=20, marginal="violin", title=f"{titolo} - Analisi di frequenza", labels={y_col: 'kg CO₂e'}, color_discrete_sequence=[colore_base])
            elif "mobile" in modo_visualizzazione:
                df_dat = df_dat.sort_values(by=x_col)
                df_dat['Media_Mobile'] = df_dat[y_col].rolling(window=7, min_periods=1).mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Dato giornaliero', marker_color=colore_base, opacity=0.4))
                fig.add_trace(go.Scatter(x=df_dat[x_col], y=df_dat['Media_Mobile'], mode='lines', name='Media mobile', line=dict(color='#111827', width=2)))
                fig.update_layout(title=f"{titolo} - Variazione su media mobile")
            elif "fissa" in modo_visualizzazione:
                media_val = df_dat[y_col].mean()
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Dato giornaliero', marker_color=colore_base, opacity=0.5))
                fig.add_hline(y=media_val, line_dash="dash", line_color="#4b5563", annotation_text=f"Media: {media_val:.2f}")
                fig.update_layout(title=f"{titolo} - Riferimento su media fissa")
            else:
                fig = px.bar(df_dat, x=x_col, y=y_col, title=titolo, labels={y_col: 'kg CO₂e', x_col: ''})
                fig.update_traces(marker_color=colore_base)
            
            fig.update_layout(height=380, font_family="Inter", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', title_font_size=13, title_font_color="#374151")
            return fig

        df_totale = df_filtrato.groupby('Data')['CO2_Totale_kg'].sum().reset_index()
        st.plotly_chart(genera_figura(df_totale, 'Data', 'CO2_Totale_kg', "Andamento complessivo impronta carbonica", "#111827"), use_container_width=True)
        
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        df_filtrato['Parametro_Title'] = df_filtrato['Parametro'].astype(str).str.title()
        df_pie = df_filtrato.groupby('Parametro_Title')['CO2_Totale_kg'].sum().reset_index()
        df_pie = df_pie[df_pie['CO2_Totale_kg'] > 0]
        
        colori_parametri = {'Materiali': "#991b1b", 'Rifiuti': "#166534", 'Trasporti': "#4338ca", 'Energia': "#ca8a04", 'Acqua': "#0369a1", 'Macchinari': "#4b5563"}
        if not df_pie.empty:
            fig_pie = px.pie(df_pie, values='CO2_Totale_kg', names='Parametro_Title', title="Incidenza strutturale per categorie operative", color='Parametro_Title', color_discrete_map=colori_parametri, hole=0.4)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
            fig_pie.update_layout(height=450, font_family="Inter", showlegend=True, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', title_font_size=13, title_font_color="#374151")
            st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        st.subheader("Analisi disaggregata per voci strutturali")
        col_grafici = st.columns(2)
        for idx, param in enumerate(df_filtrato['Parametro_Title'].dropna().unique()):
            df_p = df_filtrato[df_filtrato['Parametro_Title'] == param].groupby('Data')['CO2_Totale_kg'].sum().reset_index()
            with col_grafici[idx % 2]:
                st.plotly_chart(genera_figura(df_p, 'Data', 'CO2_Totale_kg', param, colori_parametri.get(param, '#4b5563')), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
