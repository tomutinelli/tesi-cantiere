import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io
import re

# Rimosse le librerie google genai
# PyPDF2 verrà importato localmente nella funzione per gestire il cronoprogramma

st.set_page_config(page_title="EcoSite Tracker | Analisi LCA", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1f2937;
        background-color: #f9fafb;
    }

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

    .stRadio label, .stNumberInput label, .stDateInput label, .stFileUploader label, .stSelectbox label {
        font-weight: 500; font-size: 0.9rem; color: #374151;
    }
</style>
""", unsafe_allow_html=True)

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
    st.error(f"Errore di sistema: database LCI '{percorso_lci}' non reperibile o corrotto.")
    st.stop()

def mappa_voce_a_lci(parametro, elemento_grezzo, df_inv=df_inventario):
    p_str = str(parametro).strip().lower()
    e_str = str(elemento_grezzo).strip().lower()
    voci_disponibili = df_inv[df_inv['Parametro'].str.lower() == p_str]['Elemento'].tolist()
    if not voci_disponibili: voci_disponibili = df_inv['Elemento'].tolist()
        
    for v in voci_disponibili:
        if v.lower() in e_str or e_str in v.lower(): return v
            
    if p_str == 'materiali':
        if 'calcestruzzo' in e_str or 'cls' in e_str: return 'Calcestruzzo'
        if 'acciaio' in e_str or 'ferro' in e_str: return 'Acciaio'
        if 'laterizio' in e_str or 'mattone' in e_str: return 'Laterizio'
        if 'inerti' in e_str or 'sabbia' in e_str or 'ghiaia' in e_str: return 'Inerti'
        if 'asfalto' in e_str or 'bitume' in e_str or 'strada' in e_str: return 'Asfalto/Bitume'
        if 'legno' in e_str: return 'Legno'
        if 'vetro' in e_str: return 'Vetro'
        if 'isolante' in e_str or 'lana' in e_str or 'eps' in e_str: return 'Isolante EPS'
    elif p_str == 'rifiuti':
        if 'scavo' in e_str or 'terra' in e_str: return 'Inerti / Macerie di demolizione'
        if 'calcestruzzo' in e_str: return 'Calcestruzzo di risulta'
        if 'acciaio' in e_str or 'ferro' in e_str: return 'Metallo / Acciaio di scarto'
    elif p_str in ['trasporti', 'macchinari']:
        if 'diesel' in e_str or 'gasolio' in e_str: return 'Diesel'
        if 'benzina' in e_str or 'petrol' in e_str: return 'Petrol'
        
    from difflib import get_close_matches
    match = get_close_matches(str(elemento_grezzo), voci_disponibili, n=1, cutoff=0.1)
    if match: return match[0]
    return elemento_grezzo

# ====================================================================
# NUOVO MOTORE DETERMINISTICO DI MATCHING (SENZA INTELLIGENZA ARTIFICIALE)
# ====================================================================
def elabora_materiali_bim_offline(lista_file):
    file_gantt = None
    file_abachi = []
    
    # 1. Separazione dei file (Cronoprogramma vs Abachi)
    for f in lista_file:
        f.seek(0)
        nome = f.name.lower()
        if 'crono' in nome or 'gantt' in nome or 'programma' in nome:
            file_gantt = f
        else:
            file_abachi.append(f)
            
    if not file_gantt:
        st.error("Nessun cronoprogramma trovato. Assicurati che il file abbia 'crono' o 'gantt' nel nome.")
        return pd.DataFrame()
    if not file_abachi:
        st.error("Nessun abaco trovato.")
        return pd.DataFrame()

    # 2. Parsing del Cronoprogramma
    wbs_dates = {}
    testo_gantt = ""
    ext_gantt = file_gantt.name.split('.')[-1].lower()
    
    if ext_gantt == 'pdf':
        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(file_gantt)
            for page in reader.pages:
                testo_gantt += page.extract_text() + "\n"
        except ImportError:
            st.error("Libreria PyPDF2 non trovata. Esegui 'pip install PyPDF2' nel terminale.")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"Errore nella lettura del PDF del cronoprogramma: {e}")
            return pd.DataFrame()
    else:
        try:
            testo_gantt = file_gantt.getvalue().decode('utf-8', errors='ignore')
        except: pass
            
    # Regex per intercettare pattern WBS (es. Ss_20_20, Pr_20_31_16_92) e date
    linee = testo_gantt.split('\n')
    regex_wbs = re.compile(r'([A-Za-z]+(?:_\d+)+)')
    regex_data = re.compile(r'(\d{2}/\d{2}/\d{2,4}|\d{4}-\d{2}-\d{2})')
    
    for linea in linee:
        wbs_match = regex_wbs.search(linea)
        date_matches = regex_data.findall(linea)
        if wbs_match and len(date_matches) >= 2:
            wbs_code = wbs_match.group(1).lower().strip()
            wbs_dates[wbs_code] = (date_matches[0], date_matches[1])

    if not wbs_dates:
        st.warning("Il sistema non è riuscito a rilevare codici WBS e Date nel cronoprogramma. Controlla il formato.")
        return pd.DataFrame()

    # 3. Parsing degli Abachi (Materiali e Quantità)
    df_totale_abachi = pd.DataFrame()
    for f_abaco in file_abachi:
        ext = f_abaco.name.split('.')[-1].lower()
        try:
            if ext in ['csv', 'txt']:
                df = pd.read_csv(f_abaco, sep=None, engine='python')
            elif ext in ['xls', 'xlsx']:
                df = pd.read_excel(f_abaco)
            else: continue
            
            # Ricerca flessibile delle colonne pertinenti
            wbs_col = next((c for c in df.columns if 'wbs' in str(c).lower() or 'codice' in str(c).lower() or 'code' in str(c).lower()), None)
            mat_col = next((c for c in df.columns if 'material' in str(c).lower() or 'famiglia' in str(c).lower() or 'tipo' in str(c).lower() or 'descrizione' in str(c).lower() or 'nome' in str(c).lower()), None)
            qta_col = next((c for c in df.columns if 'massa' in str(c).lower() or 'peso' in str(c).lower() or 'quant' in str(c).lower() or 'kg' in str(c).lower() or 'vol' in str(c).lower()), None)
            
            if wbs_col and mat_col and qta_col:
                df_temp = pd.DataFrame({
                    'WBS': df[wbs_col].astype(str),
                    'Elemento': df[mat_col].astype(str),
                    'Quantita': pd.to_numeric(df[qta_col].astype(str).str.replace(',','.'), errors='coerce').fillna(0)
                })
                df_totale_abachi = pd.concat([df_totale_abachi, df_temp], ignore_index=True)
        except Exception as e:
            st.error(f"Errore nella lettura dell'abaco {f_abaco.name}: {e}")
            
    if df_totale_abachi.empty:
        st.warning("Impossibile estrarre dati dagli abachi. Assicurati che contengano colonne per WBS, Materiale e Quantità/Massa.")
        return pd.DataFrame()

    # 4. Matching Deterministico WBS -> Cronoprogramma
    risultati = []
    for _, row in df_totale_abachi.iterrows():
        wbs_abaco_raw = row['WBS']
        wbs_abaco_clean_match = regex_wbs.search(wbs_abaco_raw)
        wbs_abaco = wbs_abaco_clean_match.group(1).lower().strip() if wbs_abaco_clean_match else wbs_abaco_raw.lower().strip()
        
        # Cerca corrispondenza nel dizionario delle date
        for wbs_gantt, (d_in, d_out) in wbs_dates.items():
            if wbs_gantt in wbs_abaco or wbs_abaco in wbs_gantt:
                risultati.append({
                    'Parametro': 'Materiali',
                    'Elemento': row['Elemento'],
                    'Quantita_Totale': row['Quantita'],
                    'Data_Inizio': d_in,
                    'Data_Fine': d_out
                })
                break
                
    df_out = pd.DataFrame(risultati)
    if df_out.empty:
        st.warning("Nessun match trovato tra i codici WBS degli abachi e quelli del cronoprogramma.")
        return pd.DataFrame()
        
    df_out['Data_Inizio'] = pd.to_datetime(df_out['Data_Inizio'], format='mixed', dayfirst=True, errors='coerce')
    df_out['Data_Fine'] = pd.to_datetime(df_out['Data_Fine'], format='mixed', dayfirst=True, errors='coerce')
    return df_out.dropna(subset=['Data_Inizio', 'Data_Fine'])


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
    st.markdown("""
    <div style='background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 4px; padding: 20px; margin-bottom: 20px;'>
        <h4>Analisi predittiva (Offline Mode)</h4>
        <p style='color: #4b5563; font-size: 0.9rem; margin-bottom: 10px; line-height: 1.5;'>
        Calcolo impronta carbonica tramite algoritmo deterministico. Esegue il matching diretto dei codici WBS tra cronoprogramma e abachi senza IA. La stima inserita per Macchinari, Energia e Acqua verrà distribuita equamente su tutti i giorni estratti dal programma lavori.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        stima_macchinari = st.number_input("**Macchinari** (kg diesel/giorno)", value=45.0, min_value=0.0, step=5.0)
    with col2:
        stima_energia = st.number_input("**Energia elettrica** (kWh/giorno)", value=120.0, min_value=0.0, step=10.0)
    with col3:
        stima_acqua = st.number_input("**Acqua** (litri/giorno)", value=500.0, min_value=0.0, step=50.0)
    
    files_unificati = st.file_uploader("Documentazione di cantiere: abachi e cronoprogramma (file multipli)", type=['pdf', 'txt', 'xlsx', 'csv'], accept_multiple_files=True)

    if st.button("Elabora modello predittivo"):
        if not files_unificati: 
            st.warning("È necessario caricare la documentazione per avviare il matching.")
        else:
            try:
                with st.spinner("Estrazione semantica ed elaborazione matrici WBS (Algoritmo Deterministico)..."):
                    df_estratti = elabora_materiali_bim_offline(files_unificati)
                    
                    if df_estratti.empty:
                        st.error("L'elaborazione non ha prodotto dati validi. Verifica l'integrità dei documenti e dei WBS.")
                    else:
                        righe_arricchite = []
                        giorni_unici = set()
                        
                        # 1. Distribuzione dei materiali sulle date di competenza WBS
                        for index, row in df_estratti.iterrows():
                            giorni_lavorativi = pd.bdate_range(start=row['Data_Inizio'], end=row['Data_Fine'])
                            if len(giorni_lavorativi) < 1: giorni_lavorativi = [row['Data_Inizio']]
                            quantita_giorn = row['Quantita_Totale'] / len(giorni_lavorativi)
                            
                            for d in giorni_lavorativi:
                                data_str = d.strftime('%Y-%m-%d')
                                giorni_unici.add(data_str)
                                
                                # Allocazione materiale
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Materiali', 'Elemento': row['Elemento'], 'Quantita': quantita_giorn})
                                
                                # Sfridi e logistica collegati al materiale
                                q_rif = quantita_giorn * 0.15
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Rifiuti', 'Elemento': 'Inerti / Macerie di demolizione', 'Quantita': q_rif})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (quantita_giorn/28000)*30*0.25})
                                righe_arricchite.append({'Data': data_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_rif/28000)*30*0.25})

                        # 2. Distribuzione SCELTE DELL'UTENTE (Macchinari, Energia, Acqua) spalmata su tutti i giorni attivi
                        for d in giorni_unici:
                            righe_arricchite.append({'Data': d, 'Parametro': 'Macchinari', 'Elemento': 'Diesel', 'Quantita': stima_macchinari})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': stima_energia})
                            righe_arricchite.append({'Data': d, 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': stima_acqua})

                        df_cantiere_grezzo = pd.DataFrame(righe_arricchite)
                        df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
                        st.session_state['df_cantiere'] = df_cantiere_grezzo
                        st.success("Modello LCA predittivo generato e consolidato con successo.")
            except Exception as e: st.error(f"Errore di sistema: {e}")

elif "**Fase di esecuzione**" in fase_cantiere:
    st.markdown("""
    <div style='background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 4px; padding: 20px; margin-bottom: 20px;'>
        <h4>Integrazione telemetria e monitoraggio dinamico</h4>
    </div>
    """, unsafe_allow_html=True)
    
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        files_bim = st.file_uploader("1. Abachi BIM e cronoprogramma", type=['csv', 'xlsx', 'pdf'], accept_multiple_files=True)
    with col_file2:
        files_sensori = st.file_uploader("2. Log sensoristica IoT (CSV)", type=['csv'], accept_multiple_files=True)
        
    if st.button("Unifica ed elabora monitoraggio dinamico"):
        righe_reali = []
        
        # 1. Elaborazione Matrice BIM
        if files_bim:
            try:
                with st.spinner("Allineamento matrici WBS (Offline Mode)..."):
                    df_mat = elabora_materiali_bim_offline(files_bim)
                    if not df_mat.empty:
                        for _, r in df_mat.iterrows():
                            giorni_lavorativi = pd.bdate_range(start=r['Data_Inizio'], end=r['Data_Fine'])
                            if len(giorni_lavorativi) < 1: giorni_lavorativi = [r['Data_Inizio']]
                            q_giorn = r['Quantita_Totale'] / len(giorni_lavorativi)
                            
                            for d in giorni_lavorativi:
                                d_str = d.strftime('%Y-%m-%d')
                                righe_reali.append({'Data': d_str, 'Parametro': 'Materiali', 'Elemento': r['Elemento'], 'Quantita': q_giorn})
                                q_rif = q_giorn * 0.15
                                righe_reali.append({'Data': d_str, 'Parametro': 'Rifiuti', 'Elemento': 'Inerti / Macerie di demolizione', 'Quantita': q_rif})
                                righe_reali.append({'Data': d_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_giorn/28000)*30*0.25})
                                righe_reali.append({'Data': d_str, 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': (q_rif/28000)*30*0.25})
            except Exception as e:
                st.error(f"Errore nell'elaborazione del modello BIM: {e}")

        # 2. Elaborazione Modulo Sensori
        if files_sensori:
            for f in files_sensori:
                try:
                    f.seek(0)
                    df_s = pd.read_csv(f)
                    col_lower = [c.lower() for c in df_s.columns]
                    date_col = next((c for c in df_s.columns if 'date' in c.lower() or 'time' in c.lower() or 'dat' in c.lower() or 'timestamp' in c.lower()), df_s.columns[0])
                    
                    if any('kwh' in c for c in col_lower) or any('energy' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'kwh' in c.lower() or 'energy' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Energia', 'Elemento': 'Gas naturale - NGCC', 'Quantita': r[val_col]})
                            
                    elif any('lit' in c for c in col_lower) or any('vol' in c for c in col_lower) or any('water' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'lit' in c.lower() or 'vol' in c.lower() or 'water' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Acqua', 'Elemento': 'Water supply', 'Quantita': r[val_col]})
                            
                    elif any('dist' in c for c in col_lower) or any('km' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'dist' in c.lower() or 'km' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Trasporti', 'Elemento': 'Diesel', 'Quantita': r[val_col] * 0.25})
                    
                    elif any('ore' in c for c in col_lower) or any('hour' in c for c in col_lower):
                        val_col = next(c for c in df_s.columns if 'ore' in c.lower() or 'hour' in c.lower())
                        df_agg = df_s.groupby(df_s[date_col].astype(str).str[:10])[val_col].sum().reset_index()
                        for _, r in df_agg.iterrows():
                            righe_reali.append({'Data': r[date_col], 'Parametro': 'Macchinari', 'Elemento': 'Diesel', 'Quantita': r[val_col] * 15.0})
                            
                except Exception as e:
                    st.error(f"Errore in elaborazione file telemetrico {f.name}: {e}")

        if righe_reali:
            df_cantiere_reale = pd.DataFrame(righe_reali)
            df_cantiere_reale['Quantita'] = pd.to_numeric(df_cantiere_reale['Quantita'], errors='coerce').fillna(0)
            df_cantiere_reale = df_cantiere_reale.groupby(['Data', 'Parametro', 'Elemento'], as_index=False)['Quantita'].sum()
            df_cantiere_reale['Elemento'] = df_cantiere_reale.apply(lambda r: mappa_voce_a_lci(r['Parametro'], r['Elemento']), axis=1)
            
            st.session_state['df_cantiere'] = df_cantiere_reale
            st.success("Modello dinamico unificato offline.")
        else:
            st.warning("Eseguire il caricamento di almeno un pacchetto dati per generare il calcolo.")

st.markdown("</div>", unsafe_allow_html=True)


if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere'].copy()
    
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

        col_exp1, col_exp2 = st.columns(2)
        with col_exp1: st.download_button("Esporta dati totali (CSV)", data=df_totale.to_csv(index=False).encode('utf-8'), file_name="emissioni_totali.csv", mime="text/csv")
        with col_exp2:
            output_xlsx = io.BytesIO()
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer: df_totale.to_excel(writer, index=False, sheet_name='Totale')
            st.download_button("Esporta dati totali (XLSX)", data=output_xlsx.getvalue(), file_name="emissioni_totali.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        st.subheader("Analisi disaggregata per voci strutturali")
        col_grafici = st.columns(2)
        for idx, param in enumerate(df_filtrato['Parametro_Title'].dropna().unique()):
            df_p = df_filtrato[df_filtrato['Parametro_Title'] == param].groupby('Data')['CO2_Totale_kg'].sum().reset_index()
            with col_grafici[idx % 2]:
                st.plotly_chart(genera_figura(df_p, 'Data', 'CO2_Totale_kg', param, colori_parametri.get(param, '#4b5563')), use_container_width=True)

        with st.expander("Ispezione matrice dati consolidata"):
            st.dataframe(df_filtrato.drop(columns=['Data_dt', 'Parametro_Title', 'Parametro_match', 'Elemento_match'], errors='ignore'), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
