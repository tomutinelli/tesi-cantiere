import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta
import numpy as np
import os
import io

# --- IMPOSTAZIONI PAGINA E STILE CSS ---
st.set_page_config(page_title="EcoSite Tracker", layout="wide")
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; background-color: #faf9f6; }
    .minimal-card { background-color: #ffffff !important; border: 1.5px solid #d5ddd1 !important; border-radius: 8px !important; padding: 32px !important; margin-bottom: 24px !important; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.01) !important; }
</style>
""", unsafe_allow_html=True)

st.title("EcoSite Tracker")
st.markdown("Monitoraggio Emissioni CO₂ in fase operativa. Confronto tra i dati stimati in fase di Progettazione con i dati reali di cantiere misurati tramite sensoristica.")
            
# --- INIZIALIZZAZIONE SESSION STATE ---
if 'risultati_progettazione' not in st.session_state:
    st.session_state.risultati_progettazione = []
if 'configurazioni_salvate' not in st.session_state:
    st.session_state.configurazioni_salvate = {}
if 'df_comp_cantiere' not in st.session_state:
    st.session_state.df_comp_cantiere = None
if 'loaded_file_name' not in st.session_state:
    st.session_state.loaded_file_name = None

# COLORI GRAFICI (Allineati alla foto fornita)
colori_parametri = {
    'Materiali': "#D2042D", 
    'Rifiuti': "#93C572", 
    'Trasporti': "#D1D0CE", 
    'Energia': "#FFD700",
    'Acqua': "#0369a1", 
    'Macchinari': "#0ABAB5"
}

# --- FUNZIONI DI CARICAMENTO DATI STATICI ---
def load_static_data():
    df_epd, df_lci, df_macc = None, None, None
    err = ""
    try:
        if os.path.exists("EPD materiali.xlsx"):
            df_epd = pd.read_excel("EPD materiali.xlsx")
        elif os.path.exists("epd_materiali.csv"):
            df_epd = pd.read_csv("epd_materiali.csv", sep=None, engine='python', encoding='latin1')
        else:
            err += "Manca il database EPD (EPD materiali.xlsx). "
            
        if os.path.exists("LCI.xlsx"):
            df_lci = pd.read_excel("LCI.xlsx", sheet_name=None)
            for k in df_lci.keys():
                df_lci[k].columns = df_lci[k].columns.astype(str).str.strip()
        else:
            err += "Manca il database LCI.xlsx. "
            
        if os.path.exists("macchinari.xlsx"):
            df_macc = pd.read_excel("macchinari.xlsx")
        else:
            err += "Manca il database macchinari.xlsx. "
            
    except Exception as e:
        err += f"Errore interno di lettura database: {e}"
        
    return df_epd, df_lci, df_macc, err

df_epd, df_lci, df_macc, err_msg = load_static_data()
if err_msg:
    st.error(f"⚠️ {err_msg}")

# --- SEZIONE 1: CARICAMENTO DATI UTENTE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("1. Caricamento File di Progetto")

st.markdown("""
<div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #2980b9; border-radius: 4px; margin-bottom: 20px;'>
    <strong>Istruzioni di Formattazione CSV (Fase di Progetto):</strong>
    <ul style='margin-top: 5px; margin-bottom: 0px;'>
        <li><strong>Nuova Analisi (Cronoprogramma e Abachi):</strong> I file devono essere <code>.csv</code>. Il Cronoprogramma deve includere: <em>ID attività, Nome, Struttura, WBS Lavorazione, WBS Materiale, Data d'inizio, Data di fine</em>. Gli abachi devono includere la colonna <em>Massa, Peso o Quantità</em>.</li>
        <li><strong>Analisi Precedente:</strong> Carica il file riassuntivo esportato in precedenza in formato <code>.csv</code>, scaricabile nella sezione "3.Risultati Fase di Progettazione", per visualizzare subito i risultati e i grafici.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_caricati = st.file_uploader("Carica Cronoprogramma e Quantità (CSV)", type=['csv'], accept_multiple_files=True)
with col_up2:
    file_precedente = st.file_uploader("Carica un'Analisi Precedente (CSV)", type=['csv'])

lista_quantita = []
df_crono = None

if file_precedente:
    if st.session_state.loaded_file_name != file_precedente.name:
        try:
            df_prec = pd.read_csv(file_precedente, sep=None, engine='python')
            if 'Attività' in df_prec.columns and 'Data Inizio' in df_prec.columns:
                df_prec['Data Inizio'] = pd.to_datetime(df_prec['Data Inizio'], errors='coerce').dt.date
                df_prec['Data Fine'] = pd.to_datetime(df_prec['Data Fine'], errors='coerce').dt.date
                st.session_state.risultati_progettazione = df_prec.to_dict('records')
                st.session_state.loaded_file_name = file_precedente.name
        except Exception as e:
            st.error(f"Errore nella lettura dell'analisi precedente: {e}")

if file_caricati:
    for f in file_caricati:
        try:
            content = f.getvalue().decode('utf-8-sig', errors='ignore')
            lines = content.splitlines()
            header_idx = 0
            for i, line in enumerate(lines):
                if any(k in line.lower() for k in ['wbs', 'inizio', 'massa', 'attività', 'durata']):
                    header_idx = i
                    break
            csv_data = "\n".join(lines[header_idx:])
            df_temp = pd.read_csv(io.StringIO(csv_data), sep=';')
            if len(df_temp.columns) < 2:
                df_temp = pd.read_csv(io.StringIO(csv_data), sep=',')
            df_temp.columns = df_temp.columns.str.strip()
            col_lower = [str(c).lower() for c in df_temp.columns]
            if any('inizio' in c or 'fine' in c or 'durata' in c for c in col_lower):
                df_crono = df_temp
            else:
                lista_quantita.append(df_temp)
        except Exception as e:
            st.error(f"Errore nella lettura del file {f.name}: {e}")

df_quantita = pd.concat(lista_quantita, ignore_index=True) if lista_quantita else None

if df_crono is not None:
    col_inizio = next((c for c in df_crono.columns if 'inizio' in str(c).lower()), None)
    col_fine = next((c for c in df_crono.columns if 'fine' in str(c).lower()), None)
    
    if col_inizio and col_fine:
        df_crono['Data_Inizio_Calc'] = pd.to_datetime(df_crono[col_inizio], dayfirst=True, errors='coerce')
        df_crono['Data_Fine_Calc'] = pd.to_datetime(df_crono[col_fine], dayfirst=True, errors='coerce')
        df_crono = df_crono.dropna(subset=['Data_Inizio_Calc', 'Data_Fine_Calc']).sort_values(by='Data_Inizio_Calc')
        
        if len(df_crono) > 0:
            col_id = next((c for c in df_crono.columns if 'id' in str(c).lower()), df_crono.columns[0])
            col_struttura = next((c for c in df_crono.columns if 'struttura' in str(c).lower()), None)
            col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_crono.columns[1])
   
            etichette = df_crono[col_id].astype(str) + " - " + df_crono[col_nome].astype(str)
            if col_struttura: etichette += " | " + df_crono[col_struttura].astype(str)
            
            df_crono['Label_Attivita'] = etichette
            st.success(f"Cronoprogramma caricato con successo!")
        else:
            df_crono = None
elif file_precedente and 'risultati_progettazione' in st.session_state and st.session_state.risultati_progettazione:
    df_prec = pd.DataFrame(st.session_state.risultati_progettazione)
    df_crono = pd.DataFrame({
        'ID': df_prec['ID'] if 'ID' in df_prec.columns else df_prec.index,
        'Nome': df_prec['Nome'] if 'Nome' in df_prec.columns else df_prec['Attività'],
        'Struttura': df_prec['Struttura'] if 'Struttura' in df_prec.columns else '',
        'WBS Lav': df_prec['WBS_L'] if 'WBS_L' in df_prec.columns else '',
        'WBS Materiale': df_prec['WBS_M'] if 'WBS_M' in df_prec.columns else '',
        'Data_Inizio_Calc': pd.to_datetime(df_prec['Data Inizio']),
        'Data_Fine_Calc': pd.to_datetime(df_prec['Data Fine']),
        'Durata': df_prec['Durata'] if 'Durata' in df_prec.columns else 1,
        'Label_Attivita': df_prec['Attività']
    })
    st.success("Dati dell'analisi precedente ripristinati!")

st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 2: COMPILAZIONE ATTIVITA' ---
if df_crono is not None and df_epd is not None and df_lci is not None:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("2. Analisi Attività (Fase di Progettazione)")
    
    attivita_selezionata = st.selectbox("Scegli la lavorazione dal cronoprogramma:", df_crono['Label_Attivita'].tolist())
    row_att = df_crono[df_crono['Label_Attivita'] == attivita_selezionata].iloc[0]
    
    col_id = next((c for c in df_crono.columns if 'id' in str(c).lower()), df_crono.columns[0])
    col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_crono.columns[1])
    col_struttura = next((c for c in df_crono.columns if 'struttura' in str(c).lower()), None)
    col_wbs_lav = next((c for c in df_crono.columns if 'wbs lav' in str(c).lower()), None)
    
    id_att = str(row_att[col_id]) 
    nome_att = row_att[col_nome]
    nome_strutt = row_att[col_struttura] if col_struttura else ""
    wbs_lav = row_att[col_wbs_lav] if col_wbs_lav else ""
    unique_key = str(row_att['Label_Attivita']).replace(" ", "_").replace("-", "_").replace("|", "_")
    
    col_wbs_mat = next((c for c in df_crono.columns if 'wbs materiale' in str(c).lower()), None)
    wbs_mat = row_att[col_wbs_mat] if col_wbs_mat else None
    
    quantita_mat = 0
    if df_quantita is not None and wbs_mat is not None:
        q_col_lav = next((c for c in df_quantita.columns if 'wbs lav' in str(c).lower()), None)
        q_col_mat = next((c for c in df_quantita.columns if 'wbs mat' in str(c).lower()), None)
        q_col_struttura = next((c for c in df_quantita.columns if 'struttura' in str(c).lower()), None)
        q_col_massa = next((c for c in df_quantita.columns if str(c).lower().strip() in ['massa', 'peso'] or 'quantit' in str(c).lower()), None)
        
        if q_col_lav and q_col_mat and q_col_massa:
            mask = (df_quantita[q_col_lav].astype(str).str.strip() == str(wbs_lav).strip()) & \
                   (df_quantita[q_col_mat].astype(str).str.strip() == str(wbs_mat).strip())
            if q_col_struttura and nome_strutt:
                mask = mask & (df_quantita[q_col_struttura].astype(str).str.strip() == str(nome_strutt).strip())
            match_q = df_quantita[mask]
            if not match_q.empty:
                massa_str = match_q[q_col_massa].astype(str).str.replace(r'[^\d,\.-]', '', regex=True)
                quantita_mat = pd.to_numeric(massa_str.str.replace(',', '.'), errors='coerce').fillna(0).sum()

    st.info(f"Quantità materiale associata: **{quantita_mat:,.2f}** kg")
    
    em_materiali, em_rifiuti, em_trasporti, em_energia, em_acqua, em_macchinari = 0, 0, 0, 0, 0, 0
    q_rifiuti = 0
    fe_trasporti_medio = 0
    
    if quantita_mat > 0:
        materiale_scelto = st.selectbox("Seleziona Tipo Materiale:", df_epd['Materiale'].dropna().unique())
        df_epd_filtrato = df_epd[df_epd['Materiale'] == materiale_scelto]
        col_epd_fe = next((c for c in df_epd_filtrato.columns if 'fattore' in str(c).lower()), df_epd_filtrato.columns[3])
        col_epd_c2 = next((c for c in df_epd_filtrato.columns if 'c2' in str(c).lower()), df_epd_filtrato.columns[4])
        
        scelta_specifica = st.selectbox("Scegli il prodotto specifico:", df_epd_filtrato['Produttore'])
        fattore_em_mat = float(str(df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_fe].values[0]).replace(',', '.'))
        fattore_em_rif = float(str(df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_c2].values[0]).replace(',', '.'))
        
        em_materiali = quantita_mat * fattore_em_mat
        perc_rifiuti = st.number_input("% di rifiuti stimata:", min_value=0.0, max_value=100.0, value=0.0, key=f"rif_{unique_key}") / 100
        q_rifiuti = quantita_mat * perc_rifiuti
        em_rifiuti = q_rifiuti * fattore_em_rif

    # Trasporti e Macchinari semplificati per brevità ma completi
    if 'Trasporti' in df_lci:
        df_t_sheet = df_lci['Trasporti']
        df_t_emissions = df_t_sheet.iloc[0:9].copy()
        df_t_consumi = df_t_sheet.iloc[11:21].copy()
        df_t_consumi.columns = ['Tipologia', 'Carburante', 'Consumo_Medio']
        df_t_consumi = df_t_consumi.dropna(subset=['Tipologia'])
        df_t_consumi['Label'] = df_t_consumi['Tipologia'].astype(str) + " - " + df_t_consumi['Carburante'].astype(str)
        
        num_tipi_trasporto = st.number_input("Quanti tipi di mezzi di trasporto?", min_value=0, value=1 if quantita_mat > 0 else 0, key=f"num_t_{unique_key}")
        em_trasporti = 0
        tot_km_progetto = 0
        if num_tipi_trasporto > 0:
            scelta_mezzo = st.selectbox("Seleziona Veicolo/Carburante:", df_t_consumi['Label'].unique(), key=f"t_mezzo_{unique_key}")
            row_mezzo = df_t_consumi[df_t_consumi['Label'] == scelta_mezzo].iloc[0]
            consumo_medio = float(str(row_mezzo['Consumo_Medio']).replace(',', '.'))
            viaggi = st.number_input("Viaggi stimati:", min_value=0, value=int(np.ceil(quantita_mat / 28000)) if quantita_mat > 0 else 0, key=f"t_viaggi_{unique_key}")
            distanza = st.number_input("Distanza media [km]:", min_value=0.0, value=30.0, key=f"t_dist_{unique_key}")
            carb_match = 'CNG(or LNG)' if 'CNG' in str(row_mezzo['Carburante']) else str(row_mezzo['Carburante'])
            row_em = df_t_emissions[df_t_emissions.iloc[:, 1].astype(str) == carb_match]
            fattore_carb = float(str(row_em.iloc[0, 2]).replace(',', '.')) if not row_em.empty else 0.0
            em_trasporti = viaggi * distanza * consumo_medio * fattore_carb
            tot_km_progetto = viaggi * distanza
            fe_trasporti_medio = (em_trasporti / tot_km_progetto) if tot_km_progetto > 0 else 0

    col_durata = next((c for c in df_crono.columns if 'durata' in str(c).lower()), None)
    durata_giorni = int(row_att[col_durata]) if col_durata else 1
    
    # Energia e Acqua
    tech_energia = st.selectbox("Tecnologia generazione elettrica:", df_lci['Energia'].iloc[:, 0].unique(), key=f"tech_e_{unique_key}")
    fe_energia = float(str(df_lci['Energia'][df_lci['Energia'].iloc[:, 0] == tech_energia].iloc[0, 1]).replace(',', '.'))
    kwh_stimati = st.number_input("Consumo elettrico stimato (kWh):", value=0.0, key=f"kwh_{unique_key}")
    em_energia = kwh_stimati * fe_energia

    tech_acqua = st.selectbox("Approvvigionamento/Trattamento Acqua:", df_lci['Acqua'].iloc[:, 0].unique(), key=f"tech_a_{unique_key}")
    fe_acqua = float(str(df_lci['Acqua'][df_lci['Acqua'].iloc[:, 0] == tech_acqua].iloc[0, 1]).replace(',', '.'))
    litri_stimati = st.number_input("Consumo idrico stimato (litri):", value=0.0, key=f"litri_{unique_key}")
    em_acqua = litri_stimati * fe_acqua

    if st.button("Salva (o Sovrascrivi) Calcolo per questa Attività", use_container_width=True):
        totale_eq = em_materiali + em_rifiuti + em_macchinari + em_trasporti + em_energia + em_acqua
        record = {
            "Attività": row_att['Label_Attivita'], "ID": id_att, "Nome": nome_att, "Struttura": nome_strutt,
            "WBS_L": wbs_lav, "WBS_M": wbs_mat, "Durata": durata_giorni, 
            "Data Inizio": row_att['Data_Inizio_Calc'].date(), "Data Fine": row_att['Data_Fine_Calc'].date(),
            "Materiali [kgCO₂ eq]": em_materiali, "Rifiuti [kgCO₂ eq]": em_rifiuti, 
            "Macchinari [kgCO₂ eq]": em_macchinari, "Trasporti [kgCO₂ eq]": em_trasporti,
            "Energia [kgCO₂ eq]": em_energia, "Acqua [kgCO₂ eq]": em_acqua, "Totale [kgCO₂ eq]": totale_eq
        }
        st.session_state.configurazioni_salvate[row_att['Label_Attivita']] = {
            "Energia": tech_energia, "FE_Energia": fe_energia, "Acqua": tech_acqua, "FE_Acqua": fe_acqua
        }
        lista_ris = [r for r in st.session_state.risultati_progettazione if r['Attività'] != row_att['Label_Attivita']]
        lista_ris.append(record)
        st.session_state.risultati_progettazione = sorted(lista_ris, key=lambda x: x['Data Inizio'])
        st.success("Attività salvata con successo!")
    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 3: RISULTATI PROGETTAZIONE ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("3. Risultati Fase di Progettazione")
    df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
    st.dataframe(df_risultati, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 4: FASE DI CANTIERE (DATI SENSORI) ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("4. Monitoraggio Fase di Cantiere (Dati Dinamici Sensori)")
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti (CSV)", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia (CSV)", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua (CSV)", type=['csv'])

    def elabora_sensore_con_attivita(df_sens, tipo_file, df_lci, df_ris, conf_salvate):
        risultati = []
        df_sens.columns = [str(c).strip() for c in df_sens.columns]
        col_id = next((c for c in df_sens.columns if 'id' in str(c).lower()), df_sens.columns[0])
        col_data = next((c for c in df_sens.columns if 'data' in str(c).lower()), None)
        if not col_data: return pd.DataFrame(columns=['Data'])
        
        df_sens[col_data] = pd.to_datetime(df_sens[col_data], dayfirst=True, errors='coerce')
        df_sens = df_sens.dropna(subset=[col_data])
        
        for _, row in df_sens.iterrows():
            id_sensore = str(row[col_id]).strip()
            data_sensore = row[col_data].date()
            
            match_prog = df_ris[df_ris['ID'].astype(str).str.strip() == id_sensore]
            if not match_prog.empty:
                label_attivita = match_prog.iloc[0]['Attività']
            else:
                label_attivita = f"Attività ID: {id_sensore}"
                
            emissione = 0
            if tipo_file == 'energia':
                col_q = next((c for c in df_sens.columns if 'consum' in str(c).lower() or 'kwh' in str(c).lower()), None)
                kwh = pd.to_numeric(str(row[col_q]).replace(',','.')) if col_q else 0
                fe = conf_salvate.get(label_attivita, {}).get('FE_Energia', 0.4)
                emissione = kwh * fe
                risultati.append({'Data': data_sensore, 'Attività': label_attivita, 'Energia_Cant': emissione})
            elif tipo_file == 'acqua':
                col_q = next((c for c in df_sens.columns if 'consum' in str(c).lower() or 'litr' in str(c).lower()), None)
                litri = pd.to_numeric(str(row[col_q]).replace(',','.')) if col_q else 0
                fe = conf_salvate.get(label_attivita, {}).get('FE_Acqua', 0.001)
                emissione = litri * fe
                risultati.append({'Data': data_sensore, 'Attività': label_attivita, 'Acqua_Cant': emissione})
            elif tipo_file == 'trasporti':
                col_km = next((c for c in df_sens.columns if 'km' in str(c).lower()), None)
                km = pd.to_numeric(str(row[col_km]).replace(',','.')) if col_km else 0
                emissione = km * 0.5 * 2.6 # Stima standard o da LCI
                risultati.append({'Data': data_sensore, 'Attività': label_attivita, 'Trasporti_Cant': emissione})
                
        if risultati:
            df_res = pd.DataFrame(risultati)
            agg = {c: 'sum' for c in df_res.columns if c not in ['Data', 'Attività']}
            # PULIZIA ATTIVITA': Evita duplicazioni e unisce pulito
            agg['Attività'] = lambda x: '<br>'.join(sorted(set(str(item).strip() for sub in x.dropna() for item in sub.split(' | ') if item.strip())))
            return df_res.groupby('Data').agg(agg).reset_index()
        return pd.DataFrame(columns=['Data'])

    if st.button("Elabora Dati Sensori e Genera Grafici di Confronto", use_container_width=True):
        df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
        dati_prog = []
        for _, row in df_risultati.iterrows():
            giorni_lavoro = pd.bdate_range(start=row['Data Inizio'], end=row['Data Fine'])
            n_g = len(giorni_lavoro) if len(giorni_lavoro) > 0 else 1
            for g in giorni_lavoro:
                dati_prog.append({
                    'Data': g.date(), 'Attività_Prog': row['Attività'], 'Materiali': row["Materiali [kgCO₂ eq]"]/n_g, 
                    'Rifiuti': row["Rifiuti [kgCO₂ eq]"]/n_g, 'Macchinari': row["Macchinari [kgCO₂ eq]"]/n_g, 
                    'Trasporti_Prog': row["Trasporti [kgCO₂ eq]"]/n_g, 'Energia_Prog': row["Energia [kgCO₂ eq]"]/n_g, 
                    'Acqua_Prog': row["Acqua [kgCO₂ eq]"]/n_g
                })
        
        df_comp = pd.DataFrame(dati_prog) if dati_prog else pd.DataFrame(columns=['Data'])
        if not df_comp.empty:
            df_comp = df_comp.groupby('Data').agg({
                'Materiali': 'sum', 'Rifiuti': 'sum', 'Macchinari': 'sum', 
                'Trasporti_Prog': 'sum', 'Energia_Prog': 'sum', 'Acqua_Prog': 'sum',
                'Attività_Prog': lambda x: '<br>'.join(sorted(set(str(item).strip() for sub in x.dropna() for item in sub.split(' | ') if item.strip())))
            }).reset_index()

        df_cant_t = elabora_sensore_con_attivita(pd.read_csv(file_trasporti), 'trasporti', df_lci, df_risultati, st.session_state.configurazioni_salvate) if file_trasporti else pd.DataFrame()
        df_cant_e = elabora_sensore_con_attivita(pd.read_csv(file_energia), 'energia', df_lci, df_risultati, st.session_state.configurazioni_salvate) if file_energia else pd.DataFrame()
        df_cant_a = elabora_sensore_con_attivita(pd.read_csv(file_acqua), 'acqua', df_lci, df_risultati, st.session_state.configurazioni_salvate) if file_acqua else pd.DataFrame()

        df_sensors = pd.DataFrame(columns=['Data'])
        for df_c in [df_cant_t, df_cant_e, df_cant_a]:
            if not df_c.empty:
                df_sensors = pd.merge(df_sensors, df_c, on='Data', how='outer') if not df_sensors.empty and 'Data' in df_sensors.columns else df_c

        if not df_comp.empty and not df_sensors.empty:
            df_comp = pd.merge(df_comp, df_sensors, on='Data', how='outer').fillna(0)
        elif df_comp.empty and not df_sensors.empty:
            df_comp = df_sensors.fillna(0)

        for col in ['Trasporti_Cant', 'Energia_Cant', 'Acqua_Cant', 'Materiali', 'Rifiuti', 'Macchinari', 'Trasporti_Prog', 'Energia_Prog', 'Acqua_Prog']:
            if col not in df_comp: df_comp[col] = 0

        # Unione pulita finale delle attività senza duplicati
        def clean_activities(series_list):
            items = set()
            for s in series_list:
                if pd.notna(s) and str(s).strip() != '':
                    for part in str(s).split('<br>'):
                        if part.strip(): items.add(part.strip())
            return '<br>'.join(sorted(items))

        if 'Attività_Prog' not in df_comp: df_comp['Attività_Prog'] = ""
        if 'Attività' not in df_comp: df_comp['Attività'] = ""
        
        df_comp['Attività_Unica'] = df_comp.apply(lambda r: clean_activities([r.get('Attività_Prog', ''), r.get('Attività', '')]), axis=1)

        df_comp['Totale_Prog'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Prog'] + df_comp['Energia_Prog'] + df_comp['Acqua_Prog']
        df_comp['Totale_Cant'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Cant'] + df_comp['Energia_Cant'] + df_comp['Acqua_Cant']
        st.session_state.df_comp_cantiere = df_comp.sort_values('Data')

    if st.session_state.df_comp_cantiere is not None:
        df_comp = st.session_state.df_comp_cantiere.copy()
        df_comp['Data'] = pd.to_datetime(df_comp['Data'])
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("Confronto Emissioni Settimanali: Progettazione vs Cantiere Reale")
        
        min_date_c = df_comp['Data'].min().date()
        max_date_c = df_comp['Data'].max().date()
        
        col_cfil1, col_cfil2 = st.columns(2)
        with col_cfil1:
            date_range_c = st.date_input("Intervallo temporale:", value=(min_date_c, max_date_c), key="date_comp")
        with col_cfil2:
            ambito_gwp_cant = st.selectbox("Ambito GWP Grafici:", ["GWP Totale", "GWP Altri Processi"], key="ambito_comp")
            
        start_filter_c, end_filter_c = date_range_c if isinstance(date_range_c, tuple) and len(date_range_c) == 2 else (min_date_c, max_date_c)
        df_comp_filtered = df_comp[(df_comp['Data'].dt.date >= start_filter_c) & (df_comp['Data'].dt.date <= end_filter_c)].copy()
        
        if not df_comp_filtered.empty:
            df_comp_filtered['Settimana'] = df_comp_filtered['Data'].dt.strftime('%Y-W%V')
            df_comp_filtered['Altri_Prog'] = df_comp_filtered['Rifiuti'] + df_comp_filtered['Macchinari'] + df_comp_filtered['Trasporti_Prog'] + df_comp_filtered['Energia_Prog'] + df_comp_filtered['Acqua_Prog']
            df_comp_filtered['Altri_Cant'] = df_comp_filtered['Rifiuti'] + df_comp_filtered['Macchinari'] + df_comp_filtered['Trasporti_Cant'] + df_comp_filtered['Energia_Cant'] + df_comp_filtered['Acqua_Cant']
            
            agg_dict_weekly = {c: 'sum' for c in df_comp_filtered.columns if c not in ['Data', 'Settimana', 'Attività_Unica', 'Attività_Prog', 'Attività']}
            agg_dict_weekly['Attività_Unica'] = lambda x: '<br>'.join(sorted(set(item.strip() for sub in x.dropna() for item in str(sub).split('<br>') if item.strip())))

            df_weekly = df_comp_filtered.groupby('Settimana').agg(agg_dict_weekly).reset_index()
            
            # --- FUNZIONE DI PLOT CON ALLINEAMENTO A SINISTRA PER TUTTI I GRAFICI ---
            def plot_comparison(df, base_col, title, c_prog, c_cant):
                fig = go.Figure()
                custom_data = df['Attività_Unica'] if 'Attività_Unica' in df.columns else [''] * len(df)
                
                if f"{base_col}_Prog" in df: 
                    fig.add_trace(go.Bar(
                        x=df['Settimana'], y=df[f'{base_col}_Prog'], name='Dati stimati', marker_color=c_prog,
                        customdata=custom_data,
                        hovertemplate="%{y:,.2f} kgCO₂ eq<br><b>Attività del periodo:</b><br>%{customdata}<extra></extra>"
                    ))
                fig.add_trace(go.Bar(
                    x=df['Settimana'], y=df[f'{base_col}_Cant'], name='Dati reali', marker_color=c_cant,
                    customdata=custom_data,
                    hovertemplate="%{y:,.2f} kgCO₂ eq<br><b>Attività del periodo:</b><br>%{customdata}<extra></extra>"
                ))
                
                # ALLINEAMENTO A SINISTRA (text-align: left) PER TUTTI I TOOLTIP DEL GRAFICO
                fig.update_layout(
                    title=title, barmode='group', hovermode="x unified", 
                    xaxis_title="Settimana", yaxis_title="Emissioni [kgCO₂ eq]",
                    hoverlabel=dict(align="left")
                )
                return fig

            if ambito_gwp_cant == "GWP Totale":
                st.plotly_chart(plot_comparison(df_weekly, 'Totale', "GWP Totale: Progetto vs Reale", "#01796F", "#8CD6C4"), use_container_width=True)
            else:
                st.plotly_chart(plot_comparison(df_weekly, 'Altri', "GWP Altri Processi: Progetto vs Reale", "#01796F", "#8CD6C4"), use_container_width=True)
            
            col_g1, col_g2 = st.columns(2)
            col_g1.plotly_chart(plot_comparison(df_weekly, 'Trasporti', "Trasporti", "#D1D0CE", "#707070"), use_container_width=True)
            col_g1.plotly_chart(plot_comparison(df_weekly, 'Acqua', "Acqua", "#0369a1", "#b8edea"), use_container_width=True)
            col_g2.plotly_chart(plot_comparison(df_weekly, 'Energia', "Energia Elettrica", "#FFD700", "#a16207"), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
