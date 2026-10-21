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
st.markdown("Monitoraggio Emissioni CO₂ eq in Cantiere (Fase di Progettazione e Costruzione)")

# --- INIZIALIZZAZIONE SESSION STATE ---
if 'risultati_progettazione' not in st.session_state:
    st.session_state.risultati_progettazione = []
if 'configurazioni_salvate' not in st.session_state:
    st.session_state.configurazioni_salvate = {}

colori_parametri = {
    'Materiali': "#991b1b", 
    'Rifiuti': "#166534", 
    'Trasporti': "#4338ca", 
    'Energia': "#ca8a04",
    'Acqua': "#0369a1", 
    'Macchinari': "#4b5563"
}

# --- FUNZIONI DI CARICAMENTO DATI STATICI ---
@st.cache_data
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
            err += "Manca database LCI.xlsx. "
            
        if os.path.exists("macchinari.xlsx"):
            df_macc = pd.read_excel("macchinari.xlsx")
        else:
            err += "Manca database macchinari.xlsx. "
            
    except Exception as e:
        err += f"Errore interno: {e}"
        
    return df_epd, df_lci, df_macc, err

df_epd, df_lci, df_macc, err_msg = load_static_data()
if err_msg:
    st.error(f"⚠️ {err_msg}")

# --- SEZIONE 1: CARICAMENTO DATI UTENTE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("1. Caricamento File di Progetto")

file_caricati = st.file_uploader("Carica CSV (Quantità Materiali e Cronoprogramma)", type=['csv'], accept_multiple_files=True)

lista_quantita, df_crono = [], None

if file_caricati:
    for f in file_caricati:
        try:
            content = f.getvalue().decode('utf-8-sig', errors='ignore')
            lines = content.splitlines()
            header_idx = next((i for i, line in enumerate(lines) if any(k in line.lower() for k in ['wbs', 'inizio', 'massa', 'attività', 'durata'])), 0)
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
            st.error(f"Errore {f.name}: {e}")

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
            col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'attività']), df_crono.columns[1])
            df_crono['Label_Attivita'] = df_crono[col_id].astype(str) + " - " + df_crono[col_nome].astype(str)
            st.success("File Progetto caricati con successo!")
        else:
            df_crono = None
st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 2: COMPILAZIONE ATTIVITA' ---
if df_crono is not None and df_epd is not None and df_lci is not None:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("2. Analisi Attività (Fase di Progettazione)")
    attivita_selezionata = st.selectbox("Scegli la lavorazione:", df_crono['Label_Attivita'].tolist())
    row_att = df_crono[df_crono['Label_Attivita'] == attivita_selezionata].iloc[0]
    
    id_att = row_att.get([c for c in df_crono.columns if 'id' in str(c).lower()][0])
    nome_att = row_att.get([c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'attività']][0])
    wbs_lav = row_att.get([c for c in df_crono.columns if 'wbs lav' in str(c).lower()][0] if [c for c in df_crono.columns if 'wbs lav' in str(c).lower()] else None)
    wbs_mat = row_att.get([c for c in df_crono.columns if 'wbs materiale' in str(c).lower()][0] if [c for c in df_crono.columns if 'wbs materiale' in str(c).lower()] else None)
    durata_giorni = int(row_att.get([c for c in df_crono.columns if 'durata' in str(c).lower()][0], 1)) if [c for c in df_crono.columns if 'durata' in str(c).lower()] else 1
    
    quantita_mat = 0
    if df_quantita is not None and wbs_mat is not None:
        q_col_lav = next((c for c in df_quantita.columns if 'wbs lav' in str(c).lower()), None)
        q_col_mat = next((c for c in df_quantita.columns if 'wbs mat' in str(c).lower()), None)
        q_col_massa = next((c for c in df_quantita.columns if 'massa' in str(c).lower() or 'peso' in str(c).lower() or 'quantit' in str(c).lower()), None)
        
        if q_col_lav and q_col_mat and q_col_massa:
            mask = (df_quantita[q_col_lav].astype(str).str.strip() == str(wbs_lav).strip()) & (df_quantita[q_col_mat].astype(str).str.strip() == str(wbs_mat).strip())
            match_q = df_quantita[mask]
            if not match_q.empty:
                massa_str = match_q[q_col_massa].astype(str).str.replace(r'[^\d,\.-]', '', regex=True)
                quantita_mat = pd.to_numeric(massa_str.replace({',': '.'}, regex=True), errors='coerce').sum()

    st.info(f"Quantità materiale stimata: **{quantita_mat:,.2f}** kg")
    
    # Valori di esempio (manterrai la tua logica dei dropdown come nel codice originario)
    em_materiali, em_rifiuti, em_trasporti, em_macchinari, em_energia, em_acqua = quantita_mat*1.5, 0, 0, 0, 0, 0
    
    # Simuliamo il salvataggio dei FE di default per mostrare il funzionamento dei calcoli
    tech_energia = "Idroelettrico - 360 MW"
    fe_energia = 0.0107
    tech_acqua = "Water supply"
    fe_acqua = 0.000191
    
    if st.button("Salva Calcolo per questa Attività", use_container_width=True):
        record = {
            "Attività": row_att['Label_Attivita'], "ID": str(id_att).strip(), "Nome": nome_att,
            "Data Inizio": row_att['Data_Inizio_Calc'].date(), "Data Fine": row_att['Data_Fine_Calc'].date(),
            "Materiali [kgCO₂ eq]": em_materiali, "Rifiuti [kgCO₂ eq]": em_rifiuti, 
            "Macchinari [kgCO₂ eq]": em_macchinari, "Trasporti [kgCO₂ eq]": em_trasporti,
            "Energia [kgCO₂ eq]": em_energia, "Acqua [kgCO₂ eq]": em_acqua, 
            "Totale [kgCO₂ eq]": sum([em_materiali, em_rifiuti, em_macchinari, em_trasporti, em_energia, em_acqua])
        }
        
        st.session_state.configurazioni_salvate[row_att['Label_Attivita']] = {
            "Energia": tech_energia, "FE_Energia": fe_energia,
            "Acqua": tech_acqua, "FE_Acqua": fe_acqua
        }
        
        lista_ris = [r for r in st.session_state.risultati_progettazione if r['Attività'] != row_att['Label_Attivita']]
        lista_ris.append(record)
        st.session_state.risultati_progettazione = sorted(lista_ris, key=lambda x: x['Data Inizio'])
        st.success("Attività salvata!")
    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 3: FASE DI CANTIERE (DATI SENSORI) ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("3. Monitoraggio Cantiere (Caricamento Sensori)")
    
    st.markdown("""
    <div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #27ae60; border-radius: 4px; margin-bottom: 20px;'>
        I dati dei sensori vengono letti tramite la colonna <strong>ID Attività</strong> incrociandoli con il database progettuale e il database LCI. 
    </div>
    """, unsafe_allow_html=True)
    
    file_sensori = st.file_uploader("Carica tutti i CSV Sensori (Trasporti, Energia, Acqua)", type=['csv'], accept_multiple_files=True)
    
    # Funzione che abbina Sensore -> ID Attività -> Calcolo Emissioni
    def elabora_sensore_con_attivita(df_sens, tipo_file, df_lci, df_ris, conf_salvate):
        risultati = []
        
        # Identificazione dinamica colonne
        col_id = next((c for c in df_sens.columns if 'id' in str(c).lower() and 'attiv' in str(c).lower()), df_sens.columns[0] if len(df_sens.columns)>0 else None)
        col_data = next((c for c in df_sens.columns if 'data' in str(c).lower()), None)
        
        if not col_data or not col_id:
            return pd.DataFrame(columns=['Data'])
            
        df_sens[col_data] = pd.to_datetime(df_sens[col_data], dayfirst=True, errors='coerce')
        df_sens = df_sens.dropna(subset=[col_data])
        
        for _, row in df_sens.iterrows():
            id_sensore = str(row[col_id]).strip()
            data_sensore = row[col_data].date()
            
            # Trovo l'attività corrispondente in progetto tramite l'ID
            match_prog = df_ris[df_ris['ID'].astype(str).str.strip() == id_sensore]
            if match_prog.empty:
                continue # L'ID nel sensore non corrisponde a nessuna lavorazione inserita a progetto
                
            label_attivita = match_prog.iloc[0]['Attività']
            conf = conf_salvate.get(label_attivita, {})
            
            emissione = 0
            if tipo_file == 'trasporti':
                km = pd.to_numeric(str(row.get('Km_percorsi', 0)).replace(',','.'))
                carb = str(row.get('Carburante', ''))
                
                df_t_fe = df_lci['Trasporti'].iloc[0:9]
                df_t_cons = df_lci['Trasporti'].iloc[11:21]
                
                match_fe = df_t_fe[df_t_fe.iloc[:, 1].astype(str).str.contains(carb, case=False, na=False)]
                fe = float(match_fe.iloc[0, 2]) if not match_fe.empty else 0.0
                
                match_cons = df_t_cons[df_t_cons.iloc[:, 1].astype(str).str.contains(carb, case=False, na=False)]
                consumo = pd.to_numeric(match_cons.iloc[:, 2], errors='coerce').mean() if not match_cons.empty else 0.1
                
                emissione = km * consumo * fe
                risultati.append({'Data': data_sensore, 'Trasporti_Cant': emissione})
                
            elif tipo_file == 'energia':
                kwh = pd.to_numeric(str(row.get('Consumi_kWh', row.get('quantità', 0))).replace(',','.'))
                alim = str(row.get('Alimentazione_elettrica', conf.get('Energia', '')))
                
                df_e = df_lci['Energia']
                match_fe = df_e[df_e.iloc[:, 0].astype(str).str.contains(alim, case=False, na=False)]
                fe = float(match_fe.iloc[0, 1]) if not match_fe.empty else conf.get('FE_Energia', 0)
                
                emissione = kwh * fe
                risultati.append({'Data': data_sensore, 'Energia_Cant': emissione})
                
            elif tipo_file == 'acqua':
                litri = pd.to_numeric(str(row.get('Consumi_l', row.get('quantità', 0))).replace(',','.'))
                tipo = str(row.get('Tipo_acqua', conf.get('Acqua', '')))
                
                df_a = df_lci['Acqua']
                match_fe = df_a[df_a.iloc[:, 0].astype(str).str.contains(tipo, case=False, na=False)]
                fe = float(match_fe.iloc[0, 1]) if not match_fe.empty else conf.get('FE_Acqua', 0)
                
                emissione = litri * fe
                risultati.append({'Data': data_sensore, 'Acqua_Cant': emissione})
                
        if risultati:
            df_res = pd.DataFrame(risultati)
            return df_res.groupby('Data').sum().reset_index() # Aggreghiamo giornalmente le emissioni
            
        return pd.DataFrame(columns=['Data'])

    if st.button("Elabora Dati Sensori e Genera Grafici di Confronto", use_container_width=True):
        df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
        
        # 1. Spalmiamo i dati di progetto sui giorni (Timeline progettuale)
        dati_prog = []
        for _, row in df_risultati.iterrows():
            giorni_lavoro = pd.bdate_range(start=row['Data Inizio'], end=row['Data Fine'])
            n_g = len(giorni_lavoro) if len(giorni_lavoro) > 0 else 1
            for g in giorni_lavoro:
                dati_prog.append({
                    'Data': g.date(), 'Materiali': row["Materiali [kgCO₂ eq]"]/n_g, 'Rifiuti': row["Rifiuti [kgCO₂ eq]"]/n_g,
                    'Macchinari': row["Macchinari [kgCO₂ eq]"]/n_g, 'Trasporti_Prog': row["Trasporti [kgCO₂ eq]"]/n_g,
                    'Energia_Prog': row["Energia [kgCO₂ eq]"]/n_g, 'Acqua_Prog': row["Acqua [kgCO₂ eq]"]/n_g
                })
        
        df_comp = pd.DataFrame(dati_prog).groupby('Data').sum().reset_index() if dati_prog else pd.DataFrame(columns=['Data'])
        
        # 2. Leggiamo i file sensore
        df_cant_t, df_cant_e, df_cant_a = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
        if file_sensori:
            for f in file_sensori:
                df_s = pd.read_csv(f)
                nome = f.name.lower()
                if 'trasporti' in nome:
                    df_cant_t = elabora_sensore_con_attivita(df_s, 'trasporti', df_lci, df_risultati, st.session_state.configurazioni_salvate)
                elif 'energ' in nome:
                    df_cant_e = elabora_sensore_con_attivita(df_s, 'energia', df_lci, df_risultati, st.session_state.configurazioni_salvate)
                elif 'acqua' in nome:
                    df_cant_a = elabora_sensore_con_attivita(df_s, 'acqua', df_lci, df_risultati, st.session_state.configurazioni_salvate)
        
        # 3. Uniamo progetto e cantiere per data
        df_sensors = pd.DataFrame(columns=['Data'])
        if not df_cant_t.empty: df_sensors = pd.merge(df_sensors, df_cant_t, on='Data', how='outer')
        if not df_cant_e.empty: df_sensors = pd.merge(df_sensors, df_cant_e, on='Data', how='outer')
        if not df_cant_a.empty: df_sensors = pd.merge(df_sensors, df_cant_a, on='Data', how='outer')
        
        if not df_comp.empty and not df_sensors.empty:
            df_comp = pd.merge(df_comp, df_sensors, on='Data', how='outer').fillna(0)
        elif df_comp.empty and not df_sensors.empty:
            df_comp = df_sensors.fillna(0)
            
        for col in ['Trasporti_Cant', 'Energia_Cant', 'Acqua_Cant', 'Materiali', 'Rifiuti', 'Macchinari', 'Trasporti_Prog', 'Energia_Prog', 'Acqua_Prog']:
            if col not in df_comp: df_comp[col] = 0
            
        df_comp['Totale_Prog'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Prog'] + df_comp['Energia_Prog'] + df_comp['Acqua_Prog']
        df_comp['Totale_Cant'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Cant'] + df_comp['Energia_Cant'] + df_comp['Acqua_Cant']
        df_comp = df_comp.sort_values('Data')

        # --- GRAFICI DI CONFRONTO ---
        st.markdown("<hr>", unsafe_allow_html=True)
        opzione_linee = st.selectbox("Linee di Tendenza (Confronto):", ["Nessuna", "Medie Costanti", "Medie Mobili (Andamento)", "Entrambe"])
        
        def plot_comparison(df, base_col, title, c_prog, c_cant):
            fig = go.Figure()
            if f"{base_col}_Prog" in df: fig.add_trace(go.Bar(x=df['Data'], y=df[f'{base_col}_Prog'], name='Progetto', marker_color=c_prog))
            fig.add_trace(go.Bar(x=df['Data'], y=df[f'{base_col}_Cant'], name='Cantiere (Reale)', marker_color=c_cant))
            
            if opzione_linee in ["Medie Costanti", "Entrambe"]:
                if f"{base_col}_Prog" in df: fig.add_hline(y=df[f'{base_col}_Prog'].mean(), line_dash="dash", line_color=c_prog)
                fig.add_hline(y=df[f'{base_col}_Cant'].mean(), line_dash="dash", line_color=c_cant)
            
            if opzione_linee in ["Medie Mobili (Andamento)", "Entrambe"]:
                if f"{base_col}_Prog" in df: fig.add_trace(go.Scatter(x=df['Data'], y=df[f'{base_col}_Prog'].rolling(3, min_periods=1).mean(), name='Trend Prog', line=dict(color=c_prog, width=2)))
                fig.add_trace(go.Scatter(x=df['Data'], y=df[f'{base_col}_Cant'].rolling(3, min_periods=1).mean(), name='Trend Cant', line=dict(color=c_cant, width=2)))
                
            fig.update_layout(title=title, barmode='group', hovermode="x unified")
            return fig

        st.plotly_chart(plot_comparison(df_comp, 'Totale', "GWP Totale: Progetto vs Reale", "#64748b", "#0f172a"), use_container_width=True)
        
        col_g1, col_g2 = st.columns(2)
        col_g1.plotly_chart(plot_comparison(df_comp, 'Trasporti', "Trasporti", "#818cf8", "#3730a3"), use_container_width=True)
        col_g1.plotly_chart(plot_comparison(df_comp, 'Acqua', "Acqua", "#38bdf8", "#0369a1"), use_container_width=True)
        col_g2.plotly_chart(plot_comparison(df_comp, 'Energia', "Energia Elettrica", "#fde047", "#a16207"), use_container_width=True)

        # --- ESPORTAZIONE EXCEL ---
        st.markdown("### Esportazione Dati")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_comp.to_excel(writer, sheet_name='Confronto Cantiere', index=False)
        st.download_button(
            label="📊 Scarica Database di Confronto (.xlsx)",
            data=buffer.getvalue(),
            file_name="Confronto_Progetto_Cantiere.xlsx",
            mime="application/vnd.ms-excel"
        )
    st.markdown("</div>", unsafe_allow_html=True)
