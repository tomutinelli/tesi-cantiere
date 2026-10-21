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
st.markdown("Monitoraggio Emissioni CO₂ eq in Cantiere (Progettazione vs Cantiere Reale)")

# --- INIZIALIZZAZIONE SESSION STATE ---
if 'risultati_progettazione' not in st.session_state:
    st.session_state.risultati_progettazione = []
if 'configurazioni_salvate' not in st.session_state:
    st.session_state.configurazioni_salvate = {}

# COLORI GRAFICI
colori_parametri = {
    'Materiali': "#991b1b", 'Rifiuti': "#166534", 'Trasporti': "#4338ca", 
    'Energia': "#ca8a04", 'Acqua': "#0369a1", 'Macchinari': "#4b5563"
}

# --- FUNZIONI DI CARICAMENTO DATI STATICI ---
def load_static_data():
    df_epd, df_lci, df_macc = None, None, None
    err = ""
    try:
        if os.path.exists("EPD materiali.xlsx"): df_epd = pd.read_excel("EPD materiali.xlsx")
        elif os.path.exists("epd_materiali.csv"): df_epd = pd.read_csv("epd_materiali.csv", sep=None, engine='python', encoding='latin1')
        else: err += "Manca il database EPD (EPD materiali.xlsx). "
            
        if os.path.exists("LCI.xlsx"):
            df_lci = pd.read_excel("LCI.xlsx", sheet_name=None)
            for k in df_lci.keys(): df_lci[k].columns = df_lci[k].columns.astype(str).str.strip()
        else: err += "Manca il database LCI.xlsx. "
            
        if os.path.exists("macchinari.xlsx"): df_macc = pd.read_excel("macchinari.xlsx")
        else: err += "Manca il database macchinari.xlsx. "
    except Exception as e: err += f"Errore interno di lettura database: {e}"
    return df_epd, df_lci, df_macc, err

df_epd, df_lci, df_macc, err_msg = load_static_data()
if err_msg: st.error(f"⚠️ {err_msg}")

# --- SEZIONE 1: CARICAMENTO DATI UTENTE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("1. Caricamento File di Progetto")

file_caricati = st.file_uploader("Carica i file CSV (Quantità Materiali e Cronoprogramma)", type=['csv'], accept_multiple_files=True)
lista_quantita = []
df_crono = None

if file_caricati:
    for f in file_caricati:
        try:
            content = f.getvalue().decode('utf-8-sig', errors='ignore')
            lines = content.splitlines()
            header_idx = next((i for i, line in enumerate(lines) if any(k in line.lower() for k in ['wbs', 'inizio', 'massa', 'attività', 'durata'])), 0)
            csv_data = "\n".join(lines[header_idx:])
            df_temp = pd.read_csv(io.StringIO(csv_data), sep=';')
            if len(df_temp.columns) < 2: df_temp = pd.read_csv(io.StringIO(csv_data), sep=',')
            df_temp.columns = df_temp.columns.str.strip()
            
            if any('inizio' in str(c).lower() or 'fine' in str(c).lower() for c in df_temp.columns): df_crono = df_temp
            else: lista_quantita.append(df_temp)
        except Exception as e: st.error(f"Errore lettura {f.name}: {e}")

df_quantita = pd.concat(lista_quantita, ignore_index=True) if lista_quantita else None

if df_crono is not None:
    col_inizio = next((c for c in df_crono.columns if 'inizio' in str(c).lower()), None)
    col_fine = next((c for c in df_crono.columns if 'fine' in str(c).lower()), None)
    
    if col_inizio and col_fine:
        df_crono['Data_Inizio_Calc'] = pd.to_datetime(df_crono[col_inizio], dayfirst=True, errors='coerce')
        df_crono['Data_Fine_Calc'] = pd.to_datetime(df_crono[col_fine], dayfirst=True, errors='coerce')
        df_crono = df_crono.dropna(subset=['Data_Inizio_Calc', 'Data_Fine_Calc']).sort_values(by='Data_Inizio_Calc')
        
        col_id = next((c for c in df_crono.columns if 'id' in str(c).lower()), df_crono.columns[0])
        col_struttura = next((c for c in df_crono.columns if 'struttura' in str(c).lower()), None)
        col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_crono.columns[1])
        df_crono['Label_Attivita'] = df_crono[col_id].astype(str) + " - " + df_crono[col_nome].astype(str)
        if df_quantita is not None: st.success("Cronoprogramma e materiali caricati!")
        else: st.warning("Cronoprogramma caricato, mancano materiali.")
st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 2: COMPILAZIONE ATTIVITA' ---
if df_crono is not None and df_epd is not None and df_lci is not None:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("2. Analisi Attività (Fase di Progettazione)")
    
    attivita_selezionata = st.selectbox("Scegli la lavorazione:", df_crono['Label_Attivita'].tolist())
    row_att = df_crono[df_crono['Label_Attivita'] == attivita_selezionata].iloc[0]
    
    id_att = row_att[col_id]
    nome_att = row_att[col_nome]
    nome_strutt = row_att[col_struttura] if col_struttura else ""
    col_wbs_lav = next((c for c in df_crono.columns if 'wbs lav' in str(c).lower()), None)
    col_wbs_mat = next((c for c in df_crono.columns if 'wbs mat' in str(c).lower()), None)
    wbs_lav = row_att[col_wbs_lav] if col_wbs_lav else ""
    wbs_mat = row_att[col_wbs_mat] if col_wbs_mat else None
    
    quantita_mat = 0
    if df_quantita is not None and wbs_mat is not None:
        q_col_lav = next((c for c in df_quantita.columns if 'wbs lav' in str(c).lower()), None)
        q_col_mat = next((c for c in df_quantita.columns if 'wbs mat' in str(c).lower()), None)
        q_col_massa = next((c for c in df_quantita.columns if 'massa' in str(c).lower() or 'peso' in str(c).lower() or 'quantit' in str(c).lower()), None)
        
        if q_col_lav and q_col_mat and q_col_massa:
            mask = (df_quantita[q_col_lav].astype(str).str.strip() == str(wbs_lav).strip()) & (df_quantita[q_col_mat].astype(str).str.strip() == str(wbs_mat).strip())
            match_q = df_quantita[mask]
            if not match_q.empty:
                quantita_mat = pd.to_numeric(match_q[q_col_massa].astype(str).str.replace(r'[^\d,\.-]', '', regex=True).replace({',': '.'}, regex=True), errors='coerce').fillna(0).sum()

    st.info(f"Quantità materiale: **{quantita_mat:,.2f}** kg")
    em_materiali, em_rifiuti, em_trasporti, em_energia, em_acqua, em_macchinari = 0, 0, 0, 0, 0, 0
    
    # Materiali e Rifiuti
    if quantita_mat > 0:
        materiale_scelto = st.selectbox("Tipo Materiale:", df_epd['Materiale'].dropna().unique())
        df_epd_filtrato = df_epd[df_epd['Materiale'] == materiale_scelto]
        col_epd_fe = next((c for c in df_epd_filtrato.columns if 'fattore' in str(c).lower()), df_epd_filtrato.columns[3])
        col_epd_c2 = next((c for c in df_epd_filtrato.columns if 'c2' in str(c).lower()), df_epd_filtrato.columns[4])
        
        scelta_specifica = st.selectbox("Prodotto EPD:", df_epd_filtrato['Produttore'])
        fattore_em_mat = float(str(df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_fe].values[0]).replace(',', '.'))
        fattore_em_rif = float(str(df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_c2].values[0]).replace(',', '.'))
        
        em_materiali = quantita_mat * fattore_em_mat
        perc_rifiuti = st.number_input("% Rifiuti stimata:", 0.0, 100.0, 0.0) / 100
        em_rifiuti = (quantita_mat * perc_rifiuti) * fattore_em_rif
        
        # Trasporti
        if 'Trasporti' in df_lci:
            df_t_consumi = df_lci['Trasporti'].iloc[11:21].copy().dropna(subset=[df_lci['Trasporti'].columns[0]])
            df_t_consumi.columns = ['Tipologia', 'Carburante', 'Consumo_Medio']
            df_t_consumi['Label'] = df_t_consumi['Tipologia'].astype(str) + " - " + df_t_consumi['Carburante'].astype(str)
            scelta_mezzo = st.selectbox("Mezzo Trasporto:", df_t_consumi['Label'].unique())
            consumo_medio = float(str(df_t_consumi[df_t_consumi['Label'] == scelta_mezzo]['Consumo_Medio'].values[0]).replace(',', '.'))
            fattore_carb = float(str(df_lci['Trasporti'].iloc[0:9][df_lci['Trasporti'].iloc[0:9].iloc[:,1].astype(str).str.contains(str(df_t_consumi[df_t_consumi['Label'] == scelta_mezzo]['Carburante'].values[0]).replace('CNG', 'CNG(or LNG)'), na=False)].iloc[0,2]).replace(',', '.'))
            
            viaggi = st.number_input("Viaggi stimati:", min_value=0, value=int(np.ceil(quantita_mat/28000)))
            distanza = st.number_input("Distanza media [km]:", min_value=0.0, value=30.0)
            em_trasporti = viaggi * distanza * consumo_medio * fattore_carb
            fe_trasporti_medio = (em_trasporti / (viaggi*distanza)) if viaggi*distanza > 0 else 0

    # Macchinari
    col_mac_fe = next((c for c in df_macc.columns if 'giorno' in str(c).lower()), df_macc.columns[3])
    macchinario = st.selectbox("Macchinario:", df_macc['Macchinario'].unique())
    df_macc_filt = df_macc[df_macc['Macchinario'] == macchinario]
    macc_produttore = st.selectbox("Modello Specifico:", df_macc_filt['Produttore'])
    fe_macc_8h = float(str(df_macc_filt[df_macc_filt['Produttore'] == macc_produttore][col_mac_fe].values[0]).replace(',', '.'))
    em_macchinari = st.number_input("Giorni di lavoro:", 1) * st.number_input("Num. Unità:", 1) * fe_macc_8h

    # Energia e Acqua
    col_energia_fe = df_lci['Energia'].columns[1]
    tech_energia = st.selectbox("Energia:", df_lci['Energia'][df_lci['Energia'].columns[0]].unique())
    fe_energia = float(str(df_lci['Energia'][df_lci['Energia'][df_lci['Energia'].columns[0]] == tech_energia][col_energia_fe].values[0]).replace(',', '.'))
    em_energia = st.number_input("Consumo stimato [kWh]:", 0.0) * fe_energia
    
    col_acqua_fe = df_lci['Acqua'].columns[1]
    tech_acqua = st.selectbox("Acqua:", df_lci['Acqua'][df_lci['Acqua'].columns[0]].unique())
    fe_acqua = float(str(df_lci['Acqua'][df_lci['Acqua'][df_lci['Acqua'].columns[0]] == tech_acqua][col_acqua_fe].values[0]).replace(',', '.'))
    em_acqua = st.number_input("Consumo idrico [l]:", 0.0) * fe_acqua

    if st.button("Salva Calcolo", use_container_width=True):
        st.session_state.risultati_progettazione = [r for r in st.session_state.risultati_progettazione if r['ID'] != id_att]
        st.session_state.risultati_progettazione.append({
            "Attività": row_att['Label_Attivita'], "ID": str(id_att).strip(),
            "Data Inizio": row_att['Data_Inizio_Calc'].date(), "Data Fine": row_att['Data_Fine_Calc'].date(),
            "Materiali [kgCO₂ eq]": em_materiali, "Rifiuti [kgCO₂ eq]": em_rifiuti, 
            "Macchinari [kgCO₂ eq]": em_macchinari, "Trasporti [kgCO₂ eq]": em_trasporti,
            "Energia [kgCO₂ eq]": em_energia, "Acqua [kgCO₂ eq]": em_acqua, 
            "Totale [kgCO₂ eq]": em_materiali + em_rifiuti + em_macchinari + em_trasporti + em_energia + em_acqua
        })
        st.session_state.configurazioni_salvate[str(id_att).strip()] = {
            "Consumo_mezzo": 1.0, "FE_Trasporti": fe_trasporti_medio,
            "FE_Energia": fe_energia, "FE_Acqua": fe_acqua
        }
        st.success("Salvato con successo!")
    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 3: RISULTATI PROGETTO ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("3. Sintesi Fase di Progettazione")
    df_risultati = pd.DataFrame(st.session_state.risultati_progettazione).sort_values(by='Data Inizio')
    st.dataframe(df_risultati, use_container_width=True)
    st.markdown(f"**GWP Totale Progetto:** {df_risultati['Totale [kgCO₂ eq]'].sum():,.2f} kg CO₂ eq")
    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 4: CANTIERE (DATI SENSORI) E CONFRONTO ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("4. Dati Sensori Cantiere & Confronto")
    
    st.info("""
    **Formato CSV Sensori Richiesto:** `ID attività`, `nome struttura`, `data`, `ID_sensore`, `quantità`.  
    *Nota: I valori calcolati in fase di progetto (Trasporti, Energia, Acqua) verranno sovrascritti dai dati reali dei sensori moltiplicati per i fattori emissivi scelti.*
    """)
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua", type=['csv'])
    
    # Prepara dataframe cantiere partendo da progetto (azzerando dinamici se verranno sovrascritti)
    df_cantiere = df_risultati.copy()
    df_cantiere['Trasporti [kgCO₂ eq]'] = 0.0
    df_cantiere['Energia [kgCO₂ eq]'] = 0.0
    df_cantiere['Acqua [kgCO₂ eq]'] = 0.0
    
    def applica_sensore(f_csv, col_target, key_fe):
        if f_csv:
            df_s = pd.read_csv(f_csv)
            # Ricerca flessibile delle colonne richieste
            c_id = next((c for c in df_s.columns if 'id' in str(c).lower() and 'sensore' not in str(c).lower()), df_s.columns[0])
            c_data = next((c for c in df_s.columns if 'data' in str(c).lower() or 'date' in str(c).lower()), None)
            c_qnt = next((c for c in df_s.columns if 'quantit' in str(c).lower() or 'consumo' in str(c).lower() or 'km' in str(c).lower()), df_s.columns[-1])
            
            if c_data: df_s['Data_Sensore'] = pd.to_datetime(df_s[c_data], errors='coerce').dt.date
            
            for idx, row in df_cantiere.iterrows():
                id_prog = str(row['ID']).strip()
                d_iniz, d_fin = row['Data Inizio'], row['Data Fine']
                
                if id_prog in st.session_state.configurazioni_salvate:
                    conf = st.session_state.configurazioni_salvate[id_prog]
                    mask = (df_s[c_id].astype(str).str.strip() == id_prog)
                    if c_data and 'Data_Sensore' in df_s.columns:
                        mask = mask & (df_s['Data_Sensore'] >= d_iniz) & (df_s['Data_Sensore'] <= d_fin)
                    
                    qnt_totale = pd.to_numeric(df_s[mask][c_qnt], errors='coerce').fillna(0).sum()
                    df_cantiere.loc[idx, col_target] = qnt_totale * conf[key_fe]

    applica_sensore(file_trasporti, 'Trasporti [kgCO₂ eq]', 'FE_Trasporti')
    applica_sensore(file_energia, 'Energia [kgCO₂ eq]', 'FE_Energia')
    applica_sensore(file_acqua, 'Acqua [kgCO₂ eq]', 'FE_Acqua')
    
    # Ricalcolo Totale Cantiere
    df_cantiere['Totale [kgCO₂ eq]'] = df_cantiere[['Materiali [kgCO₂ eq]', 'Rifiuti [kgCO₂ eq]', 'Macchinari [kgCO₂ eq]', 'Trasporti [kgCO₂ eq]', 'Energia [kgCO₂ eq]', 'Acqua [kgCO₂ eq]']].sum(axis=1)

    # TABELLA DELTA (DIFFERENZE)
    df_delta = pd.DataFrame({
        'Attività': df_risultati['Attività'],
        'Δ Totale [kgCO₂]': df_cantiere['Totale [kgCO₂ eq]'] - df_risultati['Totale [kgCO₂ eq]'],
        'Δ Trasporti [kgCO₂]': df_cantiere['Trasporti [kgCO₂ eq]'] - df_risultati['Trasporti [kgCO₂ eq]'],
        'Δ Energia [kgCO₂]': df_cantiere['Energia [kgCO₂ eq]'] - df_risultati['Energia [kgCO₂ eq]'],
        'Δ Acqua [kgCO₂]': df_cantiere['Acqua [kgCO₂ eq]'] - df_risultati['Acqua [kgCO₂ eq]']
    })
    st.markdown("##### Scostamento Emissioni (Cantiere Reale - Progetto)")
    st.dataframe(df_delta.style.applymap(lambda x: 'color: red' if x > 0 else 'color: green', subset=['Δ Totale [kgCO₂]', 'Δ Trasporti [kgCO₂]', 'Δ Energia [kgCO₂]', 'Δ Acqua [kgCO₂]']), use_container_width=True)

    # --- GRAFICI DI CONFRONTO ---
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("#### Grafici di Confronto (Progettazione vs Cantiere)")
    col_opt1, col_opt2 = st.columns(2)
    min_date = min(df_risultati['Data Inizio'].min(), df_cantiere['Data Inizio'].min())
    max_date = max(df_risultati['Data Fine'].max(), df_cantiere['Data Fine'].max())
    
    with col_opt1: range_date = st.date_input("Filtro Temporale:", value=(min_date, max_date))
    with col_opt2: linee = st.selectbox("Mostra Linee:", ["Istogrammi semplici", "Media Costante", "Media di Andamento (Mobile)", "Entrambe"])
    start_d, end_d = range_date if isinstance(range_date, tuple) and len(range_date) == 2 else (min_date, max_date)

    # Funzione Helper per spalmare i valori sui giorni feriali
    def aggrega_w(df, col_val):
        dati = []
        for _, r in df.iterrows():
            giorni = pd.bdate_range(r['Data Inizio'], r['Data Fine'])
            n = max(1, len(giorni))
            for g in giorni:
                if start_d <= g.date() <= end_d:
                    dati.append({"Settimana": g.strftime('%Y-W%V'), "Valore": r[col_val]/n})
        res = pd.DataFrame(dati)
        return res.groupby('Settimana')['Valore'].sum().reset_index() if not res.empty else pd.DataFrame(columns=['Settimana', 'Valore'])

    def plot_confronto(colonna, titolo):
        df_w_prog = aggrega_w(df_risultati, colonna).assign(Fase="Progettazione")
        df_w_cant = aggrega_w(df_cantiere, colonna).assign(Fase="Cantiere")
        df_plot = pd.concat([df_w_prog, df_w_cant])
        
        if df_plot.empty: return None
        
        fig = px.bar(df_plot, x="Settimana", y="Valore", color="Fase", barmode="group", title=titolo,
                     color_discrete_map={"Progettazione": "#3498db", "Cantiere": "#2ecc71"})
        
        if linee in ["Media Costante", "Entrambe"]:
            if not df_w_prog.empty: fig.add_hline(y=df_w_prog['Valore'].mean(), line_dash="dash", line_color="#2980b9")
            if not df_w_cant.empty: fig.add_hline(y=df_w_cant['Valore'].mean(), line_dash="dash", line_color="#27ae60")
        
        if linee in ["Media di Andamento (Mobile)", "Entrambe"]:
            if not df_w_prog.empty:
                df_w_prog['Media_Mobile'] = df_w_prog['Valore'].rolling(3, min_periods=1).mean()
                fig.add_trace(go.Scatter(x=df_w_prog['Settimana'], y=df_w_prog['Media_Mobile'], mode='lines', name='Andamento Prog.', line=dict(color='#2980b9', width=3)))
            if not df_w_cant.empty:
                df_w_cant['Media_Mobile'] = df_w_cant['Valore'].rolling(3, min_periods=1).mean()
                fig.add_trace(go.Scatter(x=df_w_cant['Settimana'], y=df_w_cant['Media_Mobile'], mode='lines', name='Andamento Cant.', line=dict(color='#27ae60', width=3)))
                
        fig.update_layout(yaxis_title="Emissioni [kgCO₂ eq]")
        return fig

    # Plot dei 4 grafici richiesti
    fig_tot = plot_confronto('Totale [kgCO₂ eq]', '1. Confronto TOTALE Emissioni')
    if fig_tot: st.plotly_chart(fig_tot, use_container_width=True)

    col_g1, col_g2 = st.columns(2)
    fig_tra = plot_confronto('Trasporti [kgCO₂ eq]', '2. Confronto TRASPORTI')
    if fig_tra: col_g1.plotly_chart(fig_tra, use_container_width=True)

    fig_ene = plot_confronto('Energia [kgCO₂ eq]', '3. Confronto ENERGIA')
    if fig_ene: col_g2.plotly_chart(fig_ene, use_container_width=True)
    
    col_g3, _ = st.columns(2)
    fig_acq = plot_confronto('Acqua [kgCO₂ eq]', '4. Confronto ACQUA')
    if fig_acq: col_g3.plotly_chart(fig_acq, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
