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

# COLORI GRAFICI (Allineati alla foto fornita)
colori_parametri = {
    'Materiali': "#991b1b", 
    'Rifiuti': "#166534", 
    'Trasporti': "#4338ca", 
    'Energia': "#ca8a04",
    'Acqua': "#0369a1", 
    'Macchinari': "#4b5563"
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

# RIQUADRO INFORMATIVO 1
st.markdown("""
<div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #2980b9; border-radius: 4px; margin-bottom: 20px;'>
    <strong>Istruzioni di Formattazione CSV (Fase di Progetto):</strong>
    <ul style='margin-top: 5px; margin-bottom: 0px;'>
        <li><strong>Formato Generale:</strong> I file devono essere <code>.csv</code> (separati da virgola <code>,</code> o punto e virgola <code>;</code>).</li>
        <li><strong>Cronoprogramma:</strong> Deve includere le colonne: <em>ID attività, Nome (o attività), Nome struttura, WBS Lavorazione, WBS Materiale, Data d'inizio, Data di fine</em>.</li>
        <li><strong>Abachi Quantità:</strong> Devono includere: <em>WBS Lavorazione, WBS Materiale, Struttura</em> (opzionale ma consigliata) e la colonna dei valori nominata con una di queste parole: <em>Massa, Peso, Quantità</em>.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

file_caricati = st.file_uploader("Carica i file CSV (Quantità Materiali e Cronoprogramma)", type=['csv'], accept_multiple_files=True)

lista_quantita = []
df_crono = None

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
            col_wbs = next((c for c in df_crono.columns if 'wbs lav' in str(c).lower()), None)
            col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_crono.columns[1])
   
            etichette = df_crono[col_id].astype(str) + " - " + df_crono[col_nome].astype(str)
            if col_struttura: etichette += " | " + df_crono[col_struttura].astype(str)
            
                
            df_crono['Label_Attivita'] = etichette
            
            if df_quantita is not None:
                st.success(f"Cronoprogramma caricato e {len(lista_quantita)} file dei materiali uniti con successo!")
            else:
                st.warning("Cronoprogramma caricato, ma non sono stati rilevati file delle quantità dei materiali.")
        else:
            st.error("Il cronoprogramma è stato letto, ma nessuna data era in un formato valido.")
            df_crono = None
    else:
        st.error("Impossibile trovare le colonne delle date nel cronoprogramma.")
else:
    st.info("Carica i file CSV per iniziare.")
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
    
    id_att = row_att[col_id]
    nome_att = row_att[col_nome]
    nome_strutt = row_att[col_struttura] if col_struttura else ""
    wbs_lav = row_att[col_wbs_lav] if col_wbs_lav else ""
    
    col_wbs_mat = next((c for c in df_crono.columns if 'wbs materiale' in str(c).lower()), None)
    wbs_mat = row_att[col_wbs_mat] if col_wbs_mat else None
    
    quantita_mat = 0
    if df_quantita is not None and wbs_mat is not None:
        q_col_lav = next((c for c in df_quantita.columns if 'wbs lav' in str(c).lower()), None)
        q_col_mat = next((c for c in df_quantita.columns if 'wbs mat' in str(c).lower()), None)
        
        q_col_struttura = next((c for c in df_quantita.columns if 'struttura' in str(c).lower()), None)
        
        q_col_massa = None
        for col in df_quantita.columns:
            col_lower = str(col).lower().strip()
            if col_lower == 'massa' or col_lower == 'peso' or 'quantit' in col_lower:
                q_col_massa = col
                break
        if not q_col_massa:
            q_col_massa = next((c for c in df_quantita.columns if 'volume' in str(c).lower()), None)
        
        if q_col_lav and q_col_mat and q_col_massa:
            lav_cron = str(wbs_lav).strip() if pd.notna(wbs_lav) else ""
            mat_cron = str(wbs_mat).strip() if pd.notna(wbs_mat) else ""
            struttura_cron = str(nome_strutt).strip() if pd.notna(nome_strutt) else ""
            
            lav_q_series = df_quantita[q_col_lav].fillna('').astype(str).str.strip()
            mat_q_series = df_quantita[q_col_mat].fillna('').astype(str).str.strip()
            
            mask = (lav_q_series == lav_cron) & (mat_q_series == mat_cron)
            
            if q_col_struttura and struttura_cron:
                struttura_q_series = df_quantita[q_col_struttura].fillna('').astype(str).str.strip()
                mask = mask & (struttura_q_series == struttura_cron)
            
            match_q = df_quantita[mask]
            
            if not match_q.empty:
                massa_str = match_q[q_col_massa].astype(str).str.replace(r'[^\d,\.-]', '', regex=True)
                massa_pulita = pd.to_numeric(massa_str.replace({',': '.'}, regex=True), errors='coerce').fillna(0)
                quantita_mat = massa_pulita.sum()

    
    st.info(f"Quantità materiale associata a questa attività (tramite WBS): **{quantita_mat:,.2f}** kg")
    
    em_materiali, em_rifiuti, em_trasporti, em_energia, em_acqua, em_macchinari = 0, 0, 0, 0, 0, 0
    
    # MATERIALI E RIFIUTI
    st.markdown("---")
    st.markdown("#### Materiali e Rifiuti")
    if quantita_mat > 0:
        materiale_scelto = st.selectbox("Seleziona Tipo Materiale (filtro categoria):", df_epd['Materiale'].dropna().unique())
        df_epd_filtrato = df_epd[df_epd['Materiale'] == materiale_scelto]
        
        col_epd_fe = next((c for c in df_epd_filtrato.columns if 'fattore' in str(c).lower()), df_epd_filtrato.columns[3])
        col_epd_c2 = next((c for c in df_epd_filtrato.columns if 'c2' in str(c).lower()), df_epd_filtrato.columns[4])
        col_epd_link = next((c for c in df_epd_filtrato.columns if 'epd' in str(c).lower() or 'link' in str(c).lower()), df_epd_filtrato.columns[-1])
        
        st.dataframe(
            df_epd_filtrato[['Materiale', 'Produttore', col_epd_fe, col_epd_c2, col_epd_link]],
            column_config={col_epd_link: st.column_config.LinkColumn("Link EPD")}, hide_index=True
        )
        
        scelta_specifica = st.selectbox("Scegli il prodotto specifico per il calcolo:", df_epd_filtrato['Produttore'])
        fattore_em_mat = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_fe].values[0]
        fattore_em_rif = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica][col_epd_c2].values[0]
        
        fattore_em_mat = float(str(fattore_em_mat).replace(',', '.'))
        fattore_em_rif = float(str(fattore_em_rif).replace(',', '.'))
        
        em_materiali = quantita_mat * fattore_em_mat
        
        perc_rifiuti = st.number_input("% di rifiuti stimata (sfrido):", min_value=0.0, max_value=100.0, value=0.0) / 100
        q_rifiuti = quantita_mat * perc_rifiuti
        em_rifiuti = q_rifiuti * fattore_em_rif
        
        df_sum_mat = pd.DataFrame({
            "Voce": ["Materiali", "Rifiuti"],
            "Quantità [kg]": [f"{quantita_mat:,.2f} kg", f"{q_rifiuti:,.2f}"],
            "Fattore Emissivo [kgCO₂ eq/kg]": [f"{fattore_em_mat:,.4f}", f"{fattore_em_rif:,.4f}"],
            "Emissioni [kgCO₂ eq]": [f"{em_materiali:,.2f}", f"{em_rifiuti:,.2f}"]
        })
        st.markdown("**Riepilogo Materiali e Rifiuti:**")
        st.dataframe(df_sum_mat, hide_index=True)
        
        # TRASPORTI MULTIPLI
        st.markdown("---")
        st.markdown("#### Trasporti")
        if 'Trasporti' in df_lci:
            df_t_sheet = df_lci['Trasporti']
            
            df_t_emissions = df_t_sheet.iloc[0:9].copy()
            df_t_emissions.columns = ['Tipologia', 'Carburante', 'Fattore_Emissione']
            
            df_t_consumi = df_t_sheet.iloc[11:21].copy()
            df_t_consumi.columns = ['Tipologia', 'Carburante', 'Consumo_Medio']
            df_t_consumi = df_t_consumi.dropna(subset=['Tipologia'])
            df_t_consumi['Label'] = df_t_consumi['Tipologia'].astype(str) + " - " + df_t_consumi['Carburante'].astype(str)
            
            viaggi_mat = np.ceil(quantita_mat / 28000)
            viaggi_rif = np.ceil(q_rifiuti / 28000) if q_rifiuti > 0 else 0
            viaggi_tot_stimati = int(viaggi_mat + viaggi_rif)
            
            num_tipi_trasporto = st.number_input("Quanti tipi di mezzi di trasporto verranno impiegati?", min_value=1, value=1, step=1)
            
            em_trasporti = 0
            tot_km_progetto = 0
            dati_trasporto_tabella = []
            
            for i in range(num_tipi_trasporto):
                st.markdown(f"**Mezzo di Trasporto {i+1}**")
                scelta_mezzo = st.selectbox(f"Seleziona Veicolo/Carburante {i+1}:", df_t_consumi['Label'].unique(), key=f"t_mezzo_{i}")
                row_mezzo = df_t_consumi[df_t_consumi['Label'] == scelta_mezzo].iloc[0]
                
                consumo_medio = float(str(row_mezzo['Consumo_Medio']).replace(',', '.'))
                carb_selezionato = str(row_mezzo['Carburante'])
                carb_match = 'CNG(or LNG)' if 'CNG' in carb_selezionato else carb_selezionato
                
                row_em = df_t_emissions[df_t_emissions['Carburante'].astype(str) == carb_match]
                fattore_carb = float(str(row_em['Fattore_Emissione'].values[0]).replace(',', '.')) if not row_em.empty else 0.0
                
                col_t1, col_t2 = st.columns(2)
                with col_t1:
                    viaggi = st.number_input(f"Viaggi stimati (mezzo {i+1}):", min_value=0, value=viaggi_tot_stimati if i==0 else 0, key=f"t_viaggi_{i}")
                with col_t2:
                    distanza = st.number_input(f"Distanza media a viaggio [km] (mezzo {i+1}):", min_value=0.0, value=30.0, key=f"t_dist_{i}")
                
                em_trasporto_i = viaggi * distanza * consumo_medio * fattore_carb
                em_trasporti += em_trasporto_i
                tot_km_progetto += (viaggi * distanza)
                
                dati_trasporto_tabella.append({
                    "Mezzo": scelta_mezzo,
                    "Viaggi": viaggi,
                    "Distanza [km]": distanza,
                    "Consumo [kg/km]": consumo_medio,
                    "Emissioni [kgCO₂ eq]": round(em_trasporto_i, 2)
                })

            fe_trasporti_medio = (em_trasporti / tot_km_progetto) if tot_km_progetto > 0 else 0
            
            st.markdown("**Riepilogo Trasporti:**")
            st.dataframe(pd.DataFrame(dati_trasporto_tabella), hide_index=True)
            st.write(f"**Totale Emissioni Trasporti:** {em_trasporti:,.2f} kgCO₂ eq")
            
    else:
        st.warning("Nessun materiale trovato per questa attività (Quantità = 0). Sezioni Materiali, Rifiuti e Trasporti avranno emissioni pari a 0.")
        fe_trasporti_medio = 0
        
    # MACCHINARI MULTIPLI
    st.markdown("---")
    st.markdown("#### Macchinari")
    col_mac_fe = next((c for c in df_macc.columns if 'giorno' in str(c).lower()), df_macc.columns[3] if len(df_macc.columns) > 3 else None)
    col_durata = next((c for c in df_crono.columns if 'durata' in str(c).lower()), None)
    durata_giorni = int(row_att[col_durata]) if col_durata else 1
    
    num_tipi_macc = st.number_input("Quanti tipi di macchinari verranno impiegati?", min_value=1, value=1, step=1)
    
    em_macchinari = 0
    dati_macchinari_tabella = []
    
    for i in range(num_tipi_macc):
        st.markdown(f"**Macchinario {i+1}**")
        macchinario = st.selectbox(f"Scegli Macchinario base {i+1}:", df_macc['Macchinario'].unique(), key=f"m_tipo_{i}")
        df_macc_filt = df_macc[df_macc['Macchinario'] == macchinario]
        
        if col_mac_fe:
            # TABELLA SCHEDE TECNICHE RIPRISTINATA (come richiesto)
            col_macc_link = next((c for c in df_macc_filt.columns if 'link' in str(c).lower()), df_macc_filt.columns[-1])
            st.dataframe(
                df_macc_filt[['Macchinario', 'Produttore', col_mac_fe, col_macc_link]],
                column_config={col_macc_link: st.column_config.LinkColumn("Link Schede Tecniche")},
                hide_index=True
            )
            
            macc_produttore = st.selectbox(f"Modello Specifico {i+1}:", df_macc_filt['Produttore'], key=f"m_prod_{i}")
            fe_macc_8h = float(str(df_macc_filt[df_macc_filt['Produttore'] == macc_produttore][col_mac_fe].values[0]).replace(',', '.'))
        else:
            st.warning("Impossibile trovare la colonna delle emissioni nel file macchinari.")
            fe_macc_8h = 0.0
            macc_produttore = "N/A"
            
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            num_macc = st.number_input(f"Numero di unità (Macchinario {i+1}):", min_value=1, value=1, key=f"m_unita_{i}")
        with col_m2:
            giorni_macc = st.number_input(f"Giorni di lavoro (Macchinario {i+1}):", min_value=1, value=durata_giorni, key=f"m_giorni_{i}")
            
        em_macc_i = fe_macc_8h * num_macc * giorni_macc
        em_macchinari += em_macc_i
        
        # Tabella riassuntiva Macchinari - TITOLI ORIGINALI RIPRISTINATI
        dati_macchinari_tabella.append({
            "Voce": f"Macchinario {i+1}",
            "Num. Macchine": num_macc,
            "Giorni": giorni_macc,
            "Fattore Emissivo [kgCO₂ eq/giorno]": f"{fe_macc_8h:,.2f}",
            "Emissioni [kgCO₂ eq]": f"{em_macc_i:,.2f}"
        })
        
    st.markdown("**Riepilogo Macchinari:**")
    st.dataframe(pd.DataFrame(dati_macchinari_tabella), hide_index=True)
    st.write(f"**Totale Emissioni Macchinari:** {em_macchinari:,.2f} kgCO₂ eq")
    
    # ENERGIA
    st.markdown("---")
    st.markdown("#### Energia Elettrica")
    col_energia_fe = next((c for c in df_lci['Energia'].columns if 'fattore' in str(c).lower()), df_lci['Energia'].columns[1])
    col_energia_nome = df_lci['Energia'].columns[0]
    
    tech_energia = st.selectbox("Tecnologia generazione elettrica:", df_lci['Energia'][col_energia_nome].unique())
    fe_energia = float(str(df_lci['Energia'][df_lci['Energia'][col_energia_nome] == tech_energia][col_energia_fe].values[0]).replace(',', '.'))
    kwh_stimati = st.number_input("Consumo elettrico stimato (kWh):", value=0.0)
    em_energia = kwh_stimati * fe_energia
    
    df_sum_ene = pd.DataFrame({
        "Voce": ["Energia"],
        "Consumo Totale [kWh]": [f"{kwh_stimati}"],
        "Fattore Emissivo [kgCO₂ eq/kWh]": [f"{fe_energia:,.4f}"],
        "Emissioni [kgCO₂ eq]": [f"{em_energia:,.2f}"]
    })
    st.markdown("**Riepilogo Energia:**")
    st.dataframe(df_sum_ene, hide_index=True)
    
    # ACQUA
    st.markdown("---")
    st.markdown("#### Acqua")
    col_acqua_fe = next((c for c in df_lci['Acqua'].columns if 'fattore' in str(c).lower()), df_lci['Acqua'].columns[1])
    col_acqua_nome = df_lci['Acqua'].columns[0]
    
    tech_acqua = st.selectbox("Approvvigionamento/Trattamento Acqua:", df_lci['Acqua'][col_acqua_nome].unique())
    fe_acqua = float(str(df_lci['Acqua'][df_lci['Acqua'][col_acqua_nome] == tech_acqua][col_acqua_fe].values[0]).replace(',', '.'))
    litri_stimati = st.number_input("Consumo idrico stimato (litri):", value=0.0)
    em_acqua = litri_stimati * fe_acqua
    
    df_sum_acq = pd.DataFrame({
        "Voce": ["Acqua"],
        "Consumo Totale [l]": [f"{litri_stimati}"],
        "Fattore Emissivo [kgCO₂ eq/l]": [f"{fe_acqua:,.4f}"],
        "Emissioni [kgCO₂ eq]": [f"{em_acqua:,.2f}"]
    })
    st.markdown("**Riepilogo Acqua:**")
    st.dataframe(df_sum_acq, hide_index=True)
    
    st.markdown("---")
    if st.button("Salva Calcolo per questa Attività", use_container_width=True):
        totale_eq = em_materiali + em_rifiuti + em_macchinari + em_trasporti + em_energia + em_acqua
        
        record = {
            "Attività": row_att['Label_Attivita'],
            "ID": id_att, "Nome": nome_att, "Struttura": nome_strutt, "WBS_L": wbs_lav, "WBS_M": wbs_mat,
            "Durata": durata_giorni, 
            "Data Inizio": row_att['Data_Inizio_Calc'].date(), 
            "Data Fine": row_att['Data_Fine_Calc'].date(),
            "Materiali [kgCO₂ eq]": em_materiali, "Rifiuti [kgCO₂ eq]": em_rifiuti, 
            "Macchinari [kgCO₂ eq]": em_macchinari, "Trasporti [kgCO₂ eq]": em_trasporti,
            "Energia [kgCO₂ eq]": em_energia, "Acqua [kgCO₂ eq]": em_acqua, 
            "Totale [kgCO₂ eq]": totale_eq
        }
        
        st.session_state.configurazioni_salvate[row_att['Label_Attivita']] = {
            "Carburante": "Mix" if 'Trasporti' in df_lci else None,
            "Consumo_mezzo": 1.0 if 'Trasporti' in df_lci else 0, # Impostato a 1, assorbe il fattore effettivo calcolato giù
            "FE_Trasporti": fe_trasporti_medio if 'Trasporti' in df_lci else 0,
            "Energia": tech_energia, "FE_Energia": fe_energia,
            "Acqua": tech_acqua, "FE_Acqua": fe_acqua
        }
        
        lista_ris = [r for r in st.session_state.risultati_progettazione if r['Attività'] != row_att['Label_Attivita']]
        lista_ris.append(record)
        st.session_state.risultati_progettazione = sorted(lista_ris, key=lambda x: x['Data Inizio'])
        st.success("Attività salvata con successo! Scorri in basso per i risultati aggiornati.")

    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 3: RISULTATI E GRAFICI PROGETTAZIONE ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("3. Risultati Fase di Progettazione")
    
    df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
    st.dataframe(df_risultati, use_container_width=True)
    
    # SPLIT TOTALI PROGETTAZIONE
    gwp_totale = df_risultati["Totale [kgCO₂ eq]"].sum()
    gwp_embodied = df_risultati["Materiali [kgCO₂ eq]"].sum()
    gwp_altre = gwp_totale - gwp_embodied
    
    st.markdown(f"""
    <div style='padding:15px; background-color:#e8f4f8; border-left: 5px solid #2980b9; border-radius:4px;'>
        <h4>Risultati GWP (Stima Progettuale)</h4>
        <ul style='font-size: 16px;'>
            <li><strong>GWP Embodied carbon (Solo Materiali):</strong> {gwp_embodied:,.2f} kg CO₂ eq</li>
            <li><strong>GWP Altri Processi (Trasporti, Macchinari, Energia, ecc.):</strong> {gwp_altre:,.2f} kg CO₂ eq</li>
            <li><strong style='font-size: 18px;'>GWP Totale Cantiere:</strong> <span style='font-size: 18px;'>{gwp_totale:,.2f} kg CO₂ eq</span></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("#### Filtro Temporale Grafici e Opzioni")
    col_fil1, col_fil2 = st.columns([1, 1])
    
    min_date = df_risultati['Data Inizio'].min()
    max_date = df_risultati['Data Fine'].max()
    
    with col_fil1:
        date_range = st.date_input(
            "Seleziona l'intervallo temporale:", 
            value=(min_date, max_date), min_value=min_date, max_value=max_date
        )
    with col_fil2:
        opzione_linee = st.selectbox("Linee di Tendenza:", ["Nessuna", "Media Costante", "Andamento Medio (Media Mobile)", "Entrambe"])
    
    mostra_media_costante = opzione_linee in ["Media Costante", "Entrambe"]
    mostra_media_mobile = opzione_linee in ["Andamento Medio (Media Mobile)", "Entrambe"]
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_filter, end_filter = date_range
    else:
        start_filter, end_filter = min_date, max_date

    dati_giornalieri = []
    for _, row in df_risultati.iterrows():
        giorni_lavoro = pd.bdate_range(start=row['Data Inizio'], end=row['Data Fine'])
        n_giorni = len(giorni_lavoro) if len(giorni_lavoro) > 0 else 1
        
        for g in giorni_lavoro:
            g_date = g.date()
            if start_filter <= g_date <= end_filter:
                dati_giornalieri.append({
                    "Attività": row['Attività'], 
                    "Data": g_date, 
                    "Settimana": g.strftime('%Y-W%V'),
                    "Materiali": row["Materiali [kgCO₂ eq]"]/n_giorni, "Rifiuti": row["Rifiuti [kgCO₂ eq]"]/n_giorni,
                    "Macchinari": row["Macchinari [kgCO₂ eq]"]/n_giorni, "Trasporti": row["Trasporti [kgCO₂ eq]"]/n_giorni,
                    "Energia": row["Energia [kgCO₂ eq]"]/n_giorni, "Acqua": row["Acqua [kgCO₂ eq]"]/n_giorni,
                    "Totale": row["Totale [kgCO₂ eq]"]/n_giorni
                })
    
    if dati_giornalieri:
        df_giornaliero = pd.DataFrame(dati_giornalieri)
        df_melted = df_giornaliero.melt(
            id_vars=['Settimana', 'Attività'], 
            value_vars=['Materiali', 'Rifiuti', 'Macchinari', 'Trasporti', 'Energia', 'Acqua'], 
            var_name='Parametro', value_name='Emissioni [kgCO₂ eq]'
        )
        
        df_plot = df_melted.groupby(['Settimana', 'Parametro', 'Attività'])['Emissioni [kgCO₂ eq]'].sum().reset_index()
        df_plot = df_plot[df_plot['Emissioni [kgCO₂ eq]'] > 0]
        
        df_pie = pd.DataFrame({
            "Parametro": ["Materiali", "Rifiuti", "Macchinari", "Trasporti", "Energia", "Acqua"],
            "Valore": [df_risultati["Materiali [kgCO₂ eq]"].sum(), df_risultati["Rifiuti [kgCO₂ eq]"].sum(), df_risultati["Macchinari [kgCO₂ eq]"].sum(), df_risultati["Trasporti [kgCO₂ eq]"].sum(), df_risultati["Energia [kgCO₂ eq]"].sum(), df_risultati["Acqua [kgCO₂ eq]"].sum()]
        })
        fig_pie = px.pie(df_pie, names="Parametro", values="Valore", title="Incidenza Percentuale Totale Cantiere", color="Parametro", color_discrete_map=colori_parametri, hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("#### Grafici Settimanali Dettagliati (Passa il mouse sulle barre per vedere l'Attività)")

        df_totale_assoluto = df_plot.groupby(['Settimana', 'Attività'])['Emissioni [kgCO₂ eq]'].sum().reset_index()
        df_week_tot = df_totale_assoluto.groupby('Settimana')['Emissioni [kgCO₂ eq]'].sum().reset_index()
        
        # ANDAMENTO COMPLESSIVO
        fig_totale_assoluto = px.bar(
            df_totale_assoluto, x="Settimana", y="Emissioni [kgCO₂ eq]", hover_data=["Attività"],
            title="Andamento Complessivo Totale", color_discrete_sequence=["#111827"]
        )
        if mostra_media_costante:
            media_tot = df_week_tot['Emissioni [kgCO₂ eq]'].mean()
            fig_totale_assoluto.add_hline(y=media_tot, line_dash="dash", line_color="#ef4444", annotation_text="Media Costante")
        if mostra_media_mobile:
            df_week_tot['Media_Mobile'] = df_week_tot['Emissioni [kgCO₂ eq]'].rolling(window=3, min_periods=1).mean()
            fig_totale_assoluto.add_trace(go.Scatter(x=df_week_tot['Settimana'], y=df_week_tot['Media_Mobile'], mode='lines', name='Andamento Medio', line=dict(color='#f59e0b', width=3)))
        st.plotly_chart(fig_totale_assoluto, use_container_width=True)

        # GRAFICO CUMULATO PER PARAMETRI
        fig_tot = px.bar(
            df_plot, x="Settimana", y="Emissioni [kgCO₂ eq]", color="Parametro", 
            hover_data=["Attività"], title="Andamento Settimanale Cumulato (Suddiviso per Parametro)",
            color_discrete_map=colori_parametri, barmode='stack'
        )
        if mostra_media_costante:
            fig_tot.add_hline(y=media_tot, line_dash="dash", line_color="#ef4444", annotation_text="Media Costante")
        if mostra_media_mobile:
            fig_tot.add_trace(go.Scatter(x=df_week_tot['Settimana'], y=df_week_tot['Media_Mobile'], mode='lines', name='Andamento Medio', line=dict(color='#f59e0b', width=3)))
        st.plotly_chart(fig_tot, use_container_width=True)

        st.markdown("##### Dettaglio per Singolo Parametro")
        col_g1, col_g2 = st.columns(2)
        
        parametri_lista = ['Materiali', 'Rifiuti', 'Macchinari', 'Trasporti', 'Energia', 'Acqua']
        for i, param in enumerate(parametri_lista):
            df_param = df_plot[df_plot['Parametro'] == param]
            if not df_param.empty:
                df_week_param = df_param.groupby('Settimana')['Emissioni [kgCO₂ eq]'].sum().reset_index()
                
                fig_p = px.bar(
                    df_param, x="Settimana", y="Emissioni [kgCO₂ eq]", hover_data=["Attività"],
                    title=f"Emissione: {param}", color="Parametro", color_discrete_map=colori_parametri
                )
                fig_p.update_layout(showlegend=False)
                
                if mostra_media_costante:
                    media_p = df_week_param['Emissioni [kgCO₂ eq]'].mean()
                    fig_p.add_hline(y=media_p, line_dash="dash", line_color="#ef4444")
                if mostra_media_mobile:
                    df_week_param['Media_Mobile'] = df_week_param['Emissioni [kgCO₂ eq]'].rolling(window=3, min_periods=1).mean()
                    fig_p.add_trace(go.Scatter(x=df_week_param['Settimana'], y=df_week_param['Media_Mobile'], mode='lines', name='Andamento Medio', line=dict(color='#f59e0b', width=2)))
                
                if i % 2 == 0: col_g1.plotly_chart(fig_p, use_container_width=True)
                else: col_g2.plotly_chart(fig_p, use_container_width=True)
    else:
        st.warning("Non ci sono attività nel periodo selezionato dal filtro.")

    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 4: FASE DI CANTIERE (DATI SENSORI) ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("4. Monitoraggio Fase di Cantiere (Dati Dinamici Sensori)")
    
    # RIQUADRO INFORMATIVO 2
    st.markdown("""
    <div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #27ae60; border-radius: 4px; margin-bottom: 20px;'>
        <strong>Istruzioni di Formattazione CSV (Dati Sensori):</strong>
        <ul style='margin-top: 5px; margin-bottom: 0px;'>
            <li><strong>Intestazione:</strong> Tutte le tabelle devono riportare le colonne <em>ID attività, Nome (o attività), Struttura, WBS Lavorazione</em> per potersi allacciare al calcolo in progetto.</li>
            <li><strong>Dati Specifici da inserire:</strong></li>
            <ul>
                <li>    <strong>Trasporti:</strong> Inserire colonna nominata <em>km percorsi</em>.</li>
                <li>    <strong>Energia:</strong> Inserire colonna nominata <em>consumo energetico [kWh]</em>.</li>
                <li>    <strong>Acqua:</strong> Inserire colonna nominata <em>consumo acqua [l]</em> o <em>[litri]</em>.</li>
            </ul>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti (CSV)", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia (CSV)", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua (CSV)", type=['csv'])
    
    if st.button("Ricalcola Emissioni con Dati di Cantiere"):
        df_cantiere = df_risultati.copy()
        
        if file_trasporti:
            df_t = pd.read_csv(file_trasporti)
            col_id = next((c for c in df_t.columns if 'id' in str(c).lower()), df_t.columns[1])
            col_nome = next((c for c in df_t.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_t.columns[2] if len(df_t.columns)>2 else None)
            col_str = next((c for c in df_t.columns if 'struttura' in str(c).lower()), None)
            col_wbs = next((c for c in df_t.columns if 'wbs lav' in str(c).lower()), None)
            
            etichette_t = df_t[col_id].astype(str) + " - " + df_t[col_nome].astype(str)
            if col_str: etichette_t += " | " + df_t[col_str].astype(str)
            if col_wbs: etichette_t += " (" + df_t[col_wbs].astype(str) + ")"
            
            df_t['Label_Attivita'] = etichette_t
            
            for label in df_t['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    conf = st.session_state.configurazioni_salvate[label]
                    km_totali = df_t[df_t['Label_Attivita'] == label]['km percorsi'].sum()
                    if conf['Carburante'] is not None:
                        # Nota: Consumo mezzo è salvato come 1, e FE_Trasporti incorpora il fattore moltiplicatore per ogni km
                        df_cantiere.loc[df_cantiere['Attività'] == label, 'Trasporti [kgCO₂ eq]'] = km_totali * conf['Consumo_mezzo'] * conf['FE_Trasporti']
        
        if file_energia:
            df_e = pd.read_csv(file_energia)
            col_id = next((c for c in df_e.columns if 'id' in str(c).lower()), df_e.columns[1])
            col_nome = next((c for c in df_e.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_e.columns[2] if len(df_e.columns)>2 else None)
            col_str = next((c for c in df_e.columns if 'struttura' in str(c).lower()), None)
            col_wbs = next((c for c in df_e.columns if 'wbs lav' in str(c).lower()), None)
            
            etichette_e = df_e[col_id].astype(str) + " - " + df_e[col_nome].astype(str)
            if col_str: etichette_e += " | " + df_e[col_str].astype(str)
            if col_wbs: etichette_e += " (" + df_e[col_wbs].astype(str) + ")"
            
            df_e['Label_Attivita'] = etichette_e
            for label in df_e['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    kwh_totali = df_e[df_e['Label_Attivita'] == label]['consumo energetico [kWh]'].sum()
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Energia [kgCO₂ eq]'] = kwh_totali * st.session_state.configurazioni_salvate[label]['FE_Energia']
                    
        if file_acqua:
            df_a = pd.read_csv(file_acqua)
            col_id = next((c for c in df_a.columns if 'id' in str(c).lower()), df_a.columns[1])
            col_nome = next((c for c in df_a.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_a.columns[2] if len(df_a.columns)>2 else None)
            col_str = next((c for c in df_a.columns if 'struttura' in str(c).lower()), None)
            col_wbs = next((c for c in df_a.columns if 'wbs lav' in str(c).lower()), None)
            
            etichette_a = df_a[col_id].astype(str) + " - " + df_a[col_nome].astype(str)
            if col_str: etichette_a += " | " + df_a[col_str].astype(str)
            if col_wbs: etichette_a += " (" + df_a[col_wbs].astype(str) + ")"
            
            df_a['Label_Attivita'] = etichette_a
            for label in df_a['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    l_totali = df_a[df_a['Label_Attivita'] == label]['consumo acqua [l]'].sum()
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Acqua [kgCO₂ eq]'] = l_totali * st.session_state.configurazioni_salvate[label]['FE_Acqua']
        
        df_cantiere['Totale [kgCO₂ eq]'] = df_cantiere[['Materiali [kgCO₂ eq]', 'Rifiuti [kgCO₂ eq]', 'Macchinari [kgCO₂ eq]', 'Trasporti [kgCO₂ eq]', 'Energia [kgCO₂ eq]', 'Acqua [kgCO₂ eq]']].sum(axis=1)
        
        st.dataframe(df_cantiere, use_container_width=True)
        
        # SPLIT TOTALI CANTIERE
        gwp_cant_tot = df_cantiere['Totale [kgCO₂ eq]'].sum()
        gwp_cant_embodied = df_cantiere['Materiali [kgCO₂ eq]'].sum()
        gwp_cant_altre = gwp_cant_tot - gwp_cant_embodied
        
        st.markdown(f"""
        <div style='padding:15px; background-color:#e8f8e8; border-left: 5px solid #27ae60; border-radius:4px;'>
            <h4>Risultati GWP (Dati Reali di Cantiere)</h4>
            <ul style='font-size: 16px;'>
                <li><strong>GWP Embodied carbon (Solo Materiali):</strong> {gwp_cant_embodied:,.2f} kg CO₂ eq</li>
                <li><strong>GWP Altri Processi (Trasporti, Macchinari, ecc.):</strong> {gwp_cant_altre:,.2f} kg CO₂ eq</li>
                <li><strong style='font-size: 18px;'>GWP Totale Cantiere:</strong> <span style='font-size: 18px;'>{gwp_cant_tot:,.2f} kg CO₂ eq</span></li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.subheader("Confronto: Progettazione vs Cantiere Reale")
        df_comp = pd.DataFrame({
            "Fase": ["Progettazione"]*len(df_risultati) + ["Cantiere"]*len(df_cantiere),
            "Attività": pd.concat([df_risultati['Attività'], df_cantiere['Attività']]),
            "Totale Emissioni": pd.concat([df_risultati['Totale [kgCO₂ eq]'], df_cantiere['Totale [kgCO₂ eq]']])
        })
        fig_comp = px.bar(df_comp, x="Attività", y="Totale Emissioni", color="Fase", barmode="group", title="Confronto Emissioni per Singola Attività")
        st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
