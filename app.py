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
            <li><strong>GWP Embodied carbon (Materiali):</strong> {gwp_embodied:,.2f} kg CO₂ eq</li>
            <li><strong>GWP Altri Processi (Rifiuti, Trasporti, Macchinari, Energia, Acqua):</strong> {gwp_altre:,.2f} kg CO₂ eq</li>
            <li><strong style='font-size: 18px;'>GWP Totale Cantiere:</strong> <span style='font-size: 18px;'>{gwp_totale:,.2f} kg CO₂ eq</span></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("#### Filtro Temporale Grafici e Opzioni Visualizzazione")
    col_fil1, col_fil2, col_fil3 = st.columns([1, 1, 1])
    
    min_date = df_risultati['Data Inizio'].min()
    max_date = df_risultati['Data Fine'].max()
    
    with col_fil1:
        date_range = st.date_input(
            "Seleziona l'intervallo temporale:", 
            value=(min_date, max_date), min_value=min_date, max_value=max_date
        )
    with col_fil2:
        opzione_linee = st.selectbox("Linee di Tendenza:", ["Nessuna", "Media Costante", "Andamento Medio (Media Mobile)", "Entrambe"])
    with col_fil3:
        ambito_gwp = st.selectbox("Ambito GWP Grafici:", ["GWP Totale", "GWP Embodied Carbon (Materiali)", "GWP Altri Processi"])
    
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
        
        if ambito_gwp == "GWP Embodied Carbon (Materiali)":
            df_melted = df_melted[df_melted['Parametro'] == 'Materiali']
        elif ambito_gwp == "GWP Altri Processi":
            df_melted = df_melted[df_melted['Parametro'] != 'Materiali']
            
        df_plot = df_melted.groupby(['Settimana', 'Parametro', 'Attività'])['Emissioni [kgCO₂ eq]'].sum().reset_index()
        df_plot = df_plot[df_plot['Emissioni [kgCO₂ eq]'] > 0]
        
        valori_pie = {
            "Materiali": df_risultati["Materiali [kgCO₂ eq]"].sum(),
            "Rifiuti": df_risultati["Rifiuti [kgCO₂ eq]"].sum(),
            "Macchinari": df_risultati["Macchinari [kgCO₂ eq]"].sum(),
            "Trasporti": df_risultati["Trasporti [kgCO₂ eq]"].sum(),
            "Energia": df_risultati["Energia [kgCO₂ eq]"].sum(),
            "Acqua": df_risultati["Acqua [kgCO₂ eq]"].sum()
        }
        
        if ambito_gwp == "GWP Embodied Carbon (Materiali)":
            valori_pie = {"Materiali": valori_pie["Materiali"]}
        elif ambito_gwp == "GWP Altri Processi":
            valori_pie.pop("Materiali", None)
            
        df_pie = pd.DataFrame({"Parametro": list(valori_pie.keys()), "Valore": list(valori_pie.values())})
        fig_pie = px.pie(df_pie, names="Parametro", values="Valore", title=f"Incidenza Percentuale - {ambito_gwp}", color="Parametro", color_discrete_map=colori_parametri, hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(f"#### Grafici Settimanali Dettagliati ({ambito_gwp})")

        df_totale_assoluto = df_plot.groupby(['Settimana', 'Attività'])['Emissioni [kgCO₂ eq]'].sum().reset_index()
        df_week_tot = df_totale_assoluto.groupby('Settimana')['Emissioni [kgCO₂ eq]'].sum().reset_index()
        
        fig_totale_assoluto = px.bar(
            df_totale_assoluto, x="Settimana", y="Emissioni [kgCO₂ eq]", hover_data=["Attività"],
            title=f"Andamento Complessivo - {ambito_gwp}", color_discrete_sequence=["#111827"]
        )
        if mostra_media_costante and not df_week_tot.empty:
            media_tot = df_week_tot['Emissioni [kgCO₂ eq]'].mean()
            fig_totale_assoluto.add_hline(y=media_tot, line_dash="dash", line_color="#ef4444", annotation_text="Media Costante")
        if mostra_media_mobile and not df_week_tot.empty:
            df_week_tot['Media_Mobile'] = df_week_tot['Emissioni [kgCO₂ eq]'].rolling(window=3, min_periods=1).mean()
            fig_totale_assoluto.add_trace(go.Scatter(x=df_week_tot['Settimana'], y=df_week_tot['Media_Mobile'], mode='lines', name='Andamento Medio', line=dict(color='#f59e0b', width=3)))
        st.plotly_chart(fig_totale_assoluto, use_container_width=True)

        fig_tot = px.bar(
            df_plot, x="Settimana", y="Emissioni [kgCO₂ eq]", color="Parametro", 
            hover_data=["Attività"], title=f"Andamento Settimanale Cumulato ({ambito_gwp})",
            color_discrete_map=colori_parametri, barmode='stack'
        )
        if mostra_media_costante and not df_week_tot.empty:
            fig_tot.add_hline(y=media_tot, line_dash="dash", line_color="#ef4444", annotation_text="Media Costante")
        if mostra_media_mobile and not df_week_tot.empty:
            fig_tot.add_trace(go.Scatter(x=df_week_tot['Settimana'], y=df_week_tot['Media_Mobile'], mode='lines', name='Andamento Medio', line=dict(color='#f59e0b', width=3)))
        st.plotly_chart(fig_tot, use_container_width=True)

        st.markdown("##### Dettaglio per Singolo Parametro")
        col_g1, col_g2 = st.columns(2)
        
        if ambito_gwp == "GWP Embodied Carbon (Materiali)":
            parametri_lista = ['Materiali']
        elif ambito_gwp == "GWP Altri Processi":
            parametri_lista = ['Rifiuti', 'Macchinari', 'Trasporti', 'Energia', 'Acqua']
        else:
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
    st.subheader("4. Monitoraggio Fase di Cantiere (Confronto Dati Dinamici Sensori)")
    
    # RIQUADRO INFORMATIVO 2 AGGIORNATO
    st.markdown("""
    <div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #27ae60; border-radius: 4px; margin-bottom: 20px;'>
        <strong>Istruzioni di Formattazione CSV (Dati Sensori):</strong>
        <ul style='margin-top: 5px; margin-bottom: 0px;'>
            <li><strong>Intestazione:</strong> Tutte le tabelle devono riportare le colonne <em>ID attività</em>, <em>nome struttura</em>, <em>data</em> e la <em>quantità</em> misurata (o nome affine come <em>km, consumi, litri</em>).</li>
            <li>I sensori sovrascriveranno le stime teoriche per calcolare le emissioni giornaliere reali del cantiere, conservando al contempo le scelte (FE) fatte in fase di progetto.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti (CSV)", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia (CSV)", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua (CSV)", type=['csv'])
    
    def parse_sensor_daily(file, df_risultati, conf_salvate, tipo_em):
        if file is None: return pd.DataFrame(columns=['Data', tipo_em])
        try:
            df = pd.read_csv(file)
            col_id = next((c for c in df.columns if 'id' in str(c).lower()), df.columns[0])
            col_data = next((c for c in df.columns if 'data' in str(c).lower()), None)
            col_q = next((c for c in df.columns if 'quantit' in str(c).lower() or 'consumo' in str(c).lower() or 'km' in str(c).lower() or 'litr' in str(c).lower()), df.columns[-1])
            
            if not col_data: return pd.DataFrame(columns=['Data', tipo_em])
            
            df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
            df = df.dropna(subset=[col_data])
            
            daily_em = []
            for _, row in df.iterrows():
                id_val = str(row[col_id]).strip()
                # Associazione con ID della fase progettuale
                match = df_risultati[df_risultati['ID'].astype(str).str.strip() == id_val]
                if not match.empty:
                    label = match.iloc[0]['Attività']
                    q = pd.to_numeric(str(row[col_q]).replace(',','.'), errors='coerce')
                    if pd.isna(q): q = 0
                    
                    if label in conf_salvate:
                        conf = conf_salvate[label]
                        em = 0
                        if tipo_em == 'Trasporti_Cant':
                            em = q * conf['Consumo_mezzo'] * conf['FE_Trasporti']
                        elif tipo_em == 'Energia_Cant':
                            em = q * conf['FE_Energia']
                        elif tipo_em == 'Acqua_Cant':
                            em = q * conf['FE_Acqua']
                        
                        daily_em.append({'Data': row[col_data].date(), tipo_em: em})
                        
            if not daily_em: return pd.DataFrame(columns=['Data', tipo_em])
            return pd.DataFrame(daily_em).groupby('Data')[tipo_em].sum().reset_index()
        except:
            return pd.DataFrame(columns=['Data', tipo_em])

    if st.button("Ricalcola ed Elabora Grafici di Confronto", use_container_width=True):
        df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
        conf_salvate = st.session_state.configurazioni_salvate
        
        # 1. Costruzione timeline giornaliera di PROGETTO (spalmata sui giorni lavorativi)
        dati_prog = []
        for _, row in df_risultati.iterrows():
            giorni_lavoro = pd.bdate_range(start=row['Data Inizio'], end=row['Data Fine'])
            n_g = len(giorni_lavoro) if len(giorni_lavoro) > 0 else 1
            for g in giorni_lavoro:
                dati_prog.append({
                    'Data': g.date(),
                    'Materiali': row["Materiali [kgCO₂ eq]"]/n_g,
                    'Rifiuti': row["Rifiuti [kgCO₂ eq]"]/n_g,
                    'Macchinari': row["Macchinari [kgCO₂ eq]"]/n_g,
                    'Trasporti_Prog': row["Trasporti [kgCO₂ eq]"]/n_g,
                    'Energia_Prog': row["Energia [kgCO₂ eq]"]/n_g,
                    'Acqua_Prog': row["Acqua [kgCO₂ eq]"]/n_g
                })
        
        df_prog_daily = pd.DataFrame(dati_prog).groupby('Data').sum().reset_index() if dati_prog else pd.DataFrame(columns=['Data', 'Materiali', 'Rifiuti', 'Macchinari', 'Trasporti_Prog', 'Energia_Prog', 'Acqua_Prog'])
        
        # 2. Estrapolazione dati reali giornalieri dei SENSORI
        df_cant_t = parse_sensor_daily(file_trasporti, df_risultati, conf_salvate, 'Trasporti_Cant')
        df_cant_e = parse_sensor_daily(file_energia, df_risultati, conf_salvate, 'Energia_Cant')
        df_cant_a = parse_sensor_daily(file_acqua, df_risultati, conf_salvate, 'Acqua_Cant')
        
        # 3. Unione logica (Outer Join per includere anche date fuori dal range teorico)
        df_comp = df_prog_daily.copy()
        
        df_sensors = pd.DataFrame(columns=['Data'])
        if not df_cant_t.empty: df_sensors = pd.merge(df_sensors, df_cant_t, on='Data', how='outer')
        if not df_cant_e.empty: df_sensors = pd.merge(df_sensors, df_cant_e, on='Data', how='outer')
        if not df_cant_a.empty: df_sensors = pd.merge(df_sensors, df_cant_a, on='Data', how='outer')
        
        if not df_sensors.empty:
            df_comp = pd.merge(df_comp, df_sensors, on='Data', how='outer').fillna(0)
        else:
            # Se nessun sensore è caricato, il cantiere vale 0
            df_comp['Trasporti_Cant'] = 0
            df_comp['Energia_Cant'] = 0
            df_comp['Acqua_Cant'] = 0
            df_comp = df_comp.fillna(0)
            
        # Calcolo dei TOTALI
        df_comp['Totale_Prog'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Prog'] + df_comp['Energia_Prog'] + df_comp['Acqua_Prog']
        df_comp['Totale_Cant'] = df_comp['Materiali'] + df_comp['Rifiuti'] + df_comp['Macchinari'] + df_comp['Trasporti_Cant'] + df_comp['Energia_Cant'] + df_comp['Acqua_Cant']

        # Sorte cronologico
        df_comp = df_comp.sort_values('Data')
        
        # UI Filtri Grafici Cantiere
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("#### Filtro Temporale e Opzioni di Confronto (Progettazione vs Cantiere)")
        col_cfil1, col_cfil2 = st.columns([1, 1])
        
        with col_cfil1:
            opzione_linee_c = st.selectbox("Linee di Tendenza Analitiche:", ["Nessuna", "Medie Costanti", "Medie Mobili (Andamento)", "Entrambe"], key="linee_cantiere")
        
        mostra_media_costante_c = opzione_linee_c in ["Medie Costanti", "Entrambe"]
        mostra_media_mobile_c = opzione_linee_c in ["Medie Mobili (Andamento)", "Entrambe"]
        
        def plot_comparison(df, base_col, title, color_prog="#3b82f6", color_cant="#22c55e"):
            fig = go.Figure()
            # Barre
            fig.add_trace(go.Bar(x=df['Data'], y=df[f'{base_col}_Prog'], name='Progettazione', marker_color=color_prog))
            fig.add_trace(go.Bar(x=df['Data'], y=df[f'{base_col}_Cant'], name='Cantiere (Reale)', marker_color=color_cant))
            
            # Linee Costanti
            if mostra_media_costante_c:
                mean_p = df[f'{base_col}_Prog'].mean()
                mean_c = df[f'{base_col}_Cant'].mean()
                fig.add_hline(y=mean_p, line_dash="dash", line_color=color_prog, annotation_text="Media Prog")
                fig.add_hline(y=mean_c, line_dash="dash", line_color=color_cant, annotation_text="Media Cant")
            
            # Linee Mobili (3 giorni)
            if mostra_media_mobile_c:
                ma_p = df[f'{base_col}_Prog'].rolling(window=3, min_periods=1).mean()
                ma_c = df[f'{base_col}_Cant'].rolling(window=3, min_periods=1).mean()
                fig.add_trace(go.Scatter(x=df['Data'], y=ma_p, mode='lines', name='Trend Progettazione', line=dict(color="#1e3a8a", width=2)))
                fig.add_trace(go.Scatter(x=df['Data'], y=ma_c, mode='lines', name='Trend Cantiere', line=dict(color="#14532d", width=2)))
                
            fig.update_layout(title=title, barmode='group', hovermode="x unified", xaxis_title="Data", yaxis_title="Emissioni [kgCO₂ eq]")
            return fig

        st.markdown("##### 1. Confronto Emissioni Totali Giornaliere")
        st.plotly_chart(plot_comparison(df_comp, 'Totale', "GWP Totale: Stima Progettuale vs Dato Reale", color_prog="#64748b", color_cant="#0f172a"), use_container_width=True)
        
        st.markdown("##### 2. Confronto Parametri Monitorati Dinamicamente (Sensori)")
        col_cg1, col_cg2 = st.columns(2)
        
        with col_cg1:
            st.plotly_chart(plot_comparison(df_comp, 'Trasporti', "Trasporti", color_prog="#818cf8", color_cant="#3730a3"), use_container_width=True)
            st.plotly_chart(plot_comparison(df_comp, 'Acqua', "Acqua", color_prog="#38bdf8", color_cant="#0369a1"), use_container_width=True)
            
        with col_cg2:
            st.plotly_chart(plot_comparison(df_comp, 'Energia', "Energia Elettrica", color_prog="#fde047", color_cant="#a16207"), use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
