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
if 'df_comp_cantiere' not in st.session_state:
    st.session_state.df_comp_cantiere = None

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
    
    id_att = str(row_att[col_id]) # Castato a stringa per usarlo in modo sicuro nelle key dei widget
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

    st.info(f"Quantità materiale associata a questa attività: **{quantita_mat:,.2f}** kg")
    
    # Inizializzazione variabili
    em_materiali, em_rifiuti, em_trasporti, em_energia, em_acqua, em_macchinari = 0, 0, 0, 0, 0, 0
    q_rifiuti = 0
    fe_trasporti_medio = 0
    
    # MATERIALI E RIFIUTI
    st.markdown("---")
    st.markdown("#### Materiali e Rifiuti")
    st.markdown("Le emissioni dei materiali vengono calcolate dalla fase di estrazione della materia prima fino alla creazione del prodotto finale (A1-A3). I relativi fattori emissivi si trovano nella scheda EPD del prodotto.")
    st.markdown("Le emissioni dei rifiuti vengono calcolate dalla fase di trasporto in discarica fino allo smaltimento (C2-C4). La quantità di rifiuti viene inserta da utente in % rispetto alla quantità di materiale. I relativi fattori emissivi si trovano nella scheda EPD del prodotto.")
    
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
        
        perc_rifiuti = st.number_input("% di rifiuti stimata (sfrido):", min_value=0.0, max_value=100.0, value=0.0, key=f"rif_{id_att}") / 100
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
    else:
        st.info("Nessun materiale in ingresso associato a questa attività. Emissioni per Materiali e Rifiuti pari a 0.")
        
    # TRASPORTI MULTIPLI
    st.markdown("---")
    st.markdown("#### Trasporti")
    st.markdown("Trasporto ipotizzato: max 28 t per una distanza di 30 km. L'utente potrà scegliere il tipo di carburante e i relativi fattori emissivi sono presenti nel database LCI del sito.") 
    st.markdown("In caso di mancanza del materiale associato all'attività analizzata, è comunque possibile calcolare le emissioni dovute al trasporto di materiale da demolizione in discarica.") 
    st.markdown("FONTE fattori emissivi: EMEP (European Enviroment Agency).") 
    
    if 'Trasporti' in df_lci:
        df_t_sheet = df_lci['Trasporti']
        
        df_t_emissions = df_t_sheet.iloc[0:9].copy()
        df_t_emissions.columns = ['Tipologia', 'Carburante', 'Fattore_Emissione']
        
        df_t_consumi = df_t_sheet.iloc[11:21].copy()
        df_t_consumi.columns = ['Tipologia', 'Carburante', 'Consumo_Medio']
        df_t_consumi = df_t_consumi.dropna(subset=['Tipologia'])
        df_t_consumi['Label'] = df_t_consumi['Tipologia'].astype(str) + " - " + df_t_consumi['Carburante'].astype(str)
        
        viaggi_mat = np.ceil(quantita_mat / 28000) if quantita_mat > 0 else 0
        viaggi_rif = np.ceil(q_rifiuti / 28000) if q_rifiuti > 0 else 0
        viaggi_tot_stimati = int(viaggi_mat + viaggi_rif)
        
        default_mezzi = 1 if viaggi_tot_stimati > 0 else 0
        num_tipi_trasporto = st.number_input("Quanti tipi di mezzi di trasporto verranno impiegati?", min_value=0, value=default_mezzi, step=1, key=f"num_trasporti_{id_att}")
        
        em_trasporti = 0
        tot_km_progetto = 0
        dati_trasporto_tabella = []
        
        if num_tipi_trasporto > 0:
            for i in range(num_tipi_trasporto):
                st.markdown(f"**Mezzo di Trasporto {i+1}**")
                scelta_mezzo = st.selectbox(f"Seleziona Veicolo/Carburante {i+1}:", df_t_consumi['Label'].unique(), key=f"t_mezzo_{i}_{id_att}")
                row_mezzo = df_t_consumi[df_t_consumi['Label'] == scelta_mezzo].iloc[0]
                
                consumo_medio = float(str(row_mezzo['Consumo_Medio']).replace(',', '.'))
                carb_selezionato = str(row_mezzo['Carburante'])
                carb_match = 'CNG(or LNG)' if 'CNG' in carb_selezionato else carb_selezionato
                
                row_em = df_t_emissions[df_t_emissions['Carburante'].astype(str) == carb_match]
                fattore_carb = float(str(row_em['Fattore_Emissione'].values[0]).replace(',', '.')) if not row_em.empty else 0.0
                
                col_t1, col_t2 = st.columns(2)
                with col_t1:
                    viaggi = st.number_input(f"Viaggi stimati (mezzo {i+1}):", min_value=0, value=viaggi_tot_stimati if i==0 else 0, key=f"t_viaggi_{i}_{id_att}")
                with col_t2:
                    distanza = st.number_input(f"Distanza media a viaggio [km] (mezzo {i+1}):", min_value=0.0, value=30.0, key=f"t_dist_{i}_{id_att}")
                
                em_trasporto_i = viaggi * distanza * consumo_medio * fattore_carb
                em_trasporti += em_trasporto_i
                tot_km_progetto += (viaggi * distanza)
                
                dati_trasporto_tabella.append({
                    "Mezzo": scelta_mezzo,
                    "Viaggi": viaggi,
                    "Distanza [km]": distanza,
                    "Consumo [kg/km]": consumo_medio,
                    "Fattore Emissivo [kgCO₂ eq/kg]": fattore_carb,
                    "Emissioni [kgCO₂ eq]": round(em_trasporto_i, 2)
                })

            fe_trasporti_medio = (em_trasporti / tot_km_progetto) if tot_km_progetto > 0 else 0
            
            st.markdown("**Riepilogo Trasporti:**")
            st.dataframe(pd.DataFrame(dati_trasporto_tabella), hide_index=True)
        else:
            st.info("Nessun mezzo di trasporto previsto per questa attività.")
            fe_trasporti_medio = 0
            
        st.write(f"**Totale Emissioni Trasporti:** {em_trasporti:,.2f} kgCO₂ eq")
        
    # MACCHINARI MULTIPLI
    st.markdown("---")
    st.markdown("#### Macchinari")
    st.markdown("Il calcolo delle emissioni dei macchinari si basa sul prodotto analitico tra la potenza nominale del motore, il consumo specifico di combustibile per unità di energia, il fattore di carico medio (tempo effettivo di funzionamento del macchinario nelle 8 ore lavorative), le ore lavorative giornaliere di un turno standard pari a 8 ore e il fattore emissivo del carburante.")
    col_mac_fe = next((c for c in df_macc.columns if 'giorno' in str(c).lower()), df_macc.columns[3] if len(df_macc.columns) > 3 else None)
    col_durata = next((c for c in df_crono.columns if 'durata' in str(c).lower()), None)
    durata_giorni = int(row_att[col_durata]) if col_durata else 1
    
    num_tipi_macc = st.number_input("Quanti tipi di macchinari verranno impiegati?", min_value=0, value=0, step=1, key=f"num_macc_{id_att}")
    
    em_macchinari = 0
    dati_macchinari_tabella = []
    
    if num_tipi_macc > 0:
        for i in range(num_tipi_macc):
            st.markdown(f"**Macchinario {i+1}**")
            macchinario = st.selectbox(f"Scegli Macchinario base {i+1}:", df_macc['Macchinario'].unique(), key=f"m_tipo_{i}_{id_att}")
            df_macc_filt = df_macc[df_macc['Macchinario'] == macchinario]
            
            if col_mac_fe:
                col_macc_link = next((c for c in df_macc_filt.columns if 'link' in str(c).lower()), df_macc_filt.columns[-1])
                st.dataframe(
                    df_macc_filt[['Macchinario', 'Produttore', col_mac_fe, col_macc_link]],
                    column_config={col_macc_link: st.column_config.LinkColumn("Link Schede Tecniche")},
                    hide_index=True
                )
                
                macc_produttore = st.selectbox(f"Modello Specifico {i+1}:", df_macc_filt['Produttore'], key=f"m_prod_{i}_{id_att}")
                fe_macc_8h = float(str(df_macc_filt[df_macc_filt['Produttore'] == macc_produttore][col_mac_fe].values[0]).replace(',', '.'))
            else:
                st.warning("Impossibile trovare la colonna delle emissioni nel file macchinari.")
                fe_macc_8h = 0.0
                macc_produttore = "N/A"
                
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                num_macc_input = st.number_input(f"Numero di unità (Macchinario {i+1}):", min_value=1, value=1, key=f"m_unita_{i}_{id_att}")
            with col_m2:
                giorni_macc = st.number_input(f"Giorni di lavoro (Macchinario {i+1}):", min_value=1, value=durata_giorni, key=f"m_giorni_{i}_{id_att}")
                
            em_macc_i = fe_macc_8h * num_macc_input * giorni_macc
            em_macchinari += em_macc_i
            
            dati_macchinari_tabella.append({
                "Voce": f"Macchinario {i+1}",
                "Num. Macchine": num_macc_input,
                "Giorni": giorni_macc,
                "Fattore Emissivo [kgCO₂ eq/giorno]": f"{fe_macc_8h:,.2f}",
                "Emissioni [kgCO₂ eq]": f"{em_macc_i:,.2f}"
            })
            
        st.markdown("**Riepilogo Macchinari:**")
        st.dataframe(pd.DataFrame(dati_macchinari_tabella), hide_index=True)
    else:
        st.info("Nessun macchinario previsto per questa attività.")
        
    st.write(f"**Totale Emissioni Macchinari:** {em_macchinari:,.2f} kgCO₂ eq")
    
    # ENERGIA
    st.markdown("---")
    st.markdown("#### Energia Elettrica")
    st.markdown("Il calcolo delle emissioni per l'energia elettrica si basa sul prodotto tra il consumo effettivo in kilowattora (kWh) e il fattore emissivo specifico della tecnologia di generazione selezionata.")
    st.markdown("FONTE fattori emissivi: ISPRA (Istituto Superiore per la Protezione e la Ricerca Ambientale).")
    col_energia_fe = next((c for c in df_lci['Energia'].columns if 'fattore' in str(c).lower()), df_lci['Energia'].columns[1])
    col_energia_nome = df_lci['Energia'].columns[0]
    
    tech_energia = st.selectbox("Tecnologia generazione elettrica:", df_lci['Energia'][col_energia_nome].unique(), key=f"tech_ene_{id_att}")
    fe_energia = float(str(df_lci['Energia'][df_lci['Energia'][col_energia_nome] == tech_energia][col_energia_fe].values[0]).replace(',', '.'))
    kwh_stimati = st.number_input("Consumo elettrico stimato (kWh):", value=0.0, key=f"kwh_{id_att}")
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
    st.markdown("Il calcolo moltiplica il volume consumato in litri per il fattore emissivo associato alla specifica tipologia di approvvigionamento o trattamento idrico.")
    st.markdown("FONTE fattori emissivi: GHG (GreenHouse Gas emissions).")
    col_acqua_fe = next((c for c in df_lci['Acqua'].columns if 'fattore' in str(c).lower()), df_lci['Acqua'].columns[1])
    col_acqua_nome = df_lci['Acqua'].columns[0]
    
    tech_acqua = st.selectbox("Approvvigionamento/Trattamento Acqua:", df_lci['Acqua'][col_acqua_nome].unique(), key=f"tech_acq_{id_att}")
    fe_acqua = float(str(df_lci['Acqua'][df_lci['Acqua'][col_acqua_nome] == tech_acqua][col_acqua_fe].values[0]).replace(',', '.'))
    litri_stimati = st.number_input("Consumo idrico stimato (litri):", value=0.0, key=f"litri_{id_att}")
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
            "Consumo_mezzo": 1.0 if 'Trasporti' in df_lci else 0,
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
    
    st.markdown(f'''
    <div style='padding:15px; background-color:#e8f4f8; border-left: 5px solid #2980b9; border-radius:4px;'>
        <h4>Risultati GWP (Stima Progettuale)</h4>
        <ul style='font-size: 16px;'>
            <li><strong>GWP Embodied carbon (Materiali):</strong> {gwp_embodied:,.2f} kg CO₂ eq</li>
            <li><strong>GWP Altri Processi (Rifiuti, Trasporti, Macchinari, Energia, Acqua):</strong> {gwp_altre:,.2f} kg CO₂ eq</li>
            <li><strong style='font-size: 18px;'>GWP Totale Cantiere:</strong> <span style='font-size: 18px;'>{gwp_totale:,.2f} kg CO₂ eq</span></li>
        </ul>
    </div>
    ''', unsafe_allow_html=True)
    
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
            title=f"Andamento Complessivo - {ambito_gwp}", color_discrete_sequence=["#01796F"]
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
    st.subheader("4. Monitoraggio Fase di Cantiere (Dati Dinamici Sensori)")
    
    # RIQUADRO INFORMATIVO 2
    st.markdown('''
    <div style='background-color: #eef2f3; padding: 15px; border-left: 5px solid #27ae60; border-radius: 4px; margin-bottom: 20px;'>
        <strong>Istruzioni di Formattazione CSV (Dati Sensori):</strong>
        <ul style='margin-top: 5px; margin-bottom: 0px;'>
            <li><strong>Intestazione:</strong> Tutte le tabelle devono riportare le colonne <em>ID attività</em> e <em>Data</em> per potersi allacciare al calcolo in progetto.</li>
            <li><strong>Trasporti:</strong> Colonna nominata <em>Km_percorsi</em> e colonna <em>Carburante</em> per prelevare in automatico i fattori emissivi dal database LCI.</li>
            <li><strong>Energia:</strong> Colonna <em>consumi_kWh</em> (o simile) e <em>Alimentazione_elettrica</em>.</li>
            <li><strong>Acqua:</strong> Colonna <em>consumi_l</em> (o simile) e <em>Tipo_acqua</em> (Treatment o Supply).</li>
        </ul>
    </div>
    ''', unsafe_allow_html=True)
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti (CSV)", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia (CSV)", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua (CSV)", type=['csv'])

    def elabora_sensore_con_attivita(df_sens, tipo_file, df_lci, df_ris, conf_salvate):
        risultati = []
        
        # Individua dinamicamente la colonna ID e Data
        col_id = next((c for c in df_sens.columns if 'id' in str(c).lower() and 'attiv' in str(c).lower()), None)
        if not col_id:
            col_id = next((c for c in df_sens.columns if 'id' in str(c).lower()), df_sens.columns[0])
            
        col_data = next((c for c in df_sens.columns if 'data' in str(c).lower()), None)
        
        if not col_data or not col_id:
            return pd.DataFrame(columns=['Data'])
            
        df_sens[col_data] = pd.to_datetime(df_sens[col_data], dayfirst=True, errors='coerce')
        df_sens = df_sens.dropna(subset=[col_data])
        
        for _, row in df_sens.iterrows():
            id_sensore = str(row[col_id]).strip()
            data_sensore = row[col_data].date()
            
            # Cerca l'ID nel database di progetto
            match_prog = df_ris[df_ris['ID'].astype(str).str.strip() == id_sensore]
            if match_prog.empty:
                continue 
                
            label_attivita = match_prog.iloc[0]['Attività']
            conf = conf_salvate.get(label_attivita, {})
            
            emissione = 0
            if tipo_file == 'trasporti':
                col_km = next((c for c in df_sens.columns if 'km' in str(c).lower()), None)
                km = pd.to_numeric(str(row[col_km]).replace(',','.')) if col_km else 0
                
                col_carb = next((c for c in df_sens.columns if 'carb' in str(c).lower()), None)
                carb = str(row[col_carb]) if col_carb else ''
                
                df_t_fe = df_lci['Trasporti'].iloc[0:9]
                df_t_cons = df_lci['Trasporti'].iloc[11:21]
                
                match_fe = df_t_fe[df_t_fe.iloc[:, 1].astype(str).str.contains(carb, case=False, na=False)]
                fe = float(match_fe.iloc[0, 2]) if not match_fe.empty else 0.0
                
                match_cons = df_t_cons[df_t_cons.iloc[:, 1].astype(str).str.contains(carb, case=False, na=False)]
                consumo = pd.to_numeric(match_cons.iloc[:, 2], errors='coerce').mean() if not match_cons.empty else 0.1
                
                emissione = km * consumo * fe
                risultati.append({'Data': data_sensore, 'Trasporti_Cant': emissione})
                
            elif tipo_file == 'energia':
                col_q = next((c for c in df_sens.columns if 'consum' in str(c).lower() or 'kwh' in str(c).lower()), None)
                kwh = pd.to_numeric(str(row[col_q]).replace(',','.')) if col_q else 0
                
                col_alim = next((c for c in df_sens.columns if 'alimen' in str(c).lower() or 'elettric' in str(c).lower()), None)
                alim = str(row[col_alim]) if col_alim else str(conf.get('Energia', ''))
                
                df_e = df_lci['Energia']
                match_fe = df_e[df_e.iloc[:, 0].astype(str).str.contains(alim, case=False, na=False)]
                fe = float(match_fe.iloc[0, 1]) if not match_fe.empty else conf.get('FE_Energia', 0)
                
                emissione = kwh * fe
                risultati.append({'Data': data_sensore, 'Energia_Cant': emissione})
                
            elif tipo_file == 'acqua':
                col_q = next((c for c in df_sens.columns if 'consum' in str(c).lower() or 'litr' in str(c).lower() or 'l' == str(c).lower()[-1:]), None)
                litri = pd.to_numeric(str(row[col_q]).replace(',','.')) if col_q else 0
                
                col_tipo = next((c for c in df_sens.columns if 'tipo' in str(c).lower() or 'acqua' in str(c).lower() and c != col_q), None)
                tipo = str(row[col_tipo]) if col_tipo else str(conf.get('Acqua', ''))
                
                df_a = df_lci['Acqua']
                match_fe = df_a[df_a.iloc[:, 0].astype(str).str.contains(tipo, case=False, na=False)]
                fe = float(match_fe.iloc[0, 1]) if not match_fe.empty else conf.get('FE_Acqua', 0)
                
                emissione = litri * fe
                risultati.append({'Data': data_sensore, 'Acqua_Cant': emissione})
                
        if risultati:
            # Aggrego giornalmente i dati misurati
            return pd.DataFrame(risultati).groupby('Data').sum().reset_index()
        return pd.DataFrame(columns=['Data'])
    
    # SALVO I DATI IN SESSION_STATE IN MODO CHE NON SPARISCANO AI SUCCESSIVI RERUN
    if st.button("Elabora Dati Sensori e Genera Grafici di Confronto", use_container_width=True):
        df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
        
        # 1. TIMELINE PROGETTUALE (Stime spalmate sui giorni di lavoro)
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
        
        # 2. LETTURA SENSORI CANTIERE
        df_cant_t, df_cant_e, df_cant_a = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        if file_trasporti: df_cant_t = elabora_sensore_con_attivita(pd.read_csv(file_trasporti), 'trasporti', df_lci, df_risultati, st.session_state.configurazioni_salvate)
        if file_energia: df_cant_e = elabora_sensore_con_attivita(pd.read_csv(file_energia), 'energia', df_lci, df_risultati, st.session_state.configurazioni_salvate)
        if file_acqua: df_cant_a = elabora_sensore_con_attivita(pd.read_csv(file_acqua), 'acqua', df_lci, df_risultati, st.session_state.configurazioni_salvate)
        
        # 3. MERGE DATI (Join Outer sulle Date)
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

        # Salvataggio nel session_state
        st.session_state.df_comp_cantiere = df_comp

# IL BLOCCO DEI GRAFICI ORA DIPENDE DAL SESSION_STATE E HA IL FILTRO TEMPORALE
    if st.session_state.df_comp_cantiere is not None:
        df_comp = st.session_state.df_comp_cantiere.copy()
        df_comp['Data'] = pd.to_datetime(df_comp['Data'])
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("Confronto Emissioni Settimanali: Progettazione vs Cantiere Reale")
        
        # --- FILTRO TEMPORALE E OPZIONI ---
        min_date_c = df_comp['Data'].min().date() if not df_comp.empty else pd.Timestamp.now().date()
        max_date_c = df_comp['Data'].max().date() if not df_comp.empty else pd.Timestamp.now().date()
        
        col_cfil1, col_cfil2, col_cfil3 = st.columns(3)
        with col_cfil1:
            date_range_c = st.date_input(
                "Seleziona l'intervallo temporale:", 
                value=(min_date_c, max_date_c), min_value=min_date_c, max_value=max_date_c, key="date_comp"
            )
        with col_cfil2:
            opzione_linee = st.selectbox(
                "Linee di Tendenza:", 
                ["Nessuna", "Media Costante", "Andamento Medio (Media Mobile)", "Entrambe"], 
                key="linee_comp"
            )
        with col_cfil3:
            ambito_gwp_cant = st.selectbox(
                "Ambito GWP Grafici:", 
                ["GWP Totale", "GWP Altri Processi"], 
                key="ambito_comp"
            )
            
        if isinstance(date_range_c, tuple) and len(date_range_c) == 2:
            start_filter_c, end_filter_c = date_range_c
        else:
            start_filter_c, end_filter_c = min_date_c, max_date_c
            
        # Filtro il DataFrame in base alle date scelte
        df_comp_filtered = df_comp[(df_comp['Data'].dt.date >= start_filter_c) & (df_comp['Data'].dt.date <= end_filter_c)].copy()
        
        if not df_comp_filtered.empty:
            # Creazione colonna Settimana e Raggruppamento
            df_comp_filtered['Settimana'] = df_comp_filtered['Data'].dt.strftime('%Y-W%V')
            
            # Calcolo le colonne per "Altri Processi" per Progetto e Cantiere
            df_comp_filtered['Altri_Prog'] = df_comp_filtered['Rifiuti'] + df_comp_filtered['Macchinari'] + df_comp_filtered['Trasporti_Prog'] + df_comp_filtered['Energia_Prog'] + df_comp_filtered['Acqua_Prog']
            df_comp_filtered['Altri_Cant'] = df_comp_filtered['Rifiuti'] + df_comp_filtered['Macchinari'] + df_comp_filtered['Trasporti_Cant'] + df_comp_filtered['Energia_Cant'] + df_comp_filtered['Acqua_Cant']
            
            # --- INSERIMENTO RIQUADRO RIEPILOGATIVO CONFRONTO ---
            gwp_embodied_cant = df_comp_filtered['Materiali'].sum()
            gwp_altre_prog = df_comp_filtered['Altri_Prog'].sum()
            gwp_altre_cant = df_comp_filtered['Altri_Cant'].sum()
            gwp_totale_prog = df_comp_filtered['Totale_Prog'].sum()
            gwp_totale_cant = df_comp_filtered['Totale_Cant'].sum()

            colore_confronto = "#16a34a" if gwp_totale_cant <= gwp_totale_prog else "#dc2626"

            st.markdown(f'''
            <div style='padding:15px; background-color:#e8f4f8; border-left: 5px solid #27ae60; border-radius:4px; margin-bottom: 20px;'>
                <h4>Riepilogo GWP (Periodo Selezionato): Progetto vs Cantiere Reale</h4>
                <ul style='font-size: 16px;'>
                    <li><strong>GWP Embodied carbon (Materiali invariati):</strong> {gwp_embodied_cant:,.2f} kg CO₂ eq</li>
                    <li><strong>GWP Altri Processi (Progetto):</strong> {gwp_altre_prog:,.2f} kg CO₂ eq &nbsp;|&nbsp; <strong>(Cantiere):</strong> {gwp_altre_cant:,.2f} kg CO₂ eq</li>
                    <li style='margin-top: 8px;'><strong style='font-size: 18px;'>GWP Totale - Progetto:</strong> <span style='font-size: 18px;'>{gwp_totale_prog:,.2f} kg CO₂ eq</span></li>
                    <li><strong style='font-size: 18px;'>GWP Totale - Cantiere:</strong> <span style='font-size: 18px; color: {colore_confronto};'>{gwp_totale_cant:,.2f} kg CO₂ eq</span></li>
                </ul>
            </div>
            ''', unsafe_allow_html=True)
            # --- FINE RIQUADRO ---

            cols_to_sum = [c for c in df_comp_filtered.columns if c not in ['Data', 'Settimana']]
            df_weekly = df_comp_filtered.groupby('Settimana')[cols_to_sum].sum().reset_index()
            
            # --- GRAFICI DI CONFRONTO SETTIMANALI ---
            def plot_comparison(df, base_col, title, c_prog, c_cant):
                fig = go.Figure()
                if f"{base_col}_Prog" in df: 
                    fig.add_trace(go.Bar(x=df['Settimana'], y=df[f'{base_col}_Prog'], name='Progetto', marker_color=c_prog))
                fig.add_trace(go.Bar(x=df['Settimana'], y=df[f'{base_col}_Cant'], name='Cantiere (Reale)', marker_color=c_cant))
                
                if opzione_linee in ["Media Costante", "Entrambe"]:
                    if f"{base_col}_Prog" in df: 
                        fig.add_hline(y=df[f'{base_col}_Prog'].mean(), line_dash="dash", line_color=c_prog)
                    fig.add_hline(y=df[f'{base_col}_Cant'].mean(), line_dash="dash", line_color=c_cant)
                
                if opzione_linee in ["Andamento Medio (Media Mobile)", "Entrambe"]:
                    if f"{base_col}_Prog" in df: 
                        fig.add_trace(go.Scatter(x=df['Settimana'], y=df[f'{base_col}_Prog'].rolling(3, min_periods=1).mean(), name='Trend Prog', line=dict(color=c_prog, width=2)))
                    fig.add_trace(go.Scatter(x=df['Settimana'], y=df[f'{base_col}_Cant'].rolling(3, min_periods=1).mean(), name='Trend Cant', line=dict(color=c_cant, width=2)))
                    
                fig.update_layout(title=title, barmode='group', hovermode="x unified", xaxis_title="Settimana", yaxis_title="Emissioni [kgCO₂ eq]")
                return fig

            # Logica di visualizzazione in base al menù a tendina
            if ambito_gwp_cant == "GWP Totale":
                st.plotly_chart(plot_comparison(df_weekly, 'Totale', "GWP Totale: Progetto vs Reale", "#01796F", "#8CD6C4"), use_container_width=True)
            else:
                st.plotly_chart(plot_comparison(df_weekly, 'Altri', "GWP Altri Processi: Progetto vs Reale", "#01796F", "#8CD6C4"), use_container_width=True)
            
            col_g1, col_g2 = st.columns(2)
            col_g1.plotly_chart(plot_comparison(df_weekly, 'Trasporti', "Trasporti", "#818cf8", "#3730a3"), use_container_width=True)
            col_g1.plotly_chart(plot_comparison(df_weekly, 'Acqua', "Acqua", "#0369a1", "#b8edea"), use_container_width=True)
            col_g2.plotly_chart(plot_comparison(df_weekly, 'Energia', "Energia Elettrica", "#FFD700", "#a16207"), use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)
