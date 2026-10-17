import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta
import numpy as np
import os

# --- IMPOSTAZIONI PAGINA E STILE CSS ---
st.set_page_config(page_title="EcoSite Tracker", layout="wide")
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; background-color: #faf9f6; }
    .minimal-card { background-color: #ffffff !important; border: 1.5px solid #d5ddd1 !important; border-radius: 8px !important; padding: 32px !important; margin-bottom: 24px !important; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.01) !important; }
</style>
""", unsafe_allow_html=True)

st.title("EcoSite Tracker 🌱")
st.markdown("Monitoraggio Emissioni CO₂ eq in Cantiere (Fase di Progettazione e Costruzione)")

# --- INIZIALIZZAZIONE SESSION STATE ---
if 'risultati_progettazione' not in st.session_state:
    st.session_state.risultati_progettazione = []
if 'configurazioni_salvate' not in st.session_state:
    st.session_state.configurazioni_salvate = {}

# --- FUNZIONI DI CARICAMENTO DATI STATICI ---
# Nessuna cache: legge i file freschi ad ogni aggiornamento
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
file_caricati = st.file_uploader("Carica i file CSV (Quantità Materiali e Cronoprogramma)", type=['csv'], accept_multiple_files=True)

lista_quantita = []
df_crono = None

if file_caricati:
    for f in file_caricati:
        try:
            # Prova separatore europeo
            df_temp = pd.read_csv(f, sep=';', encoding='utf-8-sig')
            if len(df_temp.columns) < 2:
                f.seek(0)
                df_temp = pd.read_csv(f, sep=',', encoding='utf-8-sig')
            
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
            # Creazione dinamica dell'etichetta del menu a tendina
            col_id = next((c for c in df_crono.columns if 'id' in str(c).lower()), df_crono.columns[0])
            col_nome = next((c for c in df_crono.columns if c.strip().lower() in ['nome', 'name', 'nome attività', 'attività']), df_crono.columns[1])
            col_struttura = next((c for c in df_crono.columns if 'struttura' in str(c).lower()), None)
            col_wbs = next((c for c in df_crono.columns if 'wbs lav' in str(c).lower()), None)
            
            etichette = df_crono[col_id].astype(str) + " - " + df_crono[col_nome].astype(str)
            if col_struttura:
                etichette += " | " + df_crono[col_struttura].astype(str)
            if col_wbs:
                etichette += " (" + df_crono[col_wbs].astype(str) + ")"
                
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
        q_col_lav = next((c for c in df_quantita.columns if 'wbs lavorazione' in str(c).lower()), None)
        q_col_mat = next((c for c in df_quantita.columns if 'wbs materiale' in str(c).lower()), None)
        q_col_massa = next((c for c in df_quantita.columns if any(k in str(c).lower() for k in ['massa', 'quantit', 'volume', 'peso'])), None)
        
        if q_col_lav and q_col_mat and q_col_massa:
            match_q = df_quantita[(df_quantita[q_col_lav].astype(str) == str(wbs_lav)) & (df_quantita[q_col_mat].astype(str) == str(wbs_mat))]
            if not match_q.empty:
                massa_pulita = pd.to_numeric(match_q[q_col_massa].astype(str).replace({',': '.'}, regex=True), errors='coerce').fillna(0)
                quantita_mat = massa_pulita.sum()
    
    st.info(f"Quantità materiale associata a questa attività (tramite WBS): **{quantita_mat:,.2f}**")
    
    with st.form(key=f"form_{attivita_selezionata}"):
        em_materiali, em_rifiuti, em_trasporti, em_energia, em_acqua, em_macchinari = 0, 0, 0, 0, 0, 0
        
        # MATERIALI E RIFIUTI
        st.markdown("#### Materiali e Rifiuti")
        if quantita_mat > 0:
            materiale_scelto = st.selectbox("Seleziona Tipo Materiale (filtro categoria):", df_epd['Materiale'].dropna().unique())
            df_epd_filtrato = df_epd[df_epd['Materiale'] == materiale_scelto]
            
            st.dataframe(
                df_epd_filtrato[['Materiale', 'Produttore', 'Fattore emissivo', 'Trattamento rifiuti', 'Link PDF']],
                column_config={"Link PDF": st.column_config.LinkColumn("Link EPD")}, hide_index=True
            )
            
            scelta_specifica = st.selectbox("Scegli il prodotto specifico per il calcolo:", df_epd_filtrato['Produttore'])
            fattore_em_mat = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica]['Fattore emissivo'].values[0]
            fattore_em_rif = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica]['Trattamento rifiuti'].values[0]
            
            fattore_em_mat = float(str(fattore_em_mat).replace(',', '.'))
            fattore_em_rif = float(str(fattore_em_rif).replace(',', '.'))
            
            em_materiali = quantita_mat * fattore_em_mat
            
            perc_rifiuti = st.number_input("% di rifiuti stimata (sfrido):", min_value=0.0, max_value=100.0, value=0.0) / 100
            q_rifiuti = quantita_mat * perc_rifiuti
            em_rifiuti = q_rifiuti * fattore_em_rif
            
            # TRASPORTI
            st.markdown("#### Trasporti")
            if 'Trasporti' in df_lci:
                carb_trasporto = st.selectbox("Carburante Trasporti:", df_lci['Trasporti']['Fuel'].dropna().unique())
                row_carb = df_lci['Trasporti'][df_lci['Trasporti']['Fuel'] == carb_trasporto].iloc[0]
                fattore_carb = float(str(row_carb['kg CO2 per kg of fuel']).replace(',', '.'))
                consumo_medio = st.number_input("Consumo medio carburante mezzo (kg/km):", value=0.25)
                
                viaggi_mat = np.ceil(quantita_mat / 28000)
                viaggi_rif = np.ceil(q_rifiuti / 28000) if q_rifiuti > 0 else 0
                viaggi_tot = viaggi_mat + viaggi_rif
                
                em_trasporti = viaggi_tot * 30 * consumo_medio * fattore_carb
        else:
            st.warning("Nessun materiale trovato per questa attività (Quantità = 0). Sezioni Materiali, Rifiuti e Trasporti avranno emissioni pari a 0.")
            carb_trasporto = None
            consumo_medio = 0
            
        # MACCHINARI
        st.markdown("#### Macchinari")
        macchinario = st.selectbox("Scegli Macchinario base:", df_macc['Macchinario'].unique())
        df_macc_filt = df_macc[df_macc['Macchinario'] == macchinario]
        st.dataframe(df_macc_filt[['Macchinario', 'Produttore', 'Fattore emissivo 8h', 'Link Scheda']], hide_index=True)
        
        macc_produttore = st.selectbox("Modello Specifico:", df_macc_filt['Produttore'])
        fe_macc_8h = float(str(df_macc_filt[df_macc_filt['Produttore'] == macc_produttore]['Fattore emissivo 8h'].values[0]).replace(',', '.'))
        
        num_macc = st.number_input("Numero di macchinari impiegati:", min_value=1, value=1)
        col_durata = next((c for c in df_crono.columns if 'durata' in str(c).lower()), None)
        durata_giorni = int(row_att[col_durata]) if col_durata else 1
        
        giorni_macc = st.number_input("Giorni di lavoro (per macchinario):", min_value=1, value=durata_giorni)
        em_macchinari = fe_macc_8h * num_macc * giorni_macc
        
        # ENERGIA
        st.markdown("#### Energia Elettrica")
        tech_energia = st.selectbox("Tecnologia generazione elettrica:", df_lci['Energia']['Tecnologia di generazione elettrica'].unique())
        fe_energia = float(str(df_lci['Energia'][df_lci['Energia']['Tecnologia di generazione elettrica'] == tech_energia]['Fattori di emissione (kg CO2eq/kWh)'].values[0]).replace(',', '.'))
        kwh_stimati = st.number_input("Consumo elettrico stimato (kWh):", value=0.0)
        em_energia = kwh_stimati * fe_energia
        
        # ACQUA
        st.markdown("#### Acqua")
        tech_acqua = st.selectbox("Approvvigionamento/Trattamento Acqua:", df_lci['Acqua']['Elemento'].unique())
        fe_acqua = float(str(df_lci['Acqua'][df_lci['Acqua']['Elemento'] == tech_acqua]['Fattore_Emissione'].values[0]).replace(',', '.'))
        litri_stimati = st.number_input("Consumo idrico stimato (litri):", value=0.0)
        em_acqua = litri_stimati * fe_acqua
        
        submit_btn = st.form_submit_button("Salva Calcolo per questa Attività")
        
        if submit_btn:
            totale_eq = em_materiali + em_rifiuti + em_macchinari + em_trasporti + em_energia + em_acqua
            
            record = {
                "Attività": row_att['Label_Attivita'],
                "ID": id_att, "Nome": nome_att, "Struttura": nome_strutt, "WBS_L": wbs_lav, "WBS_M": wbs_mat,
                "Durata": durata_giorni, 
                "Data Inizio": row_att['Data_Inizio_Calc'].date(), 
                "Data Fine": row_att['Data_Fine_Calc'].date(),
                "Materiali [kgCO2 eq]": em_materiali, "Rifiuti [kgCO2 eq]": em_rifiuti, 
                "Macchinari [kgCO2 eq]": em_macchinari, "Trasporti [kgCO2 eq]": em_trasporti,
                "Energia [kgCO2 eq]": em_energia, "Acqua [kgCO2 eq]": em_acqua, 
                "Totale [kgCO2 eq]": totale_eq
            }
            
            st.session_state.configurazioni_salvate[row_att['Label_Attivita']] = {
                "Carburante": carb_trasporto, "Consumo_mezzo": consumo_medio,
                "Energia": tech_energia, "FE_Energia": fe_energia,
                "Acqua": tech_acqua, "FE_Acqua": fe_acqua
            }
            
            lista_ris = [r for r in st.session_state.risultati_progettazione if r['Attività'] != row_att['Label_Attivita']]
            lista_ris.append(record)
            st.session_state.risultati_progettazione = sorted(lista_ris, key=lambda x: x['Data Inizio'])
            st.success("Attività salvata con successo! Vedi sotto i risultati aggiornati.")

    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 3: RISULTATI E GRAFICI PROGETTAZIONE ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("3. Risultati Fase di Progettazione")
    
    df_risultati = pd.DataFrame(st.session_state.risultati_progettazione)
    colonne_visibili = ["Attività", "Durata", "Data Inizio", "Data Fine", "Materiali [kgCO2 eq]", "Rifiuti [kgCO2 eq]", "Macchinari [kgCO2 eq]", "Trasporti [kgCO2 eq]", "Energia [kgCO2 eq]", "Acqua [kgCO2 eq]", "Totale [kgCO2 eq]"]
    st.dataframe(df_risultati[colonne_visibili], use_container_width=True)
    
    gwp_totale = df_risultati["Totale [kgCO2 eq]"].sum()
    st.markdown(f"<div style='padding:15px; background-color:#e8f4f8; border-left: 5px solid #2980b9; border-radius:4px;'><h4>GWP Totale Cantiere (Stima): {gwp_totale:,.2f} kg CO₂ eq</h4></div>", unsafe_allow_html=True)
    
    dati_giornalieri = []
    for _, row in df_risultati.iterrows():
        giorni_lavoro = pd.bdate_range(start=row['Data Inizio'], end=row['Data Fine'])
        n_giorni = len(giorni_lavoro) if len(giorni_lavoro) > 0 else 1
        
        for g in giorni_lavoro:
            dati_giornalieri.append({
                "Data": g, "Settimana": g.strftime('%Y-W%V'),
                "Materiali": row["Materiali [kgCO2 eq]"]/n_giorni, "Rifiuti": row["Rifiuti [kgCO2 eq]"]/n_giorni,
                "Macchinari": row["Macchinari [kgCO2 eq]"]/n_giorni, "Trasporti": row["Trasporti [kgCO2 eq]"]/n_giorni,
                "Energia": row["Energia [kgCO2 eq]"]/n_giorni, "Acqua": row["Acqua [kgCO2 eq]"]/n_giorni,
                "Totale": row["Totale [kgCO2 eq]"]/n_giorni
            })
    
    if dati_giornalieri:
        df_giornaliero = pd.DataFrame(dati_giornalieri)
        df_settimanale = df_giornaliero.groupby('Settimana').sum(numeric_only=True).reset_index()
        
        col1, col2 = st.columns(2)
        with col1:
            fig_bar = px.bar(df_settimanale, x="Settimana", y=["Materiali", "Rifiuti", "Macchinari", "Trasporti", "Energia", "Acqua"], title="Andamento Settimanale Emissioni per Categoria", barmode='stack')
            media_sett_totale = df_settimanale['Totale'].mean()
            fig_bar.add_hline(y=media_sett_totale, line_dash="dash", line_color="black", annotation_text="Media Costante")
            st.plotly_chart(fig_bar, use_container_width=True)
            
        with col2:
            df_pie = pd.DataFrame({
                "Parametro": ["Materiali", "Rifiuti", "Macchinari", "Trasporti", "Energia", "Acqua"],
                "Valore": [df_risultati["Materiali [kgCO2 eq]"].sum(), df_risultati["Rifiuti [kgCO2 eq]"].sum(), df_risultati["Macchinari [kgCO2 eq]"].sum(), df_risultati["Trasporti [kgCO2 eq]"].sum(), df_risultati["Energia [kgCO2 eq]"].sum(), df_risultati["Acqua [kgCO2 eq]"].sum()]
            })
            fig_pie = px.pie(df_pie, names="Parametro", values="Valore", title="Incidenza Percentuale", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 4: FASE DI CANTIERE (DATI SENSORI) ---
if st.session_state.risultati_progettazione:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("4. Monitoraggio Fase di Cantiere (Dati Dinamici Sensori)")
    
    col_s1, col_s2, col_s3 = st.columns(3)
    file_trasporti = col_s1.file_uploader("Sensori Trasporti (CSV)", type=['csv'])
    file_energia = col_s2.file_uploader("Sensori Energia (CSV)", type=['csv'])
    file_acqua = col_s3.file_uploader("Sensori Acqua (CSV)", type=['csv'])
    
    if st.button("Ricalcola Emissioni con Dati di Cantiere"):
        df_cantiere = df_risultati.copy()
        
        if file_trasporti:
            df_t = pd.read_csv(file_trasporti)
            # Ricreiamo la label dinamica se le colonne sono presenti
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
                        fattore = float(str(df_lci['Trasporti'][df_lci['Trasporti']['Fuel'] == conf['Carburante']]['kg CO2 per kg of fuel'].values[0]).replace(',', '.'))
                        df_cantiere.loc[df_cantiere['Attività'] == label, 'Trasporti [kgCO2 eq]'] = km_totali * conf['Consumo_mezzo'] * fattore
        
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
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Energia [kgCO2 eq]'] = kwh_totali * st.session_state.configurazioni_salvate[label]['FE_Energia']
                    
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
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Acqua [kgCO2 eq]'] = l_totali * st.session_state.configurazioni_salvate[label]['FE_Acqua']
        
        df_cantiere['Totale [kgCO2 eq]'] = df_cantiere[['Materiali [kgCO2 eq]', 'Rifiuti [kgCO2 eq]', 'Macchinari [kgCO2 eq]', 'Trasporti [kgCO2 eq]', 'Energia [kgCO2 eq]', 'Acqua [kgCO2 eq]']].sum(axis=1)
        
        st.dataframe(df_cantiere[colonne_visibili], use_container_width=True)
        gwp_cant = df_cantiere['Totale [kgCO2 eq]'].sum()
        st.markdown(f"<div style='padding:15px; background-color:#e8f8e8; border-left: 5px solid #27ae60; border-radius:4px;'><h4>GWP Totale Cantiere (Reale): {gwp_cant:,.2f} kg CO₂ eq</h4></div>", unsafe_allow_html=True)
        
        st.subheader("Confronto: Progettazione vs Cantiere Reale")
        df_comp = pd.DataFrame({
            "Fase": ["Progettazione"]*len(df_risultati) + ["Cantiere"]*len(df_cantiere),
            "Attività": pd.concat([df_risultati['Attività'], df_cantiere['Attività']]),
            "Totale Emissioni": pd.concat([df_risultati['Totale [kgCO2 eq]'], df_cantiere['Totale [kgCO2 eq]']])
        })
        fig_comp = px.bar(df_comp, x="Attività", y="Totale Emissioni", color="Fase", barmode="group", title="Confronto Emissioni per Singola Attività")
        st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

