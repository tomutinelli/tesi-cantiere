import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta
import numpy as np

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

# --- FUNZIONI DI CARICAMENTO DATI STATICI (DA REPOSITORY) ---
@st.cache_data
def load_static_data():
    try:
        # Carica i database statici presenti nel repository
        df_epd = pd.read_csv("epd_materiali.csv", sep=None, engine='python', encoding='latin1')
        df_lci = pd.read_excel("LCI.xlsx", sheet_name=None)
        df_macc = pd.read_excel("macchinari.xlsx", sep=None, engine='python', encoding='latin1')
        return df_epd, df_lci, df_macc
    except Exception as e:
        st.error(f"Errore nel caricamento dei database di sistema (EPD, LCI, Macchinari). Verifica che siano presenti nella cartella. Dettaglio: {e}")
        return None, None, None

df_epd, df_lci, df_macc = load_static_data()

# --- SEZIONE 1: CARICAMENTO DATI UTENTE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("1. Caricamento File di Progetto")
file_caricati = st.file_uploader("Carica i file CSV (Quantità Materiali e Cronoprogramma)", type=['csv'], accept_multiple_files=True)

df_quantita = None
df_crono = None

if file_caricati:
    for f in file_caricati:
        df_temp = pd.read_csv(f, sep=None, engine='python')
        # Riconoscimento automatico dei file in base alle colonne
        if 'Densità di massa' in df_temp.columns or 'Massa [kg]' in df_temp.columns:
            df_quantita = df_temp
        elif 'Data d’inizio' in df_temp.columns or 'Data di fine' in df_temp.columns:
            df_crono = df_temp

if df_crono is not None:
    # Ordina cronologicamente
    df_crono['Data d’inizio'] = pd.to_datetime(df_crono['Data d’inizio'])
    df_crono['Data di fine'] = pd.to_datetime(df_crono['Data di fine'])
    df_crono = df_crono.sort_values(by='Data d’inizio')
    
    # Crea etichette per il menu a tendina
    df_crono['Label_Attivita'] = df_crono['ID attività'].astype(str) + " - " + df_crono['Nome struttura'] + " (" + df_crono['WBS Lavorazione'] + ")"
    st.success("File caricati e interpretati correttamente!")
else:
    st.info("Carica il cronoprogramma per iniziare.")
st.markdown("</div>", unsafe_allow_html=True)

# --- SEZIONE 2: COMPILAZIONE ATTIVITA' ---
if df_crono is not None and df_epd is not None and df_lci is not None:
    st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
    st.subheader("2. Analisi Attività (Fase di Progettazione)")
    
    attivita_selezionata = st.selectbox("Scegli la lavorazione dal cronoprogramma:", df_crono['Label_Attivita'].tolist())
    
    # Recupero dati attività
    row_att = df_crono[df_crono['Label_Attivita'] == attivita_selezionata].iloc[0]
    id_att, nome_strutt, wbs_lav, wbs_mat = row_att['ID attività'], row_att['Nome struttura'], row_att['WBS Lavorazione'], row_att['WBS Materiale']
    
    # Recupero quantità materiale (se esiste)
    quantita_mat = 0
    if df_quantita is not None:
        match_q = df_quantita[(df_quantita['WBS Lavorazione'] == wbs_lav) & (df_quantita['WBS Materiale'] == wbs_mat)]
        if not match_q.empty:
            quantita_mat = match_q['Massa [kg]'].sum()
    
    st.info(f"Quantità materiale associata a questa attività (tramite WBS): **{quantita_mat:.2f} kg**")
    
    # --- FORM COMPILAZIONE ---
    with st.form(key=f"form_{attivita_selezionata}"):
        em_materiali = 0
        em_rifiuti = 0
        em_trasporti = 0
        em_energia = 0
        em_acqua = 0
        em_macchinari = 0
        
        # MATERIALI E RIFIUTI
        st.markdown("#### Materiali e Rifiuti")
        if quantita_mat > 0:
            # Filtra materiali EPD compatibili con l'attività (logica esemplificativa: mostriamo quelli del DB, andrebbe mappato il WBS)
            materiale_scelto = st.selectbox("Seleziona Tipo Materiale:", df_epd['Materiale'].unique())
            df_epd_filtrato = df_epd[df_epd['Materiale'] == materiale_scelto]
            
            st.dataframe(
                df_epd_filtrato[['Materiale', 'Produttore', 'Fattore emissivo', 'Trattamento rifiuti', 'Link PDF']],
                column_config={"Link PDF": st.column_config.LinkColumn("Link EPD")}, hide_index=True
            )
            
            scelta_specifica = st.selectbox("Scegli il prodotto specifico per il calcolo:", df_epd_filtrato['Produttore'])
            fattore_em_mat = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica]['Fattore emissivo'].values[0]
            fattore_em_rif = df_epd_filtrato[df_epd_filtrato['Produttore'] == scelta_specifica]['Trattamento rifiuti'].values[0]
            
            em_materiali = quantita_mat * float(fattore_em_mat)
            
            perc_rifiuti = st.number_input("% di rifiuti stimata (sfrido):", min_value=0.0, max_value=100.0, value=0.0) / 100
            q_rifiuti = quantita_mat * perc_rifiuti
            em_rifiuti = q_rifiuti * float(fattore_em_rif)
            
            # TRASPORTI (Materiali + Rifiuti)
            st.markdown("#### Trasporti")
            if 'Trasporti' in df_lci:
                carb_trasporto = st.selectbox("Carburante Trasporti:", df_lci['Trasporti']['Fuel'].dropna().unique())
                row_carb = df_lci['Trasporti'][df_lci['Trasporti']['Fuel'] == carb_trasporto].iloc[0]
                fattore_carb = float(row_carb['kg CO2 per kg of fuel'])
                consumo_medio = st.number_input("Consumo medio carburante mezzo (kg/km):", value=0.25)
                
                # Calcolo viaggi (max 28000 kg a viaggio)
                viaggi_mat = np.ceil(quantita_mat / 28000)
                viaggi_rif = np.ceil(q_rifiuti / 28000) if q_rifiuti > 0 else 0
                viaggi_tot = viaggi_mat + viaggi_rif
                
                em_trasporti = viaggi_tot * 30 * consumo_medio * fattore_carb
        else:
            st.warning("Nessun materiale trovato per questa attività. Sezioni Materiali, Rifiuti e Trasporti saltate.")
            carb_trasporto = None
            consumo_medio = 0
            
        # MACCHINARI
        st.markdown("#### Macchinari")
        macchinario = st.selectbox("Scegli Macchinario base:", df_macc['Macchinario'].unique())
        df_macc_filt = df_macc[df_macc['Macchinario'] == macchinario]
        st.dataframe(df_macc_filt[['Macchinario', 'Produttore', 'Fattore emissivo 8h', 'Link Scheda']], hide_index=True)
        
        macc_produttore = st.selectbox("Modello Specifico:", df_macc_filt['Produttore'])
        fe_macc_8h = float(df_macc_filt[df_macc_filt['Produttore'] == macc_produttore]['Fattore emissivo 8h'].values[0])
        
        num_macc = st.number_input("Numero di macchinari impiegati:", min_value=1, value=1)
        giorni_macc = st.number_input("Giorni di lavoro (per macchinario):", min_value=1, value=int(row_att['Durata']))
        em_macchinari = fe_macc_8h * num_macc * giorni_macc
        
        # ENERGIA
        st.markdown("#### Energia Elettrica")
        tech_energia = st.selectbox("Tecnologia generazione elettrica:", df_lci['Energia']['Tecnologia di generazione elettrica'].unique())
        fe_energia = float(df_lci['Energia'][df_lci['Energia']['Tecnologia di generazione elettrica'] == tech_energia]['Fattori di emissione (kg CO2eq/kWh)'].values[0])
        kwh_stimati = st.number_input("Consumo elettrico stimato (kWh):", value=0.0)
        em_energia = kwh_stimati * fe_energia
        
        # ACQUA
        st.markdown("#### Acqua")
        tech_acqua = st.selectbox("Approvvigionamento/Trattamento Acqua:", df_lci['Acqua']['Elemento'].unique())
        fe_acqua = float(df_lci['Acqua'][df_lci['Acqua']['Elemento'] == tech_acqua]['Fattore_Emissione'].values[0])
        litri_stimati = st.number_input("Consumo idrico stimato (litri):", value=0.0)
        em_acqua = litri_stimati * fe_acqua
        
        submit_btn = st.form_submit_button("Salva Calcolo per questa Attività")
        
        if submit_btn:
            totale_eq = em_materiali + em_rifiuti + em_macchinari + em_trasporti + em_energia + em_acqua
            
            record = {
                "Attività": row_att['Label_Attivita'],
                "ID": id_att, "Nome": nome_strutt, "WBS_L": wbs_lav, "WBS_M": wbs_mat,
                "Durata": row_att['Durata'], "Data Inizio": row_att['Data d’inizio'].date(), "Data Fine": row_att['Data di fine'].date(),
                "Materiali [kgCO2 eq]": em_materiali, "Rifiuti [kgCO2 eq]": em_rifiuti, 
                "Macchinari [kgCO2 eq]": em_macchinari, "Trasporti [kgCO2 eq]": em_trasporti,
                "Energia [kgCO2 eq]": em_energia, "Acqua [kgCO2 eq]": em_acqua, 
                "Totale [kgCO2 eq]": totale_eq
            }
            
            # Salva i parametri di configurazione per usarli nella fase dinamica
            st.session_state.configurazioni_salvate[row_att['Label_Attivita']] = {
                "Carburante": carb_trasporto, "Consumo_mezzo": consumo_medio,
                "Energia": tech_energia, "FE_Energia": fe_energia,
                "Acqua": tech_acqua, "FE_Acqua": fe_acqua
            }
            
            # Aggiorna i risultati (sovrascrive se l'attività esiste già)
            lista_ris = [r for r in st.session_state.risultati_progettazione if r['Attività'] != row_att['Label_Attivita']]
            lista_ris.append(record)
            st.session_state.risultati_progettazione = sorted(lista_ris, key=lambda x: x['Data Inizio'])
            st.success("Attività salvata! Vedi i risultati sotto.")

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
    
    # Distribuzione Giornaliera per Grafici
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
    
    df_giornaliero = pd.DataFrame(dati_giornalieri)
    df_settimanale = df_giornaliero.groupby('Settimana').sum(numeric_only=True).reset_index()
    
    col1, col2 = st.columns(2)
    with col1:
        fig_bar = px.bar(df_settimanale, x="Settimana", y=["Materiali", "Rifiuti", "Macchinari", "Trasporti", "Energia", "Acqua"], title="Andamento Settimanale Emissioni per Categoria", barmode='stack')
        # Linea media costante totale
        media_sett_totale = df_settimanale['Totale'].mean()
        fig_bar.add_hline(y=media_sett_totale, line_dash="dash", line_color="black", annotation_text="Media Costante")
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col2:
        df_pie = pd.DataFrame({
            "Parametro": ["Materiali", "Rifiuti", "Macchinari", "Trasporti", "Energia", "Acqua"],
            "Valore": [df_risultati["Materiali [kgCO2 eq]"].sum(), df_risultati["Rifiuti [kgCO2 eq]"].sum(), df_risultati["Macchinari [kgCO2 eq]"].sum(), df_risultati["Trasporti [kgCO2 eq]"].sum(), df_risultati["Energia [kgCO2 eq]"].sum(), df_risultati["Acqua [kgCO2 eq]"].sum()]
        })
        fig_pie = px.pie(df_pie, names="Parametro", values="Valore", title="Incidenza Percentuale per Parametro", hole=0.4)
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
        
        # Logica ricalcolo Trasporti
        if file_trasporti:
            df_t = pd.read_csv(file_trasporti)
            # Raggruppa per attività (ID univoco generato)
            df_t['Label_Attivita'] = df_t['ID attività'].astype(str) + " - " + df_t['Nome struttura'] + " (" + df_t['WBS Lavorazione'] + ")"
            for label in df_t['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    conf = st.session_state.configurazioni_salvate[label]
                    km_totali = df_t[df_t['Label_Attivita'] == label]['km percorsi'].sum()
                    if conf['Carburante'] is not None:
                        fattore = float(df_lci['Trasporti'][df_lci['Trasporti']['Fuel'] == conf['Carburante']]['kg CO2 per kg of fuel'].values[0])
                        nuovo_valore = km_totali * conf['Consumo_mezzo'] * fattore
                        df_cantiere.loc[df_cantiere['Attività'] == label, 'Trasporti [kgCO2 eq]'] = nuovo_valore
        
        # Logica ricalcolo Energia
        if file_energia:
            df_e = pd.read_csv(file_energia)
            df_e['Label_Attivita'] = df_e['ID attività'].astype(str) + " - " + df_e['Nome struttura'] + " (" + df_e['WBS Lavorazione'] + ")"
            for label in df_e['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    kwh_totali = df_e[df_e['Label_Attivita'] == label]['consumo energetico [kWh]'].sum()
                    nuovo_valore = kwh_totali * st.session_state.configurazioni_salvate[label]['FE_Energia']
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Energia [kgCO2 eq]'] = nuovo_valore
                    
        # Logica ricalcolo Acqua
        if file_acqua:
            df_a = pd.read_csv(file_acqua)
            df_a['Label_Attivita'] = df_a['ID attività'].astype(str) + " - " + df_a['Nome struttura'] + " (" + df_a['WBS Lavorazione'] + ")"
            for label in df_a['Label_Attivita'].unique():
                if label in st.session_state.configurazioni_salvate:
                    l_totali = df_a[df_a['Label_Attivita'] == label]['consumo acqua [l]'].sum()
                    nuovo_valore = l_totali * st.session_state.configurazioni_salvate[label]['FE_Acqua']
                    df_cantiere.loc[df_cantiere['Attività'] == label, 'Acqua [kgCO2 eq]'] = nuovo_valore
        
        # Ricalcola Totali Cantiere
        df_cantiere['Totale [kgCO2 eq]'] = df_cantiere[['Materiali [kgCO2 eq]', 'Rifiuti [kgCO2 eq]', 'Macchinari [kgCO2 eq]', 'Trasporti [kgCO2 eq]', 'Energia [kgCO2 eq]', 'Acqua [kgCO2 eq]']].sum(axis=1)
        
        st.dataframe(df_cantiere[colonne_visibili], use_container_width=True)
        gwp_cant = df_cantiere['Totale [kgCO2 eq]'].sum()
        st.markdown(f"<div style='padding:15px; background-color:#e8f8e8; border-left: 5px solid #27ae60; border-radius:4px;'><h4>GWP Totale Cantiere (Reale): {gwp_cant:,.2f} kg CO₂ eq</h4></div>", unsafe_allow_html=True)
        
        # --- CONFRONTO FINALE GRAFICO ---
        st.subheader("Confronto: Progettazione vs Cantiere Reale")
        df_comp = pd.DataFrame({
            "Fase": ["Progettazione"]*len(df_risultati) + ["Cantiere"]*len(df_cantiere),
            "Attività": pd.concat([df_risultati['Attività'], df_cantiere['Attività']]),
            "Totale Emissioni": pd.concat([df_risultati['Totale [kgCO2 eq]'], df_cantiere['Totale [kgCO2 eq]']])
        })
        
        fig_comp = px.bar(df_comp, x="Attività", y="Totale Emissioni", color="Fase", barmode="group", title="Confronto Emissioni per Singola Attività")
        st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
