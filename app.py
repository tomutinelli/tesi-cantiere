import streamlit as st
import pandas as pd
import pdfplumber
import json
import os
from datetime import datetime

# ==========================================
# 1. FUNZIONI DI ESTRAZIONE DATI (PDF)
# ==========================================

def estrai_cme(pdf_file):
    """
    Estrae le voci dal Computo Metrico.
    Struttura attesa (da adattare al layout reale): 
    [Codice, Descrizione Lavorazione/Materiale, Unità di Misura, Quantità, Fattore_Emissione_LCA]
    """
    dati = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tabelle = page.extract_tables()
            for tab in tabelle:
                for riga in tab:
                    if len(riga) >= 4 and riga[0]:
                        dati.append({
                            "Codice_CME": str(riga[0]).strip(),
                            "Materiale_Lavorazione": str(riga[1]).strip()[:80], # Troncato per ottimizzare i token LLM
                            "Quantita_Totale": pd.to_numeric(str(riga[3]).replace(',', '.'), errors='coerce'),
                            # Valore fittizio di default (da sostituire con database ICE o ISPRA reali)
                            "Fattore_Emissione_kgCO2": 2.5 
                        })
    return pd.DataFrame(dati).dropna(subset=['Quantita_Totale'])

def estrai_cronoprogramma(pdf_file):
    """
    Estrae le fasi e le date dal Cronoprogramma in PDF.
    Struttura attesa: [Nome Fase, Data Inizio, Data Fine]
    """
    dati = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tabelle = page.extract_tables()
            for tab in tabelle:
                for riga in tab:
                    if len(riga) >= 3 and riga[0]:
                        try:
                            inizio = pd.to_datetime(riga[1], format='%d/%m/%Y', errors='coerce')
                            fine = pd.to_datetime(riga[2], format='%d/%m/%Y', errors='coerce')
                            if pd.notnull(inizio) and pd.notnull(fine):
                                dati.append({
                                    "Nome_Fase": str(riga[0]).strip(),
                                    "Data_Inizio": inizio,
                                    "Data_Fine": fine
                                })
                        except Exception:
                            continue
                            
    df_crono = pd.DataFrame(dati)
    if not df_crono.empty:
        df_crono['Durata_Giorni'] = (df_crono['Data_Fine'] - df_crono['Data_Inizio']).dt.days + 1
    return df_crono

# ==========================================
# 2. MOTORE DI MAPPATURA (LLM)
# ==========================================

def genera_matching_materiali_fasi(lista_materiali, lista_fasi):
    """
    Invia i materiali del CME e le fasi del Cronoprogramma all'LLM.
    Restituisce un JSON che associa ogni Codice_CME al Nome_Fase pertinente.
    """
    import google.generativeai as genai
    
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.warning("API Key mancante per il matching. Imposta la variabile d'ambiente GEMINI_API_KEY.")
        return pd.DataFrame()
        
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-pro')
    
    prompt = f"""
    Sei un ingegnere edile che gestisce un cantiere.
    Hai due set di dati:
    1. Fasi del Cronoprogramma: {lista_fasi}
    2. Materiali/Lavorazioni dal Computo (CME): {lista_materiali}
    
    Devi capire dal 'Nome_Fase' quale materiale verrà utilizzato in quella fase.
    Restituisci ESCLUSIVAMENTE un JSON con questa struttura esatta:
    [
        {{"Codice_CME": "codice_associato", "Nome_Fase": "nome_fase_esatto"}}
    ]
    Non aggiungere markdown, solo il JSON puro.
    """
    
    risposta = model.generate_content(prompt)
    
    try:
        testo_json = risposta.text.replace('```json', '').replace('```', '').strip()
        mappatura = json.loads(testo_json)
        return pd.DataFrame(mappatura)
    except json.JSONDecodeError:
        st.error("Errore nella generazione del matching logico dall'IA.")
        return pd.DataFrame()

# ==========================================
# 3. DISTRIBUZIONE E CALCOLO EMISSIONI
# ==========================================

def calcola_emissioni_giornaliere(df_cme, df_crono, df_mapping):
    """
    Unisce i dataframe, calcola il rateo giornaliero e l'emissione di CO2 eq al giorno.
    """
    df_merged = df_cme.merge(df_mapping, on="Codice_CME", how="inner")
    df_final = df_merged.merge(df_crono, on="Nome_Fase", how="inner")
    
    # Distribuzione lineare (In futuro integrabile con curve di crescita Gompertz per andamenti non lineari)
    df_final['Quantita_al_Giorno'] = df_final['Quantita_Totale'] / df_final['Durata_Giorni']
    df_final['kg_CO2_Giorno'] = df_final['Quantita_al_Giorno'] * df_final['Fattore_Emissione_kgCO2']
    
    # Esplosione del calendario per il monitoraggio giornaliero
    dati_esplosi = []
    for _, row in df_final.iterrows():
        date_range = pd.date_range(start=row['Data_Inizio'], end=row['Data_Fine'])
        for data in date_range:
            dati_esplosi.append({
                'Data_Cantiere': data.strftime('%Y-%m-%d'),
                'Fase_Cronoprogramma': row['Nome_Fase'],
                'Lavorazione_CME': row['Materiale_Lavorazione'],
                'Quantita_Impiegata': round(row['Quantita_al_Giorno'], 2),
                'Emissioni_kg_CO2': round(row['kg_CO2_Giorno'], 2)
            })
                
    return pd.DataFrame(dati_esplosi)

# ==========================================
# 4. DASHBOARD UI
# ==========================================

st.set_page_config(page_title="EcoSite Tracker", layout="wide")
st.title("🌱 EcoSite Tracker: Monitoraggio Emissioni Cantiere")
st.markdown("""
Applicativo per il monitoraggio giornaliero delle emissioni di carbonio nei cantieri complessi (es. opere infrastrutturali lineari, scavi metropolitani). 
Allinea automaticamente le lavorazioni del CME alle fasi del cronoprogramma per ricavare i valori esatti di emissione in kg CO2/giorno.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Caricamento Documentazione")
    cme_pdf = st.file_uploader("Computo Metrico (PDF)", type=['pdf'], key="cme")
    crono_pdf = st.file_uploader("Cronoprogramma (PDF)", type=['pdf'], key="crono")

if cme_pdf and crono_pdf:
    with st.spinner("Estrazione dati dai PDF in corso..."):
        df_cme = estrai_cme(cme_pdf)
        df_crono = estrai_cronoprogramma(crono_pdf)
        
    with col2:
        st.subheader("Sintesi Dati Estratti")
        st.write(f"**Voci CME trovate:** {len(df_cme)}")
        st.write(f"**Fasi Cronoprogramma trovate:** {len(df_crono)}")
        
    st.divider()
        
    if st.button("🚀 Avvia Matching IA e Calcola Emissioni Giornaliere", type="primary"):
        with st.spinner("Analisi semantica in corso... L'IA sta associando i materiali alle fasi del cronoprogramma..."):
            
            lista_mat = df_cme[['Codice_CME', 'Materiale_Lavorazione']].to_dict('records')
            lista_fas = df_crono['Nome_Fase'].tolist()
            
            df_mapping = genera_matching_materiali_fasi(lista_mat, lista_fas)
            
            if not df_mapping.empty:
                st.success("Matching e allocazione completati con successo!")
                
                df_emissioni = calcola_emissioni_giornaliere(df_cme, df_crono, df_mapping)
                
                st.subheader("Cruscotto Emissioni Giornaliere")
                st.dataframe(df_emissioni, use_container_width=True)
                
                # Raggruppamento metriche per andamento globale
                df_totale_giorno = df_emissioni.groupby('Data_Cantiere')['Emissioni_kg_CO2'].sum().reset_index()
                st.line_chart(data=df_totale_giorno, x='Data_Cantiere', y='Emissioni_kg_CO2')
                
                st.download_button(
                    label="📥 Esporta Database LCA (CSV)",
                    data=df_emissioni.to_csv(index=False).encode('utf-8'),
                    file_name='ecosite_tracker_emissioni_giornaliere.csv',
                    mime='text/csv'
                )
