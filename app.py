import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# Impostazioni della pagina
st.set_page_config(page_title="Dashboard Emissioni Cantiere", layout="wide")

st.title("🏗️ Monitoraggio Emissioni CO₂e - Cantiere")

# --- SCELTA TIPO DI CANTIERE ---
st.markdown("### ⚙️ Impostazioni di Normalizzazione")
tipo_cantiere = st.radio(
    "Seleziona il tipo di cantiere per calcolare le emissioni specifiche:",
    options=[
        "Cantiere Lineare (normalizzazione per metro lineare - m)", 
        "Cantiere Standard (normalizzazione per metro quadro - m²)"
    ]
)

if "Lineare" in tipo_cantiere:
    unita = "m"
    descrizione_unita = "metro lineare"
else:
    unita = "m²"
    descrizione_unita = "metro quadro"

st.markdown(f"Questa piattaforma analizza i dati di cantiere utilizzando un **database LCI integrato** per calcolare le emissioni di **kg di CO₂ equivalente per {descrizione_unita}**.")
st.divider()

# --- LETTURA AUTOMATICA DEL FILE LCI (EXCEL) ---
# Usiamo cache_data per caricarlo una sola volta e velocizzare il sito
@st.cache_data
def carica_database_lci(percorso_file):
    if not os.path.exists(percorso_file):
        return None
        
    xls = pd.ExcelFile(percorso_file)
    df_inv = pd.DataFrame()
    
    mappatura = {
        'Materiali': {'nome_elemento': 'nome_materiale', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Rifiuti': {'nome_elemento': 'tipo_rifiuto', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Trasporti e Macchinari': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'},
        'Energia': {'nome_elemento': 'Tecnologia di generazione elettrica', 'nome_fattore': 'Fattori di emissione (kg CO2eq/kWh)'}
    }
    
    for foglio in xls.sheet_names:
        if foglio in mappatura:
            df_temp = pd.read_excel(xls, sheet_name=foglio)
            col_el = mappatura[foglio]['nome_elemento']
            col_fat = mappatura[foglio]['nome_fattore']
            
            if col_el in df_temp.columns and col_fat in df_temp.columns:
                df_temp = df_temp[[col_el, col_fat]].copy()
                df_temp.rename(columns={col_el: 'Elemento', col_fat: 'Fattore_Emissione'}, inplace=True)
                df_temp['Parametro'] = foglio 
                df_inv = pd.concat([df_inv, df_temp], ignore_index=True)
                
    df_inv = df_inv.dropna(subset=['Elemento', 'Fattore_Emissione'])
    return df_inv

# Carichiamo il database nascosto
percorso_lci = "LCI.xlsx"
df_inventario = carica_database_lci(percorso_lci)

if df_inventario is None:
    st.error("⚠️ Errore di sistema: Il file del database 'LCI.xlsx' non è stato trovato sul server. Assicurati di averlo caricato su GitHub insieme ad app.py.")
    st.stop()
else:
    st.success("✅ Database LCI (Fattori di Emissione) caricato correttamente e pronto all'uso.")

# --- CARICAMENTO DATI UTENTE ---
st.markdown("### 📂 Inserimento Dati di Progetto")
file_cantiere = st.file_uploader("Carica i Dati Giornalieri di Progetto (File .csv)", type=['csv'])

# --- ELABORAZIONE DATI ---
if file_cantiere:
    try:
        # B. Lettura del CSV di Cantiere
        df_cantiere = pd.read_csv(file_cantiere)
        
        # Gestione flessibile della colonna di normalizzazione
        if 'Metri_Lineari' in df_cantiere.columns:
            df_cantiere.rename(columns={'Metri_Lineari': 'Avanzamento'}, inplace=True)
        elif 'Metri_Quadri' in df_cantiere.columns:
            df_cantiere.rename(columns={'Metri_Quadri': 'Avanzamento'}, inplace=True)
        elif 'Avanzamento' not in df_cantiere.columns:
            st.error("Il file CSV deve contenere una colonna chiamata 'Metri_Lineari', 'Metri_Quadri' o 'Avanzamento'.")
            st.stop()
        
        df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
        
        # C. Incrocio dei dati
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.error(f"Attenzione! I seguenti elementi del CSV non sono presenti nel database LCI: {mancanti['Elemento'].unique().tolist()}")
        
        # D. Calcoli Matematici
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        df_completo['CO2_Normalizzata'] = df_completo['CO2_Totale_kg'] / df_completo['Avanzamento']
        
        st.divider()

        # ----------------------------------------
        # FILTRO TEMPORALE
        # ----------------------------------------
        st.header("📅 Filtra Dati per Periodo")
        
        min_date = df_completo['Data_dt'].min().date()
        max_date = df_completo['Data_dt'].max().date()

        date_range = st.date_input(
            "Seleziona l'intervallo di tempo da analizzare (Clicca sulla prima e sull'ultima data):",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        if len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df_completo['Data_dt'].dt.date >= start_date) & (df_completo['Data_dt'].dt.date <= end_date)
            df_filtrato = df_completo.loc[mask].copy()
        else:
            df_filtrato = df_completo.copy()
            st.warning("Seleziona una data di fine per applicare il filtro.")

        # ----------------------------------------
        # VISUALIZZAZIONE GRAFICI
        # ----------------------------------------
        st.header("📊 Analisi delle Emissioni")
        
        if df_filtrato.empty:
            st.error("Nessun dato presente nell'intervallo di date selezionato.")
        else:
            # Grafico 1: Totale
            st.subheader(f"1. Andamento delle Emissioni Totali (per {descrizione_unita})")
            df_totale = df_filtrato.groupby('Data')['CO2_Normalizzata'].sum().reset_index()
            
            fig_tot = px.bar(
                df_totale, 
                x='Data', 
                y='CO2_Normalizzata',
                title=f"Totale Emissioni dal {start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')}",
                labels={'CO2_Normalizzata': f'kg CO₂e / {unita}', 'Data': 'Giorno'},
                text_auto='.2f' 
            )
            fig_tot.update_traces(marker_color='#d62728') 
            fig_tot.update_layout(height=450, xaxis_tickangle=-45) 
            st.plotly_chart(fig_tot, use_container_width=True)

            # Grafici 2: Separati
            st.divider()
            st.subheader("2. Dettaglio per Singolo Parametro")
            
            colori_parametri = {
                'Materiali': '#1f77b4',         
                'Rifiuti': '#ff7f0e',           
                'Trasporti e Macchinari': '#2ca02c', 
                'Energia': '#ffbb78'            
            }
            
            parametri_presenti = df_filtrato['Parametro'].dropna().unique()
            col_grafici = st.columns(2)
            
            for idx, parametro in enumerate(parametri_presenti):
                df_param_singolo = df_filtrato[df_filtrato['Parametro'] == parametro].groupby('Data')['CO2_Normalizzata'].sum().reset_index()
                colore = colori_parametri.get(parametro, '#7f7f7f')
                
                fig = px.bar(
                    df_param_singolo,
                    x='Data',
                    y='CO2_Normalizzata',
                    title=f"Emissioni: {parametro}",
                    labels={'CO2_Normalizzata': f'kg CO₂e / {unita}', 'Data': ''},
                    text_auto='.2f'
                )
                fig.update_traces(marker_color=colore)
                fig.update_layout(height=350, margin=dict(t=40, b=20, l=10, r=10), xaxis_tickangle=-45)
                
                col_grafici[idx % 2].plotly_chart(fig, use_container_width=True)

            # Tabella Dati
            st.divider()
            with st.expander("Mostra i dati completi in formato tabella (Clicca per espandere)"):
                st.dataframe(df_filtrato.drop(columns=['Data_dt']))

    except Exception as e:
        st.error(f"Si è verificato un errore nell'elaborazione del CSV: {e}")
        st.info("Verifica che il CSV sia corretto e formattato adeguatamente.")
else:
    st.info("Attesa caricamento file... Carica il CSV di Progetto per avviare l'analisi.")