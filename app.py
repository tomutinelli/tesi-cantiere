import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Impostazioni della pagina
st.set_page_config(page_title="Dashboard Emissioni Cantiere", layout="wide")

st.title("🏗️ Monitoraggio Emissioni CO₂e - Cantiere Lineare")
st.markdown("Questa piattaforma incrocia i dati previsti di cantiere con l'inventario LCI per calcolare le emissioni di **kg di CO₂ equivalente per metro lineare**.")

# 1. Sezione di Caricamento File
col1, col2 = st.columns(2)
with col1:
    file_inventario = st.file_uploader("1. Carica l'Inventario LCI (File Excel .xlsx)", type=['xlsx'])
with col2:
    file_cantiere = st.file_uploader("2. Carica i Dati di Progetto (File .csv)", type=['csv'])

# 2. Elaborazione dei Dati
if file_inventario and file_cantiere:
    
    try:
        # A. Lettura dell'Inventario 
        xls = pd.ExcelFile(file_inventario)
        df_inventario = pd.DataFrame()
        
        # Mappa che "traduce" le colonne del tuo Excel
        mappatura_colonne = {
            'Materiali': {'nome_elemento': 'nome_materiale', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
            'Rifiuti': {'nome_elemento': 'tipo_rifiuto', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
            'Trasporti e Macchinari': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'},
            'Energia': {'nome_elemento': 'Tecnologia di generazione elettrica', 'nome_fattore': 'Fattori di emissione (kg CO2eq/kWh)'}
        }
        
        for foglio in xls.sheet_names:
            if foglio in mappatura_colonne:
                df_temp = pd.read_excel(xls, sheet_name=foglio)
                col_elemento = mappatura_colonne[foglio]['nome_elemento']
                col_fattore = mappatura_colonne[foglio]['nome_fattore']
                
                if col_elemento in df_temp.columns and col_fattore in df_temp.columns:
                    df_temp = df_temp[[col_elemento, col_fattore]].copy()
                    df_temp.rename(columns={col_elemento: 'Elemento', col_fattore: 'Fattore_Emissione'}, inplace=True)
                    df_temp['Parametro'] = foglio 
                    df_inventario = pd.concat([df_inventario, df_temp], ignore_index=True)
                else:
                    st.warning(f"Nel foglio '{foglio}' non ho trovato le colonne previste.")
            
        df_inventario = df_inventario.dropna(subset=['Elemento', 'Fattore_Emissione'])

        # B. Lettura del CSV di Cantiere
        df_cantiere = pd.read_csv(file_cantiere)
        
        # Convertiamo la colonna Data in formato datetime per permettere il filtraggio
        df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
        
        # C. Incrocio dei dati (Merge)
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.error(f"Attenzione! I seguenti elementi del CSV non sono stati trovati nell'inventario Excel: {mancanti['Elemento'].unique().tolist()}")
        
        # D. Calcoli Matematici
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        df_completo['CO2_kg_al_metro'] = df_completo['CO2_Totale_kg'] / df_completo['Metri_Lineari']
        
        st.success("Dati elaborati e incrociati con successo!")
        st.divider()

        # ----------------------------------------
        # FILTRO TEMPORALE (INTERVALLO DI DATE)
        # ----------------------------------------
        st.header("📅 Filtra Dati per Periodo")
        
        # Troviamo la data minima e massima nel dataset
        min_date = df_completo['Data_dt'].min().date()
        max_date = df_completo['Data_dt'].max().date()

        # Creiamo il widget del calendario
        date_range = st.date_input(
            "Seleziona l'intervallo di tempo da analizzare (Clicca sulla prima e sull'ultima data del periodo):",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        # Controlliamo che l'utente abbia selezionato un range valido (inizio e fine)
        if len(date_range) == 2:
            start_date, end_date = date_range
            # Filtriamo il dataframe originale mantenendo solo le righe comprese nel range
            mask = (df_completo['Data_dt'].dt.date >= start_date) & (df_completo['Data_dt'].dt.date <= end_date)
            df_filtrato = df_completo.loc[mask].copy()
        else:
            df_filtrato = df_completo.copy()
            st.warning("Seleziona una data di fine per applicare il filtro.")

        # ----------------------------------------
        # 3. VISUALIZZAZIONE GRAFICI (SUI DATI FILTRATI)
        # ----------------------------------------
        st.header("📊 Analisi delle Emissioni")
        
        if df_filtrato.empty:
            st.error("Nessun dato presente nell'intervallo di date selezionato.")
        else:
            # --- GRAFICO 1: TOTALE ---
            st.subheader("1. Andamento delle Emissioni Totali")
            df_totale = df_filtrato.groupby('Data')['CO2_kg_al_metro'].sum().reset_index()
            
            fig_tot = px.bar(
                df_totale, 
                x='Data', 
                y='CO2_kg_al_metro',
                title=f"Totale Emissioni dal {start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')}",
                labels={'CO2_kg_al_metro': 'kg CO₂e / m', 'Data': 'Giorno'},
                text_auto='.2f' 
            )
            fig_tot.update_traces(marker_color='#d62728') 
            fig_tot.update_layout(height=450, xaxis_tickangle=-45) 
            st.plotly_chart(fig_tot, use_container_width=True)

            # --- GRAFICI 2: SEPARATI PER PARAMETRO ---
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
                df_param_singolo = df_filtrato[df_filtrato['Parametro'] == parametro].groupby('Data')['CO2_kg_al_metro'].sum().reset_index()
                colore = colori_parametri.get(parametro, '#7f7f7f')
                
                fig = px.bar(
                    df_param_singolo,
                    x='Data',
                    y='CO2_kg_al_metro',
                    title=f"Emissioni: {parametro}",
                    labels={'CO2_kg_al_metro': 'kg CO₂e / m', 'Data': ''},
                    text_auto='.2f'
                )
                fig.update_traces(marker_color=colore)
                fig.update_layout(height=350, margin=dict(t=40, b=20, l=10, r=10), xaxis_tickangle=-45)
                
                col_grafici[idx % 2].plotly_chart(fig, use_container_width=True)

            # Tabella di riepilogo
            st.divider()
            with st.expander("Mostra i dati completi filtrati in formato tabella (Clicca per espandere)"):
                # Rimuoviamo la colonna datetime di supporto per non confondere chi legge la tabella
                st.dataframe(df_filtrato.drop(columns=['Data_dt']))

    except Exception as e:
        st.error(f"Si è verificato un errore nell'elaborazione: {e}")
        st.info("Verifica che il CSV sia corretto e che le date siano nel formato YYYY-MM-DD.")

else:
    st.info("Attesa caricamento file... Carica l'Inventario LCI e il CSV di Progetto per iniziare.")