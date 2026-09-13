import streamlit as st
import pandas as pd
import plotly.express as px

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
        # A. Lettura dell'Inventario (Fogli multipli con intestazioni diverse)
        xls = pd.ExcelFile(file_inventario)
        df_inventario = pd.DataFrame()
        
        # Questa è la mappa che "traduce" le colonne del tuo Excel specifico
        mappatura_colonne = {
            'Materiali': {'nome_elemento': 'nome_materiale', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
            'Rifiuti': {'nome_elemento': 'tipo_rifiuto', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
            'Trasporti e Macchinari': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'},
            'Energia': {'nome_elemento': 'Tecnologia di generazione elettrica', 'nome_fattore': 'Fattori di emissione (kg CO2eq/kWh)'}
        }
        
        for foglio in xls.sheet_names:
            if foglio in mappatura_colonne:
                # Legge il foglio
                df_temp = pd.read_excel(xls, sheet_name=foglio)
                
                # Recupera i nomi delle colonne da cercare per questo specifico foglio
                col_elemento = mappatura_colonne[foglio]['nome_elemento']
                col_fattore = mappatura_colonne[foglio]['nome_fattore']
                
                # Estrae solo le due colonne che ci interessano e le rinomina in modo standard
                if col_elemento in df_temp.columns and col_fattore in df_temp.columns:
                    df_temp = df_temp[[col_elemento, col_fattore]].copy()
                    df_temp.rename(columns={col_elemento: 'Elemento', col_fattore: 'Fattore_Emissione'}, inplace=True)
                    
                    # Aggiunge il nome del foglio (Parametro)
                    df_temp['Parametro'] = foglio 
                    
                    # Aggiunge i dati di questo foglio al database generale dell'inventario
                    df_inventario = pd.concat([df_inventario, df_temp], ignore_index=True)
                else:
                    st.warning(f"Nel foglio '{foglio}' non ho trovato le colonne previste.")
            
        # Rimuove eventuali righe vuote dall'inventario
        df_inventario = df_inventario.dropna(subset=['Elemento', 'Fattore_Emissione'])

        # B. Lettura del CSV di Cantiere
        df_cantiere = pd.read_csv(file_cantiere)
        
        # C. Incrocio dei dati (Merge)
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        # Avviso se ci sono elementi nel CSV non trovati nell'Excel
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.error(f"Attenzione! I seguenti elementi del CSV non sono stati trovati nell'inventario Excel: {mancanti['Elemento'].unique().tolist()}")
        
        # D. Calcoli Matematici
        # Quantità * Fattore = kg CO2 totali
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        
        # Normalizzazione per metro lineare: CO2 / Metri
        df_completo['CO2_kg_al_metro'] = df_completo['CO2_Totale_kg'] / df_completo['Metri_Lineari']
        
        st.success("Dati elaborati e incrociati con successo!")
        
        # 3. Visualizzazione Grafici
        st.divider()
        
        # Grafico 1: Emissioni divise per parametro (Grafico a barre impilate)
        df_param = df_completo.groupby(['Data', 'Parametro'])['CO2_kg_al_metro'].sum().reset_index()
        
        fig_param = px.bar(
            df_param, 
            x='Data', 
            y='CO2_kg_al_metro', 
            color='Parametro',
            title="Dettaglio giornaliero delle emissioni suddivise per parametro",
            labels={'CO2_kg_al_metro': 'kg CO₂e / metro lineare', 'Data': 'Giorno / Fase'},
            barmode='stack' 
        )
        st.plotly_chart(fig_param, use_container_width=True)

        # Grafico 2: Emissioni totali assolute per giorno
        df_totale = df_completo.groupby('Data')['CO2_kg_al_metro'].sum().reset_index()
        
        fig_tot = px.bar(
            df_totale, 
            x='Data', 
            y='CO2_kg_al_metro',
            title="Andamento delle emissioni totali normalizzate al metro lineare",
            labels={'CO2_kg_al_metro': 'kg CO₂e / metro lineare', 'Data': 'Giorno / Fase'},
            text_auto='.2f' 
        )
        fig_tot.update_traces(marker_color='indianred')
        st.plotly_chart(fig_tot, use_container_width=True)

        # Tabella di riepilogo
        with st.expander("Mostra i dati completi in formato tabella (Clicca per espandere)"):
            st.dataframe(df_completo)

    except Exception as e:
        st.error(f"Si è verificato un errore nell'elaborazione: {e}")
        st.info("Verifica che il CSV sia corretto e che non ci siano anomalie nel file Excel.")

else:
    st.info("Attesa caricamento file... Carica l'Inventario LCI e il CSV di Progetto per iniziare.")