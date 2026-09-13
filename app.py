import streamlit as st
import pandas as pd
import plotly.express as px

# Impostazioni della pagina
st.set_page_config(page_title="Dashboard Emissioni Cantiere", layout="wide")

st.title("🏗️ Monitoraggio Emissioni CO₂e - Cantiere Lineare")
st.markdown("Questa piattaforma incrocia i dati previsti di cantiere con un inventario standard per calcolare le emissioni di **kg di CO₂ equivalente per metro lineare**.")

# 1. Sezione di Caricamento File
col1, col2 = st.columns(2)
with col1:
    file_inventario = st.file_uploader("1. Carica l'Inventario Standardizzato (File Excel .xlsx)", type=['xlsx'])
with col2:
    file_cantiere = st.file_uploader("2. Carica i Dati di Progetto/Cantiere (File .csv)", type=['csv'])

# 2. Elaborazione dei Dati
if file_inventario and file_cantiere:
    
    try:
        # A. Lettura dell'Inventario (Fogli multipli)
        xls = pd.ExcelFile(file_inventario)
        df_inventario = pd.DataFrame()
        
        # Uniamo tutti i fogli dell'Excel in un'unica tabella
        for foglio in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=foglio)
            df_temp['Parametro'] = foglio 
            df_inventario = pd.concat([df_inventario, df_temp], ignore_index=True)
            
        # B. Lettura del CSV di Cantiere
        df_cantiere = pd.read_csv(file_cantiere)
        
        # C. Incrocio dei dati (Merge)
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        # Avviso se ci sono elementi nel CSV non trovati nell'Excel
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.warning(f"Attenzione! I seguenti elementi del CSV non sono stati trovati nell'inventario Excel: {mancanti['Elemento'].unique()}")
        
        # D. Calcoli Matematici
        # Quantità * Fattore = kg CO2 totali
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        
        # Normalizzazione per metro lineare: CO2 / Metri
        df_completo['CO2_kg_al_metro'] = df_completo['CO2_Totale_kg'] / df_completo['Metri_Lineari']
        
        st.success("Dati elaborati con successo!")
        
        # 3. Visualizzazione Grafici
        st.divider()
        
        # Grafico 1: Emissioni divise per parametro (Grafico a barre impilate)
        # Raggruppiamo per Data e Parametro
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
        # Raggruppiamo solo per Data
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
        st.info("Assicurati che i nomi delle colonne nel CSV siano: Data, Parametro, Elemento, Quantita, Metri_Lineari. E che i fogli Excel contengano: Elemento, Fattore_Emissione.")

else:
    st.info("Attesa caricamento file... Carica sia l'Inventario (Excel) che i Dati previsti (CSV) per generare i grafici.")