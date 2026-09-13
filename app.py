import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# Impostazioni della pagina
st.set_page_config(page_title="Dashboard Emissioni Cantiere", layout="wide")

# --- STILE CSS PERSONALIZZATO (Stile Wix / EcoSite Tracker) ---
st.markdown("""
<style>
    /* Importazione font pulito e moderno (Helvetica / Inter) */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
        color: #2b2b2b;
        background-color: #fcfcfc;
    }

    /* Nascondi intestazione e piè di pagina di Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Stile Card in stile Wix per i blocchi di contenuto */
    .wix-card {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -1px rgba(0, 0, 0, 0.01);
        margin-bottom: 20px;
    }

    /* Stile dei titoli */
    h1, h2, h3 {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
        font-weight: 600;
        color: #111827;
    }

    /* Personalizzazione pulsanti e selettori */
    .stRadio > label, .stNumberInput > label, .stDateInput > label, .stFileUploader > label {
        font-weight: 500;
        color: #374151;
    }
</style>
""", unsafe_allow_html=True)

# --- CONTENITORE PRINCIPALE STILE WIX ---
st.markdown("<h1 style='text-align: center; color: #111827; margin-bottom: 10px;'>EcoSite Tracker</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6b7280; font-size: 1.1rem; margin-bottom: 30px;'>Monitoraggio e Analisi LCA delle Emissioni in Fase di Progetto</p>", unsafe_allow_html=True)

# Blocco Guida in stile card
st.markdown("""
<div class="wix-card">
    <h3 style="margin-top: 0; color: #047857;">📖 Guida all'utilizzo</h3>
    <p style="color: #4b5563; line-height: 1.5;">
        Benvenuto nella dashboard di calcolo ambientale. Carica il file dei consumi previsti per generare l'analisi dettagliata delle emissioni normalizzate.
    </p>
    <ul style="color: #4b5563; line-height: 1.6; margin-bottom: 0;">
        <b>Requisiti CSV (4 colonne obbligatorie):</b> <code>Data</code> (AAAA-MM-GG), <code>Parametro</code> (Materiali, Rifiuti, Trasporti e Macchinari, Energia), <code>Elemento</code> (es. Calcestruzzo, Diesel) e <code>Quantita</code>.
    </ul>
</div>
""", unsafe_allow_html=True)

# --- IMPOSTAZIONI DI NORMALIZZAZIONE ---
st.markdown("<div class='wix-card'>", unsafe_allow_html=True)
st.subheader("⚙️ Parametri di Normalizzazione")
tipo_cantiere = st.radio(
    "Seleziona il tipo di opera:",
    options=[
        "Cantiere Lineare (normalizzazione per metro lineare - m)", 
        "Cantiere Standard (normalizzazione per metro quadro - m²)"
    ]
)

col_inp1, col_inp2 = st.columns([2, 1])
with col_inp1:
    if "Lineare" in tipo_cantiere:
        unita = "m"
        descrizione_unita = "metro lineare"
        dimensione_cantiere = st.number_input("Lunghezza totale dell'opera (metri):", min_value=0.1, value=100.0, step=1.0)
    else:
        unita = "m²"
        descrizione_unita = "metro quadro"
        dimensione_cantiere = st.number_input("Superficie totale dell'opera (metri quadri):", min_value=0.1, value=100.0, step=1.0)
st.markdown("</div>", unsafe_allow_html=True)

# --- LETTURA DATABASE LCI ---
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
                
    return df_inv.dropna(subset=['Elemento', 'Fattore_Emissione'])

percorso_lci = "LCI.xlsx"
df_inventario = carica_database_lci(percorso_lci)

if df_inventario is None:
    st.error("⚠️ Database 'LCI.xlsx' non trovato su GitHub.")
    st.stop()

# --- CARICAMENTO CSV UTENTE ---
st.markdown("<div class='wix-card'>", unsafe_allow_html=True)
st.subheader("📂 Caricamento Dati di Progetto")
file_cantiere = st.file_uploader("Seleziona il file .csv dei consumi giornalieri", type=['csv'])
st.markdown("</div>", unsafe_allow_html=True)

if file_cantiere:
    try:
        df_cantiere = pd.read_csv(file_cantiere)
        
        for col in ['Data', 'Parametro', 'Elemento', 'Quantita']:
            if col not in df_cantiere.columns:
                st.error(f"Errore: Manca la colonna obbligatoria '{col}' nel CSV.")
                st.stop()
        
        df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.error(f"Elementi non trovati nel database LCI: {mancanti['Elemento'].unique().tolist()}")
            st.stop()
            
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        df_completo['CO2_Normalizzata'] = df_completo['CO2_Totale_kg'] / dimensione_cantiere
        
        # --- FILTRO TEMPORALE ---
        st.markdown("<div class='wix-card'>", unsafe_allow_html=True)
        st.subheader("📅 Filtro Temporale")
        min_date = df_completo['Data_dt'].min().date()
        max_date = df_completo['Data_dt'].max().date()

        date_range = st.date_input(
            "Seleziona l'intervallo di date da visualizzare:",
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
        st.markdown("</div>", unsafe_allow_html=True)

        # --- VISUALIZZAZIONE GRAFICI ---
        st.markdown("<div class='wix-card'>", unsafe_allow_html=True)
        st.subheader("📊 Analisi e Risultati")
        st.info(f"💡 I valori grafici rappresentano l'incidenza emissiva giornaliera rapportata all'unità funzionale (diviso per **{dimensione_cantiere} {unita}**).")
        
        if df_filtrato.empty:
            st.warning("Nessun dato compreso nell'intervallo selezionato.")
        else:
            # Grafico Totale
            df_totale = df_filtrato.groupby('Data')['CO2_Normalizzata'].sum().reset_index()
            fig_tot = px.bar(
                df_totale, x='Data', y='CO2_Normalizzata',
                title=f"Andamento Emissioni Totali (kg CO₂e / {unita})",
                labels={'CO2_Normalizzata': f'kg CO₂e / {unita}', 'Data': 'Data'},
                text_auto='.2f'
            )
            fig_tot.update_traces(marker_color='#047857') # Verde professionale in stile EcoSite Tracker
            fig_tot.update_layout(height=400, font_family="Helvetica", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_tot, use_container_width=True)

            # Grafici Separati
            st.markdown("---")
            st.subheader("Dettaglio per Categoria d'Impatto")
            colori_parametri = {'Materiali': '#0284c7', 'Rifiuti': '#d97706', 'Trasporti e Macchinari': '#047857', 'Energia': '#7c3aed'}
            parametri_presenti = df_filtrato['Parametro'].dropna().unique()
            col_grafici = st.columns(2)
            
            for idx, param in enumerate(parametri_presenti):
                df_p = df_filtrato[df_filtrato['Parametro'] == param].groupby('Data')['CO2_Normalizzata'].sum().reset_index()
                fig = px.bar(
                    df_p, x='Data', y='CO2_Normalizzata',
                    title=f"{param}",
                    labels={'CO2_Normalizzata': f'kg CO₂e / {unita}', 'Data': ''},
                    text_auto='.2f'
                )
                fig.update_traces(marker_color=colori_parametri.get(param, '#374151'))
                fig.update_layout(height=320, font_family="Helvetica", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                col_grafici[idx % 2].plotly_chart(fig, use_container_width=True)

            # Tabella Dati
            with st.expander("Visualizza tabella dati dettagliata"):
                st.dataframe(df_filtrato.drop(columns=['Data_dt']), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    except Exception as e:
        st.error(foor := f"Errore nell'elaborazione del file: {e}")
else:
    st.info("In attesa del caricamento file CSV per avviare il calcolo...")