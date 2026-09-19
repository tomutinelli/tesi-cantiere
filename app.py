import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import io
from google import genai
from google.genai import types

# Impostazioni della pagina
st.set_page_config(page_title="EcoSite Tracker | LCA Dashboard", layout="wide")

# --- STILE CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1f2937;
        background-color: #faf9f6;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Card minimaliste con bordo verde pastello */
    div.minimal-card {
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        border-radius: 8px !important;
        padding: 32px !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.01) !important;
        margin-bottom: 24px !important;
    }

    h1 {
        font-weight: 600;
        letter-spacing: -0.025em;
        color: #111827;
        font-size: 2.25rem;
    }
    
    h2, h3 {
        font-weight: 500;
        letter-spacing: -0.01em;
        color: #1f2937;
    }

    .stRadio label, .stNumberInput label, .stDateInput label, .stFileUploader label, .stSelectbox label {
        font-weight: 500;
        font-size: 0.9rem;
        color: #4b5563;
    }

    /* Stile per i Bottoni Nativi (Download e Link) in Verde Pastello */
    [data-testid="stDownloadButton"] button, 
    [data-testid="stLinkButton"] button,
    [data-testid="baseButton-secondary"] {
        background-color: #ffffff !important;
        border: 1.5px solid #d5ddd1 !important;
        color: #374151 !important;
        border-radius: 8px !important;
        transition: all 0.25s ease-in-out !important;
        font-weight: 500 !important;
        width: 100% !important;
    }
    
    [data-testid="stDownloadButton"] button:hover, 
    [data-testid="stLinkButton"] button:hover,
    [data-testid="baseButton-secondary"]:hover {
        border-color: #a7b89f !important;
        background-color: #f4f7f3 !important;
        color: #111827 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- DIZIONARIO COLORI GLOBALE ---
colori_parametri = {
    'Materiali': "#B80D0D", 
    'Rifiuti': "#078303", 
    'Trasporti': "#732BB7", 
    'Energia': "#f6de03",
    'Acqua': "#53DCFE", 
    'Macchinari': "#8a929e"
}

# --- GUIDA METODOLOGICA ---
with st.expander("Note metodologiche e specifiche di utilizzo"):
    st.markdown("""
    <p style='color: #4b5563; font-size: 0.95rem; line-height: 1.6; margin-bottom: 0;'>
    Questo strumento calcola l'impronta di carbonio (kg di CO₂e) in fase di progetto incrociando i dati di consumo giornaliero con il database LCI di riferimento. 
    </p>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

# --- BLOCCO CONFIGURAZIONE ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Configurazione dell'Unità Funzionale")
st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Definisci i parametri di normalizzazione globale per l'opera.</p>", unsafe_allow_html=True)

tipo_cantiere = st.radio(
    "Tipologia di modellazione spaziale",
    options=[
        "Cantiere Lineare (normalizzazione per metro lineare - m)", 
        "Cantiere Standard (normalizzazione per metro quadro - m²)"
    ],
    label_visibility="collapsed"
)

st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
col_cfg1, col_cfg2 = st.columns([2, 1])
with col_cfg1:
    if "Lineare" in tipo_cantiere:
        unita = "m"
        dimensione_cantiere = st.number_input("Estensione longitudinale complessiva (metri)", min_value=0.1, value=100.0, step=1.0)
    else:
        unita = "m²"
        dimensione_cantiere = st.number_input("Superficie coperta complessiva (metri quadri)", min_value=0.1, value=100.0, step=1.0)
st.markdown("</div>", unsafe_allow_html=True)

# --- LETTURA DATABASE LCI ---
@st.cache_data
def carica_database_lci(percorso_file):
    if not os.path.exists(percorso_file):
        return None
    try:
        xls = pd.ExcelFile(percorso_file)
    except Exception:
        return None
        
    df_inv = pd.DataFrame()
    mappatura = {
        'Materiali': {'nome_elemento': 'nome_materiale', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Rifiuti': {'nome_elemento': 'tipo_rifiuto', 'nome_fattore': 'emissioni_kg_co2eq_kg'},
        'Energia': {'nome_elemento': 'Tecnologia di generazione elettrica', 'nome_fattore': 'Fattori di emissione (kg CO2eq/kWh)'},
        'Acqua': {'nome_elemento': 'Elemento', 'nome_fattore': 'Fattore_Emissione'},
        'Trasporti': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'},
        'Macchinari': {'nome_elemento': 'Fuel', 'nome_fattore': 'kg CO2 per kg of fuel'}
    }
    
    for foglio in xls.sheet_names:
        df_temp = pd.read_excel(xls, sheet_name=foglio)
        col_el = None
        col_fat = None
        
        if foglio in mappatura:
            col_el = mappatura[foglio]['nome_elemento']
            col_fat = mappatura[foglio]['nome_fattore']
            if foglio in ['Trasporti', 'Macchinari']:
                df_temp = df_temp.iloc[0:9].copy()
        else:
            possibili_nomi_elemento = ['Elemento', 'Macchinario', 'Nome', 'Acqua', 'Tipo', 'Fuel']
            possibili_nomi_fattore = ['Fattore_Emissione', 'kg CO2eq', 'Emissione', 'emissioni_kg_co2eq_kg', 'kg co2 per kg of fuel']
            
            for c in df_temp.columns:
                if any(x.lower() in str(c).lower() for x in possibili_nomi_elemento) and col_el is None:
                    col_el = c
                if any(x.lower() in str(c).lower() for x in possibili_nomi_fattore) and col_fat is None:
                    col_fat = c
        
        if col_el in df_temp.columns and col_fat in df_temp.columns:
            df_temp = df_temp[[col_el, col_fat]].copy()
            df_temp.rename(columns={col_el: 'Elemento', col_fat: 'Fattore_Emissione'}, inplace=True)
            df_temp['Parametro'] = foglio 
            df_temp['Fattore_Emissione'] = pd.to_numeric(df_temp['Fattore_Emissione'], errors='coerce')
            df_inv = pd.concat([df_inv, df_temp], ignore_index=True)
                
    return df_inv.dropna(subset=['Elemento', 'Fattore_Emissione'])

percorso_lci = "LCI.xlsx"
df_inventario = carica_database_lci(percorso_lci)

if df_inventario is None or df_inventario.empty:
    st.error(f"Errore critico: Il database '{percorso_lci}' non è reperibile o non è valido sul server.")
    st.stop()

# --- BLOCCO INPUT DATI (IA vs MANUALE) ---
st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Caricamento Dataset di Progetto")

tab1, tab2 = st.tabs(["✨ Elaborazione Intelligente (IA)", "📂 Caricamento CSV Manuale"])

with tab1:
    st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Carica i tuoi elaborati (PDF, TXT, Excel). L'IA estrarrà i dati, mapperà i materiali e calcolerà i trasporti automaticamente.</p>", unsafe_allow_html=True)
    
    file_computo = st.file_uploader("Computo Metrico (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'], key="ia_comp")
    file_cronoprogramma = st.file_uploader("Cronoprogramma / Gantt (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'], key="ia_crono")
    file_trasporti = st.file_uploader("Note distanze trasporti (TXT o PDF)", type=['txt', 'pdf', 'xlsx', 'csv'], key="ia_trasp")

    if st.button("Elabora e Normalizza con IA", key="btn_ia"):
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("Chiave API mancante nei Secrets. Inserisci GEMINI_API_KEY per utilizzare questa funzione.")
        elif not file_computo:
            st.warning("Carica almeno il file del computo metrico per procedere.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.spinner("L'intelligenza artificiale sta analizzando gli elaborati..."):
                    contents = []
                    
                    for file_obj in [file_computo, file_cronoprogramma, file_trasporti]:
                        if file_obj is not None:
                            estensione = file_obj.name.split('.')[-1].lower()
                            if estensione == 'pdf':
                                mime = 'application/pdf'
                            elif estensione == 'xlsx':
                                mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            elif estensione == 'csv':
                                mime = 'text/csv'
                            else:
                                mime = 'text/plain'
                            
                            contents.append(
                                types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime)
                            )
                    
                    # --- FIX PROMPT CON VOCI DEL DATABASE ---
                    elementi_validi = df_inventario['Elemento'].unique().tolist()
                    lista_elementi_str = ", ".join([str(e) for e in elementi_validi])
                    
                    prompt_sistema = f"""
                    Sei un esperto ingegnere edile e analista LCA. 
                    Il tuo compito è analizzare i documenti di progetto forniti in input (computi, cronoprogrammi, note trasporti) e generarne un'unica tabella CSV pulita.
                    
                    Il CSV finale DEVE avere esattamente queste 4 intestazioni di colonna:
                    Data,Parametro,Elemento,Quantita
                    
                    Regole ferree:
                    1. 'Data': Formato AAAA-MM-GG. Se c'è un cronoprogramma, distribuisci le quantità nelle date corrette.
                    2. 'Parametro': Scegli tra: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
                    3. 'Elemento': DEVI TASSATIVAMENTE usare SOLO uno dei seguenti nomi esatti presenti nel database: 
                    [{lista_elementi_str}]. 
                    Non inventare nomi nuovi. Leggi la descrizione nel computo e scegli il nome dalla lista qui sopra che corrisponde meglio semanticamente.
                    4. 'Quantita': Valore numerico (per trasporti: massa in tonnellate x km).
                    
                    Restituisci ESCLUSIVAMENTE il codice CSV grezzo, pronto per pd.read_csv().
                    """
                    contents.append(prompt_sistema)
                    
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=contents
                    )
                    
                    csv_testo = response.text.strip()
                    if csv_testo.startswith("```"):
                        csv_testo = csv_testo.split("```")[1]
                        if csv_testo.startswith("csv"):
                            csv_testo = csv_testo[3:].strip()
                    
                    df_cantiere = pd.read_csv(io.StringIO(csv_testo))
                    st.session_state['df_cantiere'] = df_cantiere
                    st.success("Documenti elaborati e normalizzati con successo dall'IA!")
            except Exception as e:
                st.error(f"Errore durante l'elaborazione con l'IA: {e}")

with tab2:
    st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Importa il file in formato CSV contenente la serie temporale dei consumi già formattata (Data, Parametro, Elemento, Quantita).</p>", unsafe_allow_html=True)
    file_csv_manuale = st.file_uploader("Seleziona file CSV", type=['csv'], label_visibility="collapsed", key="csv_manuale")
    if file_csv_manuale:
        try:
            df_cantiere = pd.read_csv(file_csv_manuale)
            st.session_state['df_cantiere'] = df_cantiere
            st.success("File CSV caricato correttamente!")
        except Exception as e:
            st.error(f"Errore nella lettura del file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# =====================================================================
# ELABORAZIONE E GRAFICI (Si attiva solo se i dati sono caricati)
# =====================================================================
if 'df_cantiere' in st.session_state:
    df_cantiere = st.session_state['df_cantiere']
    
    # --- VISUALIZZAZIONE ANTEPRIMA DATI ---
    with st.expander("👀 Visualizza Anteprima Dati Input (Utile per verifica IA)"):
        st.dataframe(df_cantiere, use_container_width=True)
        csv_input = df_cantiere.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Scarica Dati Input Normalizzati (CSV)",
            data=csv_input,
            file_name="dati_input_cantiere.csv",
            mime="text/csv",
            key="dl_input_csv"
        )
    
    try:
        # Validazione Colonne
        for col in ['Data', 'Parametro', 'Elemento', 'Quantita']:
            if col not in df_cantiere.columns:
                st.error(f"Errore di struttura: La colonna '{col}' risulta assente.")
                st.stop()
        
        df_cantiere['Data_dt'] = pd.to_datetime(df_cantiere['Data'], format='%Y-%m-%d', errors='coerce')
        df_completo = pd.merge(df_cantiere, df_inventario, on=['Parametro', 'Elemento'], how='left')
        
        mancanti = df_completo[df_completo['Fattore_Emissione'].isnull()]
        if not mancanti.empty:
            st.warning(f"Attenzione: Elementi non riconosciuti nel database LCI e calcolati a zero: {mancanti['Elemento'].unique().tolist()}")
            df_completo['Fattore_Emissione'] = df_completo['Fattore_Emissione'].fillna(0)
            
        df_completo['CO2_Totale_kg'] = df_completo['Quantita'] * df_completo['Fattore_Emissione']
        df_completo['CO2_Normalizzata'] = df_completo['CO2_Totale_kg'] / dimensione_cantiere
        
        # --- FILTRO TEMPORALE ---
        st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
        st.subheader("Filtro Temporale")
        st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 15px;'>Seleziona l'arco temporale di interesse per l'analisi.</p>", unsafe_allow_html=True)
        
        min_date = df_completo['Data_dt'].min().date()
        max_date = df_completo['Data_dt'].max().date()

        date_range = st.date_input(
            "Intervallo temporale",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            label_visibility="collapsed"
        )

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df_completo['Data_dt'].dt.date >= start_date) & (df_completo['Data_dt'].dt.date <= end_date)
            df_filtrato = df_completo.loc[mask].copy()
        elif isinstance(date_range, date):
            mask = (df_completo['Data_dt'].dt.date == date_range)
            df_filtrato = df_completo.loc[mask].copy()
        else:
            df_filtrato = df_completo.copy()
            
        st.markdown("</div>", unsafe_allow_html=True)

        # --- VISUALIZZAZIONE GRAFICI ---
        st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
        
        # LINK WIX FULLSCREEN
        col_title, col_link = st.columns([2, 1])
        with col_title:
            st.subheader("Risultati Analitici")
        with col_link:
            # Ricorda di inserire qui il tuo link reale di Streamlit per l'esportazione su Wix
            link_app = "https://INSERISCI-QUI-IL-TUO-LINK.streamlit.app" 
            st.link_button("Apri a schermo intero per esportare", link_app)
            
        st.markdown(f"<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 25px;'>I valori visualizzati esprimono l'incidenza normalizzata rispetto all'unità funzionale complessiva (<b>{dimensione_cantiere} {unita}</b>).</p>", unsafe_allow_html=True)
        
        if df_filtrato.empty:
            st.warning("Nessuna evidenza registrata nell'intervallo temporale selezionato.")
        else:
            modo_visualizzazione = st.selectbox(
                "Modalità di visualizzazione grafica",
                options=[
                    "Standard (Barre temporali)", 
                    "Linea Media (Andamento con media mobile)",
                    "Linea Media Fissa (Orizzontale)"
                ],
                index=0
            )

            def genera_figura(df_dat, x_col, y_col, titolo, colore_base):
                df_dat = df_dat.copy()
                if df_dat.empty:
                    return go.Figure()
                    
                if "Media mobile" in modo_visualizzazione or "Media (Andamento" in modo_visualizzazione:
                    df_dat = df_dat.sort_values(by=x_col)
                    df_dat['Media_Mobile'] = df_dat[y_col].rolling(window=7, min_periods=1).mean()
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Valore Giornaliero', marker_color=colore_base, opacity=0.4))
                    fig.add_trace(go.Scatter(x=df_dat[x_col], y=df_dat['Media_Mobile'], mode='lines', name='Linea Media (7gg)', line=dict(color='#111827', width=3)))
                    fig.update_layout(title=f"{titolo} - Andamento con Media Mobile")
                elif "Fissa" in modo_visualizzazione:
                    media_val = df_dat[y_col].mean()
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=df_dat[x_col], y=df_dat[y_col], name='Valore Giornaliero', marker_color=colore_base, opacity=0.5))
                    fig.add_hline(
                        y=media_val, line_dash="dash", line_color="#27ae60", 
                        annotation_text=f"Media Fissa: {media_val:.2f}", annotation_position="top right"
                    )
                    fig.update_layout(title=f"{titolo} - Con Linea Media Fissa nell'Intervallo")
                else:
                    fig = px.bar(
                        df_dat, x=x_col, y=y_col,
                        title=titolo, labels={y_col: f'kg CO₂e / {unita}', x_col: ''}, text_auto='.2f'
                    )
                    fig.update_traces(marker_color=colore_base)
                
                fig.update_layout(
                    height=380, font_family="Inter", 
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    title_font_size=14, title_font_color="#374151"
                )
                return fig

            # 1. Grafico Totale
            df_totale = df_filtrato.groupby('Data')['CO2_Normalizzata'].sum().reset_index()
            fig_tot = genera_figura(df_totale, 'Data', 'CO2_Normalizzata', f"Andamento Complessivo (kg CO₂e / {unita})", "#0B0752")
            st.plotly_chart(fig_tot, use_container_width=True)
            
            # 2. Diagramma a Torta (Incidenza Percentuale)
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            df_pie = df_filtrato.groupby('Parametro')['CO2_Normalizzata'].sum().reset_index()
            
            fig_pie = px.pie(
                df_pie, 
                values='CO2_Normalizzata', 
                names='Parametro',
                title="Incidenza Percentuale delle Categorie sulle Emissioni Totali",
                color='Parametro',
                color_discrete_map=colori_parametri,
                hole=0.4 # Grafico ad anello (Donut chart)
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
            fig_pie.update_layout(
                height=450, 
                font_family="Inter", 
                showlegend=True, 
                title_font_size=14, 
                title_font_color="#374151",
                plot_bgcolor='rgba(0,0,0,0)', 
                paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_pie, use_container_width=True)

            # Pulsanti Download Dati Totali
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                csv_totale = df_totale.to_csv(index=False).encode('utf-8')
                st.download_button("Scarica dati totali (CSV)", data=csv_totale, file_name="emissioni_totali.csv", mime="text/csv", key="dl_tot_csv")
            with col_exp2:
                output_xlsx = io.BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                    df_totale.to_excel(writer, index=False, sheet_name='Totale Giornaliero')
                st.download_button("Scarica dati totali (XLSX)", data=output_xlsx.getvalue(), file_name="emissioni_totali.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_tot_xlsx")

            st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
            st.subheader("Disaggregazione per Categoria")
            
            parametri_presenti = df_filtrato['Parametro'].dropna().unique()
            col_grafici = st.columns(2)
            
            for idx, param in enumerate(parametri_presenti):
                df_p = df_filtrato[df_filtrato['Parametro'] == param].groupby('Data')['CO2_Normalizzata'].sum().reset_index()
                colore_cat = colori_parametri.get(param, '#4b5563')
                
                fig_cat = genera_figura(df_p, 'Data', 'CO2_Normalizzata', f"{param}", colore_cat)
                
                with col_grafici[idx % 2]:
                    st.plotly_chart(fig_cat, use_container_width=True)
                    csv_cat = df_p.to_csv(index=False).encode('utf-8')
                    st.download_button(f"Scarica dati {param} (CSV)", data=csv_cat, file_name=f"dati_{param.lower()}.csv", mime="text/csv", key=f"dl_{param}")

            # Tabella Dati Globale
            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
            with st.expander("Esporta / Visualizza matrice dati completa calcolata"):
                st.dataframe(df_filtrato.drop(columns=['Data_dt']), use_container_width=True)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_filtrato.drop(columns=['Data_dt']).to_excel(writer, index=False, sheet_name='Dettaglio Completo')
                st.download_button("Scarica dataset completo calcolato (XLSX)", data=excel_buffer.getvalue(), file_name="dataset_completo_lca.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_full_xlsx")

        st.markdown("</div>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Si è verificato un errore durante l'elaborazione dei dati LCA: {e}")
