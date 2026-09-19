from google import genai
import pandas as pd
import io
import streamlit as st

# Inizializzazione del client Gemini (legge la chiave dai secrets di Streamlit)
# Nel file .streamlit/secrets.toml salverai: GEMINI_API_KEY="tua_chiave"
client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY"))

st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Importazione Intelligente dei Documenti di Progetto")
st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Carica i tuoi elaborati (Computo, Cronoprogramma e note Trasporti). L'IA li normalizzerà automaticamente.</p>", unsafe_allow_html=True)

# Upload multipli per i diversi elaborati
file_computo = st.file_uploader("Computo Metrico (Excel o CSV)", type=['xlsx', 'csv'])
file_cronoprogramma = st.file_uploader("Cronoprogramma / Gantt (Excel, CSV o TXT)", type=['xlsx', 'csv', 'txt'])
file_trasporti = st.file_uploader("Note distanze trasporti (TXT o PDF/Excel)", type=['txt', 'xlsx', 'csv'])

# Pulsante per avviare l'elaborazione dell'IA
if st.button("Elabora e Normalizza con IA"):
    if not file_computo:
        st.warning("Carica almeno il file del computo metrico per procedere.")
    else:
        with st.spinner("L'intelligenza artificiale sta analizzando gli elaborati e strutturando i dati..."):
            
            # Prepariamo i file da inviare a Gemini
            contents = []
            
            # Leggiamo i file caricati e li passiamo al modello
            for file_obj in [file_computo, file_cronoprogramma, file_trasporti]:
                if file_obj is not None:
                    contents.append(file_obj.getvalue())
            
            prompt_sistema = """
            Sei un esperto ingegnere edile e analista LCA. 
            Il tuo compito è analizzare i documenti di progetto forniti in input (che possono includere computi metrici, cronoprogrammi e distanze di trasporto) e generarne un'unica tabella CSV pulita.
            
            Il CSV finale DEVE avere esattamente queste 4 intestazioni di colonna, senza eccezioni:
            Data,Parametro,Elemento,Quantita
            
            Regole ferree:
            1. 'Data': Formato AAAA-MM-GG. Se c'è un cronoprogramma, distribuisci le quantità nelle date corrette delle fasi. Altrimenti usa la data di inizio progetto.
            2. 'Parametro': Deve appartenere tassativamente a una di queste categorie del database LCI: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
            3. 'Elemento': Il nome dell'elemento o del combustibile coerente con i fattori di emissione.
            4. 'Quantita': Il valore numerico (per i trasporti, calcola la massa in tonnellate moltiplicata per i chilometri indicati nelle note, oppure esprimi l'unità corretta in base alla categoria).
            
            Restituisci ESCLUSIVAMENTE il codice CSV grezzo, senza blocchi di codice Markdown (niente ```csv ... ```), pronto per essere letto direttamente da Pandas con pd.read_csv().
            """
            
            contents.append(prompt_sistema)
            
            try:
                # Chiamata al modello Gemini (usiamo flash per velocità ed efficienza sui dati testuali/tabulari)
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=contents
                )
                
                csv_testo = response.text.strip()
                # Pulizia di sicurezza da eventuali blocchi markdown se l'IA li inserisce comunque
                if csv_testo.startswith("```"):
                    csv_testo = csv_testo.split("```")[1]
                    if csv_testo.startswith("csv"):
                        csv_testo = csv_testo[3:].strip()
                
                # Convertiamo il testo generato dall'IA in un DataFrame utilizzabile dal resto dell'app
                df_cantiere = pd.read_csv(io.StringIO(csv_testo))
                
                # Salviamo il dataframe nella sessione di Streamlit così resta in memoria
                st.session_state['df_cantiere_ia'] = df_cantiere
                st.success("Documenti elaborati e normalizzati con successo dall'IA!")
                
            except Exception as e:
                st.error(f"Errore durante l'elaborazione con l'IA: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# Se l'IA ha generato il dataframe, lo passiamo al tuo flusso esistente
if 'df_cantiere_ia' in st.session_state:
    df_cantiere = st.session_state['df_cantiere_ia']
    
    # DA QUI IN POI IL RESTO DEL TUO CODICE RIMANE IDENTICO:
    # merge con LCI, filtri temporali, grafici Plotly e pulsanti di esportazione!
