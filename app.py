from google import genai
from google.genai import types
import pandas as pd
import io
import streamlit as st

# Inizializzazione del client Gemini
client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY"))

st.markdown("<div class='minimal-card'>", unsafe_allow_html=True)
st.subheader("Importazione Intelligente dei Documenti di Progetto")
st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Carica i tuoi elaborati. Puoi usare PDF, Excel, CSV o TXT. L'IA li normalizzerà automaticamente.</p>", unsafe_allow_html=True)

# AGGIORNATO: Ora accetta pdf, txt, xlsx e csv in tutti i campi per massima flessibilità
file_computo = st.file_uploader("Computo Metrico (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'])
file_cronoprogramma = st.file_uploader("Cronoprogramma / Gantt (PDF, TXT, Excel o CSV)", type=['pdf', 'txt', 'xlsx', 'csv'])
file_trasporti = st.file_uploader("Note distanze trasporti (TXT o PDF)", type=['txt', 'pdf', 'xlsx', 'csv'])

if st.button("Elabora e Normalizza con IA"):
    if not file_computo:
        st.warning("Carica almeno il file del computo metrico per procedere.")
    else:
        with st.spinner("L'intelligenza artificiale sta analizzando gli elaborati..."):
            
            contents = []
            
            # Prepariamo i file assegnando il corretto formato (MIME type) per farli leggere a Gemini
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
                    
                    # Converte il file letto da Streamlit in un formato comprensibile all'IA
                    contents.append(
                        types.Part.from_bytes(data=file_obj.getvalue(), mime_type=mime)
                    )
            
            prompt_sistema = """
            Sei un esperto ingegnere edile e analista LCA. 
            Il tuo compito è analizzare i documenti di progetto forniti in input (che possono includere computi metrici, cronoprogrammi e distanze di trasporto) e generarne un'unica tabella CSV pulita.
            
            Il CSV finale DEVE avere esattamente queste 4 intestazioni di colonna, senza eccezioni:
            Data,Parametro,Elemento,Quantita
            
            Regole ferree:
            1. 'Data': Formato AAAA-MM-GG. Se c'è un cronoprogramma, distribuisci le quantità nelle date corrette delle fasi. Altrimenti usa la data di inizio progetto.
            2. 'Parametro': Deve appartenere tassativamente a una di queste categorie: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
            3. 'Elemento': Il nome dell'elemento o del combustibile coerente con i fattori di emissione.
            4. 'Quantita': Il valore numerico (per i trasporti, calcola la massa in tonnellate e moltiplicala per i chilometri indicati nelle note).
            
            Restituisci ESCLUSIVAMENTE il codice CSV grezzo, senza blocchi di codice Markdown, pronto per pd.read_csv().
            """
            contents.append(prompt_sistema)
            
            try:
                # Gemini 1.5 Flash (o 2.5 Flash) è perfetto per elaborare velocemente PDF e testi lunghi
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=contents
                )
                
                csv_testo = response.text.strip()
                if csv_testo.startswith("```"):
                    csv_testo = csv_testo.split("```")[1]
                    if csv_testo.startswith("csv"):
                        csv_testo = csv_testo[3:].strip()
                
                df_cantiere = pd.read_csv(io.StringIO(csv_testo))
                st.session_state['df_cantiere_ia'] = df_cantiere
                st.success("Documenti elaborati e normalizzati con successo dall'IA!")
                
            except Exception as e:
                st.error(f"Errore durante l'elaborazione con l'IA: {e}")

st.markdown("</div>", unsafe_allow_html=True)
