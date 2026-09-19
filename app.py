with tab1:
    st.markdown("""
    <div style='background-color: #f4f7f3; border: 1.5px solid #d5ddd1; border-radius: 8px; padding: 20px; margin-bottom: 20px;'>
        <h4 style='color: #111827; margin-top: 0; font-size: 1.1rem; font-weight: 600;'>📖 Guida Operativa: Procedura per l'Elaborazione con IA</h4>
        <ol style='color: #4b5563; font-size: 0.9rem; line-height: 1.6; margin-bottom: 0; padding-left: 20px;'>
            <li><b>Prepara la documentazione di cantiere:</b> Raccogli i file relativi al progetto (es. il computo metrico estimativo in PDF o Excel, il cronoprogramma o diagramma di Gantt delle lavorazioni, e le note relative alle distanze dei trasporti).</li>
            <li><b>Carica i file nei campi sottostanti:</b> Trascina o seleziona i documenti nei rispettivi riquadri dedicati. Il computo metrico è obbligatorio per estrarre le quantità principali.</li>
            <li><b>Avvia l'analisi semantica:</b> Clicca sul pulsante <i>"Elabora e Normalizza con IA"</i>. L'intelligenza artificiale leggerà i file, estrarrà le date, i parametri e le quantità grezze.</li>
            <li><b>Mappatura automatica LCI:</b> Il motore Python interno intercetterà le descrizioni estratte dall'IA e le assocerà istantaneamente e in modo blindato alle voci ufficiali presenti nel file <code>LCI.xlsx</code>.</li>
            <li><b>Visualizza i risultati:</b> Subito dopo l'elaborazione, potrai controllare l'anteprima della tabella normalizzata, scaricare il CSV pulito e visualizzare i grafici analitici aggiornati.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<p style='color: #6b7280; font-size: 0.9rem; margin-bottom: 20px;'>Carica i tuoi elaborati (PDF, TXT, Excel o CSV). L'IA estrarrà i dati e il motore Python li mapperà automaticamente sul database LCI.</p>", unsafe_allow_html=True)
    
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
                with st.spinner("L'intelligenza artificiale sta analizzando i documenti..."):
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
                    
                    prompt_sistema = \"\"\"
                    Sei un ingegnere edile e analista LCA. Estrai i dati dai documenti forniti (computi, cronoprogrammi, note trasporti) e restituisci una tabella CSV pulita con queste esatte 4 intestazioni di colonna:
                    Data,Parametro,Elemento,Quantita
                    
                    - Data: Formato AAAA-MM-GG.
                    - Parametro: Scegli tassativamente tra: Materiali, Rifiuti, Energia, Acqua, Trasporti, Macchinari.
                    - Elemento: Riporta la descrizione dell'elemento o materiale trovata nel computo.
                    - Quantita: Valore numerico (per i trasporti, calcola la massa in tonnellate moltiplicata per i chilometri).
                    
                    Restituisci ESCLUSIVAMENTE il codice CSV grezzo, senza blocchi Markdown, pronto per pd.read_csv().
                    \"\"\"
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
                        elif csv_testo.startswith("\n"):
                            csv_testo = csv_testo.strip()
                    
                    df_cantiere_grezzo = pd.read_csv(io.StringIO(csv_testo))
                    
                    # Correttore automatico intestazioni
                    colonne_mappa = {}
                    for c in df_cantiere_grezzo.columns:
                        c_low = str(c).strip().lower()
                        if 'data' in c_low or 'date' in c_low or 'giorno' in c_low:
                            colonne_mappa[c] = 'Data'
                        elif 'param' in c_low or 'categ' in c_low:
                            colonne_mappa[c] = 'Parametro'
                        elif 'elem' in c_low or 'material' in c_low or 'voce' in c_low:
                            colonne_mappa[c] = 'Elemento'
                        elif 'quant' in c_low or 'val' in c_low or 'qt' in c_low or 'amount' in c_low:
                            colonne_mappa[c] = 'Quantita'
                    
                    df_cantiere_grezzo.rename(columns=colonne_mappa, inplace=True)
                    
                    if 'Parametro' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Parametro'] = 'Materiali'
                    if 'Data' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Data'] = str(date.today())
                    if 'Elemento' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Elemento'] = 'Calcestruzzo'
                    if 'Quantita' not in df_cantiere_grezzo.columns:
                        df_cantiere_grezzo['Quantita'] = 1.0

                    # Motore di traduzione semantica Python verso il database LCI
                    def mappa_voce_a_lci(parametro, elemento_grezzo):
                        p_str = str(parametro).strip()
                        e_str = str(elemento_grezzo).strip().lower()
                        
                        voci_disponibili = df_inventario[df_inventario['Parametro'].str.lower() == p_str.lower()]['Elemento'].tolist()
                        if not voci_disponibili:
                            voci_disponibili = df_inventario['Elemento'].tolist()
                            
                        for v in voci_disponibili:
                            if v.lower() in e_str or e_str in v.lower():
                                return v
                                
                        if p_str.lower() == 'materiali':
                            if 'calcestruzzo' in e_str or 'cls' in e_str: return 'Calcestruzzo'
                            if 'acciaio' in e_str or 'ferro' in e_str: return 'Acciaio'
                            if 'laterizio' in e_str or 'mattone' in e_str: return 'Laterizio'
                            if 'inerti' in e_str or 'sabbia' in e_str or 'ghiaia' in e_str: return 'Inerti'
                            if 'asfalto' in e_str or 'bitume' in e_str: return 'Asfalto/Bitume'
                            if 'legno' in e_str: return 'Legno'
                            if 'vetro' in e_str: return 'Vetro'
                            if 'isolante' in e_str or 'lana' in e_str or 'eps' in e_str: return 'Isolante EPS'
                        elif p_str.lower() == 'rifiuti':
                            if 'scavo' in e_str or 'terra' in e_str: return 'Inerti / Macerie di demolizione'
                            if 'calcestruzzo' in e_str: return 'Calcestruzzo di risulta'
                            if 'acciaio' in e_str or 'ferro' in e_str: return 'Metallo / Acciaio di scarto'
                            if 'legno' in e_str: return 'Legno da cantiere'
                        elif p_str.lower() in ['trasporti', 'macchinari']:
                            if 'diesel' in e_str or 'gasolio' in e_str: return 'Diesel'
                            if 'benzina' in e_str or 'petrol' in e_str: return 'Petrol'
                            
                        match = get_close_matches(str(elemento_grezzo), voci_disponibili, n=1, cutoff=0.1)
                        if match:
                            return match[0]
                            
                        return elemento_grezzo

                    df_cantiere_grezzo['Elemento'] = df_cantiere_grezzo.apply(
                        lambda row: mappa_voce_a_lci(row['Parametro'], row['Elemento']), axis=1
                    )
                    
                    st.session_state['df_cantiere'] = df_cantiere_grezzo
                    st.success("Documenti elaborati e normalizzati con successo dall'IA!")
            except Exception as e:
                st.error(f"Errore durante l'elaborazione con l'IA: {e}")
