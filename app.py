import os
import time
from datetime import datetime, timedelta
import streamlit as st
import pandas as pd
import qrcode
import cv2
import numpy as np
from PIL import Image
import textwrap
import json
import io
from fpdf import FPDF
import matplotlib.pyplot as plt
import base64
import numpy as np
import plotly.express as px

# Import per l'estrazione del testo dai file locali
from pypdf import PdfReader
from docx import Document


# ==================================================================
# --- 1. CONFIGURAZIONI INIZIALI E DATABASE LOCALI ---
# ==================================================================
st.set_page_config(
    page_title="Piattaforma Centralizzata HSE System", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

FILE_NEAR_MISS = "near_miss.csv"
FILE_ANALISI_NM = "analisi_near_miss.csv"
FILE_SCADENZARIO = "scadenzario.xlsx"
DIR_CONFORMITA = "documenti_conformita"
DIR_IMMAGINI_ANALISI = "Immagini_Analisi_Near_Miss"
DIR_ANALISI = os.path.dirname(os.path.abspath(__file__))

if not os.path.exists(DIR_CONFORMITA):
    os.makedirs(DIR_CONFORMITA)

if not os.path.exists(DIR_IMMAGINI_ANALISI):
    os.makedirs(DIR_IMMAGINI_ANALISI)

COLONNE_SCADENZARIO = [
    "Adempimento", 
    "Regolamento", 
    "Cosa produrre", 
    "Rilasciato", 
    "In vigore durante (mesi)", 
    "Scadenza", 
    "Responsabile", 
    "Note"
]
# Salva normativa di conformità legislativa - Aggiungi normativa
def salva_normativa(chiave, titolo, url, descrizione):
    db = {}
    if os.path.exists("normative.json"):
        try:
            with open("normative.json", "r") as f:
                db = json.load(f)
        except Exception:
            db = {}
    
    db[chiave.strip()] = {
        "titolo": titolo.strip(),
        "url": url.strip(),
        "desc": descrizione.strip()
    }
    
    with open("normative.json", "w") as f:
        json.dump(db, f, indent=4)


# ==================================================================
# --- 2. FUNZIONI DI SUPPORTO E CALCOLO AUTOMATICO SCADENZE ---
# ==================================================================
def estrai_testo_da_file(percorso_file):
    estensione = os.path.splitext(percorso_file)[1].lower()
    testo = ""
    try:
        if estensione == ".pdf":
            reader = PdfReader(percorso_file)
            for page in reader.pages:
                t = page.extract_text()
                if t: testo += t + "\n"
        elif estensione in [".docx", ".doc"]:
            doc = Document(percorso_file)
            for p in doc.paragraphs:
                testo += p.text + "\n"
        elif estensione in [".xlsx", ".xls", ".csv"]:
            df = pd.read_excel(percorso_file) if "xls" in estensione else pd.read_csv(percorso_file, sep=";")
            testo += df.to_string()
    except Exception as e:
        return f"[Errore lettura {os.path.basename(percorso_file)}: {e}]"
    return testo

def cerca_nei_documenti(query, cartella):
    file_presenti = [os.path.join(cartella, f) for f in os.listdir(cartella) if os.path.isfile(os.path.join(cartella, f))]
    if not file_presenti:
        return []
    
    parole_chiave = [w.lower() for w in query.split() if len(w) > 2]
    if not parole_chiave:
        return "specifica"

    riscontri = []
    for fp in file_presenti:
        contenuto = estrai_testo_da_file(fp)
        if contenuto.strip():
            paragrafi = [p.strip() for p in contenuto.split("\n") if len(p.strip()) > 15]
            for i, para in enumerate(paragrafi):
                p_lower = para.lower()
                conteggio = sum(1 for kw in parole_chiave if kw in p_lower)
                if conteggio > 0:
                    rilevanza = int((conteggio / len(parole_chiave)) * 100)
                    riscontri.append({
                        "nome_file": os.path.basename(fp),
                        "percorso_completo": fp,
                        "fonte": f"{os.path.basename(fp)} (Paragrafo {i+1})",
                        "testo": para,
                        "score": rilevanza
                    })

    if not riscontri:
        return []

    return sorted(riscontri, key=lambda x: x["score"], reverse=True)[:3]

def estrai_info_checklist(percorso_file):
    nome_base = os.path.basename(percorso_file)
    titolo_ricavato = os.path.splitext(nome_base)[0].replace("_", " ")
    testo = estrai_testo_da_file(percorso_file)
    linee = [l.strip() for l in testo.split("\n") if l.strip()]
    
    regolamento = "Non identificato (Consultare file)"
    documenti_produrre = "Vedere testo integrale"
    settore_aziendale = "Aziendale Generale"
    
    settori_mappa = {
        "produzione": "Produzione / Officina",
        "logistica": "Logistica / Magazzino",
        "uffici": "Uffici / Amministrazione",
        "manutenzione": "Manutenzione / Impianti",
        "cantiere": "Cantieri / Esterni",
        "laboratorio": "Laboratorio / Qualità",
        "commerciale": "Commerciale / Vendite"
    }
    
    testo_lower_completo = testo.lower()
    for chiave, nome_settore in settori_mappa.items():
        if chiave in testo_lower_completo:
            settore_aziendale = nome_settore
            break

    for linea in linee:
        l_low = linea.lower()
        if "d.lgs" in l_low or "decreto" in l_low or "legge" in l_low or "regolamento" in l_low:
            if len(linea) < 100:
                regolamento = linea
                break
                
    for linea in linee:
        l_low = linea.lower()
        if "allegato" in l_low or "produrre" in l_low or "documentazione" in l_low or "certificato" in l_low:
            if len(linea) < 120 and len(linea) > 20:
                documenti_produrre = linea
                break
                
    return titolo_ricavato, regolamento, documenti_produrre, settore_aziendale

def calcola_data_scadenza(row):
    scadenza_attuale = str(row.get("Scadenza", "")).strip()
    if scadenza_attuale and scadenza_attuale.lower() not in ["nan", "none", "", "<na>", "nat"]:
        return scadenza_attuale
        
    rilasciato = str(row.get("Rilasciato", "")).strip()
    mesi = str(row.get("In vigore durante (mesi)", "")).strip()
    
    if not rilasciato or rilasciato.lower() in ["nan", "none", ""] or not mesi or mesi.lower() in ["nan", "none", "", "0"]:
        return ""
        
    try:
        rilasciato_pulito = rilasciato.replace("-", "/").split()[0]
        data_rilascio = datetime.strptime(rilasciato_pulito, "%d/%m/%Y")
        
        mesi_da_aggiungere = int(float(mesi))
        nuovo_mese = data_rilascio.month + mesi_da_aggiungere
        nuovo_anno = data_rilascio.year + (nuovo_mese - 1) // 12
        nuovo_mese = (nuovo_mese - 1) % 12 + 1
        
        nuovo_giorno = min(data_rilascio.day, [31, 29 if nuovo_anno % 4 == 0 and (nuovo_anno % 100 != 0 or nuovo_anno % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][nuovo_mese - 1])
        data_scadenza = datetime(nuovo_anno, nuovo_mese, nuovo_giorno)
        return data_scadenza.strftime("%d/%m/%Y")
    except:
        return ""

def evidenzia_righe_scadenza(row):
    val = row.get("Scadenza", "")
    if pd.isna(val) or not val or str(val).strip() in ["", "nan", "None", "<NA>"]:
        return [''] * len(row)
    
    anno_corrente = datetime.now().year
    anno_successivo = anno_corrente + 1
    
    try:
        parti = str(val).strip().replace("-", "/").split("/")
        if len(parti) == 3:
            anno_scadenza = int(parti[-1].split()[0])
            if anno_scadenza == anno_corrente:
                return ['background-color: #ffcccc; color: #cc0000; font-weight: bold;'] * len(row)
            elif anno_scadenza == anno_successivo:
                return ['background-color: #fff2cc; color: #b68500; font-weight: bold;'] * len(row)
    except:
        pass
    return [''] * len(row)

# Funzione nativa per la generazione di QR Code (Ritorna un'immagine PIL)
def genera_qr_nativo(data):
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color='black', back_color='white')
def decodifica_qr_opencv(image_file):
    # Converti l'immagine in un formato leggibile da OpenCV
    img = Image.open(image_file).convert('RGB')
    img_array = np.array(img)

    #Decodifica tramite OpenCV
    detector = cv2.QRCodeDetector()
    valore, punti, qr_code = detector.detectAndDecode(img_array)

    return valore if valore else None

# ==================================================================
# --- 3. INTESTAZIONE DELLA PIATTAFORMA ---
# ==================================================================
st.title("SECURITY AND HSE SYSTEM PLATFORM")
st.markdown("### Sistema di Gestione Integrato Locale e Multiutente")
st.markdown("---")

# --- 1. Inizializzazione dello Stato (Inseriscilo dopo st.set_page_config) ---
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Home Dashboard"

# --- 2. Creazione della barra di navigazione ---
# Usiamo st.radio in orizzontale per simulare le tab
tab_list = [
    "Home Dashboard", 
    "Segnalazione Near Miss", 
    "Scadenzario Adempimenti",
    "Analisi Segnalazioni Near Miss", 
    "Conformità Legislativa",
    "Analisi - Fase 2",
    "KPI",
    "Piano Miglioramento",
    "Stima Costo Economico",
    "Skill Matrix",
    "Controllo DPI"
]

# Impostiamo la navigazione
nav = st.radio(
    "Navigazione", 
    options=tab_list, 
    index=tab_list.index(st.session_state.active_tab), 
    horizontal=True, 
    label_visibility="collapsed"
)

# Aggiorniamo lo stato
st.session_state.active_tab = nav
st.markdown("---") # Linea di separazione estetica

# ==================================================================
# --- SEZIONE 1: HOME DASHBOARD ---
# ==================================================================
if nav == "Home Dashboard":
    st.header("Quadro Generale di Controllo HSE")
    st.markdown("Benvenuto nel menu principale. Qui trovi i dati riassuntivi estratti in tempo reale dai database locali.")
    st.markdown(" ")
    
    colA, colB, colC, colD = st.columns(4)
    with colA:
        tot_nm = len(pd.read_csv(FILE_NEAR_MISS, sep=";")) if os.path.exists(FILE_NEAR_MISS) else 0
        st.metric("Segnalazioni Ricevute", tot_nm)
    with colB:
        tot_an = len(pd.read_csv(FILE_ANALISI_NM, sep=";")) if os.path.exists(FILE_ANALISI_NM) else 0
        st.metric("Analisi Trattate dal RSPP", tot_an)
    with colC:
        tot_scad = len(pd.read_excel(FILE_SCADENZARIO)) if os.path.exists(FILE_SCADENZARIO) else 0
        st.metric("Adempimenti in Registro Scadenze", tot_scad)
    with colD:
        tot_cost = len

# ==================================================================
# --- SEZIONE 2: SEGNALAZIONE NEAR MISS ---
# ==================================================================
if nav == "Segnalazione Near Miss":
    st.info(
        "Near miss (mancato infortunio): incidente avvenuto nel luogo di lavoro che non ha recato danno fisico al lavoratore, pur avendone il potenziale.\n"
        "Esempi: caduta di materiale imballato durante movimentazione con carrello elevatore; improvvisa fuoriuscita di liquido da tubazione; lavoratore scivola su pavimento bagnato senza riportare danni.\n\n"
        "Non conformità: situazione di pericolo che non genera alcun incidente/infortunio ma rilevabile su procedure operative, attrezzature, ambienti di lavoro, dpi.\n"
        "Esempi: macchinario senza protezione, casco di sicurezza non indossato, area di lavoro priva di percorsi sicuri."
    )
    
    st.subheader("MODULO S.NM.NC - Segnalazione Near Miss o Non Conformità")
    
    st.markdown("#### Inserisce immagine (Facoltativo)")
    opzione_immagine = st.radio(
        "Scegli la modalità di inserimento immagine:", 
        ["Nessuna immagine", "Carica file", "Scatta foto col cellulare/webcam"],
        key="scelta_media_reattiva"
    )
    
    immagine_salvata_nome = "Nessuna"
    if opzione_immagine == "Carica file":
        file_img = st.file_uploader("Scegli un file immagine", type=["png", "jpg", "jpeg"], key="uploader_reattivo")
        if file_img:
            immagine_salvata_nome = file_img.name
    elif opzione_immagine == "Scatta foto col cellulare/webcam":
        foto_scattata = st.camera_input("Scatta una foto della criticità", key="camera_reattiva")
        if foto_scattata:
            immagine_salvata_nome = f"scatto_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            
    st.markdown("---")

    with st.form("form_segnalazione_near_miss", clear_on_submit=True):
        col_tipo, col_segnalatore = st.columns(2)
        with col_tipo:
            tipo_evento = st.radio("Tipo evento", ["Near Miss", "Non Conformità"])
        with col_segnalatore:
            segnalatore = st.text_input("Segnalatore (inserire mansione o nome cognome)")
            
        col_sesso, col_eta, col_data = st.columns(3)
        with col_sesso:
            sesso = st.radio("Sesso", ["Maschio", "Femmina"])
        with col_eta:
            fascia_eta = st.radio("Fascia di Età", ["<18 anni", "18-30 anni", "31-50 anni", "51-67 anni"])
        with col_data:
            data_evento = st.date_input("Data (formato gg/mm/aaaa)")
            
        col_luogo, col_reparto = st.columns(2)
        with col_luogo:
            luogo = st.radio("Luogo", ["In Azienda", "In itinere", "In missione"])
        with col_reparto:
            reparto_aziendale = st.text_input("Reparto (se è In Azienda)")
            
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fascia_oraria = st.radio("Fascia oraria di accadimento", ["0-6", "6-12", "12-18", "18-24"])
        with col_f2:
            fascia_lavoratore = st.text_input("Fascia oraria per il lavoratore (1, 2, 3 ora Max. 8 ore)")
            
        descrizione = st.text_area("Descrizione dell'evento o della criticità (campo a testo libero)")

        st.markdown("#### Possibili cause dell'evento / In caso di Non Conformità selezionare la tipologia")
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            c_err_proc = st.checkbox("Errore procedurale (disattenzione, scarsa conoscenza procedure operative, ...)")
            c_prob_comm = st.checkbox("Problema di comunicazione (lingua, incertezza nei ruoli e/o compiti, ...)")
            c_manc_proc = st.checkbox("Mancanza/inadeguatezza di procedure operative")
            c_manc_prot = st.checkbox("Mancanza di protezioni sull'attrezzatura")
            c_car_prot = st.checkbox("Carenza (inadeguatezza) di protezioni sull'attrezzatura")
            c_anom_guasto = st.checkbox("Anomalia/guasto in avviamento/arresto/esercizio (funzionamento)")
            c_unica_attrez = st.checkbox("Unica attrezzatura disponibile ma non idonea alla lavorazione")
            c_ass_attrez = st.checkbox("Assenza di attrezzature idonee alla lavorazione")
            c_stocc_err = st.checkbox("Stoccaggio/etichettatura errato di materiali")
            c_prob_mat = st.checkbox("Problema legato alle caratteristiche/trasformazioni di materiali")
            c_segnal_inad = st.checkbox("Segnaletica di sicurezza/Cartellonistica inadeguata o assente")
            c_ass_perc = st.checkbox("Assenza o inadeguatezza di percorsi in sicurezza, vie di transito, uscite di emergenza")
            
        with col_c2:
            c_illum_inad = st.checkbox("Illuminazione non idonea o assente")
            c_ass_barr = st.checkbox("Assenza o inadeguatezza di barriere, protezioni, parapetti, armature")
            c_spazi_inad = st.checkbox("Spazi inadeguati su postazioni di lavoro")
            c_ass_stocc = st.checkbox("Assenza o inadeguatezza di aree di stoccaggio")
            c_pres_liq = st.checkbox("Presenza imprevista di liquidi (acqua, olio, ...)")
            c_pres_gas = st.checkbox("Presenza imprevista di gas, vapori")
            c_crit_imp = st.checkbox("Criticità su impianti generali a supporto dell'area di lavoro")
            c_pres_elett = st.checkbox("Presenza di elettricità/linea elettrica accessibile")
            c_rumore = st.checkbox("Livelli di rumorosità inadeguati")
            c_manc_dpi = st.checkbox("Mancato uso o uso errato di DPI")
            c_dpi_non_forn = st.checkbox("DPI non fornito")
            c_dpi_inad = st.checkbox("DPI inadeguato")
            
        altro_specificare = st.text_input("Altro (specificare campo a testo libero)")
        storico_riscontro = st.radio(
            "In base alla tua esperienza lavorativa, la situazione rilevata o osservata si è già presentata in passato anche recente?",
            ["Sì frequentemente", "Sì raramente", "No"]
        )
        valutazioni_proposte = st.text_area("Valutazioni / azioni / proposte di miglioramento (campo a testo libero)")
        
        submit_modulo = st.form_submit_button("Registra ed Invia Segnalazione", use_container_width=True)
        
        if submit_modulo:
            if not descrizione.strip():
                st.error("Errore: La descrizione dell'evento è obbligatoria per effettuare il salvataggio.")
            else:
                cause_selezionate = []
                mappa_cause = {
                    "Errore procedurale": c_err_proc, "Problema comunicazione": c_prob_comm, "Mancanza procedure": c_manc_proc,
                    "Mancanza protezioni": c_manc_prot, "Carenza protezioni": c_car_prot, "Anomalia guasto": c_anom_guasto,
                    "Unica attrezzatura non idonea": c_unica_attrez, "Assenza attrezzature idonee": c_ass_attrez,
                    "Stoccaggio errato": c_stocc_err, "Problema materiali": c_prob_mat, "Segnaletica inadeguata": c_segnal_inad,
                    "Inadeguatezza percorsi": c_ass_perc, "Illuminazione inadeguata": c_illum_inad, "Assenza barriere": c_ass_barr,
                    "Spazi inadeguati": c_spazi_inad, "Assenza aree stoccaggio": c_ass_stocc, "Presenza liquidi": c_pres_liq,
                    "Presenza gas": c_pres_gas, "Criticità impianti": c_crit_imp, "Presenza elettricità": c_pres_elett,
                    "Rumorosità": c_rumore, "Mancato uso DPI": c_manc_dpi, "DPI non fornito": c_dpi_non_forn, "DPI inadeguato": c_dpi_inad
                }
                for nome_c, var_c in mappa_cause.items():
                    if var_c:
                        cause_selezionate.append(nome_c)
                if altro_specificare.strip():
                    cause_selezionate.append(f"Altro: {altro_specificare.strip()}")

                nuovo_record = {
                    "Data Segnalazione": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                    "Tipo Evento": tipo_evento,
                    "Segnalatore": segnalatore.strip() if segnalatore.strip() else "Anonimo",
                    "Sesso": sesso,
                    "Fascia Eta": fascia_eta,
                    "Data Evento Real": data_evento.strftime("%d/%m/%Y"),
                    "Luogo": luogo,
                    "Reparto": reparto_aziendale.strip() if reparto_aziendale.strip() else "N/D",
                    "Fascia Oraria Accadimento": fascia_oraria,
                    "Fascia Oraria Lavoratore": fascia_lavoratore.strip() if fascia_lavoratore.strip() else "N/D",
                    "Descrizione": descrizione.strip(),
                    "Immagine Allegata": immagine_salvata_nome,
                    "Cause Rilevate": ", ".join(cause_selezionate),
                    "Verificato in Passato": storico_riscontro,
                    "Proposte Miglioramento": valutazioni_proposte.strip(),
                    "Firma Presa in Carico": "Da firmare"
                }
                
                df_nuovo = pd.DataFrame([nuovo_record])
                if not os.path.isfile(FILE_NEAR_MISS):
                    df_nuovo.to_csv(FILE_NEAR_MISS, index=False, sep=";")
                else:
                    df_nuovo.to_csv(FILE_NEAR_MISS, mode='a', header=False, index=False, sep=";")
                st.success("Segnalazione acquisita con successo nel file CSV locale!")
                time.sleep(0.5)
                st.rerun()

# ==================================================================
# --- SEZIONE 3: SCADENZARIO ADEMPIMENTI ---
# ==================================================================
if nav == "Scadenzario Adempimenti":
    st.header("Registro Scadenzario Adempimenti Aziendali")
    st.markdown("Sezione Protetta — Modifiche simultanee salvate in tempo reale sul file Excel locale")
    
    if "autenticato_scadenze" not in st.session_state:
        st.session_state.autenticato_scadenze = False
    if "df_scadenzario_state" not in st.session_state:
        st.session_state.df_scadenzario_state = None
        
    if not st.session_state.autenticato_scadenze:
        pwd_scad = st.text_input("Inserisci la password di sblocco Scadenzario", type="password", key="pwd_scad_tab")
        if st.button("Convalida Password Scadenzario", use_container_width=True):
            if pwd_scad == "hse2026":
                st.session_state.autenticato_scadenze = True
                
                if os.path.exists(FILE_SCADENZARIO):
                    try:
                        df_caricato = pd.read_excel(FILE_SCADENZARIO)
                        for col in COLONNE_SCADENZARIO:
                            if col not in df_caricato.columns:
                                df_caricato[col] = ""
                        df_caricato = df_caricato[COLONNE_SCADENZARIO]
                    except Exception as e:
                        st.error(f"Errore lettura file: {e}")
                        df_caricato = pd.DataFrame(columns=COLONNE_SCADENZARIO)
                else:
                    df_caricato = pd.DataFrame(columns=COLONNE_SCADENZARIO)
                
                df_caricato = df_caricato.fillna("")
                for col in df_caricato.columns:
                    df_caricato[col] = df_caricato[col].astype(str).replace(["nan", "None", "<NA>", "NaT"], "").str.strip()
                
                st.session_state.df_scadenzario_state = df_caricato
                st.rerun()
            else:
                st.error("Password non corretta. Accesso negato.")
                
    if st.session_state.autenticato_scadenze and st.session_state.df_scadenzario_state is not None:
        st.success("Accesso Concesso alle scadenze.")
        st.markdown("---")
        
        st.subheader("1. Modifica ed Inserimento Dati")
        df_editor_input = st.session_state.df_scadenzario_state.copy()
        
        df_modificato = st.data_editor(
            df_editor_input,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="editor_scadenzario_pure_data"
        )
        
        if st.button("Aggiorna Calcoli Automatici e Salva nel File Excel", use_container_width=True):
            try:
                df_elaborazione = df_modificato.copy()
                df_elaborazione = df_elaborazione.fillna("")
                for col in df_elaborazione.columns:
                    df_elaborazione[col] = df_elaborazione[col].astype(str).replace(["nan", "None", "<NA>", "NaT"], "").str.strip()
                
                df_elaborazione["Scadenza"] = df_elaborazione.apply(calcola_data_scadenza, axis=1)
                df_elaborazione["In vigore durante (mesi)"] = pd.to_numeric(df_elaborazione["In vigore durante (mesi)"], errors='coerce').fillna(0).astype(int)
                
                df_elaborazione.to_excel(FILE_SCADENZARIO, index=False)
                st.session_state.df_scadenzario_state = df_elaborazione
                st.success("File Excel salvato!")
                time.sleep(0.4)
                st.rerun()
            except Exception as err:
                st.error(f"Errore durante il salvataggio: {err}")

        st.markdown("---")
        st.subheader("2. Registro Alert & Scadenze Effettive (Vista Finale)")
        df_vista_alert = st.session_state.df_scadenzario_state.copy()
        
        if not df_vista_alert.empty:
            df_vista_alert["Scadenza"] = df_vista_alert.apply(calcola_data_scadenza, axis=1)
            styler_colorato = df_vista_alert.style.apply(evidenzia_righe_scadenza, axis=1)
            st.dataframe(styler_colorato, use_container_width=True, hide_index=True)
        else:
            st.info("Nessun adempimento presente nel registro.")

# ==================================================================
# --- SEZIONE 4: ANALISI SEGNALAZIONI NEAR MISS ---
# ==================================================================
if nav == "Analisi Segnalazioni Near Miss":
    st.header("Analisi approfondita delle segnalazioni Near Miss")
    
    if "autenticato_rspp" not in st.session_state:
        st.session_state.autenticato_rspp = False
        
    if not st.session_state.autenticato_rspp:
        pwd_rspp = st.text_input("Inserisci la Password di Accesso", type="password", key="pwd_rspp_tab")
        if st.button("Convalida Accesso", use_container_width=True):
            if pwd_rspp == "hse2026":
                st.session_state.autenticato_rspp = True
                st.rerun()
            else:
                st.error("Credenziali errate.")
                
    if st.session_state.autenticato_rspp:
        st.success("Autenticato")
        df_segnalazioni = pd.read_csv(FILE_NEAR_MISS, sep=";") if os.path.exists(FILE_NEAR_MISS) else pd.DataFrame()
        df_analisi = pd.read_csv(FILE_ANALISI_NM, sep=";") if os.path.exists(FILE_ANALISI_NM) else pd.DataFrame()
        
        if "sub_sezione_rspp" not in st.session_state:
            st.session_state.sub_sezione_rspp = "compilazione"
            
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("Apri Nuovo Modulo Analisi", use_container_width=True):
                st.session_state.sub_sezione_rspp = "compilazione"
        with col_m2:
            if st.button("Commento e firma RSPP", use_container_width=True):
                st.session_state.sub_sezione_rspp = "firma"
                
        st.markdown("---")
        
        if st.session_state.sub_sezione_rspp == "compilazione":
            opzioni_tendina = ["Nessun collegamento (Crea analisi indipendente)"]
            if not df_segnalazioni.empty:
                for idx, row in df_segnalazioni.iterrows():
                    opzioni_tendina.append(f"{row.get('Data Segnalazione','N/D')} | {row.get('Tipo Evento','N/D')} | {row.get('Segnalatore','N/D')}")
                    
            selezione_nm = st.selectbox("Seleziona una segnalazione a cui allacciarti:", opzioni_tendina)
            desc_def = ""
            if selezione_nm != "Nessun collegamento (Crea analisi indipendente)":
                indice = opzioni_tendina.index(selezione_nm) - 1
                if indice < len(df_segnalazioni):
                    desc_def = str(df_segnalazioni.iloc[indice].get('Descrizione', ''))
                st.info("Testo caricato.")
            
            st.markdown("#### Inserimento immagine o allegato di supporto per l'Analisi (Facoltativo)")
            opzione_media_analisi = st.radio(
                "Scegli la modalità di inserimento file/immagine per l'analisi:", 
                ["Nessun file", "Carica file locale", "Scatta foto istantanea"],
                key="scelta_media_analisi_rspp"
            )
            
            allegato_analisi_nome = "Nessuna"
            if opzione_media_analisi == "Carica file locale":
                file_img_an = st.file_uploader("Scegli un file per l'analisi", type=["png", "jpg", "jpeg", "pdf", "docx"], key="uploader_analisi_rspp")
                if file_img_an:
                    allegato_analisi_nome = file_img_an.name
                    with open(os.path.join(DIR_IMMAGINI_ANALISI, file_img_an.name), "wb") as f_local:
                        f_local.write(file_img_an.getbuffer())
                    st.caption(f"File '{file_img_an.name}' salvato in {DIR_IMMAGINI_ANALISI}/")
            elif opzione_media_analisi == "Scatta foto istantanea":
                foto_scattata_an = st.camera_input("Scatta una foto della verifica tecnica", key="camera_analisi_rspp")
                if foto_scattata_an:
                    allegato_analisi_nome = f"analisi_scatto_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                    with open(os.path.join(DIR_IMMAGINI_ANALISI, allegato_analisi_nome), "wb") as f_local:
                        f_local.write(foto_scattata_an.getbuffer())
                    st.caption(f"Foto '{allegato_analisi_nome}' archiviata in {DIR_IMMAGINI_ANALISI}/")
            
            st.markdown("---")
            with st.form("form_analisi_sup"):
                descrizione_finale = st.text_area("Integrazione dell'evento", value=desc_def)
                incidente_selezionato = st.multiselect("Incidente potenziale:", ["Caduta dall’alto", "Ribaltamento mezzo", "Contatto elettrico", "Tagli, punture", "Altro"])
                attivita_selezionata = st.multiselect("Attività svolta:", ["Lavori manuali", "Azionamento macchine", "Manutenzione", "Altro"])
                cause_selezionate = st.multiselect("Cause radice:", ["Errore procedurale", "Illuminazione inadeguata", "Mancanza/Uso errato DPI", "Altro"])
                storico_eventi = st.radio("Già verificato in passato?", ["Sì frequentemente", "Sì raramente", "No"])
                criticita_selezionate = st.multiselect("Criticità:", ["Vigilanza/Coordinamento", "Emergenze e Antincendio", "Formazione carente", "Nessuna"])
                
                colX, colY = st.columns(2)
                with colX:
                    danno_strutture = st.radio("Danno a strutture", ["nessuno", "lieve", "medio", "notevole"])
                    danno_produttivita = st.radio("Danno produttivo", ["nessuna", "breve", "media", "rilevante"])
                with colY:
                    danno_persone = st.radio("Danno potenziale persone", ["nessuno", "lieve", "grave", "gravissimo"])
                    frequenza = st.radio("Frequenza stimata", ["rara", "frequente", "molto frequente"])
                    
                if st.form_submit_button("Salva Modulo Direzione"):
                    try:
                        nuova_an = {
                            "Data Analisi": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                            "Segnalazione Collegata": selezione_nm,
                            "Descrizione": descrizione_finale,
                            "Incidente": ", ".join(incidente_selezionato),
                            "Attività": ", ".join(attivita_selezionata),
                            "Cause": ", ".join(cause_selezionate),
                            "Storico": storico_eventi,
                            "Criticità": ", ".join(criticita_selezionate),
                            "Danno Strutture": danno_strutture,
                            "Danno Produttività": danno_produttivita,
                            "Danno Persone": danno_persone,
                            "Frequenza": frequenza,
                            "Commento RSPP": "",
                            "Firma RSPP (Stato)": "Non Firmato",
                            "Allegato Analisi": allegato_analisi_nome
                        }
                        df_n = pd.DataFrame([nuova_an])
                        if not os.path.isfile(FILE_ANALISI_NM):
                            df_n.to_csv(FILE_ANALISI_NM, index=False, sep=";")
                        else:
                            df_n.to_csv(FILE_ANALISI_NM, mode='a', header=False, index=False, sep=";")
                        st.success("Analisi e relativi allegati salvati correttamente nel database locale!")
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as e: st.error(f"Errore: {e}")
                    
        elif st.session_state.sub_sezione_rspp == "firma":
            if df_analisi.empty:
                st.warning("Nessuna analisi presente.")
            else:
                opzioni_r = []
                mappatura = {}
                for idx, r in df_analisi.iterrows():
                    testo_o = f"Analisi del {r.get('Data Analisi','N/D')} | Collegamento: {r.get('Segnalazione Collegata','')}"
                    opzioni_r.append(testo_o)
                    mappatura[testo_o] = idx
                    
                scelta_rec = st.selectbox("Scegli la riga da integrare:", opzioni_r)
                idx_sel = mappatura[scelta_rec]
                
                comm_pre = str(df_analisi.at[idx_sel, "Commento RSPP"]) if "Commento RSPP" in df_analisi.columns and pd.notna(df_analisi.at[idx_sel, "Commento RSPP"]) else ""
                comm_in = st.text_area("Note / Commenti del Professionista:", value=comm_pre)
                file_f = st.file_uploader("Carica Firma Grafica", type=["png","jpg","jpeg"])
                
                if file_f: st.image(file_f, width=150)
                if st.button("Salva ed Applica Modifiche in Riga"):
                    if not comm_in.strip(): st.error("Inserire un commento.")
                    else:
                        df_analisi.at[idx_sel, "Commento RSPP"] = comm_in.strip()
                        df_analisi.at[idx_sel, "Firma RSPP (Stato)"] = f"Firmato ({file_f.name})" if file_f else "Testo Convalidato"
                        df_analisi.to_csv(FILE_ANALISI_NM, index=False, sep=";")
                        st.success("Record aggiornato in linea!")
                        time.sleep(0.5)
                        st.rerun()

# ==================================================================
# --- SEZIONE 5: CONFORMITÀ LEGISLATIVA (RIFERIMENTI REALI) ---
# ==================================================================
if nav == "Conformità Legislativa":
    st.header("Conformità Legislativa e Archivio Normativo")
    
    if "autenticato_upload_legale" not in st.session_state:
        st.session_state.autenticato_upload_legale = False
        
    col_bot, col_upload = st.columns([1, 1])
    
    # -----------------------------------------------------------
    # AREA PUBBLICA (col_bot)
    # -----------------------------------------------------------
    with col_bot:
        st.subheader("Assistente di Scansione Istantanea")
        st.caption("Accesso Pubblico - Ricerca libera e lettura QR Code.")
        # 1. Ricerca Documentale
        prompt_utente = st.text_input("Scrivi l'oggetto della richiesta (es: requisiti antincendio, obblighi, nomine):", key="txt_prompt_sup")
        if st.button("Interroga Archivio Normativo", use_container_width=True):
            if prompt_utente.strip():
                with st.spinner("Scansione testi in corso..."):
                    risultati_ricerca = cerca_nei_documenti(prompt_utente, DIR_CONFORMITA)
                    if risultati_ricerca == "specifica":
                        st.warning("Inserisci una query di ricerca più dettagliata.")
                    elif not risultati_ricerca:
                        st.info("Nessun paragrafo corrisponde alla ricerca basata su parole chiave.")
                    else:
                        st.markdown("---")
                        st.markdown("### Riscontri della Ricerca:")
                        for r in risultati_ricerca:
                            st.markdown(f"**Fonte:** {r['fonte']} (Rilevanza: {r['score']}%)")
                            st.info(f"> {r['testo'][:1000]}")
                            try:
                                with open(r['percorso_completo'], "rb") as file_da_aprire:
                                    dati_file = file_da_aprire.read()
                                st.download_button(
                                    label=f"Apri / Scarica File Originale ({r['nome_file']})",
                                    data=dati_file,
                                    file_name=r['nome_file'],
                                    mime="application/octet-stream",
                                    key=f"dl_{r['fonte'].replace(' ','_')}_{time.time()}"
                                )
                            except Exception as e_file:
                                st.error(f"Impossibile agganciare l'anteprima del file: {e_file}")
            else:
                st.warning("Inserire un testo di ricerca prima di premere il pulsante.")
                
        # -----------------------------------------------------------
        # POSIZIONAMENTO PUBBLICO E STRUTTURAZIONE NATIVA QR CODE
        # -----------------------------------------------------------
        st.markdown("---")
        st.subheader("Lettura Link tramite QR Code")
        st.caption("Strumento pubblico di codifica rapida per la documentazione HSE aziendale.")
        
       # Scelta della modalità di input
        metodo_input = st.radio(
        "Scegli come inserire il codice:", 
        ["Carica file dal dispositivo", "Fotocamera"], 
        horizontal=True,
        key="qr_pub_metodo"
        )

        immagine_input = None

        # Gestione logica in base alla scelta
        if metodo_input == "Carica file dal dispositivo":
            immagine_input = st.file_uploader(
                "Seleziona un'immagine (png, jpg, jpeg):", 
                type=["png", "jpg", "jpeg"], 
                key="up_qr_pub"
            )
        else:
            immagine_input = st.camera_input("Scatta una foto del QR Code:", key="cam_qr_pub")
        
        risultato_pub = None

        # Elaborazione dell'immagine (funziona con entrambi gli input)
        if immagine_input:
        # La funzione decodifica_qr_opencv gestisce l'oggetto file_like (file o stream camera)
            risultato_pub = decodifica_qr_opencv(immagine_input)
    
        if risultato_pub:
            st.success("Codice QR decodificato con successo!")
            st.code(risultato_pub, language="text")
        
            if risultato_pub.startswith("http://") or risultato_pub.startswith("https://"):
                st.markdown(f"[Clicca qui per aprire il link rilevato]({risultato_pub})")
        else:
            st.warning("Nessun codice QR valido rilevato nell'immagine.")

        # -----------------------------------------------------------
        # AREA PRIVATA / AMMINISTRATIVA (col_upload)
        # -----------------------------------------------------------        
    with col_upload:
        st.subheader("Area Amministrazione Archivio")
        
        if not st.session_state.autenticato_upload_legale:
                st.caption("Questa funzione richiede i privileges di Amministrazione per inserire nuovi testi ed estrarre la check-list.")
                pwd_leg = st.text_input("Inserisci la password amministrativa", type="password", key="pwd_leg_tab")
                if st.button("Sblocca Funzioni Avanzate", use_container_width=True):
                    if pwd_leg == "hse2026":
                        st.session_state.autenticato_upload_legale = True
                        st.rerun()
                    else:
                        st.error("Password di conformità errata. Accesso negato.")
        else:
            st.success("Modalità di amministrazione attiva.")
            if st.button("Chiudi sessione Amministratore (Blocca)", use_container_width=True):
                st.session_state.autenticato_upload_legale = False
                st.rerun()
                # --- Funzioni Admin ---
            tab_admin1, tab_admin2 = st.tabs(["Documenti & Check-list", "Gestione QR & Normativa"])
            
            with tab_admin1:
                st.markdown("---")
                st.subheader("Caricamento Documenti")
                file_caricati = st.file_uploader(
                    "Trascina qui Decreti, Testi Unici o Tabelle Word/Excel da archiviare sul disco locale:",
                    type=["pdf", "csv", "xlsx", "xls", "docx", "doc", "jpg", "jpeg", "png"],
                    accept_multiple_files=True,
                    key="uploader_leg_sup"
                    )
                if file_caricati:
                    for f in file_caricati:
                        save_path = os.path.join(DIR_CONFORMITA, f.name)
                        with open(save_path, "wb") as f_out:
                            f_out.write(f.getbuffer())
                    st.success("File salvati!")
                
                # Check-list
                st.markdown("---")
                st.subheader("Estrazione Check-list Documentale")
            
                if st.button("Genera ed Elabora Check-list CSV", use_container_width=True, key="btn_genera_checklist_doc"):
                    file_presenti = [os.path.join(DIR_CONFORMITA, f) for f in os.listdir(DIR_CONFORMITA) if os.path.isfile(os.path.join(DIR_CONFORMITA, f))]
                    if not file_presenti:
                        st.warning("L'archivio è vuoto. Carica almeno un documento.")
                    else:
                         with st.spinner("Analisi ed estrazione strutturata..."):
                            righe_checklist = []
                            for fp in file_presenti:
                                 timestamp_creazione = os.path.getmtime(fp)
                                 data_caricamento = datetime.fromtimestamp(timestamp_creazione).strftime("%d-%m-%Y")
                                 titolo, regolamento, doc_produrre, settore_aziendale = estrai_info_checklist(fp)
                            
                                 righe_checklist.append({
                                    "Quando fu caricata": data_caricamento,
                                    "Titolo": titolo,
                                     "Settore aziendale": settore_aziendale,
                                     "Regolamento o legge di riferimento": regolamento,
                                     "Documenti da produrre o allegati": doc_produrre
                                })
                            df_checklist = pd.DataFrame(righe_checklist)
                            st.session_state["cached_df_checklist"] = df_checklist
                            st.success("Check-list aggiornata in memoria.")

                    if "cached_df_checklist" in st.session_state:
                        df_visualizza = st.session_state["cached_df_checklist"]
                        st.dataframe(df_visualizza, use_container_width=True, hide_index=True)
                        csv_buffer = df_visualizza.to_csv(index=False, sep=";").encode('utf-8')
                        st.download_button(
                            label="Scarica Check-list in formato CSV",
                            data=csv_buffer,
                            file_name=f"checklist_conformita_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                        
            # -----------------------------------------------------------
            # GENERA QRCode
            # -----------------------------------------------------------    
            with tab_admin2:
                # Generazione QR
                st.markdown("#### Genera QR Code")
                opzione_qr = st.radio("Seleziona operazione QR:", ["Genera QR Code", "Leggi/Decodifica QR Code"], key="radio_opzioni_qr_pubblico")
                img_ottenuta = None
        
                if opzione_qr == "Genera QR Code":
                    testo_da_convertire = st.text_input("Inserisci l'URL o il testo da inserire nel QR Code:", placeholder="https://www.gazzettaufficiale.it/")
                    if testo_da_convertire.strip():
                        img_ottenuta = genera_qr_nativo(testo_da_convertire)
                
                    # Esegui la generazione e il salvataggio SOLO se il testo è presente
                    if testo_da_convertire.strip():
                        img_ottenuta = genera_qr_nativo(testo_da_convertire)
        
                    # Verifica che img_ottenuta non sia None (buona pratica)
                        if img_ottenuta:
                        # Salvataggio temporaneo per consentire la visualizzazione e il download
                            percorso_temp_qr = "temp_generated_qr.png"
                            img_ottenuta.save(percorso_temp_qr)
                
                            st.image(percorso_temp_qr, width=220, caption="Codice QR generato con successo!")
                
                            with open(percorso_temp_qr, "rb") as f_qr:
                                st.download_button(
                                    label="Scarica Immagine QR Code (.png)",
                                    data=f_qr.read(),
                                    file_name="qr_code_hse.png",
                                    mime="image/png",
                                    use_container_width=True
                                )
                    
                elif opzione_qr == "Leggi/Decodifica QR Code":
                    file_qr_caricato = st.file_uploader("Carica un'immagine contenente un codice QR:", type=["png", "jpg", "jpeg"], key="uploader_file_qr_nativo")
                    if file_qr_caricato:
                        risultato_priv = decodifica_qr_opencv(file_qr_caricato)
                        if risultato_priv:
                            st.success("Codice QR decodificato con successo!")
                            st.code(risultato_priv, language="text")
                    
                            if risultato_priv.startswith("http://") or risultato_priv.startswith("https://"):
                                st.markdown(f"[Clicca qui per aprire il link rilevato]({risultato_priv})")
                        else:
                            st.warning("Nessun codice QR valido rilevato all'interno dell'immagine caricata.")

                # -----------------------------------------------------------
                # VERIFICA AGGIORNAMENTI COGENTI CON LINK REALI
                # -----------------------------------------------------------
                st.markdown("---")
                st.subheader("Verifica Aggiornamenti Cogenti (Italia & UE)")
                st.caption("Il sistema incrocia i dati estratti con il contesto operativo e le fonti ufficiali della gazzetta italiana ed europea.")
            
                quadro_libero = st.text_area(
                    "Fornisci dettagli sul contesto aziendale o processi specifici (es: smart working, rifiuti pericolosi, imballaggi plastici, marketing ecologico):",
                    placeholder="Inserisci qui note libere per la verifica..."
                )
            
                if st.button("Genera ed Elabora Check-list CSV", use_container_width=True, key="btn_genera_checklist_cogenti"):
                    file_presenti = [os.path.join(DIR_CONFORMITA, f) for f in os.listdir(DIR_CONFORMITA) if os.path.isfile(os.path.join(DIR_CONFORMITA, f))]
                    righe = []
                    for fp in file_presenti:
                        titolo, reg, settore, data = estrai_info_checklist(fp)
                        righe.append({"Titolo Documento": titolo, "Regolamento": reg, "Settore": settore, "Data Caricamento": data})
                    st.session_state["cached_df_checklist"] = pd.DataFrame(righe)
                    st.rerun()

                if "cached_df_checklist" in st.session_state:
                    st.dataframe(st.session_state["cached_df_checklist"], use_container_width=True)
                    csv = st.session_state["cached_df_checklist"].to_csv(index=False).encode('utf-8')
                    st.download_button("Scarica Check-list CSV", csv, "checklist.csv", "text/csv")

                # 4. Motore Confronto Legislativo Dinamico (Con Gap-Analysis)
                st.markdown("---")
                st.subheader("Confronto Legislativo & Stato Conformità")
            
                with st.expander("Aggiungi nuova norma al database interno"):
                    k = st.text_input("Parola chiave (es: rifiuti):")
                    t = st.text_input("Titolo:")
                    u = st.text_input("URL ufficiale:")
                    d = st.text_area("Descrizione della norma:")
                    if st.button("Salva nel database"): 
                        salva_normativa(k, t, u, d)
                        st.success("Norma salvata!")
            
                note_libere = st.text_area("Analisi testo libero / criticità:", placeholder="es: gestione rifiuti pericolosi, amianto, ecc.")
            
                if st.button("Avvia Analisi di Conformità", use_container_width=True):
                    # 1. Caricamento Database Normativo
                    db = {}
                    if os.path.exists("normative.json"):
                        with open("normative.json", "r") as f:
                            try: db = json.load(f)
                            except: db = {}
                
                    # 2. Preparazione dati (Checklist e File)
                    txt_checklist = st.session_state["cached_df_checklist"].to_string() if "cached_df_checklist" in st.session_state else ""
                    txt_docs = " ".join(os.listdir(DIR_CONFORMITA)) if os.path.exists(DIR_CONFORMITA) else ""
                    txt_totale_analisi = (note_libere + " " + txt_checklist + " " + txt_docs).lower()
                
                    st.markdown("### Risultati Stato Conformità")
                
                    trovato = False
                    for chiave, info in db.items():
                        # Se la parola chiave è presente nell'analisi o nella checklist/file
                        if chiave.lower() in txt_totale_analisi:
                            trovato = True
                        
                        # LOGICA DI CONFORMITÀ: 
                        # Verifichiamo se la parola chiave è presente nei file caricati (conformità documentale)
                        is_conforme = chiave.lower() in (txt_checklist + txt_docs).lower()
                        
                        # Layout Risultato
                        col_stato, col_info = st.columns([1, 4])
                        
                        with col_stato:
                            if is_conforme:
                                st.success("CONFORME")
                                st.caption("Documentazione trovata")
                            else:
                                st.error("NON CONFORME")
                                st.caption("Documentazione mancante")
                        
                        with col_info:
                            with st.expander(f"Dettaglio: {info['titolo']}", expanded=True):
                                st.write(f"**Descrizione:** {info['desc']}")
                                st.link_button("Vai alla fonte ufficiale", info['url'])
                                
                                if not is_conforme:
                                    st.warning(f"Attenzione: Il tema '{chiave}' è rilevato nel contesto, ma non risulta tra i documenti archiviati nella checklist.")
                
                    if not trovato:
                        st.info("Nessun tema normativo specifico del database rilevato nei documenti o nelle note.")
# ==================================================================
# --- SEZIONE 6: Analisi - Fase 2 ---
# ==================================================================
if nav == "Analisi - Fase 2":
    st. header ("Analisi - Fase 2 delle segnalazioni near miss")
            # --- SEZIONE 2: ANALISI ISHIKAWA ---
    if "autenticato_fase2" not in st.session_state:
        st.session_state.autenticato_fase2 = False
        
    if not st.session_state.autenticato_fase2:
        pwd_fase2 = st.text_input("Inserisci la Password di Accesso", type="password", key="pwd_fase2_tab")
        if st.button("Convalida Accesso", use_container_width=True):
            if pwd_fase2 == "hse2026":
                st.session_state.autenticato_fase2 = True
                st.rerun()
            else:
                st.error("Credenziali errate.")
    if st.session_state.autenticato_fase2:

        # Caricamento e unione dati per dropdown
        opzioni = ["Nessuna (Nuova analisi)"]
            
        # Leggi Near Miss
        if os.path.exists(FILE_NEAR_MISS):
            df_nm = pd.read_csv(FILE_NEAR_MISS, sep=";")
            for idx, r in df_nm.iterrows():
                opzioni.append(f"NM | {r.get('Data Segnalazione', 'N/D')} | {r.get('Tipo Evento', 'Evento')}")
            
        # Leggi Analisi già fatte
        if os.path.exists(FILE_ANALISI_NM):
            df_an = pd.read_csv(FILE_ANALISI_NM, sep=";")
            for idx, r in df_an.iterrows():
                opzioni.append(f"AN | {r.get('Data Analisi', 'N/D')} | Collegamento: {r.get('Segnalazione Collegata', 'Analisi')}")
            
        scelta_rif = st.selectbox("Seleziona evento/analisi collegata:", opzioni)

        # Campi Input (Ishikawa)
        c_ma, c_mb = st.columns(2)
        with c_ma:
            macchina = st.text_area("Macchina / Infrastruttura", key="inp_m")
            materiale = st.text_area("Materiale", key="inp_mat")
        with c_mb:
            metodo = st.text_area("Metodo / Procedura", key="inp_met")
            manodopera = st.text_area("Manodopera / Comportamento", key="inp_man")
            
        w1 = st.text_input("Perché 1:", key="w1")
        w2 = st.text_input("Perché 2:", key="w2")
        w3 = st.text_input("Perché 3:", key="w3")
        w4 = st.text_input("Perché 4:", key="w4")
        w5 = st.text_input("Perché 5:", key="w5")
        conc = st.text_area("Conclusioni:", key="conc")

        if st.button("Genera Report e File"):
            temp_img = os.path.join(DIR_ANALISI, "temp_ishikawa.png")
                
            # Matplotlib setup
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.axis('off')
            ax.plot([1, 9], [0, 0], color='black', lw=3)
                
            # Aggiunta branchie (Ishikawa)
            def wrap(text): return "\n".join(textwrap.wrap(str(text), width=20))
            branches = [
                ((2, 0), (1.5, 1.5), f"Macchina:\n{wrap(macchina)}"),
                ((4, 0), (3.5, -1.5), f"Metodo:\n{wrap(metodo)}"),
                ((6, 0), (6.5, 1.5), f"Materiale:\n{wrap(materiale)}"),
                ((8, 0), (8.5, -1.5), f"Manodopera:\n{wrap(manodopera)}")
            ]
            for start, end, label in branches:
                ax.plot([start[0], end[0]], [start[1], end[1]], 'k-', lw=2)
                ax.text(end[0], end[1], label, fontsize=9, ha='center', va='center', bbox=dict(facecolor='white', alpha=0.8))
                
            plt.savefig(temp_img, bbox_inches='tight', dpi=300)
            plt.close()
            st.pyplot(fig)
            
            # Lettura dell'immagine PNG creata in binario per il salvataggio in sessione
            with open(temp_img, "rb") as img_file:
                png_bytes = img_file.read()

            # Raccolta metadati richiesti (Data generazione e file associati)
            data_generazione = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            file_obbligatorio = FILE_ANALISI_NM
            file_facoltativo = FILE_NEAR_MISS if os.path.exists(FILE_NEAR_MISS) else "Non utilizzato / Assente"

            # Conversione dell'immagine in stringa Base64 per l'inclusione nel JSON
            png_base64 = base64.b64encode(png_bytes).decode('utf-8')

            # Dati strutturati per Export
            dati_report = {
                "Data Generazione Report": data_generazione,
                "File Obbligatorio Associato": file_obbligatorio,
                "File Facoltativo Associato": file_facoltativo,
                "Riferimento Selezionato": scelta_rif,
                "4M": {
                    "Macchina": macchina, 
                    "Materiale": materiale, 
                    "Metodo": metodo, 
                    "Manodopera": manodopera
                },
                "5Whys": [w1, w2, w3, w4, w5],
                "Conclusioni": conc
            }

            # 1. PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, "Report Analisi Near Miss", ln=True, align='C')
            pdf.ln(10) # Spazio
        
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 6, f"Data Generazione: {data_generazione}", ln=True)
            pdf.cell(0, 6, f"File Obbligatorio: {file_obbligatorio}", ln=True)
            pdf.cell(0, 6, f"File Facoltativo: {file_facoltativo}", ln=True)
            pdf.cell(0, 6, f"Riferimento: {scelta_rif}", ln=True)
            pdf.ln(5)

            # Scrittura Analisi 4M
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "Analisi 4M:", ln=True)
            pdf.set_font("Arial", size=12)
            for key, val in dati_report["4M"].items():
                pdf.cell(0, 10, f"- {key}: {val}", ln=True)
            pdf.ln(5)

            # Scrittura 5Whys
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "5Whys:", ln=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 10, f"Perchè 1: {dati_report['5Whys'][0]}", ln=True)
            pdf.ln(5)

            # Scrittura Conclusioni
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "Conclusioni:", ln=True)
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, dati_report['Conclusioni'])

            # Inserimento visivo del grafico PNG nel PDF
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 10, "Diagramma di Ishikawa:", ln=True)
            pdf.image(temp_img, w=180)

            # Output del file in bytes
            pdf_output = bytes(pdf.output(dest='S')) 
            st.session_state.pdf_bytes = pdf_output

            # 2. XLSX
            import io
            buffer_xlsx = io.BytesIO()
            pd.DataFrame([dati_report["4M"]]).to_excel(buffer_xlsx, index=False)
            st.session_state.xlsx_bytes = buffer_xlsx.getvalue()

            # 3. JSON (Include i metadati, i testi e l'immagine codificata in Base64)
            st.session_state.json_bytes = json.dumps(dati_report, indent=4, ensure_ascii=False).encode('utf-8')
                
            # 4. PNG (File immagine separato)
            st.session_state.png_bytes = png_bytes

            st.session_state.report_ready = True
            st.success("Report generato con successo!")

        # Bottoni Download
        if st.session_state.get("report_ready"):
            col_d1, col_d2, col_d3 = st.columns(3)
            col_d1.download_button("Scarica PDF", st.session_state.pdf_bytes, "Report.pdf")
            col_d2.download_button("Scarica XLSX", st.session_state.xlsx_bytes, "Dati.xlsx")
            col_d3.download_button("Scarica JSON", st.session_state.json_bytes, "Dati.json")

# ==================================================================
# --- SEZIONE 7: KPI ---
# ==================================================================
if nav == "KPI":
    st.header("KPI & Indicatori di Performance HSE")
    
    if "autenticato_kpi" not in st.session_state:
        st.session_state.autenticato_kpi = False
        
    if not st.session_state.autenticato_kpi:
        pwd_kpi = st.text_input("Inserisci la Password di Accesso per la sezione KPI", type="password", key="pwd_kpi_sec")
        if st.button("Convalida Accesso KPI", use_container_width=True):
            if pwd_kpi == "hse2026":
                st.session_state.autenticato_kpi = True
                st.rerun()
            else:
                st.error("Password errata.")
        
    if st.session_state.autenticato_kpi:
        def render_sezione_kpi():
            st.title("Sezione KPI Aziendali")
            st.markdown(
                "Gestione centralizzata, monitoraggio e storico delle performance di"
                " sicurezza sul lavoro."
            )
            st.markdown("---")

            oggi = datetime.today().date()
            data_str = oggi.strftime('%Y-%m-%d')

            # Creazione della cartella KPI se non esiste
            os.makedirs("KPI", exist_ok=True)

            # Percorsi per la persistenza automatica dei dati tra i refresh
            path_master_pers = os.path.join("KPI", "master_kpi_definitions.csv")
            path_storico_pers = os.path.join("KPI", "storico_misure.csv")

            # ==========================================
            # ARCHIVIO CONDIVISO UNICO (PRESENTE IN ENTRAMBE LE TAB)
            # ==========================================
            if "t1_kpi_definitions" not in st.session_state:
                if os.path.exists(path_master_pers):
                    try:
                        st.session_state.t1_kpi_definitions = pd.read_csv(path_master_pers).fillna("")
                    except Exception:
                        st.session_state.t1_kpi_definitions = pd.DataFrame()
                else:
                    st.session_state.t1_kpi_definitions = pd.DataFrame([
                        {
                            "ID_KPI": "NM-01",
                            "Nome_KPI": "Numero Totale Near Miss Segnalati",
                            "Categoria": "Volume",
                            "Unita_Misura": "Numero",
                            "Cadenza": "Settimanale",
                            "Formula": "Conteggio assoluto segnalazioni validate",
                            "Parametro_1": "Segnalazioni valide",
                            "Parametro_2": "",
                            "Parametro_3": "",
                            "Baseline": "Media storica di riferimento: 10",
                        },
                        {
                            "ID_KPI": "NM-02",
                            "Nome_KPI": "Tasso di Risoluzione Near Miss",
                            "Categoria": "Efficacia",
                            "Unita_Misura": "%",
                            "Cadenza": "Settimanale",
                            "Formula": "(Parametro_1 / Parametro_2) * 100",
                            "Parametro_1": "Near Miss Chiusi",
                            "Parametro_2": "Near Miss Totali",
                            "Parametro_3": "",
                            "Baseline": "Target minimo aziendale: 75%",
                        },
                        {
                            "ID_KPI": "NM-03",
                            "Nome_KPI": "Tempo Medio di Presa in Carico",
                            "Categoria": "Reattività",
                            "Unita_Misura": "Giorni",
                            "Cadenza": "Bisettimanale",
                            "Formula": "Parametro_1 / Parametro_2",
                            "Parametro_1": "Somma giorni attesa",
                            "Parametro_2": "Numero totale pratiche",
                            "Parametro_3": "",
                            "Baseline": "Target massimo accettabile: 2 giorni",
                        },
                    ])
                    st.session_state.t1_kpi_definitions.to_csv(path_master_pers, index=False, encoding='utf-8')

            # Storico delle misurazioni periodiche nel tempo
            if "t2_storico_misure" not in st.session_state:
                if os.path.exists(path_storico_pers):
                    try:
                        df_loaded = pd.read_csv(path_storico_pers)
                        df_loaded["Data_Monitoraggio"] = pd.to_datetime(df_loaded["Data_Monitoraggio"]).dt.date
                        st.session_state.t2_storico_misure = df_loaded.fillna("")
                    except Exception:
                        st.session_state.t2_storico_misure = pd.DataFrame()
                else:
                    st.session_state.t2_storico_misure = pd.DataFrame([
                        {
                            "Data_Monitoraggio": oggi - timedelta(days=7),
                            "ID_KPI": "NM-01",
                            "Nome_KPI": "Numero Totale Near Miss Segnalati",
                            "Valore_Registrato": 12.0,
                            "Dettaglio_Fattori": "Segnalazioni valide: 12",
                        },
                        {
                            "Data_Monitoraggio": oggi,
                            "ID_KPI": "NM-01",
                            "Nome_KPI": "Numero Totale Near Miss Segnalati",
                            "Valore_Registrato": 14.0,
                            "Dettaglio_Fattori": "Segnalazioni valide: 14",
                        },
                        {
                            "Data_Monitoraggio": oggi - timedelta(days=7),
                            "ID_KPI": "NM-02",
                            "Nome_KPI": "Tasso di Risoluzione Near Miss",
                            "Valore_Registrato": 70.0,
                            "Dettaglio_Fattori": "Near Miss Chiusi: 35.0 / Near Miss Totali: 50.0",
                        },
                        {
                            "Data_Monitoraggio": oggi,
                            "ID_KPI": "NM-02",
                            "Nome_KPI": "Tasso di Risoluzione Near Miss",
                            "Valore_Registrato": 78.5,
                            "Dettaglio_Fattori": "Near Miss Chiusi: 47.0 / Near Miss Totali: 60.0",
                        },
                    ])
                    st.session_state.t2_storico_misure.to_csv(path_storico_pers, index=False, encoding='utf-8')

            # --- NAVIGAZIONE PRINCIPALE STABILE (PERSISTENTE) ---
            if "kpi_main_nav" not in st.session_state:
                st.session_state.kpi_main_nav = "📊 Definizione Master KPI"

            st.session_state.kpi_main_nav = st.radio(
                "Seleziona Sezione Principale KPI",
                ["📊 Definizione Master KPI", "⚙️ Strumento di Monitoraggio & Storico"],
                horizontal=True,
                key="radio_kpi_main_nav_selector"
            )
            st.markdown("---")

            # ==========================================
            # TAB 1: DEFINIZIONE MASTER KPI
            # ==========================================
            if st.session_state.kpi_main_nav == "📊 Definizione Master KPI":
                st.subheader("📋 Gestione Master dei KPI, Formule e Parametri")
                st.markdown(
                    "I KPI inseriti o modificati qui vengono riconosciuti e sincronizzati"
                    " automaticamente nella `tab_kpi_2`."
                )
                st.markdown("---")

                edited_t1_kpi = st.data_editor(
                    st.session_state.t1_kpi_definitions,
                    num_rows="dynamic",
                    key="t1_kpi_editor",
                    use_container_width=True,
                )

                if st.button("💾 Salva Modifiche Master (Tab 1)", key="t1_save_btn"):
                    st.session_state.t1_kpi_definitions = edited_t1_kpi.copy()
                    st.success("Definizioni e parametri sincronizzati con successo!")
                    st.rerun()

                st.markdown("---")
                # Sezione Download e Salvataggio Tab 1
                st.markdown("##### 📥 Esportazione Tabella Master KPI")
                nome_file_t1 = f"{data_str}_Tab1-MasterKPI.csv"
                path_t1 = os.path.join("KPI", nome_file_t1)
                csv_t1 = st.session_state.t1_kpi_definitions.to_csv(index=False).encode('utf-8')

                col_exp_1, col_exp_2 = st.columns(2)
                with col_exp_1:
                    if st.button("📁 Salva in cartella 'KPI' (Tab 1)", key="btn_save_folder_t1", use_container_width=True):
                        st.session_state.t1_kpi_definitions.to_csv(path_t1, index=False, encoding='utf-8')
                        st.success(f"Salvato con successo in: {path_t1}")
                with col_exp_2:
                    st.download_button(
                        label="⬇️ Scarica CSV Master KPI",
                        data=csv_t1,
                        file_name=nome_file_t1,
                        mime="text/csv",
                        key="dl_btn_t1",
                        use_container_width=True
                    )

                st.markdown("---")
                
                # --- ELIMINAZIONE PROTETTA MASTER KPI ---
                st.markdown("##### 🗑️ Eliminazione Master KPI (Richiede Password)")
                with st.form("form_elimina_master_kpi"):
                    kpi_list_names = st.session_state.t1_kpi_definitions["Nome_KPI"].tolist() if not st.session_state.t1_kpi_definitions.empty else []
                    kpi_da_eliminare = st.selectbox("Seleziona KPI da eliminare permanentemente", kpi_list_names, key="sel_kpi_del")
                    pwd_del_kpi = st.text_input("Inserisci la password di sezione per confermare l'eliminazione", type="password", key="pwd_del_kpi_field")
                    
                    if st.form_submit_button("Conferma ed Elimina Master KPI", use_container_width=True):
                        if pwd_del_kpi == "hse2026":
                            if kpi_da_eliminare:
                                # Rimuovi da Master KPI
                                st.session_state.t1_kpi_definitions = st.session_state.t1_kpi_definitions[
                                    st.session_state.t1_kpi_definitions["Nome_KPI"] != kpi_da_eliminare
                                ]
                                # Rimuovi anche dallo storico collegato
                                st.session_state.t2_storico_misure = st.session_state.t2_storico_misure[
                                    st.session_state.t2_storico_misure["Nome_KPI"] != kpi_da_eliminare
                                ]
                                # Aggiorna i file persistenti su disco
                                st.session_state.t1_kpi_definitions.to_csv(path_master_pers, index=False, encoding='utf-8')
                                st.session_state.t2_storico_misure.to_csv(path_storico_pers, index=False, encoding='utf-8')

                                st.success(f"KPI '{kpi_da_eliminare}' eliminato con successo!")
                                st.rerun()
                        else:
                            st.error("Password errata. Impossibile procedere con l'eliminazione.")

                st.markdown("---")
                st.metric(
                    "KPI Attivi nel Sistema", len(st.session_state.t1_kpi_definitions)
                )

            # ==========================================
            # TAB 2: STRUMENTO DI MONITORAGGIO E STORICO
            # ==========================================
            else:
                st.subheader(
                    "Strumento di Monitoraggio (Riconoscimento Automatico KPI da Tab 1)"
                )
                st.markdown(
                    "Qui vengono letti gli stessi KPI definiti nella Tab 1, consentendo di"
                    " inserire i fattori di calcolo e registrare nuove misurazioni nel"
                    " tempo."
                )
                st.markdown("---")

                # Acquisizione diretta degli stessi KPI della Tab 1
                df_kpi_master = st.session_state.t1_kpi_definitions.copy()

                def calcola_prossima_data_approssimata(row):
                    storico_kpi = st.session_state.t2_storico_misure[
                        st.session_state.t2_storico_misure["ID_KPI"] == row["ID_KPI"]
                    ]
                    if not storico_kpi.empty:
                        ultima_data = pd.to_datetime(
                            storico_kpi["Data_Monitoraggio"].max()
                        ).date()
                        cad = str(row["Cadenza"])
                        if "Settimanale" in cad:
                            delta_gg = 7
                        elif "Bisettimanale" in cad:
                            delta_gg = 14
                        elif "Mensile" in cad:
                            delta_gg = 30
                        elif "Trimestrale" in cad:
                            delta_gg = 90
                        elif "Semestrale" in cad:
                            delta_gg = 180
                        else:
                            delta_gg = 365
                        return ultima_data + timedelta(days=delta_gg)
                    else:
                        return oggi + timedelta(days=7)
                    
                if not df_kpi_master.empty:
                    df_kpi_master["Prossima_Data"] = df_kpi_master.apply(
                        calcola_prossima_data_approssimata, axis=1
                    )

                    def calcola_stato_t2(row):
                        delta = (row["Prossima_Data"] - oggi).days
                        if delta < 0:
                            return "Scaduto", "🔴"
                        elif delta <= 3:
                            return "In Scadenza", "🟡"
                        else:
                            return "Regolare", "🟢"

                    df_kpi_master["Stato_Testo"], df_kpi_master["Stato_Colore"] = zip(
                        *df_kpi_master.apply(calcola_stato_t2, axis=1)
                    )
                    df_kpi_master["Giorni_Rimasti"] = df_kpi_master["Prossima_Data"].apply(
                        lambda x: (x - oggi).days
                    )

                # --- SOTTO-NAVIGAZIONE STABILE PER IL MONITORAGGIO ---
                if "kpi_sub_nav" not in st.session_state:
                    st.session_state.kpi_sub_nav = "📈 Dashboard & Grafici"

                st.session_state.kpi_sub_nav = st.radio(
                    "Seleziona Sottosezione Monitoraggio",
                    [
                        "📈 Dashboard & Grafici",
                        "🕒 Registra Nuovo Monitoraggio",
                        "🛠️ Aggiungi Nuovo KPI & Configurazione",
                    ],
                    horizontal=True,
                    key="radio_kpi_sub_nav_selector"
                )
                st.markdown("---")

                # --- SOTTO-TAB 1: MONITORAGGIO E GRAFICI ---
                if st.session_state.kpi_sub_nav == "📈 Dashboard & Grafici":
                    if not df_kpi_master.empty:
                        scadenze_critiche_t2 = df_kpi_master[
                            df_kpi_master["Giorni_Rimasti"] <= 3
                        ]

                        if not scadenze_critiche_t2.empty:
                            st.warning(
                                f"⚠️ Attenzione: {len(scadenze_critiche_t2)} KPI richiedono un"
                                " monitoraggio imminente!"
                            )
                        else:
                            st.success("Tutti i KPI condivisi sono regolari.")

                        st.markdown("#### Tabella di Controllo KPI Sincronizzati")

                        def color_status_t2(val):
                            if val == "Scaduto":
                                return (
                                    "background-color: #ffcccc; color: #900c3f; font-weight: bold;"
                                )
                            elif val == "In Scadenza":
                                return (
                                    "background-color: #fff3cd; color: #856404; font-weight: bold;"
                                )
                            else:
                                return "background-color: #d4edda; color: #155724;"

                        df_controllo_display = df_kpi_master[[
                                "Stato_Colore",
                                "ID_KPI",
                                "Nome_KPI",
                                "Categoria",
                                "Unita_Misura",
                                "Cadenza",
                                "Prossima_Data",
                                "Formula",
                                "Baseline",
                                "Stato_Testo",
                            ]]
                        st.dataframe(
                            df_controllo_display
                            .style.map(color_status_t2, subset=["Stato_Testo"])
                            .format({"Prossima_Data": lambda t: t.strftime("%d/%m/%Y")}),
                            use_container_width=True,
                        )
                        # Esportazione Tabella di Controllo KPI (Sottosezione Tab2 - Controllo)
                        st.markdown("##### 📥 Esportazione Tabella di Controllo KPI")
                        nome_file_controllo = f"{data_str}_Tab2-ControlloKPI.csv"
                        path_controllo = os.path.join("KPI", nome_file_controllo)
                        csv_controllo = df_controllo_display.to_csv(index=False).encode('utf-8')

                        col_c1, col_c2 = st.columns(2)
                        with col_c1:
                            if st.button("📁 Salva in cartella 'KPI' (Controllo KPI)", key="btn_save_folder_ctrl", use_container_width=True):
                                df_controllo_display.to_csv(path_controllo, index=False, encoding='utf-8')
                                st.success(f"Salvato con successo in: {path_controllo}")
                        with col_c2:
                            st.download_button(
                                label="⬇️ Scarica CSV Controllo KPI",
                                data=csv_controllo,
                                file_name=nome_file_controllo,
                                mime="text/csv",
                                key="dl_btn_ctrl",
                                use_container_width=True
                            )

                        st.markdown("---")
                        st.markdown("#### 📊 Evoluzione Storica dei KPI")
                        kpi_selezionato_storico = st.selectbox(
                            "Seleziona KPI per visualizzare il grafico temporale:",
                            df_kpi_master["Nome_KPI"],
                            key="t2_sel_storico_chart",
                        )
                        df_storico_filtrato = st.session_state.t2_storico_misure[
                            st.session_state.t2_storico_misure["Nome_KPI"]
                            == kpi_selezionato_storico
                        ]

                        if not df_storico_filtrato.empty:
                            df_chart = df_storico_filtrato.set_index("Data_Monitoraggio")[
                                ["Valore_Registrato"]
                            ]
                            st.line_chart(df_chart)
                        else:
                            st.info("Nessuno storico registrato per questo KPI.")
                        st.info("Nessun KPI presente.")

                # --- SOTTO-TAB 2: REGISTRA NUOVO MONITORAGGIO PERIODICO ---
                elif st.session_state.kpi_sub_nav == "🕒 Registra Nuovo Monitoraggio":
                    st.markdown(
                        "#### 🕒 Aggiungi Nuovo Monitoraggio (Nuova Riga nello Storico)"
                    )
                    st.markdown(
                        "I KPI sottostanti sono riconosciuti direttamente da quelli definiti"
                        " nella `tab_kpi_1`."
                    )

                    if not df_kpi_master.empty:
                        kpi_sel_upd = st.selectbox(
                            "Seleziona KPI Riconosciuto",
                            df_kpi_master["Nome_KPI"],
                            key="t2_sel_kpi_storico",
                        )

                        kpi_row = df_kpi_master[
                            df_kpi_master["Nome_KPI"] == kpi_sel_upd
                        ].iloc[0]
                        formula_corrente = str(kpi_row["Formula"])
                        cadenza_corrente = str(kpi_row["Cadenza"])
                        id_kpi_corrente = str(kpi_row["ID_KPI"])
                        p1_nome = str(kpi_row["Parametro_1"])
                        p2_nome = str(kpi_row["Parametro_2"])
                        p3_nome = str(kpi_row["Parametro_3"])

                        st.info(
                            f"**Formula:** `{formula_corrente}` | **Baseline:**"
                            f" `{kpi_row['Baseline']}`"
                        )

                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            valori_fattori = []
                            dettagli_lista = []

                            if p1_nome and p1_nome.lower() != "nan" and p1_nome.strip() != "":
                                v1 = st.number_input(
                                    f"Parametro 1: {p1_nome}",
                                    value=10.0,
                                    format="%.2f",
                                    key="t2_p1_val",
                                )
                                valori_fattori.append(v1)
                                dettagli_lista.append(f"{p1_nome}: {v1}")

                            if p2_nome and p2_nome.lower() != "nan" and p2_nome.strip() != "":
                                v2 = st.number_input(
                                    f"Parametro 2: {p2_nome}",
                                    value=15.0,
                                    format="%.2f",
                                    key="t2_p2_val",
                                )
                                valori_fattori.append(v2)
                                dettagli_lista.append(f"{p2_nome}: {v2}")

                            if p3_nome and p3_nome.lower() != "nan" and p3_nome.strip() != "":
                                v3 = st.number_input(
                                    f"Parametro 3: {p3_nome}",
                                    value=0.0,
                                    format="%.2f",
                                    key="t2_p3_val",
                                )
                                valori_fattori.append(v3)
                                dettagli_lista.append(f"{p3_nome}: {v3}")

                            if len(valori_fattori) >= 2 and valori_fattori[1] > 0:
                                if "%" in str(kpi_row["Unita_Misura"]):
                                    valore_calcolato = round(
                                        (valori_fattori[0] / valori_fattori[1]) * 100, 2
                                    )
                                else:
                                    valore_calcolato = round(
                                        valori_fattori[0] / valori_fattori[1], 2
                                    )
                            elif len(valori_fattori) == 1:
                                valore_calcolato = valori_fattori[0]
                            else:
                                valore_calcolato = st.number_input(
                                    "Valore Misurato", value=0.0, format="%.2f", key="t2_fallback_v"
                                )
                                dettagli_lista.append(f"Valore Assoluto: {valore_calcolato}")

                        dettaglio_fattori_str = " / ".join(dettagli_lista)

                        st.metric(
                            "Valore Risultante Calcolato",
                            f"{valore_calcolato} {kpi_row['Unita_Misura']}",
                        )

                        with col_f2:
                            data_monitoraggio = st.date_input(
                                "Data monitoraggio", value=oggi, key="t2_data_mon_input"
                            )

                            if "Settimanale" in cadenza_corrente:
                                gg_agg = 7
                            elif "Bisettimanale" in cadenza_corrente:
                                gg_agg = 14
                            elif "Mensile" in cadenza_corrente:
                                gg_agg = 30
                            elif "Trimestrale" in cadenza_corrente:
                                gg_agg = 90
                            elif "Semestre" in cadenza_corrente:
                                gg_agg = 180
                            else:
                                gg_agg = 365

                            prossima_scadenza_calc = data_monitoraggio + timedelta(days=gg_agg)
                            st.write(
                                f"📅 **Cadenza:** {cadenza_corrente} -> Nuova scadenza stimata:"
                                f" **{prossima_scadenza_calc.strftime('%d/%m/%Y')}**"
                            )

                        if st.button(
                            "💾 Registra Nuovo Monitoraggio nello Storico",
                            key="t2_btn_registra_storico",
                        ):
                            nuova_riga_storico = pd.DataFrame([{
                                "Data_Monitoraggio": data_monitoraggio,
                                "ID_KPI": id_kpi_corrente,
                                "Nome_KPI": kpi_sel_upd,
                                "Valore_Registrato": valore_calcolato,
                                "Dettaglio_Fattori": dettaglio_fattori_str,
                            }])
                            st.session_state.t2_storico_misure = pd.concat(
                                [st.session_state.t2_storico_misure, nuova_riga_storico],
                                ignore_index=True,
                            )
                            # Salva permanentemente su disco
                            st.session_state.t2_storico_misure.to_csv(path_storico_pers, index=False, encoding='utf-8')

                            st.success(
                                "Nuovo monitoraggio aggiunto con successo nello storico!"
                            )
                            st.rerun()

                        st.markdown("---")
                        st.markdown("#### Tabella Storico Monitoraggi")
                        df_storico_display = st.session_state.t2_storico_misure.sort_values(
                            by="Data_Monitoraggio", ascending=False
                        )
                        st.dataframe(df_storico_display, use_container_width=True)

                        # Esportazione Tabella Storico Monitoraggi (Sottosezione Tab2 - Storico)
                        st.markdown("##### 📥 Esportazione Tabella Storico Monitoraggi")
                        nome_file_storico = f"{data_str}_Tab2-StoricoMisure.csv"
                        path_storico = os.path.join("KPI", nome_file_storico)
                        csv_storico = df_storico_display.to_csv(index=False).encode('utf-8')

                        col_s1, col_s2 = st.columns(2)
                        with col_s1:
                            if st.button("📁 Salva in cartella 'KPI' (Storico Misure)", key="btn_save_folder_storico", use_container_width=True):
                                df_storico_display.to_csv(path_storico, index=False, encoding='utf-8')
                                st.success(f"Salvato con successo in: {path_storico}")
                        with col_s2:
                            st.download_button(
                                label="⬇️ Scarica CSV Storico Misure",
                                data=csv_storico,
                                file_name=nome_file_storico,
                                mime="text/csv",
                                key="dl_btn_storico",
                                use_container_width=True
                            )
                            
                        st.markdown("---")
                        
                        # --- ELIMINAZIONE PROTETTA REGISTRO STORICO ---
                        st.markdown("##### 🗑️ Eliminazione Registro di Monitoraggio (Richiede Password)")
                        if not st.session_state.t2_storico_misure.empty:
                            storico_temp_del = st.session_state.t2_storico_misure.copy()
                            storico_temp_del["Etichetta_Riga"] = (
                                storico_temp_del["Data_Monitoraggio"].astype(str)
                                + " | "
                                + storico_temp_del["Nome_KPI"]
                                + " | Valore: "
                                + storico_temp_del["Valore_Registrato"].astype(str)
                            )
                            
                            with st.form("form_elimina_registro_storico"):
                                riga_selezionata_del = st.selectbox(
                                    "Seleziona il registro di monitoraggio da eliminare",
                                    storico_temp_del["Etichetta_Riga"].tolist(),
                                    key="sel_storico_del"
                                )
                                pwd_del_storico = st.text_input("Inserisci la password di sezione per confermare l'eliminazione", type="password", key="pwd_del_storico_field")
                                
                                if st.form_submit_button("Conferma ed Elimina Registro", use_container_width=True):
                                    if pwd_del_storico == "hse2026":
                                        idx_da_rimuovere = storico_temp_del[
                                            storico_temp_del["Etichetta_Riga"] == riga_selezionata_del
                                        ].index
                                        st.session_state.t2_storico_misure = st.session_state.t2_storico_misure.drop(idx_da_rimuovere).reset_index(drop=True)
                                        # Aggiorna il file persistente su disco
                                        st.session_state.t2_storico_misure.to_csv(path_storico_pers, index=False, encoding='utf-8')

                                        st.success("Registro di monitoraggio eliminato con successo!")
                                        st.rerun()
                                    else:
                                        st.error("Password errata. Impossibile procedere con l'eliminazione.")
                        else:
                            st.info("Nessun registro di monitoraggio disponibile per l'eliminazione.")
                    else:
                        st.info("Nessun KPI disponibile.")

                # --- SOTTO-TAB 3: AGGIUNGI NUOVO KPI & CONFIGURAZIONE ---
                else:
                    st.markdown("#### 🛠️ Aggiungi un Nuovo KPI (Sincronizzato con Tab 1)")
                    st.markdown(
                        "I KPI inseriti qui saranno aggiunti all'archivio comune e visibili"
                        " in entrambe le tab."
                    )

                    with st.form("t2_form_nuovo_kpi"):
                        c_a1, c_a2 = st.columns(2)
                        with c_a1:
                            nid = st.text_input("ID KPI (es. NM-04)", value="NM-04", key="t2_nid")
                            nnome = st.text_input(
                                "Nome KPI", value="Indice Chiusura Prescrizioni", key="t2_nnome"
                            )
                            ncat = st.selectbox(
                                "Categoria",
                                [
                                    "Volume",
                                    "Efficacia",
                                    "Reattività",
                                    "Cultura HSE",
                                    "Prevenzione",
                                    "Ambiente",
                                ],
                                key="t2_ncat",
                            )
                            nmis = st.selectbox(
                                "Unità di Misura", ["%", "Numero", "Giorni"], key="t2_nmis"
                            )
                            ncad = st.selectbox(
                                "Cadenza",
                                ["Settimanale", "Bisettimanale", "Mensile", "Trimestrale", "Semestrale", "Annuale"],
                                key="t2_ncad",
                            )
                        with c_a2:
                            nform = st.text_input(
                                "Formula (es. (Parametro_1 / Parametro_2) * 100)",
                                value="(Parametro_1 / Parametro_2) * 100",
                                key="t2_nform",
                            )
                            np1 = st.text_input(
                                "Parametro 1 (Fattore)",
                                value="Valore Numeratore",
                                key="t2_np1",
                            )
                            np2 = st.text_input(
                                "Parametro 2 (Fattore)",
                                value="Valore Denominatore",
                                key="t2_np2",
                            )
                            np3 = st.text_input(
                                "Parametro 3 (Opzionale)", value="", key="t2_np3"
                            )

                        # Gestione Baseline (Sì o No con Text Area dedicata)
                        nbase_choice = st.selectbox(
                            "Baseline Esistente", ["No", "Sì"], key="t2_nbase_choice"
                        )
                        nbaseline_text = st.text_area(
                            "Valore o Descrizione della Baseline",
                            value=(
                                "Inserisci qui i dettagli della baseline o target iniziale..."
                                if nbase_choice == "Sì"
                                else "Nessuna baseline storica precedente."
                            ),
                            key="t2_nbaseline_textarea",
                        )

                        if st.form_submit_button("Crea e Sincronizza Nuovo KPI"):
                            if nid in st.session_state.t1_kpi_definitions["ID_KPI"].values:
                                st.error(f"L'ID KPI '{nid}' esiste già.")
                            else:
                                nuova_riga_kpi = pd.DataFrame([{
                                    "ID_KPI": nid,
                                    "Nome_KPI": nnome,
                                    "Categoria": ncat,
                                    "Unita_Misura": nmis,
                                    "Cadenza": ncad,
                                    "Formula": nform,
                                    "Parametro_1": np1,
                                    "Parametro_2": np2,
                                    "Parametro_3": np3,
                                    "Baseline": nbaseline_text,
                                }])
                                st.session_state.t1_kpi_definitions = pd.concat(
                                    [st.session_state.t1_kpi_definitions, nuova_riga_kpi],
                                    ignore_index=True,
                                )
                                # Salva permanentemente su disco
                                st.session_state.t1_kpi_definitions.to_csv(path_master_pers, index=False, encoding='utf-8')

                                st.success(
                                    "Nuovo KPI creato e riconosciuto correttamente in entrambe le"
                                    " tab!"
                                )
                        st.rerun()
        render_sezione_kpi()
# ==================================================================
# --- SEZIONE 8: Piano di Miglioramento ---
# ==================================================================
if nav == "Piano Miglioramento":
    st.header("Piano di Miglioramento HSE")
    
    # Sottosezioni della Sezione 8
    sotto_sec = st.radio(
        "Seleziona Sottosezione", 
        ["Documentazione di Riferimento", "Valutazione del rischio", "Azioni Piano di Miglioramento"], 
        horizontal=True, 
        key="radio_sotto_sec_8"
    )
    
    if sotto_sec == "Documentazione di Riferimento":
        st.subheader("Consultazione Istruzione Operativa")
        st.markdown("Consulta o scarica il documento PDF ufficiale relativo al piano di miglioramento.")
        
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base_dir = os.getcwd()
            
        file_pdf_path = os.path.join(base_dir, "Piano_Miglioramento", "Istruzione_Piano_Miglioramento.pdf")
        
        if not os.path.exists(file_pdf_path):
            percorsi_alternativi = [
                os.path.join("Piano_Miglioramento", "Istruzione_Piano_Miglioramento.pdf"),
                os.path.join("APP HSE", "Piano_Miglioramento", "Istruzione_Piano_Miglioramento.pdf")
            ]
            for p in percorsi_alternativi:
                if os.path.exists(p):
                    file_pdf_path = os.path.abspath(p)
                    break
        
        if os.path.exists(file_pdf_path):
            with open(file_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            
            st.download_button(
                label="📥 Scarica / Apri Istruzione Piano di Miglioramento (.pdf)",
                data=pdf_bytes,
                file_name="Istruzione_Piano_Miglioramento.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
            try:
                import base64
                base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)
            except Exception as e:
                st.info("Utilizza il pulsante di download sopra per consultare il documento nel lettore PDF del tuo computer.")
        else:
            st.error("Il file PDF 'Istruzione_Piano_Miglioramento.pdf' non è stato trovato.")
                
    elif sotto_sec == "Valutazione del rischio":
        st.subheader("Valutazione del Rischio e Collegamento Analisi")
        st.markdown("Inserisci la password per accedere all'area di collegamento e alla tabella dinamica di valutazione del rischio.")
        
        # Gestione autenticazione sottosezione
        if "auth_val_rischio" not in st.session_state:
            st.session_state.auth_val_rischio = False
            
        if not st.session_state.auth_val_rischio:
            pwd_vr = st.text_input("Inserisci la password per la Valutazione del Rischio", type="password", key="pwd_vr_input")
            if st.button("Verifica Password", use_container_width=True):
                if pwd_vr == "hse2026":
                    st.session_state.auth_val_rischio = True
                    st.success("Accesso autorizzato!")
                    st.rerun()
                else:
                    st.error("Password errata.")
        
        if st.session_state.auth_val_rischio:
            st.markdown("---")
            st.markdown("### 🔗 Area di Collegamento con 'Analisi - Fase 2'")
            st.markdown("Seleziona e visualizza i dettagli del report di analisi e il Near Miss collegato prima di procedere con la valutazione.")

            # Caricamento e unione dati per dropdown
            opzioni = ["Nessuna (Nuova analisi)"]
                                
            # Leggi Near Miss
            if os.path.exists(FILE_NEAR_MISS):
                df_nm = pd.read_csv(FILE_NEAR_MISS, sep=";")
                for idx, r in df_nm.iterrows():
                    opzioni.append(f"NM | {r.get('Data Segnalazione', 'N/D')} | {r.get('Tipo Evento', 'Evento')}")
                                
            # Leggi Analisi già fatte
            if os.path.exists(FILE_ANALISI_NM):
                df_an = pd.read_csv(FILE_ANALISI_NM, sep=";")
                for idx, r in df_an.iterrows():
                    opzioni.append(f"AN | {r.get('Data Analisi', 'N/D')} | Collegamento: {r.get('Segnalazione Collegata', 'Analisi')}")
                                
            scelta_rif = st.selectbox("Seleziona evento/analisi collegata:", opzioni)
            
            
            file_analisi_target = globals().get("FILE_ANALISI_NM", "analisi_fase_2.csv")
            file_near_miss_target = globals().get("FILE_NEAR_MISS", "near_miss.csv")
            
            df_an_f2 = pd.read_csv(file_analisi_target, sep=";") if os.path.exists(file_analisi_target) else pd.DataFrame()
            df_nm = pd.read_csv(file_near_miss_target, sep=";") if os.path.exists(file_near_miss_target) else pd.DataFrame()
            
            if not df_an_f2.empty:
                opzioni_analisi = [f"Report #{idx+1} (Data/Ora: {df_an_f2.iloc[idx].get('Data/Ora', 'N/D')})" for idx in range(len(df_an_f2))]
                scelta_analisi = st.selectbox("Seleziona il Report di Analisi - Fase 2 da consultare/collegare", opzioni_analisi, key="seleziona_report_analisi_s8")
                
                # Salviamo la scelta in session_state per renderla disponibile anche nell'altra sottosezione
                st.session_state.report_analisi_selezionato = scelta_analisi
                
                idx_selezionato = opzioni_analisi.index(scelta_analisi)
                riga_selezionata = df_an_f2.iloc[idx_selezionato]
                
                col_info1, col_info2 = st.columns(2)
                with col_info1:
                    st.info(f"**Data e Ora Elaborazione Report:** {riga_selezionata.get('Data/Ora', 'Non disponibile')}")
                with col_info2:
                    st.info(f"**Analisi Segnalazione Near Miss:** {riga_selezionata.get('Analisi RSPP', riga_selezionata.iloc[1] if len(riga_selezionata) > 1 else 'N/D')}")
                
                if "ID Near Miss" in riga_selezionata and not df_nm.empty:
                    id_nm = riga_selezionata["ID Near Miss"]
                    st.markdown(f"**Dettaglio Near Miss Collegato (ID: {id_nm}):**")
                    st.write(riga_selezionata.to_dict())
            else:
                st.warning("Nessun report di 'Analisi - Fase 2' trovato nei sistemi operativi.")
                st.session_state.report_analisi_selezionato = "Report_Generico"
                
            st.markdown("---")
            st.markdown("### Tabella Dinamica di Valutazione del Rischio")
            st.markdown("I parametri numerici (**Probabilità**, **Gravità**, **Sensibilità**, **Controllo**) accettano esclusivamente numeri interi compresi tra **1 e 3**.")
            
            # Inizializzazione DataFrame Valutazione Rischio in session_state se non presente
            if "df_vr_session" not in st.session_state:
                dir_pm_path = os.path.join("APP HSE", "Piano_Miglioramento")
                os.makedirs(dir_pm_path, exist_ok=True)
                file_excel_vr = os.path.join(dir_pm_path, "valutazione_rischio.xlsx")
                if os.path.exists(file_excel_vr):
                    st.session_state.df_vr_session = pd.read_excel(file_excel_vr)
                else:
                    st.session_state.df_vr_session = pd.DataFrame(columns=[
                        "Rischio", "Probabilità", "Gravità", "Sensibilità", "Controllo", "Significatività"
                    ])
                    st.session_state.df_vr_session.loc[0] = ["Esempio rischio infortuni", 2, 2, 2, 2, 4]
                
            for col in ["Probabilità", "Gravità", "Sensibilità", "Controllo"]:
                if col in st.session_state.df_vr_session.columns:
                    st.session_state.df_vr_session[col] = pd.to_numeric(st.session_state.df_vr_session[col], errors='coerce').fillna(1).astype(int)
            
            st.session_state.df_vr_session["Significatività"] = ((st.session_state.df_vr_session["Gravità"] * st.session_state.df_vr_session["Probabilità"] * st.session_state.df_vr_session["Sensibilità"]) / st.session_state.df_vr_session["Controllo"].replace(0, 1)).round().astype(int)
            
            edited_df = st.data_editor(
                st.session_state.df_vr_session,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Rischio": st.column_config.TextColumn("Rischio", required=True),
                    "Probabilità": st.column_config.NumberColumn("Probabilità", min_value=1, max_value=3, step=1, format="%d"),
                    "Gravità": st.column_config.NumberColumn("Gravità", min_value=1, max_value=3, step=1, format="%d"),
                    "Sensibilità": st.column_config.NumberColumn("Sensibilità", min_value=1, max_value=3, step=1, format="%d"),
                    "Controllo": st.column_config.NumberColumn("Controllo", min_value=1, max_value=3, step=1, format="%d"),
                    "Significatività": st.column_config.NumberColumn("Significatività", disabled=True, format="%d")
                },
                key="editor_valutazione_rischio"
            )
            
            edited_df["Significatività"] = ((edited_df["Probabilità"] * edited_df["Gravità"] * edited_df["Sensibilità"]) / edited_df["Controllo"].replace(0, 1)).round().astype(int)
            st.session_state.df_vr_session = edited_df
            
            st.markdown("### Vista Consolidata con Indicatori di Rischio Colorati")
            st.markdown("Legenda colori Significatività: **Verde** (1-3) | **Giallo** (4-6) | **Rosso** (7-9)")
            
            def color_cells(val):
                try:
                    v = int(val)
                    if 1 <= v <= 3:
                        return 'background-color: #d4edda; color: #155724; font-weight: bold;'
                    elif 4 <= v <= 6:
                        return 'background-color: #fff3cd; color: #856404; font-weight: bold;'
                    elif 7 <= v <= 9:
                        return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                except:
                    pass
                return ''

            styled_df = edited_df.style.map(color_cells, subset=["Significatività"])
            st.dataframe(styled_df, use_container_width=True)

            # --- Salvataggio e Download della Tabella Dinamica (Senza Colori) ---
            if 'edited_df' in locals() and not edited_df.empty:
                try:
                    base_dir = os.path.dirname(os.path.abspath(__file__))
                except NameError:
                    base_dir = os.getcwd()
        
            # Percorso della cartella: APP HSE -> Piano_Miglioramento -> Valutazione Rischio 
            target_val_rischio_dir = os.path.join(base_dir, "APP HSE", "Piano_Miglioramento", "Valutazione Rischio ")
            if not os.path.exists(os.path.join(base_dir, "APP HSE")):
                target_val_rischio_dir = os.path.join(base_dir, "Piano_Miglioramento", "Valutazione Rischio ")
        
            os.makedirs(target_val_rischio_dir, exist_ok=True)
    
            # Generazione nome file dinamico con timestamp per farlo variare a ogni salvataggio
            from datetime import datetime
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name_dinamico = f"Report_di_Analisi_Fase_2_da_consultare_collegare_Valutazione_Rischio_Dinamica_{timestamp_str}.csv"
            full_path_dinamico = os.path.join(target_val_rischio_dir, file_name_dinamico)
    
            # Esportazione in CSV standard (senza formattazione o colori)
            csv_data_dinamico = edited_df.to_csv(index=False, sep=";")
    
            # Salvataggio automatico nella cartella dedicata
            with open(full_path_dinamico, "w", encoding="utf-8") as f:
                f.write(csv_data_dinamico)
        
            # Pulsante Streamlit per il download (corretto)
            st.download_button(
                label="📥 Scarica Tabella Dinamica in formato CSV (.csv)",
                data=csv_data_dinamico,
                file_name=file_name_dinamico,
                mime="text/csv",
                use_container_width=True
            )
            st.success(f"File salvato automaticamente nella cartella: `{target_val_rischio_dir}`")
        else:
            st.info("Compila o genera la tabella dinamica per abilitare il download e il salvataggio del file.")
            
                
    elif sotto_sec == "Azioni Piano di Miglioramento":
        st.subheader("Azioni Piano di Miglioramento")
        st.markdown("Gestione delle azioni intraprese, delle tipologie di intervento e del relativo follow-up.")
        
        # Caricamento e unione dati per dropdown
        opzioni = ["Nessuna (Nuova analisi)"]
                    
        # Leggi Near Miss
        if os.path.exists(FILE_NEAR_MISS):
            df_nm = pd.read_csv(FILE_NEAR_MISS, sep=";")
            for idx, r in df_nm.iterrows():
                opzioni.append(f"NM | {r.get('Data Segnalazione', 'N/D')} | {r.get('Tipo Evento', 'Evento')}")
                    
        # Leggi Analisi già fatte
        if os.path.exists(FILE_ANALISI_NM):
            df_an = pd.read_csv(FILE_ANALISI_NM, sep=";")
            for idx, r in df_an.iterrows():
                opzioni.append(f"AN | {r.get('Data Analisi', 'N/D')} | Collegamento: {r.get('Segnalazione Collegata', 'Analisi')}")
                    
        scelta_rif = st.selectbox("Seleziona evento/analisi collegata:", opzioni)

        st.markdown("---")
        st.markdown("### 1. AZIONI INTRAPRESE")
        
        st.markdown("#### Azioni immediate di rimedio")
        st.markdown("*Confronto con campo “Valutazioni / azioni / proposte di miglioramento” in modulo segnalazione*")
        
        if "azioni_immediate_txt" not in st.session_state:
            st.session_state.azioni_immediate_txt = ""
            
        azioni_immediate = st.text_area(
            "Descrivi le azioni immediate di rimedio", 
            value=st.session_state.azioni_immediate_txt,
            key="txt_azioni_immediate_rimedio_input", 
            placeholder="Inserisci la descrizione delle azioni immediate..."
        )
        st.session_state.azioni_immediate_txt = azioni_immediate
        
        st.markdown("#### Azioni di miglioramento (correttive, preventive) - Tipologia intervento")
        st.markdown("Utilizza la tabella sottostante per aggiungere le righe e selezionare la tipologia di intervento e la descrizione.")
        
        lista_tipologie = [
            "Tecnico",
            "Formazione / Addestramento",
            "Informazione / Comunicazione / Partecipazione",
            "Definizione / revisione delle procedure e istruzioni lavorative",
            "Verifica applicazione procedure / istruzioni / comportamenti",
            "Altro (specificare)"
        ]
        
        if "df_tipologie_session" not in st.session_state:
            st.session_state.df_tipologie_session = pd.DataFrame(columns=["Tipologia di Intervento", "Descrizione"])
            st.session_state.df_tipologie_session.loc[0] = ["Tecnico", "Esempio intervento tecnico correttivo"]
            
        edited_tipologie = st.data_editor(
            st.session_state.df_tipologie_session,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Tipologia di Intervento": st.column_config.SelectboxColumn(
                    "Tipologia di Intervento",
                    options=lista_tipologie,
                    required=True
                ),
                "Descrizione": st.column_config.TextColumn("Descrizione", required=True)
            },
            key="editor_tipologie_interventi"
        )
        st.session_state.df_tipologie_session = edited_tipologie
        
        st.markdown("---")
        st.markdown("### 2. FOLLOW UP AZIONI INTRAPRESE")
        st.markdown("Tabella di monitoraggio, pianificazione e verifica delle azioni correttive e preventive.")
        
        if "df_followup_session" not in st.session_state:
            st.session_state.df_followup_session = pd.DataFrame(columns=[
                "Azione / Descrizione",
                "Responsabile attuazione",
                "Accountable attuazione",
                "Entro il",
                "Firma presa in carico",
                "Data attuazione",
                "Verifica attuazione",
                "Data e firma"
            ])
            st.session_state.df_followup_session.loc[0] = [
                "1° - Esempio azione correttiva",
                "Mario Rossi",
                "Luigi Verdi",
                "2026-12-31",
                "Presa in carico",
                "2026-06-15",
                "Verificato e conforme",
                "2026-06-20 - M. Rossi"
            ]
            
        edited_followup = st.data_editor(
            st.session_state.df_followup_session,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Azione / Descrizione": st.column_config.TextColumn("Azione / Descrizione", required=True),
                "Responsabile attuazione": st.column_config.TextColumn("Responsabile attuazione"),
                "Accountable attuazione": st.column_config.TextColumn("Accountable attuazione"),
                "Entro il": st.column_config.TextColumn("Entro il (Scadenza)"),
                "Firma presa in carico": st.column_config.TextColumn("Firma presa in carico"),
                "Data attuazione": st.column_config.TextColumn("Data attuazione"),
                "Verifica attuazione": st.column_config.TextColumn("Verifica attuazione"),
                "Data e firma": st.column_config.TextColumn("Data e firma")
            },
            key="editor_follow_up_azioni"
        )
        st.session_state.df_followup_session = edited_followup
        
        st.markdown("---")
        # UNICO PULSANTE DI AGGIORNAMENTO / SALVATAGGIO
        if st.button("💾 Aggiorna e Salva Tutti i Dati del Piano di Miglioramento", use_container_width=True):
            # Cartella esistente Piano_Miglioramento dentro APP HSE
            dir_dest = "Piano_Miglioramento"
            os.makedirs(dir_dest, exist_ok=True)
            
            # Generazione nome file basato sul pattern: [evento/analisi collegata]_Piano Miglioramento
            report_sel = st.session_state.get("report_analisi_selezionato", "Report_Generico")
            nome_file_pulito = "".join([c if c.isalnum() or c in (' ', '_', '-') else '_' for c in report_sel]).strip()
            nome_file_xlsx = f"{nome_file_pulito}_Piano Miglioramento.xlsx"
            percorso_completo = os.path.join(dir_dest, nome_file_xlsx)
            
            # Preparazione del DataFrame per Azioni Intraprese (incluse le azioni immediate di rimedio)
            df_azioni_intraprese = pd.DataFrame({
                "Tipologia Contenuto": ["Azioni immediate di rimedio"],
                "Descrizione / Dettaglio": [st.session_state.get("azioni_immediate_txt", "")]
            })
            
            # Scrittura dei dati in un unico file Excel con più fogli distinti
            try:
                with pd.ExcelWriter(percorso_completo, engine='openpyxl') as writer:

                    # 1. Foglio: Valutazione del rischio
                    df_vr_to_save = st.session_state.get("df_vr_session", pd.DataFrame())
                    df_vr_to_save.to_excel(writer, sheet_name="Valutazione del rischio", index=False)
                    
                    # 2. Foglio: Azioni intraprese (include Azioni immediate di rimedio)
                    df_azioni_intraprese.to_excel(writer, sheet_name="Azioni intraprese", index=False)
                    
                    # 3. Foglio: Azioni di miglioramento - Tipologia intervento
                    df_tip_to_save = st.session_state.get("df_tipologie_session", pd.DataFrame())
                    df_tip_to_save.to_excel(writer, sheet_name="Tipologie intervento", index=False)
                    
                    # 4. Foglio: Follow up azioni intraprese
                    df_follow_to_save = st.session_state.get("df_followup_session", pd.DataFrame())
                    df_follow_to_save.to_excel(writer, sheet_name="Follow-up azioni", index=False)
                
                st.success(f"Dati salvati e aggiornati con successo nella cartella esistente:\n`{percorso_completo}`")
            except Exception as e:
                st.error(f"Errore durante il salvataggio del file: {e}")

# ==================================================================
# --- SEZIONE 9: Stima Costo Economico ---
# ==================================================================
if nav == "Stima Costo Economico":
    st.header("Stima Costo Economico del Near Miss")
    
    # Gestione autenticazione per la Sezione 9
    if "auth_stima_economico" not in st.session_state:
        st.session_state.auth_stima_economico = False
        
    if not st.session_state.auth_stima_economico:
        st.markdown("Inserisci la password per accedere all'area di stima del costo economico.")
        pwd_sec9 = st.text_input("Password Sezione 9", type="password", key="pwd_sec9_input")
        if st.button("Verifica Password", use_container_width=True, key="btn_verify_pwd_sec9"):
            if pwd_sec9 == "hse2026":
                st.session_state.auth_stima_economico = True
                st.success("Accesso autorizzato!")
                st.rerun()
            else:
                st.error("Password errata.")
    
    if st.session_state.auth_stima_economico:
        # Sottosezioni della Sezione 9
        sotto_sec_9 = st.radio(
            "Seleziona Sottosezione", 
            ["Documentazione di Riferimento", "Calcolo economico NM"], 
            horizontal=True, 
            key="radio_sotto_sec_9"
        )
        
        if sotto_sec_9 == "Documentazione di Riferimento":
            st.subheader("Consultazione Documento Stima Economica")
            st.markdown("Consulta o scarica il documento PDF ufficiale relativo alla stima economica del near miss.")
            
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            except NameError:
                base_dir = os.getcwd()
                
            file_pdf_path = os.path.join(base_dir, "stima_economica", "Stima Economica del Near Miss.pdf")
            
            if not os.path.exists(file_pdf_path):
                percorsi_alternativi = [
                    os.path.join("stima_economica", "Stima Economica del Near Miss.pdf"),
                    os.path.join("APP HSE", "stima_economica", "Stima Economica del Near Miss.pdf")
                ]
                for p in percorsi_alternativi:
                    if os.path.exists(p):
                        file_pdf_path = os.path.abspath(p)
                        break
            
            if os.path.exists(file_pdf_path):
                with open(file_pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                
                st.download_button(
                    label="📥 Scarica / Apri Documento Stima Economica (.pdf)",
                    data=pdf_bytes,
                    file_name="Stima Economica del Near Miss.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                try:
                    import base64
                    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                except Exception as e:
                    st.info("Utilizza il pulsante di download sopra per consultare il documento nel lettore PDF del tuo computer.")
            else:
                st.error("Il file PDF 'Stima Economica del Near Miss.pdf' non è stato trovato nella cartella 'stima_economica'.")
                
        elif sotto_sec_9 == "Calcolo economico NM":
            st.subheader("Calcolo economico NM - Tabella Dinamica e Parametri")
            st.markdown("Configura i parametri di inquadramento (Manodopera) e il fatturato dell'anno precedente (Vendite e Reputazione) per i calcoli automatici.")

            # Caricamento e unione dati per dropdown
            opzioni = ["Nessuna (Nuova analisi)"]
                                            
            # Leggi Near Miss
            if os.path.exists(FILE_NEAR_MISS):
                df_nm = pd.read_csv(FILE_NEAR_MISS, sep=";")
                for idx, r in df_nm.iterrows():
                    opzioni.append(f"NM | {r.get('Data Segnalazione', 'N/D')} | {r.get('Tipo Evento', 'Evento')}")
                                            
            # Leggi Analisi già fatte
            if os.path.exists(FILE_ANALISI_NM):
                df_an = pd.read_csv(FILE_ANALISI_NM, sep=";")
                for idx, r in df_an.iterrows():
                    opzioni.append(f"AN | {r.get('Data Analisi', 'N/D')} | Collegamento: {r.get('Segnalazione Collegata', 'Analisi')}")
                                            
            scelta_rif = st.selectbox("Seleziona evento/analisi collegata:", opzioni)
            
            # Parametri specifici per la gestione della Manodopera
            st.markdown("#### Configurazione Condizioni Area Manodopera")
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                inquadramento_mansione = st.selectbox(
                    "Inquadramento", 
                    ["Q", "AS", "A", "B1", "B2S", "B2", "C1S", "C1", "C2", "C3", "D1", "D2", "E"], 
                    key="inquadramento_sel"
                )
            with col_m2:
                inf_1gg = st.selectbox("Infortunio sul lavoro 1gg", ["No", "Sì"], key="inf_1gg_sel")
                inf_2_4gg = st.selectbox("Infortunio sul lavoro 2<4gg", ["No", "Sì"], key="inf_2_4gg_sel")
            with col_m3:
                inf_5gg = st.selectbox("Infortunio sul lavoro >=5gg", ["No", "Sì"], key="inf_5gg_sel")
            
            inquadramento_mapping = {
                "Q": 2882.91, "AS": 2873.55, "A": 2521.80, "B1": 2292.49, "B2S": 2234.61, "B2": 2159.89,
                "C1S": 2034.87, "C1": 1960.11, "C2": 1826.94, "C3": 1732.11, "D1": 1656.25, "D2": 1561.08, "E": 1456.61
            }
            valore_inquadramento = inquadramento_mapping.get(inquadramento_mansione, 2159.89)
            
            perc_indennita_val = 0.0
            ral_mansione_calc = 0.0
            
            if inf_1gg == "Sì":
                perc_indennita_val = 100.0
                ral_mansione_calc = (valore_inquadramento / 30.0) * 3.0 * (perc_indennita_val / 100.0)
            elif inf_2_4gg == "Sì":
                perc_indennita_val = 40.0
                ral_mansione_calc = ((valore_inquadramento / 30.0) * 3.0 * (perc_indennita_val / 100.0)) + (valore_inquadramento / 30.0)
            elif inf_5gg == "Sì":
                perc_indennita_val = 25.0
                ral_mansione_calc = ((valore_inquadramento / 30.0) * 3.0 * (perc_indennita_val / 100.0)) + (valore_inquadramento / 30.0)
            else:
                ral_mansione_calc = valore_inquadramento

            # Parametri specifici per Vendite e Reputazione (Fatturato Anno Precedente)
            st.markdown("#### Configurazione Condizioni Aree Vendite e Reputazione")
            col_v1, col_v2 = st.columns(2)
            with col_v1:
                fatturato_vendite = st.number_input("Fatturato anno precedente (Vendite) [€]", value=1000000.0, step=10000.0, key="fat_vendite_input")
                vendite_1pct = st.selectbox("Diminuzione vendite 1% (Vendite)", ["No", "Sì"], key="vendite_1_sel")
                vendite_5pct = st.selectbox("Diminuzione vendite 5% (Vendite)", ["No", "Sì"], key="vendite_5_sel")
            with col_v2:
                fatturato_reputazione = st.number_input("Fatturato anno precedente (Reputazione) [€]", value=1000000.0, step=10000.0, key="fat_reputazione_input")
                rep_10pct = st.selectbox("Diminuzione fatturato 10% (Reputazione)", ["No", "Sì"], key="rep_10_sel")
                rep_15pct = st.selectbox("Diminuzione fatturato 15% (Reputazione)", ["No", "Sì"], key="rep_15_sel")

            # Calcoli condizionali Vendite
            val_vendite_1 = fatturato_vendite * (1.0 / 100.0) if vendite_1pct == "Sì" else 0.0
            val_vendite_5 = fatturato_vendite * (5.0 / 100.0) if vendite_5pct == "Sì" else 0.0

            # Calcoli condizionali Reputazione
            val_rep_10 = fatturato_reputazione * (10.0 / 100.0) if rep_10pct == "Sì" else 0.0
            val_rep_15 = fatturato_reputazione * (15.0 / 100.0) if rep_15pct == "Sì" else 0.0

            # Inizializzazione DataFrame Tabella Dinamica in session_state
            if "df_calcolo_economico_nm" not in st.session_state:
                data_nm = [
                    # Macchinari
                    ["Macchinari", "Ricambio attrezzatura", 0.0],
                    ["Macchinari", "Ricambio pezzo macchinario", 0.0],
                    ["Macchinari", "Cambio attrezzatura", 0.0],
                    ["Macchinari", "Cambio Macchinario", 0.0],
                    # Manodopera
                    ["Manodopera", "Infortunio sul lavoro 1gg", 0.0],
                    ["Manodopera", "Infortunio sul lavoro 2<4gg", 0.0],
                    ["Manodopera", "Infortunio sul lavoro >=5gg", 0.0],
                    ["Manodopera", "Percentuale di indennità", 0.0],
                    ["Manodopera", "RAL per mansione", 0.0],
                    ["Manodopera", "Formazione neoassunto", 0.0],
                    ["Manodopera", "Addestramento personale", 0.0],
                    # Materiali
                    ["Materiali", "Carta", 0.0],
                    ["Materiali", "Amido", 0.0],
                    ["Materiali", "DPI", 0.0],
                    ["Materiali", "Immobiliare", 0.0],
                    ["Materiali", "Vari", 0.0],
                    # Metodo
                    ["Metodo", "Nuova procedura-istruzione", 0.0],
                    ["Metodo", "Periodo adattamento", 0.0],
                    ["Metodo", "Redistribuzione aziendale", 0.0],
                    # Vendite
                    ["Vendite", "Diminuzione delle vendite del 1% rispetto all’anno precedente", 0.0],
                    ["Vendite", "Diminuzione del fatturato del 5% rispetto all’anno precedente", 0.0],
                    # Reputazione
                    ["Reputazione", "Diminuzione del fatturato del 10% rispetto all’anno precedente", 0.0],
                    ["Reputazione", "Diminuzione del fatturato del 15% rispetto all’anno precedente", 0.0],
                    # Sanzioni
                    ["Sanzioni", "Sanzioni amministrative / penali (scaglionate)", 0.0]
                ]
                st.session_state.df_calcolo_economico_nm = pd.DataFrame(data_nm, columns=[
                    "Area d'impatto", "Sottocategoria", "Stima costo (€)"
                ])
            
            df_ce = st.session_state.df_calcolo_economico_nm
            
            # Assegnazione automatica dei valori calcolati nel DataFrame
            df_ce.loc[df_ce["Sottocategoria"] == "Percentuale di indennità", "Stima costo (€)"] = perc_indennita_val
            df_ce.loc[df_ce["Sottocategoria"] == "RAL per mansione", "Stima costo (€)"] = ral_mansione_calc
            df_ce.loc[df_ce["Sottocategoria"] == "Diminuzione delle vendite del 1% rispetto all’anno precedente", "Stima costo (€)"] = val_vendite_1
            df_ce.loc[df_ce["Sottocategoria"] == "Diminuzione del fatturato del 5% rispetto all’anno precedente", "Stima costo (€)"] = val_vendite_5
            df_ce.loc[df_ce["Sottocategoria"] == "Diminuzione del fatturato del 10% rispetto all’anno precedente", "Stima costo (€)"] = val_rep_10
            df_ce.loc[df_ce["Sottocategoria"] == "Diminuzione del fatturato del 15% rispetto all’anno precedente", "Stima costo (€)"] = val_rep_15
            
            # Editor della tabella dinamica
            edited_ce = st.data_editor(
                df_ce,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Area d'impatto": st.column_config.TextColumn("Area d'impatto", disabled=True),
                    "Sottocategoria": st.column_config.TextColumn("Sottocategoria", disabled=True),
                    "Stima costo (€)": st.column_config.NumberColumn("Valore / Costo (€ o %)", min_value=0.0, step=10.0, format="%.2f")
                },
                key="editor_calcolo_economico_nm"
            )
            
            st.session_state.df_calcolo_economico_nm = edited_ce
            
            # Esclusione della riga "Percentuale di indennità" dal calcolo dei costi economici monetari
            df_costi_monetari = edited_ce[edited_ce["Sottocategoria"] != "Percentuale di indennità"]
            
            # Calcolo automatico del totale generale e per area
            st.markdown("### Riepilogo Costi per Area e Totale Generale")
            
            totale_generale = df_costi_monetari["Stima costo (€)"].sum()
            
            col_tot1, col_tot2 = st.columns(2)
            with col_tot1:
                st.metric(label="💰 STIMA ECONOMICA TOTALE", value=f"€ {totale_generale:,.2f}")
                
            with col_tot2:
                riepilogo_aree = df_costi_monetari.groupby("Area d'impatto")["Stima costo (€)"].sum()
                st.markdown("**Totali parziali per Area (esclusa % indennità):**")
                for area, val in riepilogo_aree.items():
                    st.text(f"- {area}: € {val:,.2f}")
            
            st.markdown("---")
            
            # Preparazione del DataFrame da esportare in CSV con l'evento/analisi collegata inclusa
            df_export = edited_ce.copy()
            df_export.insert(0, "Evento / Analisi Collegata", scelta_rif)
            
            riga_totale = pd.DataFrame([[scelta_rif, "TOTALE GENERALE", "SOMMA TUTTI I COSTI", totale_generale]], columns=df_export.columns)
            df_export = pd.concat([df_export, riga_totale], ignore_index=True)
            
            # Definizione del percorso di salvataggio nella cartella richiesta: APP HSE / Stima_Economica / Report_Stima_Economica
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            except NameError:
                base_dir = os.getcwd()
                
            target_report_dir = os.path.join(base_dir, "Stima_Economica", "Report_Stima_Economica")
            if not os.path.exists(target_report_dir):
                target_report_dir_alt = os.path.join(base_dir, "APP HSE", "Stima_Economica", "Report_Stima_Economica")
                if os.path.exists(os.path.join(base_dir, "APP HSE")) or "APP HSE" in base_dir:
                    target_report_dir = target_report_dir_alt
            
            os.makedirs(target_report_dir, exist_ok=True)
            
            # Generazione del nome file dinamico pulito
            import re
            clean_rif = re.sub(r'[\\/*?:"<>|]', "", scelta_rif)
            clean_rif = clean_rif.replace(" ", "_")
            file_name_export = f"{clean_rif}_Stima_Costo_Economico_NM.csv"
            
            full_file_path = os.path.join(target_report_dir, file_name_export)
            
            # Salvataggio automatico del file CSV nella cartella dedicata
            df_export.to_csv(full_file_path, index=False, sep=";")
            
            csv_data_ce = df_export.to_csv(index=False, sep=";")
            st.download_button(
                label="Scarica Tabella Calcolo Economico NM con Totale in formato CSV (.csv)",
                data=csv_data_ce,
                file_name=file_name_export,
                mime="text/csv",
                use_container_width=True
            )
            st.success(f"File salvato con successo nella cartella: `{target_report_dir}`")
# ==================================================================
# --- SEZIONE: Skill Matrix ---
# ==================================================================
if nav == "Skill Matrix":
    st.header("Skill Matrix - Gestione e Autovalutazione")
    
    # Gestione autenticazione per la Sezione Skill Matrix
    if "auth_skill_matrix" not in st.session_state:
        st.session_state.auth_skill_matrix = False
    sotto_sec_sm = st.radio(
        "Seleziona Sottosezione",
        ["Autovalutazione Skill Matrix", "Skill Matrix"],
        horizontal=True,
        key="radio_sotto_sec_sm"
    )
        
        # Gestione percorsi base
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base_dir = os.getcwd()
            
    skill_matrix_dir = os.path.join(base_dir, "Skill_Matrix")
    if not os.path.exists(skill_matrix_dir):
        skill_matrix_dir_alt = os.path.join(base_dir, "APP HSE", "Skill_Matrix")
        if os.path.exists(os.path.join(base_dir, "APP HSE")) or "APP HSE" in base_dir:
            skill_matrix_dir = skill_matrix_dir_alt
                
    #SOTTOSEZIONE PUBBLICA- AUTOVALUTAZIONE SKILL MATRIX
    if sotto_sec_sm == "Autovalutazione Skill Matrix":
        st.subheader("Autovalutazione Skill Matrix")
        st.markdown("Consulta o scarica il documento PDF di autovalutazione e compila il form sottostante.")
            
        # 1. Download / Visualizzazione PDF "Skill Matrix Autovalutazione.pdf"
        file_pdf_sm = os.path.join(skill_matrix_dir, "Skill Matrix Autovalutazione.pdf")
            
        if os.path.exists(file_pdf_sm):
            with open(file_pdf_sm, "rb") as f:
                pdf_bytes = f.read()
                
            st.download_button(
                label="📥 Scarica / Apri PDF 'Skill Matrix Autovalutazione'",
                data=pdf_bytes,
                file_name="Skill Matrix Autovalutazione.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.warning("Il file PDF 'Skill Matrix Autovalutazione.pdf' non è stato trovato nella cartella 'Skill_Matrix'.")
            
        st.markdown("---")
        st.markdown("#### Form di Autovalutazione")
            
        # 2. Form di compilazione con i nuovi campi richiesti
        with st.form("form_autovalutazione_skill"):
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                nome_utente = st.text_input("Nome")
            with col_f2:
                cognome_utente = st.text_input("Cognome")
                
            col_f3, col_f4 = st.columns(2)
            with col_f3:
                inquadramento_mansione = st.text_input("Inquadramento-Mansione")
            with col_f4:
                ambito_lavorativo = st.selectbox(
                    "Ambito lavorativo", 
                    ["Uffici", "Produzione", "Manutenzione", "Stoccaggio MP", "Trasporto-Logistica", "Commerciale"]
                )
                
            data_compilazione = st.date_input("Data Autovalutazione", value=datetime.today())
                
            st.markdown("##### AUTOVALUTAZIONE DELLA SKILL MATRIX (Punteggio tra 1 e 5)")
                
            q1 = st.slider("Quanto conosci dei processi produttivi della tua mansione?", 1, 5, 3, key="q1")
            q2 = st.slider("Hai un buon rapporto con i colleghi di reparto?", 1, 5, 3, key="q2")
            q3 = st.slider("Valuta le tue capacità di interfacciarti con fornitori e/o clienti", 1, 5, 3, key="q3")
            q4 = st.slider("Quanto conosci dei processi produttivi del cartone e scatole?", 1, 5, 3, key="q4")
            q5 = st.slider("Quanto conosci il processo di pallettizzazione del prodotto?", 1, 5, 3, key="q5")
            q6 = st.slider("Hai competenze legali o tecniche?", 1, 5, 3, key="q6")
            q7 = st.slider("Quanto sei in grado di individuare i rischi-fabbisogni negli ambienti lavorativi", 1, 5, 3, key="q7")
            q8 = st.slider("Valuta la tua capacità d’adattamento", 1, 5, 3, key="q8")
            q9 = st.slider("Valuta le tue capacità comunicative", 1, 5, 3, key="q9")
            q10 = st.slider("Sei una persona precisa nel lavoro che svolge?", 1, 5, 3, key="q10")
            q11 = st.slider("Riesci a persuadere agli altri per svolgere delle attività o fare cambiamenti?", 1, 5, 3, key="q11")
            q12 = st.slider("Hai un’analisi critico del contesto lavorativo? (sai cosa funzione e cosa si potrebbe migliorare)", 1, 5, 3, key="q12")
            q13 = st.slider("Sei a conoscenza delle turnazioni di lavoro, come vengono comunicate e le dinamiche lavorative all’interno di ogni turno?", 1, 5, 3, key="q13")
            q14 = st.slider("Ci sono altre persone sotto la tua responsabilità o supervisione?", 1, 5, 3, key="q14")
                
            submitted_form = st.form_submit_button("Invia e Salva Autovalutazione", use_container_width=True)
                
            if submitted_form:
                 if not nome_utente.strip() or not cognome_utente.strip():
                    st.error("Inserisci obbligatoriamente Nome e Cognome prima di procedere.")
                 else:
                    autoval_dir = os.path.join(skill_matrix_dir, "Autovalutazione")
                    os.makedirs(autoval_dir, exist_ok=True)
                        
                    import re
                    clean_name = re.sub(r'[\\/*?:"<>|]', "", f"{nome_utente.strip()}_{cognome_utente.strip()}")
                    clean_date = re.sub(r'[\\/*?:"<>|]', "", str(data_compilazione))
                    file_name_csv = f"{clean_name}_{clean_date}_Autovalutazione_SkillMatrix.csv"
                    full_csv_path = os.path.join(autoval_dir, file_name_csv)
                        
                    dati_form = {
                        "Campo": [
                            "Nome", "Cognome", "Inquadramento-Mansione", "Ambito lavorativo", "Data Autovalutazione",
                            "Processi produttivi mansione", "Rapporto colleghi", "Interfaccia fornitori-clienti",
                            "Processi cartone-scatole", "Processo pallettizzazione", "Competenze legali-tecniche",
                            "Individuazione rischi-fabbisogni", "Capacità d'adattamento", "Capacità comunicative",
                            "Precisione lavoro", "Persuasione", "Analisi critica contesto", "Turnazioni", "Responsabilità supervisione"
                        ],
                        "Valore": [
                            nome_utente.strip(), cognome_utente.strip(), inquadramento_mansione.strip(), ambito_lavorativo, str(data_compilazione),
                            q1, q2, q3, q4, q5, q6, q7, q8, q9, q10, q11, q12, q13, q14
                        ]
                    }
                    df_res = pd.DataFrame(dati_form)
                    df_res.to_csv(full_csv_path, index=False, sep=";")
                        
                    st.success(f"Autovalutazione salvata con successo! File: `{file_name_csv}` nella cartella `Autovalutazione`.")
        # =========================================================
        # SOTTOSEZIONE RISERVATA - SKILL MATRIX (Richiede Password)
        # =========================================================                 
        elif:
            if not st.session_state.auth_skill_matrix:
                st.markdown("Inserisci la password per accedere alla sezione Skill Matrix.")
                pwd_skill = st.text_input("Password Skill Matrix", type="password", key="pwd_skill_input")
                if st.button("Verifica Password", use_container_width=True, key="btn_verify_pwd_skill"):
                    if pwd_skill == "hse2026":
                        st.session_state.auth_skill_matrix = True
                        st.success("Accesso autorizzato!")
                        st.rerun()
                    else:
                        st.error("Password errata.")
    
            if sotto_sec_sm == "Skill Matrix":
                st.subheader("Skill Matrix - Panoramica Generale e Tabella Dinamica")
                st.markdown(
                    "A ogni competenza si attribuirà un punteggio del 1 al 5. "
                    "Le persone saranno anche classificate a seconda del area lavoratoriva d'appartenenza. "
                    "Posteriormente, se si ritiene opportuno, le persone volontarie che apparterrano al Comitato "
                    "per la sicurezza di quel specifico Near Miss, potranno essere selezionate oltre che per le "
                    "competenze, anche per l'ambito lavorativo d'appartenenza."
                )
                st.markdown("---")
                st.markdown("Visualizza e modifica direttamente i dati aggregati delle autovalutazioni inserite dal personale.")
                
                autoval_dir = os.path.join(skill_matrix_dir, "Autovalutazione")
                if os.path.exists(autoval_dir):
                    files_csv = [f for f in os.listdir(autoval_dir) if f.endswith(".csv")]
                    if files_csv:
                        lista_righe_tabella = []
                        
                        for f_csv in files_csv:
                            f_path = os.path.join(autoval_dir, f_csv)
                            try:
                                df_temp = pd.read_csv(f_path, sep=";")
                                
                                def get_val(campo_str, default_val):
                                    res = df_temp.loc[df_temp["Campo"] == campo_str, "Valore"]
                                    return res.values[0] if not res.empty else default_val
    
                                lista_righe_tabella.append({
                                    "Nome": str(get_val("Nome", "N/D")),
                                    "Cognome": str(get_val("Cognome", "N/D")),
                                    "Inquadramento-Mansione": str(get_val("Inquadramento-Mansione", "")),
                                    "Ambito lavorativo": str(get_val("Ambito lavorativo", "Produzione")),
                                    "Data Autovalutazione": str(get_val("Data Autovalutazione", "")),
                                    "Processi produttivi mansione": float(get_val("Processi produttivi mansione", 3)),
                                    "Rapporto colleghi": float(get_val("Rapporto colleghi", 3)),
                                    "Interfaccia fornitori-clienti": float(get_val("Interfaccia fornitori-clienti", 3)),
                                    "Processi cartone-scatole": float(get_val("Processi cartone-scatole", 3)),
                                    "Processo pallettizzazione": float(get_val("Processo pallettizzazione", 3)),
                                    "Competenze legali-tecniche": float(get_val("Competenze legali-tecniche", 3)),
                                    "Individuazione rischi-fabbisogni": float(get_val("Individuazione rischi-fabbisogni", 3)),
                                    "Capacità d'adattamento": float(get_val("Capacità d'adattamento", 3)),
                                    "Capacità comunicative": float(get_val("Capacità comunicative", 3)),
                                    "Precisione lavoro": float(get_val("Precisione lavoro", 3)),
                                    "Persuasione": float(get_val("Persuasione", 3)),
                                    "Analisi critica contesto": float(get_val("Analisi critica contesto", 3)),
                                    "Turnazioni": float(get_val("Turnazioni", 3)),
                                    "Responsabilità supervisione": float(get_val("Responsabilità supervisione", 3)),
                                    "File Sorgente": f_csv
                                })
                            except Exception:
                                pass
                                
                        if lista_righe_tabella:
                            df_master_sm = pd.DataFrame(lista_righe_tabella)
                            
                            st.markdown("#### Tabella Panoramica Modificabile")
                            st.info("Puoi modificare i valori direttamente nella tabella sottostante. Clicca sui pulsanti in basso per salvare o scaricare.")
                            
                            edited_master_sm = st.data_editor(
                                df_master_sm,
                                use_container_width=True,
                                num_rows="dynamic",
                                column_config={
                                    "Ambito lavorativo": st.column_config.SelectboxColumn(
                                        "Ambito lavorativo",
                                        options=["Uffici", "Produzione", "Manutenzione", "Stoccaggio MP", "Trasporto-Logistica", "Commerciale"],
                                        required=True
                                    ),
                                    "File Sorgente": st.column_config.TextColumn("File Sorgente", disabled=True)
                                },
                                key="editor_skill_matrix_generale"
                            )
                            
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                if st.button("💾 Salva Modifiche Tabella Skill Matrix", use_container_width=True, key="btn_save_master_sm"):
                                    salvati_ok = True
                                    for idx, row in edited_master_sm.iterrows():
                                        f_src = row.get("File Sorgente", "")
                                        if pd.isna(f_src) or not f_src:
                                            import re
                                            c_name = re.sub(r'[\\/*?:"<>|]', "", f"{row['Nome']}_{row['Cognome']}")
                                            c_date = re.sub(r'[\\/*?:"<>|]', "", str(row['Data Autovalutazione']))
                                            f_src = f"{c_name}_{c_date}_Autovalutazione_SkillMatrix.csv"
                                            
                                        f_path = os.path.join(autoval_dir, f_src)
                                        
                                        df_single_updated = pd.DataFrame({
                                            "Campo": [
                                                "Nome", "Cognome", "Inquadramento-Mansione", "Ambito lavorativo", "Data Autovalutazione",
                                                "Processi produttivi mansione", "Rapporto colleghi", "Interfaccia fornitori-clienti",
                                                "Processi cartone-scatole", "Processo pallettizzazione", "Competenze legali-tecniche",
                                                "Individuazione rischi-fabbisogni", "Capacità d'adattamento", "Capacità comunicative",
                                                "Precisione lavoro", "Persuasione", "Analisi critica contesto", "Turnazioni", "Responsabilità supervisione"
                                            ],
                                            "Valore": [
                                                row["Nome"], row["Cognome"], row["Inquadramento-Mansione"], row["Ambito lavorativo"], row["Data Autovalutazione"],
                                                row["Processi produttivi mansione"], row["Rapporto colleghi"], row["Interfaccia fornitori-clienti"],
                                                row["Processi cartone-scatole"], row["Processo pallettizzazione"], row["Competenze legali-tecniche"],
                                                row["Individuazione rischi-fabbisogni"], row["Capacità d'adattamento"], row["Capacità comunicative"],
                                                row["Precisione lavoro"], row["Persuasione"], row["Analisi critica contesto"], row["Turnazioni"], row["Responsabilità supervisione"]
                                            ]
                                        })
                                        try:
                                            df_single_updated.to_csv(f_path, index=False, sep=";")
                                        except Exception:
                                            salvati_ok = False
                                            
                                    if salvati_ok:
                                        st.success("Modifiche salvate con successo nei file CSV di autovalutazione!")
                                    else:
                                        st.warning("Salvataggio completato con alcune eccezioni.")
                            
                            with col_btn2:
                                csv_master_data = edited_master_sm.to_csv(index=False, sep=";")
                                st.download_button(
                                    label="📥 Scarica Tabella Master in formato CSV",
                                    data=csv_master_data,
                                    file_name="Skill_Matrix_Panoramica_Generale.csv",
                                    mime="text/csv",
                                    use_container_width=True
                                )
                        else:
                            st.info("Nessun dato valido trovato nei file CSV.")
                else:
                    st.info("Nessuna autovalutazione completata e salvata al momento.")
            else:
                st.info("La cartella delle autovalutazioni non è ancora stata creata.")
