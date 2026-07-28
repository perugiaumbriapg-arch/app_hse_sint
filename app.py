import os
import time
from datetime import datetime, timedelta
import streamlit as st
import pandas as pd
import urllib.parse
import urllib.request
import numpy as np
from PIL import Image
import textwrap
import json
import io

# Gestione sicura dell'importazione FPDF / FPDF2 per ambienti Cloud
try:
    from fpdf import FPDF
except ImportError:
    try:
        from fpdf2 import FPDF
    except ImportError:
        FPDF = None

import matplotlib
matplotlib.use('Agg')  # Modalità non interattiva per ambienti cloud/server
import matplotlib.pyplot as plt
import base64
import plotly.express as px
import qrcode
import cv2

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
            with open("normative.json", "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception:
            db = {}
    
    db[chiave.strip()] = {
        "titolo": titolo.strip(),
        "url": url.strip(),
        "desc": descrizione.strip()
    }
    
    with open("normative.json", "w", encoding="utf-8") as f:
        json.dump(db, f, indent=4, ensure_ascii=False)


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
            df = pd.read_excel(percorso_file) if "xls" in estensione else pd.read_csv(percorso_file, sep=";", encoding="utf-8")
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
    try:
        img = Image.open(image_file).convert('RGB')
        img_array = np.array(img)
        detector = cv2.QRCodeDetector()
        valore, punti, qr_code = detector.detectAndDecode(img_array)
        return valore if valore else None
    except Exception:
        return None

# ==================================================================
# --- 3. INTESTAZIONE DELLA PIATTAFORMA ---
# ==================================================================
st.title("SECURITY AND HSE SYSTEM PLATFORM")
st.markdown("### Sistema di Gestione Integrato Locale e Multiutente")
st.markdown("---")

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Home Dashboard"

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

nav = st.radio(
    "Navigazione", 
    options=tab_list, 
    index=tab_list.index(st.session_state.active_tab) if st.session_state.active_tab in tab_list else 0, 
    horizontal=True, 
    label_visibility="collapsed"
)

st.session_state.active_tab = nav
st.markdown("---")

# ==================================================================
# --- SEZIONE 1: HOME DASHBOARD ---
# ==================================================================
if nav == "Home Dashboard":
    st.header("Quadro Generale di Controllo HSE")
    st.markdown("Benvenuto nel menu principale. Qui trovi i dati riassuntivi estratti in tempo reale dai database locali.")
    st.markdown(" ")
    
    colA, colB, colC = st.columns(3)
    with colA:
        tot_nm = len(pd.read_csv(FILE_NEAR_MISS, sep=";", encoding="utf-8")) if os.path.exists(FILE_NEAR_MISS) else 0
        st.metric("Segnalazioni Ricevute", tot_nm)
    with colB:
        tot_an = len(pd.read_csv(FILE_ANALISI_NM, sep=";", encoding="utf-8")) if os.path.exists(FILE_ANALISI_NM) else 0
        st.metric("Analisi Trattate dal RSPP", tot_an)
    with colC:
        tot_scad = len(pd.read_excel(FILE_SCADENZARIO)) if os.path.exists(FILE_SCADENZARIO) else 0
        st.metric("Adempimenti in Registro Scadenze", tot_scad)

# ==================================================================
# --- SEZIONE 2: SEGNALAZIONE NEAR MISS ---
# ==================================================================
elif nav == "Segnalazione Near Miss":
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
                    df_nuovo.to_csv(FILE_NEAR_MISS, index=False, sep=";", encoding="utf-8")
                else:
                    df_nuovo.to_csv(FILE_NEAR_MISS, mode='a', header=False, index=False, sep=";", encoding="utf-8")
                st.success("Segnalazione acquisita con successo nel file CSV locale!")
                time.sleep(0.5)
                st.rerun()

# ==================================================================
# --- SEZIONE 3: SCADENZARIO ADEMPIMENTI ---
# ==================================================================
elif nav == "Scadenzario Adempimenti":
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
elif nav == "Analisi Segnalazioni Near Miss":
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
        df_segnalazioni = pd.read_csv(FILE_NEAR_MISS, sep=";", encoding="utf-8") if os.path.exists(FILE_NEAR_MISS) else pd.DataFrame()
        df_analisi = pd.read_csv(FILE_ANALISI_NM, sep=";", encoding="utf-8") if os.path.exists(FILE_ANALISI_NM) else pd.DataFrame()
        
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
                            df_n.to_csv(FILE_ANALISI_NM, index=False, sep=";", encoding="utf-8")
                        else:
                            df_n.to_csv(FILE_ANALISI_NM, mode='a', header=False, index=False, sep=";", encoding="utf-8")
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
                        df_analisi.to_csv(FILE_ANALISI_NM, index=False, sep=";", encoding="utf-8")
                        st.success("Record aggiornato in linea!")
                        time.sleep(0.5)
                        st.rerun()

# ==================================================================
# --- SEZIONE 5: CONFORMITÀ LEGISLATIVA ---
# ==================================================================
elif nav == "Conformità Legislativa":
    st.header("Conformità Legislativa e Archivio Normativo")
    
    if "autenticato_upload_legale" not in st.session_state:
        st.session_state.autenticato_upload_legale = False
        
    col_bot, col_upload = st.columns([1, 1])
    
    with col_bot:
        st.subheader("Assistente di Scansione Istantanea")
        st.caption("Accesso Pubblico - Ricerca libera e lettura QR Code.")
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
                
        st.markdown("---")
        st.subheader("Lettura Link tramite QR Code")
        st.caption("Strumento pubblico di codifica rapida per la documentazione HSE aziendale.")
        
        metodo_input = st.radio(
            "Scegli come inserire il codice:", 
            ["Carica file dal dispositivo", "Fotocamera"], 
            horizontal=True,
            key="qr_pub_metodo"
        )

        immagine_input = None
        if metodo_input == "Carica file dal dispositivo":
            immagine_input = st.file_uploader(
                "Seleziona un'immagine (png, jpg, jpeg):", 
                type=["png", "jpg", "jpeg"], 
                key="up_qr_pub"
            )
        else:
            immagine_input = st.camera_input("Scatta una foto del QR Code:", key="cam_qr_pub")
        
        risultato_pub = None
        if immagine_input:
            risultato_pub = decodifica_qr_opencv(immagine_input)
    
        if risultato_pub:
            st.success("Codice QR decodificato con successo!")
            st.code(risultato_pub, language="text")
            if risultato_pub.startswith("http://") or risultato_pub.startswith("https://"):
                st.markdown(f"[Clicca qui per aprire il link rilevato]({risultato_pub})")
        else:
            st.warning("Nessun codice QR valido rilevato nell'immagine.")
        
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
                        csv_buffer = df_visualizza.to_csv(index=False, sep=";", encoding='utf-8').encode('utf-8')
                        st.download_button(
                            label="Scarica Check-list in formato CSV",
                            data=csv_buffer,
                            file_name=f"checklist_conformita_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                        
            with tab_admin2:
                st.markdown("#### Genera QR Code")
                opzione_qr = st.radio(
                    "Seleziona operazione QR:", 
                    ["Genera QR Code", "Leggi/Decodifica QR Code"], 
                    key="radio_opzioni_qr_pubblico"
                )
                
                if opzione_qr == "Genera QR Code":
                    testo_da_convertire = st.text_input(
                        "Inserisci l'URL o il testo da inserire nel QR Code:", 
                        placeholder="https://www.gazzettaufficiale.it/", 
                        key="input_testo_qr_gen"
                    )
                    if testo_da_convertire.strip():
                        qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={urllib.parse.quote(testo_da_convertire.strip())}"
                        try:
                            with urllib.request.urlopen(qr_url) as response:
                                qr_bytes = response.read()
                            st.image(qr_bytes, width=220, caption="Codice QR generato con successo!")
                            st.download_button(
                                label="Scarica Immagine QR Code (.png)",
                                data=qr_bytes,
                                file_name="qr_code_hse.png",
                                mime="image/png",
                                use_container_width=True,
                                key="dl_btn_qr_code"
                            )
                        except Exception as e:
                            st.error(f"Errore di connessione durante la generazione del QR Code: {e}")
                else:
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

                st.markdown("---")
                st.subheader("Verifica Aggiornamenti Cogenti (Italia & UE)")
                with st.expander("Aggiungi nuova norma al database interno"):
                    k = st.text_input("Parola chiave (es: rifiuti):", key="norm_k")
                    t = st.text_input("Titolo:", key="norm_t")
                    u = st.text_input("URL ufficiale:", key="norm_u")
                    d = st.text_area("Descrizione della norma:", key="norm_d")
                    if st.button("Salva nel database", key="btn_save_norm"): 
                        salva_normativa(k, t, u, d)
                        st.success("Norma salvata!")
            
                note_libere = st.text_area("Analisi testo libero / criticità:", placeholder="es: gestione rifiuti pericolosi, amianto, ecc.", key="note_libere_conf")
                if st.button("Avvia Analisi di Conformità", use_container_width=True, key="btn_avvia_conf"):
                    db = {}
                    if os.path.exists("normative.json"):
                        with open("normative.json", "r", encoding="utf-8") as f:
                            try: db = json.load(f)
                            except: db = {}
                
                    txt_checklist = st.session_state["cached_df_checklist"].to_string() if "cached_df_checklist" in st.session_state else ""
                    txt_docs = " ".join(os.listdir(DIR_CONFORMITA)) if os.path.exists(DIR_CONFORMITA) else ""
                    txt_totale_analisi = (note_libere + " " + txt_checklist + " " + txt_docs).lower()
                
                    st.markdown("### Risultati Stato Conformità")
                    trovato = False
                    for chiave, info in db.items():
                        if chiave.lower() in txt_totale_analisi:
                            trovato = True
                        is_conforme = chiave.lower() in (txt_checklist + txt_docs).lower()
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
elif nav == "Analisi - Fase 2":
    st.header("Analisi - Fase 2 delle segnalazioni near miss")
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
        opzioni = ["Nessuna (Nuova analisi)"]
        if os.path.exists(FILE_NEAR_MISS):
            df_nm = pd.read_csv(FILE_NEAR_MISS, sep=";", encoding="utf-8")
            for idx, r in df_nm.iterrows():
                opzioni.append(f"NM | {r.get('Data Segnalazione', 'N/D')} | {r.get('Tipo Evento', 'Evento')}")
        if os.path.exists(FILE_ANALISI_NM):
            df_an = pd.read_csv(FILE_ANALISI_NM, sep=";", encoding="utf-8")
            for idx, r in df_an.iterrows():
                opzioni.append(f"AN | {r.get('Data Analisi', 'N/D')} | Collegamento: {r.get('Segnalazione Collegata', 'Analisi')}")
            
        scelta_rif = st.selectbox("Seleziona evento/analisi collegata:", opzioni)

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
            if FPDF is None:
                st.error("Errore: la libreria 'fpdf2' non risulta ancora configurata nel sistema. Controlla il file requirements.txt.")
                st.stop()
                
            temp_img = os.path.join(DIR_ANALISI, "temp_ishikawa.png")
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.axis('off')
            ax.plot([1, 9], [0, 0], color='black', lw=3)
                
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
            plt.close(fig)
            st.pyplot(fig)
            
            with open(temp_img, "rb") as img_file:
                png_bytes = img_file.read()

            data_generazione = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            file_obbligatorio = FILE_ANALISI_NM
            file_facoltativo = FILE_NEAR_MISS if os.path.exists(FILE_NEAR_MISS) else "Non utilizzato / Assente"

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

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, "Report Analisi Near Miss", ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 6, f"Data Generazione: {data_generazione}", ln=True)
            pdf.cell(0, 6, f"File Obbligatorio: {file_obbligatorio}", ln=True)
            pdf.cell(0, 6, f"File Facoltativo: {file_facoltativo}", ln=True)
            pdf.cell(0, 6, f"Riferimento: {scelta_rif}", ln=True)
            pdf.ln(5)

            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "Analisi 4M:", ln=True)
            pdf.set_font("Arial", size=12)
            for key, val in dati_report["4M"].items():
                pdf.cell(0, 10, f"- {key}: {val}", ln=True)
            pdf.ln(5)

            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "5Whys:", ln=True)
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 10, f"Perchè 1: {dati_report['5Whys'][0]}", ln=True)
            pdf.ln(5)

            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "Conclusioni:", ln=True)
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, dati_report['Conclusioni'])

            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 10, "Diagramma di Ishikawa:", ln=True)
            pdf.image(temp_img, w=180)

            st.session_state.pdf_bytes = bytes(pdf.output(dest='S'))
            buffer_xlsx = io.BytesIO()
            pd.DataFrame([dati_report["4M"]]).to_excel(buffer_xlsx, index=False)
            st.session_state.xlsx_bytes = buffer_xlsx.getvalue()
            st.session_state.json_bytes = json.dumps(dati_report, indent=4, ensure_ascii=False).encode('utf-8')
            st.session_state.png_bytes = png_bytes
            st.session_state.report_ready = True
            st.success("Report generato con successo!")

        if st.session_state.get("report_ready"):
            col_d1, col_d2, col_d3 = st.columns(3)
            col_d1.download_button("Scarica PDF", st.session_state.pdf_bytes, "Report.pdf")
            col_d2.download_button("Scarica XLSX", st.session_state.xlsx_bytes, "Dati.xlsx")
            col_d3.download_button("Scarica JSON", st.session_state.json_bytes, "Dati.json")

# ==================================================================
# --- SEZIONE 7: KPI ---
# ==================================================================
elif nav == "KPI":
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
            st.markdown("Gestione centralizzata, monitoraggio e storico delle performance di sicurezza sul lavoro.")
            st.markdown("---")

            oggi = datetime.today().date()
            data_str = oggi.strftime('%Y-%m-%d')
            os.makedirs("KPI", exist_ok=True)

            path_master_pers = os.path.join("KPI", "master_kpi_definitions.csv")
            path_storico_pers = os.path.join("KPI", "storico_misure.csv")

            if "t1_kpi_definitions" not in st.session_state:
                if os.path.exists(path_master_pers):
                    try:
                        st.session_state.t1_kpi_definitions = pd.read_csv(path_master_pers, encoding="utf-8").fillna("")
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
                    ])
                    st.session_state.t1_kpi_definitions.to_csv(path_master_pers, index=False, encoding='utf-8')

            if "t2_storico_misure" not in st.session_state:
                if os.path.exists(path_storico_pers):
                    try:
                        df_loaded = pd.read_csv(path_storico_pers, encoding="utf-8")
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
                    ])
                    st.session_state.t2_storico_misure.to_csv(path_storico_pers, index=False, encoding='utf-8')

            if "kpi_main_nav" not in st.session_state:
                st.session_state.kpi_main_nav = "📊 Definizione Master KPI"

            st.session_state.kpi_main_nav = st.radio(
                "Seleziona Sezione Principale KPI",
                ["📊 Definizione Master KPI", "⚙️ Strumento di Monitoraggio & Storico"],
                horizontal=True,
                key="radio_kpi_main_nav_selector"
            )
            st.markdown("---")

            if st.session_state.kpi_main_nav == "📊 Definizione Master KPI":
                st.subheader("📋 Gestione Master dei KPI, Formule e Parametri")
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
                nome_file_t1 = f"{data_str}_Tab1-MasterKPI.csv"
                path_t1 = os.path.join("KPI", nome_file_t1)
                csv_t1 = st.session_state.t1_kpi_definitions.to_csv(index=False, encoding='utf-8').encode('utf-8')

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
            else:
                st.subheader("Strumento di Monitoraggio (Riconoscimento Automatico KPI da Tab 1)")
                df_kpi_master = st.session_state.t1_kpi_definitions.copy()
                st.info("Sezione Monitoraggio KPI attiva.")
                if not df_kpi_master.empty:
                    st.dataframe(df_kpi_master, use_container_width=True)
                else:
                    st.warning("Nessun KPI definito.")

        render_sezione_kpi()

# ==================================================================
# --- SEZIONE 8: Piano di Miglioramento ---
# ==================================================================
elif nav == "Piano Miglioramento":
    st.header("Piano di Miglioramento HSE")
    
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
        else:
            st.info("File PDF 'Istruzione_Piano_Miglioramento.pdf' non trovato sul disco. Puoi caricarlo nella cartella di lavoro.")
                
    elif sotto_sec == "Valutazione del rischio":
        st.subheader("Valutazione del Rischio e Collegamento Analisi")
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
            st.markdown("### Tabella Dinamica di Valutazione del Rischio")
            st.markdown("I parametri numerici (**Probabilità**, **Gravità**, **Sensibilità**, **Controllo**) accettano numeri interi tra 1 e 3.")
            
            if "df_vr_session" not in st.session_state:
                st.session_state.df_vr_session = pd.DataFrame([
                    {"Rischio": "Rischio movimentazione carrelli", "Probabilità": 2, "Gravità": 2, "Sensibilità": 2, "Controllo": 2, "Significatività": 16}
                ])
            
            df_vr_edited = st.data_editor(st.session_state.df_vr_session, num_rows="dynamic", key="editor_vr_s8", use_container_width=True)
            if st.button("Ricalcola e Salva Valutazione Rischio", use_container_width=True):
                try:
                    for col in ["Probabilità", "Gravità", "Sensibilità", "Controllo"]:
                        df_vr_edited[col] = pd.to_numeric(df_vr_edited[col], errors='coerce').fillna(1).astype(int)
                    df_vr_edited["Significatività"] = df_vr_edited["Probabilità"] * df_vr_edited["Gravità"] * df_vr_edited["Sensibilità"] * df_vr_edited["Controllo"]
                    st.session_state.df_vr_session = df_vr_edited
                    st.success("Valutazione del rischio aggiornata con successo!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore nel calcolo: {e}")
    else:
        st.subheader("Azioni Piano di Miglioramento")
        st.markdown("Pianifica, assegna e monitora le azioni correttive e preventive scaturite dai Near Miss e dalle valutazioni dei rischi.")
        
        if "df_azioni_pm" not in st.session_state:
            st.session_state.df_azioni_pm = pd.DataFrame([
                {"Azione": "Installare protezioni aggiuntive su nastro trasportatore", "Responsabile": "RSPP / Manutenzione", "Scadenza": "30/09/2026", "Stato": "In Corso"}
            ])
            
        df_azioni_edited = st.data_editor(st.session_state.df_azioni_pm, num_rows="dynamic", key="editor_azioni_pm", use_container_width=True)
        if st.button("Salva Piano Azioni", use_container_width=True):
            st.session_state.df_azioni_pm = df_azioni_edited
            st.success("Piano azioni salvato correttamente!")

# ==================================================================
# --- SEZIONE 9: Stima Costo Economico ---
# ==================================================================
elif nav == "Stima Costo Economico":
    st.header("Stima del Costo Economico degli Incidenti & Near Miss")
    st.markdown("Valutazione finanziaria dei costi diretti e indiretti legati agli eventi di mancato infortunio, infortuni e non conformità.")
    
    with st.form("form_stima_costi"):
        col1, col2 = st.columns(2)
        with col1:
            n_eventi = st.number_input("Numero stimato di eventi/anno", min_value=1, value=5)
            costo_diretto_medio = st.number_input("Costo diretto medio per evento (€)", min_value=0.0, value=1500.0)
        with col2:
            costo_indiretto_medio = st.number_input("Costo indiretto medio (fermo impianto, perdita produttività) (€)", min_value=0.0, value=3500.0)
            budget_prev = st.number_input("Budget annuale prevenzione DPI e Formazione (€)", min_value=0.0, value=10000.0)
            
        calcola_costi = st.form_submit_button("Calcola Impatto Economico", use_container_width=True)
        if calcola_costi:
            tot_costi_eventi = n_eventi * (costo_diretto_medio + costo_indiretto_medio)
            st.success(f"Costo stimato totale annuo derivante dagli eventi: **€ {tot_costi_eventi:,.2f}**")
            st.info(f"Rapporto investimento prevenzione / costo eventi: **{(budget_prev / (tot_costi_eventi if tot_costi_eventi > 0 else 1)) * 100:.1f}%**")

# ==================================================================
# --- SEZIONE 10: Skill Matrix ---
# ==================================================================
elif nav == "Skill Matrix":
    st.header("Skill Matrix & Formazione del Personale")
    st.markdown("Matrice delle competenze, abilitazioni e scadenze formative per ciascun lavoratore.")
    
    if "df_skill_matrix" not in st.session_state:
        st.session_state.df_skill_matrix = pd.DataFrame([
            {"Nominativo": "Mario Rossi", "Mansione": "Magazziniere", "Formazione Generale": "Valida", "Muletto / Carrello": "Valido (Scad. 2027)", "Antincendio": "Valido", "Primo Soccorso": "In Scadenza"},
            {"Nominativo": "Luigi Bianchi", "Mansione": "Operaio Produzione", "Formazione Generale": "Valida", "Muletto / Carrello": "Non Abilitato", "Antincendio": "Valido", "Primo Soccorso": "Valido"}
        ])
        
    df_skill_edited = st.data_editor(st.session_state.df_skill_matrix, num_rows="dynamic", key="editor_skill_matrix", use_container_width=True)
    if st.button("Salva Skill Matrix", use_container_width=True):
        st.session_state.df_skill_matrix = df_skill_edited
        st.success("Skill Matrix aggiornata con successo!")

# ==================================================================
# --- SEZIONE 11: Controllo DPI ---
# ==================================================================
elif nav == "Controllo DPI":
    st.header("Controllo e Consegna Dispositivi di Protezione Individuale (DPI)")
    st.markdown("Registro di consegna e verifica periodica dell'integrità dei DPI assegnati ai lavoratori.")
    
    if "df_dpi" not in st.session_state:
        st.session_state.df_dpi = pd.DataFrame([
            {"Nominativo": "Mario Rossi", "DPI Consegnato": "Scarpe Antinfortunistiche S3", "Data Consegna": "15/01/2025", "Stato": "Integro / In Uso", "Firma Ricezione": "Firmato"}
        ])
        
    df_dpi_edited = st.data_editor(st.session_state.df_dpi, num_rows="dynamic", key="editor_dpi", use_container_width=True)
    if st.button("Salva Registro DPI", use_container_width=True):
        st.session_state.df_dpi = df_dpi_edited
        st.success("Registro DPI aggiornato e salvato con successo!")
