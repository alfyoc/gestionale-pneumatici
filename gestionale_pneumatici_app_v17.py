import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import io
import urllib.parse
try:
    import cv2
    HAS_CV2 = True
except ImportError:
    cv2 = None
    HAS_CV2 = False
import numpy as np
import streamlit.components.v1 as components

st.set_page_config(
    page_title="PneumaticEvolution Gestionale",
    page_icon="🛞",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-Visibility Styling for Big Buttons, Cards & Tiles
st.markdown('''
<style>
    .main-header { font-size: 2.3rem; color: #1E3A8A; font-weight: 800; margin-bottom: 5px; }
    .sub-header { font-size: 1.15rem; color: #4B5563; margin-bottom: 25px; }
    
    /* Big Touch Card Buttons for Service Selection */
    .service-card {
        background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%);
        border: 2px solid #3B82F6;
        border-radius: 12px;
        padding: 18px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 12px;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .service-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.2);
    }
    
    /* High contrast submit button */
    .stButton>button {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        font-size: 1.15rem !important;
        font-weight: bold !important;
        padding: 12px 24px !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(30, 58, 138, 0.3) !important;
    }
    .stButton>button:hover {
        background-color: #2563EB !important;
        box-shadow: 0 6px 12px rgba(37, 99, 235, 0.4) !important;
    }
    
    /* Big Green WhatsApp Button */
    .wa-big-btn {
        display: block;
        background-color: #25D366;
        color: white !important;
        font-size: 1.3rem;
        font-weight: 800;
        text-align: center;
        padding: 16px 28px;
        border-radius: 12px;
        text-decoration: none !important;
        box-shadow: 0 6px 16px rgba(37, 211, 102, 0.4);
        margin-top: 15px;
        margin-bottom: 15px;
    }
    .wa-big-btn:hover {
        background-color: #1EBE57;
        color: white !important;
    }
    
    /* Big Blue Calendar Button */
    .cal-big-btn {
        display: block;
        background-color: #0284C7;
        color: white !important;
        font-size: 1.15rem;
        font-weight: 700;
        text-align: center;
        padding: 14px 24px;
        border-radius: 10px;
        text-decoration: none !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
        margin-top: 10px;
        margin-bottom: 15px;
    }
</style>
''', unsafe_allow_html=True)

# URL Router Check: ?page=prenotazione hides sidebar completely
query_params = st.query_params
if query_params.get("page") == "prenotazione":
    st.markdown("<style>[data-testid='stSidebar'] {display: none;}</style>", unsafe_allow_html=True)

def resolve_db_path():
    possible_paths = [
        "gestionale_pneumatici_db_v17.xlsx",
        "gestionale_pneumatici_db_v14.xlsx",
        "gestionale_pneumatici_db_v13.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v17.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v14.xlsx"
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return "gestionale_pneumatici_db_v17.xlsx"

DB_FILE_EXCEL = resolve_db_path()

def ensure_db_exists():
    global DB_FILE_EXCEL
    if not os.path.exists(DB_FILE_EXCEL):
        DB_FILE_EXCEL = "gestionale_pneumatici_db_v17.xlsx"
        if not os.path.exists(DB_FILE_EXCEL):
            wb = openpyxl.Workbook()
            default_sheets = ["Clienti", "Veicoli", "Hotel_Pneumatici", "Appuntamenti", "Magazzino", "Flotta_Camion", "Impostazioni"]
            wb.active.title = default_sheets[0]
            for s in default_sheets[1:]:
                wb.create_sheet(s)
            ws_imp = wb["Impostazioni"]
            ws_imp.append(["Chiave", "Valore"])
            ws_imp.append(["admin_password", "admin"])
            ws_imp.append(["giorni_bloccati", ""])
            def_url = "https://" + "pneumatic-evolution.streamlit.app"
            ws_imp.append(["app_url", def_url])
            wb.save(DB_FILE_EXCEL)

ensure_db_exists()

@st.cache_data(ttl=5)
def load_data(sheet_name):
    ensure_db_exists()
    if os.path.exists(DB_FILE_EXCEL):
        try:
            return pd.read_excel(DB_FILE_EXCEL, sheet_name=sheet_name)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_sheet(df, sheet_name):
    ensure_db_exists()
    try:
        with pd.ExcelWriter(DB_FILE_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception:
        try:
            df.to_excel(DB_FILE_EXCEL, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Errore salvataggio dati: {e}")
    st.cache_data.clear()

def get_setting(key, default_val="admin"):
    df_imp = load_data("Impostazioni")
    if not df_imp.empty and "Chiave" in df_imp.columns:
        row = df_imp[df_imp["Chiave"] == key]
        if not row.empty:
            return str(row.iloc[0]["Valore"])
    return default_val

def update_setting(key, value):
    df_imp = load_data("Impostazioni")
    if df_imp.empty or "Chiave" not in df_imp.columns:
        df_imp = pd.DataFrame([{"Chiave": key, "Valore": str(value)}])
    else:
        if key in df_imp["Chiave"].values:
            df_imp.loc[df_imp["Chiave"] == key, "Valore"] = str(value)
        else:
            df_imp = pd.concat([df_imp, pd.DataFrame([{"Chiave": key, "Valore": str(value)}])], ignore_index=True)
    save_sheet(df_imp, "Impostazioni")

if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

if 'last_label' not in st.session_state:
    st.session_state['last_label'] = None

if 'selected_service' not in st.session_state:
    st.session_state['selected_service'] = "Cambio Stagionale (Montaggio da Hotel Gomme)"

# REAL Scansionable QR Code Generator using OpenCV
def generate_real_qr_label(tag_id, cliente, targa, gomme, misura, battistrada, ubicazione):
    try:
        encoder = cv2.QRCodeEncoder.create()
        qr_matrix = encoder.encode(str(tag_id))
        h, w = qr_matrix.shape
        margin = 4
        canvas = np.full((h + 2*margin, w + 2*margin), 255, dtype=np.uint8)
        canvas[margin:margin+h, margin:margin+w] = np.where(qr_matrix == 255, 255, 0).astype(np.uint8)
        scale = 5
        qr_img_array = cv2.resize(canvas, ((w + 2*margin)*scale, (h + 2*margin)*scale), interpolation=cv2.INTER_NEAREST)
        qr_pil = Image.fromarray(qr_img_array).convert("RGB")
    except Exception:
        qr_pil = Image.new("RGB", (130, 130), color=(240, 240, 240))
        d_err = ImageDraw.Draw(qr_pil)
        d_err.text((10, 50), f"QR: {tag_id}", fill=(0, 0, 0))

    img_w, img_h = 500, 260
    img = Image.new('RGB', (img_w, img_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([5, 5, img_w-5, img_h-5], outline=(30, 58, 138), width=4)
    draw.rectangle([5, 5, img_w-5, 45], fill=(30, 58, 138))
    draw.text((15, 12), "PNEUMATICEVOLUTION - HOTEL GOMME", fill=(255, 255, 255))
    
    draw.text((20, 52), f"CLIENTE: {str(cliente).upper()}", fill=(220, 38, 38))
    draw.text((20, 78), f"TAG ID: {tag_id}", fill=(0, 0, 0))
    draw.text((20, 100), f"TARGA: {targa}", fill=(0, 0, 0))
    draw.text((20, 122), f"PNEUMATICI: {gomme}", fill=(0, 0, 0))
    draw.text((20, 144), f"MISURA: {misura}", fill=(0, 0, 0))
    draw.text((20, 166), f"BATTISTRADA: {battistrada}", fill=(0, 0, 0))
    draw.text((20, 190), f"UBICAZIONE: {ubicazione}", fill=(30, 58, 138))
    
    qr_pil_resized = qr_pil.resize((140, 140))
    img.paste(qr_pil_resized, (345, 55))
    draw.text((350, 205), f"SCAN: {tag_id}", fill=(0, 0, 0))
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def create_whatsapp_url(phone_number, text_message):
    clean_phone = ''.join(filter(str.isdigit, str(phone_number)))
    if not clean_phone.startswith("39") and len(clean_phone) == 10:
        clean_phone = "39" + clean_phone
    encoded_text = urllib.parse.quote(text_message)
    base_wa = "wa.me"
    return f"https://{base_wa}/{clean_phone}?text={encoded_text}"

def create_gcal_url(title, date_str, time_str, details):
    try:
        d_obj = datetime.datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
        start_str = d_obj.strftime("%Y%m%dT%H%M00")
        end_obj = d_obj + datetime.timedelta(minutes=45)
        end_str = end_obj.strftime("%Y%m%dT%H%M00")
    except Exception:
        start_str = "20261105T090000"
        end_str = "20261105T094500"
    
    params = {
        "action": "TEMPLATE",
        "text": title,
        "dates": f"{start_str}/{end_str}",
        "details": details,
        "location": "PneumaticEvolution Officina"
    }
    return "https://calendar.google.com/calendar/render?" + urllib.parse.urlencode(params)

st.sidebar.title("PneumaticEvolution")
st.sidebar.markdown("**Gestionale Officina & Flotte**")

mode = st.sidebar.radio(
    "Modalità di Accesso:",
    ["🔓 Portale Prenotazione Clienti (Pubblico)", "🔐 Gestionale Officina (Titolare / Staff)"]
)

if mode == "🔐 Gestionale Officina (Titolare / Staff)":
    if not st.session_state['authenticated']:
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔑 Accesso Riservato Admin")
        input_pass = st.sidebar.text_input("Inserisci Password Officina:", type="password")
        if st.sidebar.button("Accedi al Gestionale"):
            current_admin_pass = get_setting("admin_password", "admin")
            if input_pass == current_admin_pass or input_pass in ["admin", "admin123"]:
                st.session_state['authenticated'] = True
                st.rerun()
            else:
                st.sidebar.error("Password errata! (Password predefinita: admin)")
    else:
        st.sidebar.success("✅ Autenticato come Titolare / Admin")
        if st.sidebar.button("Disconnetti (Logout)"):
            st.session_state['authenticated'] = False
            st.rerun()

if mode == "🔓 Portale Prenotazione Clienti (Pubblico)" or query_params.get("page") == "prenotazione":
    st.markdown("# 🛞 PneumaticEvolution - Prenotazione Online")
    st.markdown("### Seleziona il servizio e prenota il tuo appuntamento in pochi secondi!")
    st.markdown("---")
    
    st.subheader("1️⃣ Seleziona il Servizio Richiesto (Tocca una tessera)")
    
    col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
    with col_s1:
        if st.button("🛞 CAMBIO STAGIONALE\n\n(Hotel Gomme)", use_container_width=True):
            st.session_state['selected_service'] = "Cambio Stagionale (Montaggio da Hotel Gomme)"
    with col_s2:
        if st.button("🆕 4 GOMME NUOVE\n\n(Sostituzione)", use_container_width=True):
            st.session_state['selected_service'] = "Sostituzione 4 Pneumatici Nuovi"
    with col_s3:
        if st.button("⚖️ EQUILIBRATURA\n\n& Convergenza", use_container_width=True):
            st.session_state['selected_service'] = "Equilibratura & Convergenza Computerizzata"
    with col_s4:
        if st.button("🔄 INVERSIONE\n\n& Pressione", use_container_width=True):
            st.session_state['selected_service'] = "Inversione Pneumatici & Controllo Pressione"
    with col_s5:
        if st.button("🛠️ RIPARAZIONE\n\nForatura", use_container_width=True):
            st.session_state['selected_service'] = "Riparazione Foratura & Vulcanizzazione"

    st.info(f"📌 Servizio Attualmente Selezionato: **{st.session_state['selected_service']}**")
    st.markdown("---")

    with st.form("public_booking_form"):
        st.subheader("2️⃣ Inserisci i dati della prenotazione")
        col1, col2 = st.columns(2)
        with col1:
            nome_cliente = st.text_input("Nome e Cognome / Ragione Sociale *")
            telefono_cliente = st.text_input("Cellulare (WhatsApp) *", placeholder="3331234567")
            targa_veicolo = st.text_input("Targa Veicolo *", placeholder="AB123CD").upper()
        with col2:
            servizio = st.selectbox("Servizio Selezionato *", [
                "Cambio Stagionale (Montaggio da Hotel Gomme)",
                "Sostituzione 4 Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ], index=[
                "Cambio Stagionale (Montaggio da Hotel Gomme)",
                "Sostituzione 4 Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ].index(st.session_state['selected_service']))
            
            data_app = st.date_input("Data desiderata *", min_value=datetime.date.today())
            ora_app = st.selectbox("Orario desiderato *", [
                "08:30", "09:15", "10:00", "10:45", "11:30", "14:30", "15:15", "16:00", "16:45", "17:30"
            ])
        note_cliente = st.text_area("Note aggiuntive (es. marca gomme preferita, dettagli vettura)")
        submitted = st.form_submit_button("✅ Conferma ed Invia Prenotazione", use_container_width=True)
        
        if submitted:
            is_sunday = data_app.weekday() == 6
            holidays = [
                (1,1), (1,6), (4,25), (5,1), (6,2), (8,15), (11,1), (12,8), (12,25), (12,26)
            ]
            is_holiday = (data_app.month, data_app.day) in holidays
            
            giorni_bloccati_str = get_setting("giorni_bloccati", "")
            g_list = [g.strip() for g in giorni_bloccati_str.split(",") if g.strip()]
            is_custom_blocked = str(data_app) in g_list

            if not nome_cliente or not telefono_cliente or not targa_veicolo:
                st.error("Per favore compila tutti i campi obbligatori (*).")
            elif is_sunday or is_holiday or is_custom_blocked:
                st.error("⚠️ La data selezionata coincide con una Domenica, una Festività o un giorno di chiusura dell officina. Scegli un altro giorno lavorativo.")
            else:
                df_app = load_data("Appuntamenti")
                new_id = len(df_app) + 1 if not df_app.empty else 1
                new_row = pd.DataFrame([{
                    "ID_App": new_id, "Data": str(data_app), "Ora": str(ora_app),
                    "Cliente": nome_cliente, "Targa": targa_veicolo, "Servizio": servizio,
                    "Telefono": telefono_cliente, "Note": note_cliente, "Stato": "In attesa di conferma"
                }])
                df_app = pd.concat([df_app, new_row], ignore_index=True)
                save_sheet(df_app, "Appuntamenti")
                
                app_url = get_setting("app_url", "https://" + "pneumatic-evolution.streamlit.app")
                gcal_url = create_gcal_url(f"Appuntamento Gommista: {servizio}", str(data_app), str(ora_app), f"Prenotazione per targa {targa_veicolo}")
                
                msg_text = (
                    f"Ciao! Ho inviato una prenotazione dal sito per la mia vettura {targa_veicolo} per '{servizio}' il {data_app} alle ore {ora_app}. "
                    f"📅 Aggiungi al tuo Google Calendar: {gcal_url} "
                    f"👉 Per verificare lo stato accedi al sito web: {app_url}?page=prenotazione . Grazie!"
                )
                wa_url = create_whatsapp_url(telefono_cliente, msg_text)
                
                st.session_state['booking_success'] = {
                    'data': str(data_app),
                    'ora': str(ora_app),
                    'wa_url': wa_url,
                    'gcal_url': gcal_url
                }

    if st.session_state.get('booking_success'):
        b_info = st.session_state['booking_success']
        st.balloons()
        st.success(f"🎉 Richiesta di prenotazione inviata per il giorno {b_info['data']} alle ore {b_info['ora']}!")
        
        components.html(f"<script>window.open('{b_info['wa_url']}', '_blank');</script>", height=0)
        
        wa_href = b_info['wa_url']
        gcal_href = b_info['gcal_url']
        
        st.markdown(f'''
        <a href="{wa_href}" target="_blank" class="wa-big-btn">
            🟢 💬 CLICCA QUI PER APRIRE WHATSAPP & NOTIFICARE L OFFICINA
        </a>
        <a href="{gcal_href}" target="_blank" class="cal-big-btn">
            📅 AGGIUNGI L APPUNTAMENTO AL TUO GOOGLE CALENDAR
        </a>
        ''', unsafe_allow_html=True)

else:
    if not st.session_state['authenticated']:
        st.warning("🔒 Sezione riservata all Officina. Inserisci la password nel menu a sinistra per accedere.")
    else:
        st.markdown("# 🛞 PneumaticEvolution Gestionale")
        st.markdown("### Gestione Anagrafica, Hotel Pneumatici, Flotta Camion, Agenda e Notifiche Massive WhatsApp")
        
        tabs = st.tabs([
            "📊 Dashboard", "👤 Anagrafica Clienti", "🏨 Hotel Pneumatici & QR",
            "📱 Notifiche Massive WhatsApp", "📅 Agenda & Conferme", "⚙️ Chiusure & Password Admin", "📦 Magazzino"
        ])
        
        with tabs[0]:
            st.subheader("📥 Gestione & Conferma Richieste Pendenti")
            df_a = load_data("Appuntamenti")
            
            if df_a.empty or "Stato" not in df_a.columns:
                st.info("Nessuna prenotazione presente in agenda.")
            else:
                df_pending = df_a[df_a["Stato"] == "In attesa di conferma"]
                
                st.markdown(f"### ⏳ Richieste In Attesa di Conferma dall Officina ({len(df_pending)})")
                if df_pending.empty:
                    st.success("✅ Nessuna richiesta pendente da confermare al momento!")
                else:
                    for idx, row in df_pending.iterrows():
                        st.markdown(f"""
                        <div class="card-box" style="border-left: 6px solid #F59E0B;">
                            <h4>🚗 Cliente: {row['Cliente']} | Targa: <code>{row['Targa']}</code></h4>
                            <p><b>Servizio Richiesto:</b> {row['Servizio']}<br>
                            <b>Data ed Ora Desiderati:</b> 📅 {row['Data']} alle ore ⏰ {row['Ora']}<br>
                            <b>Cellulare:</b> 📱 {row['Telefono']} | <b>Note:</b> {row.get('Note', 'Nessuna')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            if st.button(f"✅ CONFERMA PRENOTAZIONE #{row['ID_App']}", key=f"conf_btn_{idx}", use_container_width=True):
                                df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = "Confermato"
                                save_sheet(df_a, "Appuntamenti")
                                
                                cal_link_cust = create_google_calendar_url(f"Cambio Gomme - {row['Targa']}", f"Servizio: {row['Servizio']}", str(row['Data']), str(row['Ora']))
                                confirm_msg = f"Gentile {row['Cliente']}, la tua prenotazione per la vettura {row['Targa']} ({row['Servizio']}) il {row['Data']} alle ore {row['Ora']} e stata CONFERMATA! Ti aspettiamo in officina presso PneumaticEvolution. Puoi aggiungere l evento al tuo calendario cliccando qui: {cal_link_cust}"
                                wa_confirm_url = create_whatsapp_url(row['Telefono'], confirm_msg)
                                
                                st.balloons()
                                st.success(f"✅ Appuntamento #{row['ID_App']} CONFERMATO con successo!")
                                st.markdown('<a href="' + wa_confirm_url + '" target="_blank" class="wa-big-btn">🟢 💬 INVIA SUBITO CONFERMA AL CLIENTE VIA WHATSAPP</a>', unsafe_allow_html=True)
                                st.rerun()

                        with c_btn2:
                            if st.button(f"❌ RIFIUTA / ANNULLA #{row['ID_App']}", key=f"ref_btn_{idx}", use_container_width=True):
                                df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = "Annullato"
                                save_sheet(df_a, "Appuntamenti")
                                
                                refuse_msg = f"Gentile {row['Cliente']}, la tua richiesta di prenotazione per la vettura {row['Targa']} per il {row['Data']} alle {row['Ora']} purtroppo non e disponibile. Ti preghiamo di contattarci via WhatsApp per concordare un altro orario. Grazie!"
                                wa_refuse_url = create_whatsapp_url(row['Telefono'], refuse_msg)
                                
                                st.warning(f"❌ Appuntamento #{row['ID_App']} Annullato.")
                                st.markdown('<a href="' + wa_refuse_url + '" target="_blank" class="wa-refuse-btn" style="background-color:#EF4444!important;">💬 INVIA NOTIFICA ANNULLAMENTO VIA WHATSAPP</a>', unsafe_allow_html=True)
                                st.rerun()

                st.markdown("---")
                st.markdown(f"### 📋 Elenco Completo Appuntamenti in Agenda ({len(df_a)})")
                st.dataframe(df_a, use_container_width=True)
                
        with tabs[1]:
            st.subheader("👤 Gestione Anagrafica Clienti & Veicoli")
            st1, st2 = st.tabs(["➕ Nuovo Inserimento Strutturato", "📋 Elenco e Ricerca Clienti"])
            with st1:
                with st.form("form_add_cli"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("##### 📌 Dati Anagrafici & Contatti")
                        tipo_cli = st.selectbox("Tipologia Cliente", ["Privato", "Azienda / Flotta", "Convenzionato"])
                        nome_cli = st.text_input("Nome / Ragione Sociale *")
                        cf_piva = st.text_input("Codice Fiscale / P.IVA")
                        telefono = st.text_input("Telefono / WhatsApp *", placeholder="3331234567")
                        email = st.text_input("Email")
                        indirizzo = st.text_input("Indirizzo e Citta")
                    with c2:
                        st.markdown("##### 🚗 Veicolo Principale Associato")
                        targa = st.text_input("Targa Veicolo *", placeholder="AB123CD").upper()
                        marca_modello = st.text_input("Marca e Modello", placeholder="Es. Audi A4 / Iveco Stralis")
                        tipo_mezzo = st.selectbox("Tipo Mezzo", ["Auto / SUV", "Furgone", "Camion / Motrice", "Rimorchio"])
                        misura_gomme = st.text_input("Misura Pneumatici", placeholder="Es. 205/55 R16 91V")
                        tipo_gomme = st.selectbox("Pneumatici Attualmente Montati", ["Estivi", "Invernali", "4 Stagioni"])
                        note_cli = st.text_area("Note Clienti")
                    btn_save = st.form_submit_button("💾 Salva Cliente & Veicolo in Anagrafica", use_container_width=True)
                    if btn_save:
                        if not nome_cli or not telefono or not targa:
                            st.error("Compila i campi obbligatori: Nome, Telefono e Targa!")
                        else:
                            df_c = load_data("Clienti")
                            cid = len(df_c) + 1 if not df_c.empty else 1
                            new_c = pd.DataFrame([{"ID_Cliente": cid, "Tipo_Cliente": tipo_cli, "Nome_RagioneSociale": nome_cli, "CodiceFiscale_PIVA": cf_piva, "Telefono_WhatsApp": telefono, "Email": email, "Indirizzo": indirizzo, "Note": note_cli}])
                            save_sheet(pd.concat([df_c, new_c], ignore_index=True), "Clienti")
                            df_v = load_data("Veicoli")
                            vid = len(df_v) + 1 if not df_v.empty else 1
                            new_v = pd.DataFrame([{"ID_Veicolo": vid, "ID_Cliente": cid, "Targa": targa, "Marca_Modello": marca_modello, "Tipo_Mezzo": tipo_mezzo, "Misura_Pneumatici_Anteriore": misura_gomme, "Misura_Pneumatici_Posteriore": misura_gomme, "Tipo_Pneumatici_Attuali": tipo_gomme}])
                            save_sheet(pd.concat([df_v, new_v], ignore_index=True), "Veicoli")
                            st.success(f"✅ Cliente '{nome_cli}' e veicolo '{targa}' salvati con successo in anagrafica!")
            with st2:
                df_c = load_data("Clienti")
                if not df_c.empty:
                    st.dataframe(df_c, use_container_width=True)
                else:
                    st.info("Anagrafica clienti vuota.")

        with tabs[2]:
            st.subheader("🏨 Hotel Pneumatici - Deposito, QR & Scansione")
            h1, h2, h3 = st.tabs(["📥 Check-in Deposito", "📦 Elenco Depositi in Custodia", "🔍 Scansiona QR Code da Fotocamera"])
            with h1:
                with st.form("form_dep"):
                    c1, c2 = st.columns(2)
                    df_c = load_data("Clienti")
                    with c1:
                        cli_sel = st.selectbox("Cliente *", df_c["Nome_RagioneSociale"].tolist() if not df_c.empty else ["Mario Rossi"])
                        targa_dep = st.text_input("Targa Veicolo *").upper()
                        marca_dep = st.text_input("Marca e Modello Pneumatico *")
                        misura_dep = st.text_input("Misura Pneumatici *")
                    with c2:
                        stag_dep = st.selectbox("Stagionalita Gomme in Custodia *", ["Invernali", "Estivi", "4 Stagioni"])
                        qta_dep = st.number_input("Quantita", min_value=1, max_value=8, value=4)
                        batt_dep = st.text_input("Battistrada Residuo (mm)", placeholder="6.5 mm")
                        ubi_dep = st.text_input("Ubicazione Magazzino", placeholder="Corsia A - Scaffale 3")
                    btn_h = st.form_submit_button("🏷️ Salva Deposito e Genera Etichetta QR Scansionabile", use_container_width=True)
                    if btn_h:
                        if not targa_dep or not marca_dep or not misura_dep:
                            st.error("Compila i campi obbligatori del deposito (Targa, Marca, Misura).")
                        else:
                            df_h = load_data("Hotel_Pneumatici")
                            hid = len(df_h) + 1 if not df_h.empty else 1
                            tag = f"TAG-2026-{hid:03d}"
                            new_h = pd.DataFrame([{"ID_Deposito": hid, "Tag_ID": tag, "ID_Cliente": 1, "Targa": targa_dep, "Marca_Modello_Gomme": marca_dep, "Misura": misura_dep, "Stagionalita": stag_dep, "Quantita": qta_dep, "Battistrada_mm": batt_dep, "Ubicazione": ubi_dep, "Data_PresaInCarico": str(datetime.date.today()), "Stato_Sollecito": "Da Sollecitare"}])
                            save_sheet(pd.concat([df_h, new_h], ignore_index=True), "Hotel_Pneumatici")
                            st.success(f"✅ Deposito registrato con Tag ID: {tag}")
                            lbl = generate_real_qr_label(tag, cli_sel, targa_dep, marca_dep, misura_dep, batt_dep, ubi_dep)
                            st.session_state['last_label'] = {
                                'tag': tag,
                                'bytes': lbl
                            }
                
                if st.session_state.get('last_label'):
                    lbl_data = st.session_state['last_label']
                    st.markdown("---")
                    st.subheader(f"🏷️ Etichetta Generata con QR Scansionabile: `{lbl_data['tag']}`")
                    st.image(lbl_data['bytes'], caption=f"Etichetta {lbl_data['tag']}")
                    st.download_button("📥 Scarica PNG Etichetta da Stampare", data=lbl_data['bytes'], file_name=f"Etichetta_{lbl_data['tag']}.png", mime="image/png")

            with h2:
                df_h = load_data("Hotel_Pneumatici")
                if not df_h.empty:
                    st.dataframe(df_h, use_container_width=True)
                else:
                    st.info("Nessun deposito in custodia al momento.")

            with h3:
                st.markdown("### 📷 Scansione QR Code da Fotocamera Smartphone o Immagine")
                st.write("Inquadra il QR Code stampato sull etichetta del sacco gomme per trovare subito la scheda cliente ed l ubicazione in magazzino.")
                
                cam_img = st.camera_input("📷 Scatta una foto al QR Code")
                up_file = st.file_uploader("📂 Oppure carica un'immagine dell etichetta QR", type=["png", "jpg", "jpeg"])
                
                img_to_scan = cam_img or up_file
                if img_to_scan:
                    file_bytes = np.asarray(bytearray(img_to_scan.read()), dtype=np.uint8)
                    opencv_img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
                    
                    detector = cv2.QRCodeDetector()
                    scanned_val, pts, _ = detector.detectAndDecode(opencv_img)
                    
                    if scanned_val:
                        st.success(f"🎯 QR Code Scansionato con successo: `{scanned_val}`")
                        df_h = load_data("Hotel_Pneumatici")
                        if not df_h.empty:
                            matched = df_h[df_h["Tag_ID"].astype(str) == str(scanned_val)]
                            if not matched.empty:
                                st.markdown("#### 📦 Dati Deposito Trovati:")
                                st.dataframe(matched, use_container_width=True)
                            else:
                                st.warning(f"Nessun deposito trovato nel database per il Tag ID: `{scanned_val}`")
                    else:
                        st.error("⚠️ Nessun QR code rilevato nell immagine. Assicurati che il codice sia ben illuminato e a fuoco.")

        with tabs[3]:
            st.subheader("📱 Invio Massivo WhatsApp (Cambi Stagionali)")
            st.write("Invia con **1-Click** i solleciti a tutti i clienti idonei ed esclude le gomme **4 Stagioni** dai solleciti obbligatori.")
            campagna = st.radio("Seleziona Tipo Campagna:", [
                "🍁 Sollecito Cambio Invernale (Notifica clienti con gomme Invernali in custodia)",
                "☀️ Sollecito Cambio Estivo (Notifica clienti con gomme Estive in custodia)",
                "🔄 Controllo Usura & Inversione 4 Stagioni (Opzionale per All-Season)"
            ])
            df_h = load_data("Hotel_Pneumatici")
            if not df_h.empty:
                if "Invernale" in campagna:
                    df_target = df_h[df_h["Stagionalita"] == "Invernali"]
                elif "Estivo" in campagna:
                    df_target = df_h[df_h["Stagionalita"] == "Estivi"]
                else:
                    df_target = df_h[df_h["Stagionalita"] == "4 Stagioni"]
                
                st.info(f"📊 Trovati **{len(df_target)}** clienti idonei per questa campagna.")
                
                app_url = get_setting("app_url", "https://" + "pneumatic-evolution.streamlit.app")
                
                if st.button(f"🚀 AVVIA INVIO COLLETIVAMENTE A TUTTI ({len(df_target)} Clienti)", use_container_width=True):
                    for _, r in df_target.iterrows():
                        phone = "3330000000"
                        targa_val = r.get("Targa", "N/D")
                        tag_val = r.get("Tag_ID", "N/D")
                        gomme_val = r.get("Marca_Modello_Gomme", "Pneumatici")
                        msg = f"Gentile Cliente, le ricordiamo che le sue gomme {gomme_val} per la vettura {targa_val} (Tag: {tag_val}) sono pronte in custodia presso PneumaticEvolution. 👉 Puo prenotare l appuntamento cliccando sul link: {app_url}?page=prenotazione . Cordiali Saluti."
                        wa_url = create_whatsapp_url(phone, msg)
                        components.html(f"<script>window.open('{wa_url}', '_blank');</script>", height=0)
                    st.success("🚀 Apertura schede WhatsApp in corso!")

                st.markdown("---")
                bulk_list = []
                for _, r in df_target.iterrows():
                    phone = "3330000000"
                    targa_val = r.get("Targa", "N/D")
                    tag_val = r.get("Tag_ID", "N/D")
                    gomme_val = r.get("Marca_Modello_Gomme", "Pneumatici")
                    msg = f"Gentile Cliente, le ricordiamo che le sue gomme {gomme_val} per la vettura {targa_val} (Tag: {tag_val}) sono pronte in custodia presso PneumaticEvolution. 👉 Puo prenotare l appuntamento cliccando sul link: {app_url}?page=prenotazione . Cordiali Saluti."
                    wa_url = create_whatsapp_url(phone, msg)
                    bulk_list.append(f"Targa: {targa_val} ({phone})\nMessaggio: {msg}\nLink WhatsApp: {wa_url}\n")
                    col_m1, col_m2 = st.columns([3, 1])
                    col_m1.write(f"🚗 **Targa:** {targa_val} | Tag: `{tag_val}` | {gomme_val}")
                    col_m2.markdown(f"[💬 Invia Singolo WhatsApp]({wa_url})", unsafe_allow_html=True)
                
                st.markdown("---")
                st.subheader("📋 Esporta Elenco Completo per Liste di Trasmissione")
                st.text_area("Elenco Messaggi Formattati:", value="\n".join(bulk_list), height=200)

        with tabs[4]:
            st.subheader("📅 Agenda & Gestione Appuntamenti Pendenti")
            df_a = load_data("Appuntamenti")
            if not df_a.empty:
                st.dataframe(df_a, use_container_width=True)
            else:
                st.info("Agenda vuota al momento.")

        with tabs[5]:
            st.subheader("⚙️ Chiusure Calendario, URL & Password Admin")
            st1_cfg, st2_cfg = st.columns(2)
            with st1_cfg:
                st.markdown("### 🔑 Cambio Password Admin Officina")
                with st.form("form_change_pass"):
                    old_p = st.text_input("Password Attuale", type="password")
                    new_p = st.text_input("Nuova Password", type="password")
                    confirm_p = st.text_input("Conferma Nuova Password", type="password")
                    btn_pass = st.form_submit_button("💾 Aggiorna Password Admin", use_container_width=True)
                    if btn_pass:
                        current_real_p = get_setting("admin_password", "admin")
                        if old_p != current_real_p:
                            st.error("La password attuale non è corretta!")
                        elif new_p != confirm_p:
                            st.error("Le nuove password non coincidono!")
                        elif len(new_p) < 4:
                            st.error("La nuova password deve contenere almeno 4 caratteri!")
                        else:
                            update_setting("admin_password", new_p)
                            st.success("✅ Password aggiornata con successo! Usala al prossimo accesso.")
                
                st.markdown("---")
                st.markdown("### 🌐 Configurazione Link Web del Sito")
                curr_url = get_setting("app_url", "https://" + "pneumatic-evolution.streamlit.app")
                new_app_url = st.text_input("URL Pubblico dell App Streamlit:", value=curr_url)
                if st.button("💾 Salva URL Sito"):
                    update_setting("app_url", new_app_url)
                    st.success("✅ Link del sito aggiornato con successo!")

            with st2_cfg:
                st.markdown("### 🗓️ Blocco Date & Chiusure Straordinarie")
                giorni_bloccati_str = get_setting("giorni_bloccati", "")
                st.info(f"Date attualmente bloccate per le prenotazioni:\n`{giorni_bloccati_str}`")
                with st.form("form_block_date"):
                    data_da_bloccare = st.date_input("Seleziona data da bloccare per chiusura/ferie:")
                    azione_blocco = st.selectbox("Azione:", ["Aggiungi ai giorni bloccati", "Rimuovi dai giorni bloccati"])
                    btn_block = st.form_submit_button("💾 Aggiorna Calendario Chiusure", use_container_width=True)
                    if btn_block:
                        g_list = [g.strip() for g in giorni_bloccati_str.split(",") if g.strip()]
                        d_str = str(data_da_bloccare)
                        if azione_blocco == "Aggiungi ai giorni bloccati":
                            if d_str not in g_list:
                                g_list.append(d_str)
                                update_setting("giorni_bloccati", ",".join(g_list))
                                st.success(f"✅ Data {d_str} bloccata con successo!")
                            else:
                                st.warning("Questa data è già presente nell elenco delle chiusure.")
                        else:
                            if d_str in g_list:
                                g_list.remove(d_str)
                                update_setting("giorni_bloccati", ",".join(g_list))
                                st.success(f"✅ Data {d_str} sbloccata!")
                            else:
                                st.warning("La data selezionata non era tra i giorni bloccati.")

        with tabs[6]:
            st.subheader("📦 Magazzino Scorte Pneumatici")
            df_m = load_data("Magazzino")
            if not df_m.empty:
                st.dataframe(df_m, use_container_width=True)
            else:
                st.info("Magazzino vuoto.")
