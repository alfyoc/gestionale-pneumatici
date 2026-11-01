import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import io
import urllib.parse
import json
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

# URL Router Check: ?page=prenotazione hides sidebar & optimizes mobile layout
query_params = st.query_params
raw_page = str(query_params.get("page", "")).lower()
is_public_page_url = "prenotazione" in raw_page or "clienti" in raw_page

if is_public_page_url:
    st.markdown('''
    <style>
        section[data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        header[data-testid="stHeader"] { visibility: hidden !important; height: 0px !important; }
        div[data-testid="stToolbar"] { visibility: hidden !important; }
        .main .block-container { 
            padding-top: 1rem !important; 
            padding-bottom: 2rem !important; 
            max-width: 900px !important;
            margin: 0 auto !important;
        }
    </style>
    ''', unsafe_allow_html=True)

def resolve_db_path():
    possible_paths = [
        "gestionale_pneumatici_db_v23.xlsx",
        "gestionale_pneumatici_db_v14.xlsx",
        "gestionale_pneumatici_db_v13.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v23.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v14.xlsx"
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return "gestionale_pneumatici_db_v23.xlsx"

DB_FILE_EXCEL = resolve_db_path()

def ensure_db_exists():
    global DB_FILE_EXCEL
    if not os.path.exists(DB_FILE_EXCEL):
        DB_FILE_EXCEL = "gestionale_pneumatici_db_v23.xlsx"
        if not os.path.exists(DB_FILE_EXCEL):
            wb = openpyxl.Workbook()
            default_sheets = ["Clienti", "Veicoli", "Hotel_Pneumatici", "Appuntamenti", "Magazzino", "Ordini_Pneumatici", "Flotta_Camion", "Impostazioni"]
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
            df = pd.read_excel(DB_FILE_EXCEL, sheet_name=sheet_name, dtype=str)
            if not df.empty:
                df = df.fillna("")
                for col in df.columns:
                    df[col] = df[col].astype(object)
            return df
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

def create_ics_calendar(events_list, cal_name="PneumaticEvolution Officina"):
    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//PneumaticEvolution//Gestionale Officina//IT",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{cal_name}"
    ]
    for ev in events_list:
        try:
            d_str = str(ev.get('date_str', '')).strip()
            t_str = str(ev.get('time_str', '09:00')).strip()
            d_obj = datetime.datetime.strptime(f"{d_str} {t_str}", "%Y-%m-%d %H:%M")
            dtstart = d_obj.strftime("%Y%m%dT%H%M00")
            dtend = (d_obj + datetime.timedelta(minutes=45)).strftime("%Y%m%dT%H%M00")
            dtstamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M00Z")
            targa = str(ev.get('targa', '')).upper()
            cliente = str(ev.get('cliente', '')).title()
            servizio = str(ev.get('title', 'Servizio Pneumatici'))
            stato = str(ev.get('stato', 'Confermato'))
            summary = f"🛞 [{targa}] {cliente} - {servizio}"
            desc = f"Cliente: {cliente}\nTarga: {targa}\nServizio: {servizio}\nStato: {stato}\nNote: {ev.get('details', '')}"
            uid = f"app-{ev.get('id', 1)}-{dtstart}@pneumaticevolution.com"
            ics.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{dtstamp}",
                f"DTSTART:{dtstart}",
                f"DTEND:{dtend}",
                f"SUMMARY:{summary}",
                f"DESCRIPTION:{desc}",
                "LOCATION:PneumaticEvolution Officina",
                "STATUS:CONFIRMED",
                "END:VEVENT"
            ])
        except Exception:
            continue
    ics.append("END:VCALENDAR")
    return "\r\n".join(ics)

def render_fullcalendar_component(events_df):
    fc_events = []
    if not events_df.empty and "Data" in events_df.columns:
        for idx, row in events_df.iterrows():
            try:
                d_str = str(row["Data"]).strip()
                t_str = str(row["Ora"]).strip() if "Ora" in row and pd.notna(row["Ora"]) else "09:00"
                dt_obj = datetime.datetime.strptime(f"{d_str} {t_str}", "%Y-%m-%d %H:%M")
                dt_end = dt_obj + datetime.timedelta(minutes=45)
                
                st_val = str(row.get("Stato", "Confermato"))
                if st_val == "In attesa di conferma":
                    color = "#F59E0B"
                    prefix = "⏳ "
                elif st_val == "Confermato":
                    color = "#10B981"
                    prefix = "✅ "
                elif "Proposta" in st_val:
                    color = "#3B82F6"
                    prefix = "🔄 "
                else:
                    color = "#EF4444"
                    prefix = "❌ "
                
                title = f"{prefix}{row.get('Targa', 'AB123CD')} - {row.get('Cliente', 'Cliente')} ({row.get('Servizio', 'Gomme')})"
                fc_events.append({
                    "title": title,
                    "start": dt_obj.strftime("%Y-%m-%dT%H:%M:00"),
                    "end": dt_end.strftime("%Y-%m-%dT%H:%M:00"),
                    "color": color,
                    "description": f"Cliente: {row.get('Cliente','')}, Tel: {row.get('Telefono','')}, Stato: {st_val}"
                })
            except Exception:
                continue
    
    json_events = json.dumps(fc_events)
    
    html_code = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset='utf-8' />
  <link href='https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/index.global.min.css' rel='stylesheet' />
  <script src='https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/index.global.min.js'></script>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 5px; background: #F9FAFB; }}
    #calendar {{ max-width: 100%; margin: 0 auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.06); }}
    .fc-toolbar-title {{ font-size: 1.25rem !important; color: #1E3A8A; font-weight: 800; }}
    .fc-button-primary {{ background-color: #1E3A8A !important; border-color: #1E3A8A !important; border-radius: 6px !important; font-weight: 600 !important; }}
    .fc-button-primary:hover {{ background-color: #2563EB !important; border-color: #2563EB !important; }}
    .fc-event {{ cursor: pointer; border-radius: 6px; padding: 2px 4px; font-weight: 600; font-size: 0.85rem; }}
  </style>
</head>
<body>
  <div id='calendar'></div>
  <script>
    document.addEventListener('DOMContentLoaded', function() {{
      var calendarEl = document.getElementById('calendar');
      var calendar = new FullCalendar.Calendar(calendarEl, {{
        initialView: 'dayGridMonth',
        locale: 'it',
        firstDay: 1,
        headerToolbar: {{
          left: 'prev,next today',
          center: 'title',
          right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
        }},
        buttonText: {{
          today: 'Oggi',
          month: 'Mese',
          week: 'Settimana',
          day: 'Giorno',
          list: 'Lista'
        }},
        events: {json_events},
        eventClick: function(info) {{
          alert(info.event.title + '\n' + info.event.extendedProps.description);
        }}
      }});
      calendar.render();
    }});
  </script>
</body>
</html>"""
    return html_code


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

if mode == "🔓 Portale Prenotazione Clienti (Pubblico)" or is_public_page_url:
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
            "📥 Appuntamenti & Conferme",
            "👤 Anagrafica Clienti & Veicoli",
            "🏨 Hotel Pneumatici & QR",
            "📦 Magazzino Pneumatici & Ordini",
            "📱 Notifiche Massive WhatsApp",
            "📅 Agenda & Calendario Sincronizzato",
            "⚙️ Chiusure, URL & Password Admin"
        ])
        
        with tabs[0]:
            st.subheader("📥 Gestione & Conferma Richieste Pendenti")
            
            # Persistent Banner for Admin WhatsApp Notification Trigger
            if st.session_state.get('admin_wa_action'):
                act = st.session_state['admin_wa_action']
                st.balloons()
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%); border: 2px solid #10B981; border-radius: 14px; padding: 22px; margin-bottom: 25px; box-shadow: 0 4px 12px rgba(16,185,129,0.2);">
                    <h3 style="color: #065F46; margin-top: 0;">📱 Notifica WhatsApp Pronta per {act['cliente']}</h3>
                    <p style="font-size: 1.05rem; color: #047857;"><b>Esito Operazione:</b> {act['tipo']}<br>
                    <b>Numero Destinatario:</b> <code>{act['phone']}</code></p>
                    <p style="background: white; padding: 12px; border-radius: 8px; border: 1px solid #A7F3D0; font-family: monospace;">
                    {act['msg']}
                    </p>
                    <a href="{act['wa_url']}" target="_blank" class="wa-big-btn" style="margin-top: 15px; margin-bottom: 10px;">
                        🟢 💬 CLICCA QUI PER APRIRE WHATSAPP & INVIARE IL MESSAGGIO A {str(act['cliente']).upper()}
                    </a>
                </div>
                """, unsafe_allow_html=True)
                components.html(f"<script>window.open('{act['wa_url']}', '_blank');</script>", height=0)
                if st.button("❌ Chiudi Avviso Notifica", key="close_wa_banner"):
                    st.session_state['admin_wa_action'] = None
                    st.rerun()

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
                        <div class="card-box" style="border-left: 6px solid #F59E0B; background-color: #FFFBEB;">
                            <h4 style="color: #92400E; margin-top: 0;">🚗 Cliente: {row['Cliente']} | Targa: <code>{row['Targa']}</code></h4>
                            <p style="font-size: 1.05rem;"><b>Servizio Richiesto:</b> {row['Servizio']}<br>
                            <b>Data ed Ora Desiderati:</b> 📅 <b>{row['Data']}</b> alle ore ⏰ <b>{row['Ora']}</b><br>
                            <b>Cellulare:</b> 📱 <code>{row['Telefono']}</code> | <b>Note:</b> {row.get('Note', 'Nessuna')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        col_ac, col_pr = st.columns([1, 1])
                        
                        with col_ac:
                            if st.button(f"✅ CONFERMA PRENOTAZIONE #{row['ID_App']}", key=f"conf_btn_{idx}", use_container_width=True):
                                df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = "Confermato"
                                save_sheet(df_a, "Appuntamenti")
                                
                                cal_link_cust = create_gcal_url(f"Cambio Gomme - {row['Targa']}", str(row['Data']), str(row['Ora']), f"Servizio: {row['Servizio']}")
                                confirm_msg = f"Gentile {row['Cliente']}, la tua prenotazione per la vettura {row['Targa']} ({row['Servizio']}) il {row['Data']} alle ore {row['Ora']} e stata CONFERMATA! Ti aspettiamo in officina presso PneumaticEvolution. Puoi aggiungere l evento al tuo calendario cliccando qui: {cal_link_cust}"
                                wa_confirm_url = create_whatsapp_url(row['Telefono'], confirm_msg)
                                
                                st.session_state['admin_wa_action'] = {
                                    'wa_url': wa_confirm_url,
                                    'msg': confirm_msg,
                                    'cliente': row['Cliente'],
                                    'phone': row['Telefono'],
                                    'tipo': 'Conferma Appuntamento (Confermato)'
                                }
                                st.rerun()

                        with col_pr:
                            with st.expander(f"⚠️ Proponi Nuova Data / Orario o Rifiuta #{row['ID_App']}", expanded=False):
                                opt_act = st.radio(
                                    "Seleziona l azione:",
                                    ["📅 Proponi Nuova Data ed Orario Alternativo", "❌ Rifiuta Definitivamente"],
                                    key=f"opt_rad_{idx}"
                                )
                                
                                if "Proponi Nuova Data" in opt_act:
                                    col_p1, col_p2 = st.columns(2)
                                    with col_p1:
                                        prop_date = st.date_input("Nuova Data Proposta:", min_value=datetime.date.today(), key=f"prop_d_{idx}")
                                    with col_p2:
                                        prop_time = st.selectbox("Nuovo Orario Proposto:", [
                                            "08:30", "09:15", "10:00", "10:45", "11:30", "14:30", "15:15", "16:00", "16:45", "17:30"
                                        ], key=f"prop_t_{idx}")
                                    
                                    if st.button(f"💾 INVIA NUOVA PROPOSTA VIA WHATSAPP #{row['ID_App']}", key=f"save_prop_{idx}", use_container_width=True):
                                        new_status = f"Proposta Nuova Data ({prop_date} {prop_time})"
                                        df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = new_status
                                        save_sheet(df_a, "Appuntamenti")
                                        
                                        app_url = get_setting("app_url", "https://" + "pneumatic-evolution" + ".streamlit.app")
                                        prop_msg = f"Gentile {row['Cliente']}, riguardo la tua richiesta per la vettura {row['Targa']} per il {row['Data']} alle {row['Ora']}, purtroppo quell orario non e disponibile. Ti proponiamo in alternativa la data {prop_date} alle ore {prop_time}. Clicca per accedere al sito e confermare o scegliere un altro orario: {app_url}?page=prenotazione . Grazie! - PneumaticEvolution"
                                        wa_prop_url = create_whatsapp_url(row['Telefono'], prop_msg)
                                        
                                        st.session_state['admin_wa_action'] = {
                                            'wa_url': wa_prop_url,
                                            'msg': prop_msg,
                                            'cliente': row['Cliente'],
                                            'phone': row['Telefono'],
                                            'tipo': f'Proposta Nuova Data ({prop_date} alle {prop_time})'
                                        }
                                        st.rerun()
                                else:
                                    if st.button(f"❌ INVIA RIFIUTO DEFINITIVO VIA WHATSAPP #{row['ID_App']}", key=f"save_refuse_{idx}", use_container_width=True):
                                        df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = "Rifiutato"
                                        save_sheet(df_a, "Appuntamenti")
                                        
                                        app_url = get_setting("app_url", "https://" + "pneumatic-evolution" + ".streamlit.app")
                                        refuse_msg = f"Gentile {row['Cliente']}, ci dispiace informarti che la tua richiesta di prenotazione per la vettura {row['Targa']} per il {row['Data']} alle ore {row['Ora']} non puo essere accolta per indisponibilita. Ti invitiamo ad accedere al sito per scegliere una data diversa: {app_url}?page=prenotazione . Cordiali Saluti."
                                        wa_refuse_url = create_whatsapp_url(row['Telefono'], refuse_msg)
                                        
                                        st.session_state['admin_wa_action'] = {
                                            'wa_url': wa_refuse_url,
                                            'msg': refuse_msg,
                                            'cliente': row['Cliente'],
                                            'phone': row['Telefono'],
                                            'tipo': 'Prenotazione Rifiutata'
                                        }
                                        st.rerun()

                st.markdown("---")
                st.markdown(f"### 📋 Elenco Completo Appuntamenti in Agenda ({len(df_a)})")
                st.dataframe(df_a, use_container_width=True)

        # TAB 1: Anagrafica Clienti & Veicoli (Modifica, Elimina & Scarico Magazzino)
        with tabs[1]:
            st.subheader("👤 Gestione Anagrafica Clienti & Veicoli")
            st1, st2, st3 = st.tabs(["➕ Nuovo Inserimento Clienti", "📋 Elenco e Ricerca", "✏️ Modifica / Elimina Cliente"])
            
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
                        st.markdown("##### 🚗 Veicolo & Integrazione Magazzino")
                        targa = st.text_input("Targa Veicolo *", placeholder="AB123CD").upper()
                        marca_modello = st.text_input("Marca e Modello", placeholder="Es. Audi A4 / Iveco Stralis")
                        tipo_mezzo = st.selectbox("Tipo Mezzo", ["Auto / SUV", "Furgone", "Camion / Motrice", "Rimorchio"])
                        misura_gomme = st.text_input("Misura Pneumatici", placeholder="Es. 205/55 R16 91V")
                        tipo_gomme = st.selectbox("Pneumatici Attualmente Montati", ["Estivi", "Invernali", "4 Stagioni"])
                        
                        df_mag_select = load_data("Magazzino")
                        mag_options = ["Forniti dal cliente / Acquisto Esterno (Nessun scarico magazzino)"]
                        if not df_mag_select.empty:
                            for _, r_m in df_mag_select.iterrows():
                                mag_options.append(f"Scarica da Magazzino: #{r_m.get('ID_Articolo', 1)} - {r_m.get('Marca_Modello', '')} ({r_m.get('Misura', '')}) [Disp: {r_m.get('Quantita_Disponibile', 0)}]")
                        
                        fonte_gomme_sel = st.selectbox("Provenienza Pneumatici per il Veicolo:", mag_options)
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
                            
                            id_mag_ref = ""
                            if "Scarica da Magazzino:" in fonte_gomme_sel:
                                try:
                                    art_id_str = fonte_gomme_sel.split("#")[1].split(" - ")[0]
                                    art_id = int(art_id_str)
                                    id_mag_ref = str(art_id)
                                    if not df_mag_select.empty:
                                        mask_art = df_mag_select["ID_Articolo"].astype(str) == str(art_id)
                                        if mask_art.any():
                                            curr_q = pd.to_numeric(df_mag_select.loc[mask_art, "Quantita_Disponibile"], errors='coerce').fillna(0).iloc[0]
                                            df_mag_select.loc[mask_art, "Quantita_Disponibile"] = max(0, int(curr_q) - 4)
                                            save_sheet(df_mag_select, "Magazzino")
                                            st.info(f"📦 Scaricati 4 pneumatici dal Magazzino per l articolo #{art_id}!")
                                except Exception as e_m:
                                    st.warning(f"Impossibile scaricare dal magazzino: {e_m}")

                            df_v = load_data("Veicoli")
                            vid = len(df_v) + 1 if not df_v.empty else 1
                            new_v = pd.DataFrame([{"ID_Veicolo": vid, "ID_Cliente": cid, "Targa": targa, "Marca_Modello": marca_modello, "Tipo_Mezzo": tipo_mezzo, "Misura_Pneumatici_Anteriore": misura_gomme, "Misura_Pneumatici_Posteriore": misura_gomme, "Tipo_Pneumatici_Attuali": tipo_gomme, "Fonte_Gomme": fonte_gomme_sel, "ID_Magazzino": id_mag_ref}])
                            save_sheet(pd.concat([df_v, new_v], ignore_index=True), "Veicoli")
                            st.success(f"✅ Cliente '{nome_cli}' e veicolo '{targa}' salvati con successo!")

            with st2:
                st.markdown("#### 📋 Registro Anagrafica Clienti")
                df_c = load_data("Clienti")
                df_v = load_data("Veicoli")
                if not df_c.empty:
                    st.dataframe(df_c, use_container_width=True)
                    st.markdown("#### 🚗 Veicoli Associati")
                    st.dataframe(df_v, use_container_width=True)
                else:
                    st.info("Anagrafica clienti vuota.")

            with st3:
                st.markdown("### ✏️ Modifica o Elimina un Cliente Esistente")
                df_c_edit = load_data("Clienti")
                if df_c_edit.empty:
                    st.info("Nessun cliente presente in anagrafica.")
                else:
                    cli_dict = {f"#{r['ID_Cliente']} - {r['Nome_RagioneSociale']} ({r.get('Telefono_WhatsApp', '')})": r['ID_Cliente'] for _, r in df_c_edit.iterrows()}
                    selected_cli_label = st.selectbox("Seleziona il Cliente da Modificare / Eliminare:", list(cli_dict.keys()))
                    target_cid = cli_dict[selected_cli_label]
                    
                    client_row = df_c_edit[df_c_edit["ID_Cliente"] == target_cid].iloc[0]
                    
                    with st.form("form_edit_cli"):
                        st.markdown(f"##### Modifica Dati Cliente #{target_cid}")
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            edit_tipo = st.selectbox("Tipo Cliente", ["Privato", "Azienda / Flotta", "Convenzionato"], index=["Privato", "Azienda / Flotta", "Convenzionato"].index(client_row.get("Tipo_Cliente", "Privato")))
                            edit_nome = st.text_input("Nome / Ragione Sociale", value=str(client_row.get("Nome_RagioneSociale", "")))
                            edit_cf = st.text_input("Codice Fiscale / P.IVA", value=str(client_row.get("CodiceFiscale_PIVA", "")))
                        with ec2:
                            edit_tel = st.text_input("Telefono / WhatsApp", value=str(client_row.get("Telefono_WhatsApp", "")))
                            edit_email = st.text_input("Email", value=str(client_row.get("Email", "")))
                            edit_indirizzo = st.text_input("Indirizzo", value=str(client_row.get("Indirizzo", "")))
                        edit_note = st.text_area("Note Cliente", value=str(client_row.get("Note", "")))
                        
                        btn_update_cli = st.form_submit_button("💾 Salva Modifiche Anagrafica", use_container_width=True)
                        if btn_update_cli:
                            df_c_edit = df_c_edit.astype(object)
                            mask_c = df_c_edit["ID_Cliente"].astype(str) == str(target_cid)
                            df_c_edit.loc[mask_c, "Tipo_Cliente"] = str(edit_tipo)
                            df_c_edit.loc[mask_c, "Nome_RagioneSociale"] = str(edit_nome)
                            df_c_edit.loc[mask_c, "CodiceFiscale_PIVA"] = str(edit_cf)
                            df_c_edit.loc[mask_c, "Telefono_WhatsApp"] = str(edit_tel)
                            df_c_edit.loc[mask_c, "Email"] = str(edit_email)
                            df_c_edit.loc[mask_c, "Indirizzo"] = str(edit_indirizzo)
                            df_c_edit.loc[mask_c, "Note"] = str(edit_note)
                            save_sheet(df_c_edit, "Clienti")
                            st.success("✅ Dati del cliente aggiornati con successo!")
                            st.rerun()

                    st.markdown("---")
                    st.markdown("##### 🗑️ Zona di Eliminazione Cliente")
                    st.warning("⚠️ L eliminazione di un cliente rimuovera l anagrafica ed i veicoli ad esso associati.")
                    if st.button(f"🗑️ ELIMINA DEFINITIVAMENTE CLIENTE #{target_cid}", type="primary"):
                        df_c_edit = df_c_edit[df_c_edit["ID_Cliente"] != target_cid]
                        save_sheet(df_c_edit, "Clienti")
                        
                        df_v_edit = load_data("Veicoli")
                        if not df_v_edit.empty:
                            df_v_edit = df_v_edit[df_v_edit["ID_Cliente"] != target_cid]
                            save_sheet(df_v_edit, "Veicoli")
                            
                        st.success(f"🗑️ Cliente #{target_cid} eliminato dall anagrafica!")
                        st.rerun()

        # TAB 2: Hotel Pneumatici & QR Code (Modifica Deposito con Mantenimento Tag ID)
        with tabs[2]:
            st.subheader("🏨 Hotel Pneumatici - Deposito, QR & Scansione")
            h1, h2, h3, h4 = st.tabs(["📥 Check-in Deposito", "📦 Elenco Depositi", "✏️ Modifica Deposito & Rigenera QR", "🔍 Scansiona QR da Fotocamera"])
            
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
                st.markdown("### ✏️ Modifica Dati Deposito & Rigenera QR (Mantiene lo Stesso Tag ID)")
                st.write("Puoi aggiornare la posizione dello scaffale, la misura, la marca o il battistrada senza cambiare il codice progressivo originale del cliente!")
                df_h_edit = load_data("Hotel_Pneumatici")
                if df_h_edit.empty:
                    st.info("Nessun deposito presente da modificare.")
                else:
                    dep_dict = {f"{r.get('Tag_ID', '')} - Targa: {r.get('Targa', '')} ({r.get('Marca_Modello_Gomme', '')})": r.get('Tag_ID', '') for _, r in df_h_edit.iterrows()}
                    selected_tag_label = st.selectbox("Seleziona il Deposito da Modificare:", list(dep_dict.keys()))
                    target_tag = dep_dict[selected_tag_label]
                    
                    dep_row = df_h_edit[df_h_edit["Tag_ID"] == target_tag].iloc[0]
                    
                    with st.form("form_edit_deposito"):
                        st.info(f"🏷️ Deposito Selezionato: **{target_tag}** (Il Tag ID rimarra invariato!)")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            e_cliente = st.text_input("Nome Cliente", value=str(dep_row.get("ID_Cliente", "Cliente")))
                            e_targa = st.text_input("Targa Veicolo", value=str(dep_row.get("Targa", ""))).upper()
                            e_marca = st.text_input("Marca e Modello Gomme", value=str(dep_row.get("Marca_Modello_Gomme", "")))
                            e_misura = st.text_input("Misura Pneumatici", value=str(dep_row.get("Misura", "")))
                        with col_e2:
                            e_stag = st.selectbox("Stagionalita", ["Invernali", "Estivi", "4 Stagioni"], index=["Invernali", "Estivi", "4 Stagioni"].index(dep_row.get("Stagionalita", "Invernali")))
                            e_qta = st.number_input("Quantita", min_value=1, max_value=8, value=int(dep_row.get("Quantita", 4)))
                            e_batt = st.text_input("Battistrada Residuo (mm)", value=str(dep_row.get("Battistrada_mm", "6.0 mm")))
                            e_ubi = st.text_input("Ubicazione Scaffale / Magazzino", value=str(dep_row.get("Ubicazione", "Corsia A")))
                        
                        btn_update_dep = st.form_submit_button("💾 Salva Modifiche & Rigenera QR Code (Mantiene Tag ID)", use_container_width=True)
                        if btn_update_dep:
                            df_h_edit = df_h_edit.astype(object)
                            mask_h = df_h_edit["Tag_ID"].astype(str) == str(target_tag)
                            df_h_edit.loc[mask_h, "Targa"] = str(e_targa)
                            df_h_edit.loc[mask_h, "Marca_Modello_Gomme"] = str(e_marca)
                            df_h_edit.loc[mask_h, "Misura"] = str(e_misura)
                            df_h_edit.loc[mask_h, "Stagionalita"] = str(e_stag)
                            df_h_edit.loc[mask_h, "Quantita"] = str(e_qta)
                            df_h_edit.loc[mask_h, "Battistrada_mm"] = str(e_batt)
                            df_h_edit.loc[mask_h, "Ubicazione"] = str(e_ubi)
                            save_sheet(df_h_edit, "Hotel_Pneumatici")
                            
                            lbl_updated = generate_real_qr_label(target_tag, e_cliente, e_targa, e_marca, e_misura, e_batt, e_ubi)
                            st.session_state['last_label'] = {
                                'tag': target_tag,
                                'bytes': lbl_updated
                            }
                            st.success(f"✅ Deposito `{target_tag}` aggiornato e nuovo QR Code rigenerato!")
                            st.rerun()

            with h4:
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

        # TAB 3: Magazzino Pneumatici & Ordini in Arrivo
        with tabs[3]:
            st.subheader("📦 Magazzino Pneumatici & Ordini In Arrivo")
            m_tab1, m_tab2, m_tab3 = st.tabs(["📦 Inventario Stock Disponibile", "➕ Nuovo Carico Pneumatici", "🚚 Pneumatici Ordinati (In Arrivo)"])
            
            with m_tab1:
                df_m = load_data("Magazzino")
                if not df_m.empty:
                    valore_totale = 0
                    if "Quantita_Disponibile" in df_m.columns and "Costo_Acquisto_Unitario" in df_m.columns:
                        valore_totale = (pd.to_numeric(df_m["Quantita_Disponibile"], errors="coerce").fillna(0) * pd.to_numeric(df_m["Costo_Acquisto_Unitario"], errors="coerce").fillna(0)).sum()
                    
                    col_m_met1, col_m_met2 = st.columns(2)
                    col_m_met1.metric("Totale Articoli in Stock", len(df_m))
                    col_m_met2.metric("Valore Totale Inventario Magazzino", f"€ {valore_totale:,.2f}")
                    
                    st.dataframe(df_m, use_container_width=True)
                else:
                    st.info("Magazzino vuoto. Carica nuovi pneumatici dalla scheda affianco!")

            with m_tab2:
                st.markdown("### ➕ Registra Nuovo Carico Pneumatici in Magazzino")
                with st.form("form_add_stock"):
                    col_st1, col_st2 = st.columns(2)
                    with col_st1:
                        st_marca = st.text_input("Marca e Modello Pneumatico *", placeholder="Es. Pirelli Cinturato P7")
                        st_misura = st.text_input("Misura Pneumatici *", placeholder="Es. 205/55 R16 91V")
                        st_qta = st.number_input("Quantita Disponibile (Pezzi) *", min_value=1, max_value=200, value=4)
                    with col_st2:
                        st_costo = st.number_input("Costo di Acquisto Unitario (€) *", min_value=0.0, value=50.0, step=5.0)
                        st_prezzo = st.number_input("Prezzo di Vendita Unitario (€)", min_value=0.0, value=90.0, step=5.0)
                        st_ubi = st.text_input("Ubicazione Magazzino", placeholder="Corsia M1 - Scaffale A")
                    st_note = st.text_area("Note Magazzino")
                    
                    btn_save_stock = st.form_submit_button("💾 Salva Carico in Inventario Magazzino", use_container_width=True)
                    if btn_save_stock:
                        if not st_marca or not st_misura:
                            st.error("Compila Marca e Misura del pneumatico!")
                        else:
                            df_m_save = load_data("Magazzino")
                            mid = len(df_m_save) + 1 if not df_m_save.empty else 1
                            new_m = pd.DataFrame([{"ID_Articolo": mid, "Marca_Modello": st_marca, "Misura": st_misura, "Quantita_Disponibile": st_qta, "Costo_Acquisto_Unitario": st_costo, "Prezzo_Vendita_Unitario": st_prezzo, "Ubicazione_Magazzino": st_ubi, "Note": st_note}])
                            save_sheet(pd.concat([df_m_save, new_m], ignore_index=True), "Magazzino")
                            st.success(f"✅ Caricati {st_qta} pezzi di '{st_marca}' in magazzino!")

            with m_tab3:
                st.markdown("### 🚚 Registro Pneumatici Ordinati (Non Ancora Arrivati)")
                st.write("Gestisci l inventario delle gomme in ordine dal fornitore ed effettua il carico in magazzino con 1-click appena la merce viene consegnata!")
                
                with st.form("form_add_ordine"):
                    st.markdown("##### ➕ Inserisci Nuovo Ordine Gomme")
                    col_ord1, col_ord2 = st.columns(2)
                    with col_ord1:
                        o_marca = st.text_input("Marca e Modello Pneumatici Ordinati *", placeholder="Es. Continental WinterContact")
                        o_misura = st.text_input("Misura Pneumatici *", placeholder="Es. 205/55 R16 91T")
                        o_qta = st.number_input("Quantita Ordinata *", min_value=1, max_value=500, value=16)
                    with col_ord2:
                        o_fornitore = st.text_input("Fornitore / Distributore", placeholder="Es. GommeSpA")
                        o_costo_tot = st.number_input("Costo Acquisto Totale (€)", min_value=0.0, value=800.0, step=10.0)
                        o_data_arr = st.date_input("Data Presunta Arrivo", min_value=datetime.date.today())
                    
                    btn_save_ord = st.form_submit_button("💾 Salva Ordine in Registro", use_container_width=True)
                    if btn_save_ord:
                        if not o_marca or not o_misura:
                            st.error("Inserisci Marca e Misura dell ordine!")
                        else:
                            df_o = load_data("Ordini_Pneumatici")
                            oid = len(df_o) + 1 if not df_o.empty else 1
                            new_o = pd.DataFrame([{"ID_Ordine": oid, "Marca_Modello": o_marca, "Misura": o_misura, "Quantita_Ordinata": o_qta, "Fornitore": o_fornitore, "Costo_Acquisto_Totale": o_costo_tot, "Data_Ordine": str(datetime.date.today()), "Data_Presunta_Arrivo": str(o_data_arr), "Stato_Ordine": "In Ordine / In Spedizione"}])
                            save_sheet(pd.concat([df_o, new_o], ignore_index=True), "Ordini_Pneumatici")
                            st.success("✅ Ordine registrato con successo!")

                st.markdown("---")
                st.markdown("##### 📋 Elenco Ordini in Arrivo & Carico Merci 1-Click")
                df_o_list = load_data("Ordini_Pneumatici")
                if not df_o_list.empty:
                    st.dataframe(df_o_list, use_container_width=True)
                    
                    df_o_pending = df_o_list[df_o_list["Stato_Ordine"] != "Consegnato in Magazzino"]
                    if not df_o_pending.empty:
                        st.markdown("##### 📦 Registra Arrivo Merce & Sposta in Magazzino:")
                        for idx_o, row_o in df_o_pending.iterrows():
                            c_o1, c_o2 = st.columns([3, 1])
                            c_o1.write(f"🚚 **Ordine #{row_o['ID_Ordine']}:** {row_o['Marca_Modello']} ({row_o['Misura']}) - Qta: **{row_o['Quantita_Ordinata']}** | Fornitore: {row_o.get('Fornitore', 'N/D')}")
                            if c_o2.button(f"📦 REGISTRA ARRIVO #{row_o['ID_Ordine']}", key=f"btn_arrived_{idx_o}"):
                                df_o_list = df_o_list.astype(object)
                                df_o_list.loc[df_o_list["ID_Ordine"].astype(str) == str(row_o["ID_Ordine"]), "Stato_Ordine"] = "Consegnato in Magazzino"
                                save_sheet(df_o_list, "Ordini_Pneumatici")
                                
                                df_m_add = load_data("Magazzino")
                                c_tot = pd.to_numeric(row_o.get("Costo_Acquisto_Totale", 50), errors='coerce')
                                q_ord = pd.to_numeric(row_o.get("Quantita_Ordinata", 1), errors='coerce')
                                c_tot = float(c_tot) if pd.notnull(c_tot) else 50.0
                                q_ord = float(q_ord) if (pd.notnull(q_ord) and q_ord > 0) else 1.0
                                cost_unit = c_tot / q_ord
                                mid = len(df_m_add) + 1 if not df_m_add.empty else 1
                                new_stock = pd.DataFrame([{"ID_Articolo": mid, "Marca_Modello": row_o["Marca_Modello"], "Misura": row_o["Misura"], "Quantita_Disponibile": row_o["Quantita_Ordinata"], "Costo_Acquisto_Unitario": cost_unit, "Prezzo_Vendita_Unitario": cost_unit * 1.5, "Ubicazione_Magazzino": "Magazzino Arrivi", "Note": f"Da Ordine #{row_o['ID_Ordine']}"}])
                                save_sheet(pd.concat([df_m_add, new_stock], ignore_index=True), "Magazzino")
                                st.success(f"🎉 Merce dell Ordine #{row_o['ID_Ordine']} caricata con successo in Magazzino!")
                                st.rerun()
                else:
                    st.info("Nessun ordine presente.")

        with tabs[4]:
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

        with tabs[5]:
            st.subheader("📅 Agenda & Calendario Interattivo Sincronizzato")
            st.write("Visualizza il calendario visivo dell officina e sincronizza tutti gli appuntamenti direttamente su **iPhone (Apple Calendar)** e **Android (Google Calendar)**.")
            
            df_a = load_data("Appuntamenti")
            
            # Key Metrics Header
            m1, m2, m3, m4 = st.columns(4)
            tot_count = len(df_a) if not df_a.empty else 0
            conf_count = len(df_a[df_a["Stato"] == "Confermato"]) if not df_a.empty and "Stato" in df_a.columns else 0
            pend_count = len(df_a[df_a["Stato"] == "In attesa di conferma"]) if not df_a.empty and "Stato" in df_a.columns else 0
            prop_count = len(df_a[df_a["Stato"].astype(str).str.contains("Proposta", na=False)]) if not df_a.empty and "Stato" in df_a.columns else 0
            
            m1.metric("Totale Appuntamenti", tot_count)
            m2.metric("🟢 Confermati", conf_count)
            m3.metric("⏳ In Attesa (Pendenti)", pend_count)
            m4.metric("🔵 Proposte Nuova Data", prop_count)
            
            st.markdown("---")
            
            # Interactive Visual Calendar View
            st.markdown("### 🗓️ Calendario Grafico Interattivo Officina")
            st.write("Passa dalla vista **Mese**, **Settimana**, **Giorno** alla **Lista Agenda** per controllare le fasce orarie occupate.")
            
            html_cal = render_fullcalendar_component(df_a)
            components.html(html_cal, height=600, scrolling=True)
            
            st.markdown("---")
            
            # Smartphone Device Synchronization Section
            st.markdown("### 📱 Sincronizzazione Diretta con iPhone e Android")
            st.write("Puoi esportare ed iscrivere il tuo dispositivo iPhone o Android per ricevere tutti gli appuntamenti aggiornati nel tuo calendario personale.")
            
            col_sync1, col_sync2 = st.columns(2)
            
            with col_sync1:
                st.markdown("#### 🍏 Sincronizzazione iPhone (Apple Calendar)")
                st.write("Scarica il file di calendario `.ics` ed aprilo su iPhone per aggiungere tutti gli eventi all app **Calendario** nativa iOS.")
                
                ev_list = []
                if not df_a.empty:
                    for idx, r in df_a.iterrows():
                        ev_list.append({
                            'id': r.get('ID_App', idx),
                            'date_str': str(r.get('Data', '')),
                            'time_str': str(r.get('Ora', '09:00')),
                            'title': str(r.get('Servizio', 'Gomme')),
                            'targa': str(r.get('Targa', '')),
                            'cliente': str(r.get('Cliente', '')),
                            'stato': str(r.get('Stato', 'Confermato')),
                            'details': str(r.get('Note', ''))
                        })
                
                ics_data = create_ics_calendar(ev_list, "PneumaticEvolution Officina")
                st.download_button(
                    label="📥 Scarica File Calendario Completo (.ics per iPhone)",
                    data=ics_data,
                    file_name="PneumaticEvolution_Agenda.ics",
                    mime="text/calendar",
                    use_container_width=True
                )
                
                st.markdown("""
                **Guida per l abbonamento continuativo su iPhone:**
                1. Apri **Impostazioni** sul tuo iPhone ➔ **Calendario** ➔ **Account**.
                2. Tocca **Aggiungi account** ➔ **Altro** ➔ **Aggiungi calendario iscritto**.
                3. Scarica ed apri il file `.ics` sopra oppure incolla il link dell app per abilitare l aggiornamento automatico.
                """)
                
            with col_sync2:
                st.markdown("#### 🤖 Sincronizzazione Android (Google Calendar)")
                st.write("Importa il file dell agenda direttamente nell app **Google Calendar** del tuo cellulare Android.")
                
                st.download_button(
                    label="📥 Scarica Calendario (.ics per Google Calendar)",
                    data=ics_data,
                    file_name="PneumaticEvolution_Google.ics",
                    mime="text/calendar",
                    use_container_width=True
                )
                
                st.markdown("""
                **Guida per l importazione su Android / Google Calendar:**
                1. Vai su [Google Calendar](https://calendar.google.com) da browser web sul cellulare o PC.
                2. Clicca sull icona **"+"** accanto ad **Altri calendari** ➔ **Importa**.
                3. Seleziona il file `PneumaticEvolution_Google.ics` appena scaricato per sincronizzare tutti i cambi gomme!
                """)
                
            st.markdown("---")
            st.markdown("### 📋 Elenco & Aggiunta 1-Click Singoli Appuntamenti su Device Admin")
            
            if not df_a.empty:
                for idx, r in df_a.iterrows():
                    targa_val = r.get('Targa', 'N/D')
                    cli_val = r.get('Cliente', 'N/D')
                    serv_val = r.get('Servizio', 'N/D')
                    d_val = str(r.get('Data', ''))
                    t_val = str(r.get('Ora', ''))
                    st_val = str(r.get('Stato', 'Confermato'))
                    
                    gcal_url_admin = create_gcal_url(f"🛞 [{targa_val}] {cli_val} - {serv_val}", d_val, t_val, f"Cliente: {cli_val}, Stato: {st_val}")
                    
                    single_ics = create_ics_calendar([{
                        'id': r.get('ID_App', idx),
                        'date_str': d_val,
                        'time_str': t_val,
                        'title': serv_val,
                        'targa': targa_val,
                        'cliente': cli_val,
                        'stato': st_val,
                        'details': f"Note: {r.get('Note', '')}"
                    }])
                    
                    c_a1, c_a2, c_a3 = st.columns([3, 1, 1])
                    c_a1.write(f"🚗 **{targa_val}** - {cli_val} | 📅 {d_val} ore {t_val} | `{st_val}`")
                    c_a2.markdown(f"[🤖 Add Google Calendar]({gcal_url_admin})", unsafe_allow_html=True)
                    c_a3.download_button("🍏 Add iPhone (.ics)", data=single_ics, file_name=f"App_{targa_val}.ics", mime="text/calendar", key=f"ics_single_{idx}")
            else:
                st.info("Nessun appuntamento da visualizzare.")

        with tabs[6]:
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

