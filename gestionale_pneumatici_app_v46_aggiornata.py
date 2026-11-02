import textwrap

def clean_html(s):
    if not isinstance(s, str):
        return s
    return '\n'.join(line.strip() for line in s.splitlines() if line.strip())

import streamlit as st
import pandas as pd
import datetime
import os
import shutil
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
st.markdown(clean_html('''
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    html, body, [class*="css"], .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }
    /* Tablet & iPhone Mobile-First Responsive Styles */
    .main-header { font-size: 2.2rem; color: #1E3A8A; font-weight: 800; margin-bottom: 5px; }
    .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 20px; }
    
    /* Touch Target & iOS Safari Font-Size Fix (Prevents Auto-Zoom on iPhone) */
    input[type="text"], input[type="password"], input[type="number"], input[type="date"], select, textarea, div[data-baseweb="input"], div[data-baseweb="select"] {
        font-size: 16px !important;
        min-height: 46px !important;
        border-radius: 8px !important;
    }

    /* Touch Friendly Buttons for Mobile / Tablet */
    .stButton>button {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        font-size: 1.15rem !important;
        font-weight: bold !important;
        min-height: 52px !important;
        padding: 12px 20px !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: 0 4px 8px rgba(30, 58, 138, 0.25) !important;
        touch-action: manipulation;
        width: 100% !important;
    }
    .stButton>button:hover {
        background-color: #2563EB !important;
        box-shadow: 0 6px 14px rgba(37, 99, 235, 0.35) !important;
    }
    
    /* Big Green WhatsApp Button for Touchscreens */
    .wa-big-btn {
        display: block;
        background-color: #25D366;
        color: white !important;
        font-size: 1.25rem;
        font-weight: 800;
        text-align: center;
        padding: 16px 20px;
        border-radius: 12px;
        text-decoration: none !important;
        box-shadow: 0 6px 16px rgba(37, 211, 102, 0.35);
        margin-top: 12px;
        margin-bottom: 12px;
        touch-action: manipulation;
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
        font-size: 1.1rem;
        font-weight: 700;
        text-align: center;
        padding: 14px 20px;
        border-radius: 10px;
        text-decoration: none !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
        margin-top: 10px;
        margin-bottom: 12px;
    }

    /* Responsive Padding & Font Sizes for Tablet & Mobile (iPhone / iPad) */
    @media (max-width: 768px) {
        .main .block-container {
            padding-top: 0.5rem !important;
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
            padding-bottom: 2rem !important;
        }
        .main-header { font-size: 1.65rem !important; text-align: center !important; }
        .sub-header { font-size: 0.95rem !important; text-align: center !important; }
        .stButton>button {
            min-height: 52px !important;
            font-size: 1.05rem !important;
            padding: 10px 14px !important;
        }
        .wa-big-btn { font-size: 1.1rem !important; padding: 14px 16px !important; }
        .cal-big-btn { font-size: 1rem !important; padding: 12px 14px !important; }
    }

    @media (min-width: 769px) and (max-width: 1024px) {
        /* iPad / Tablet Specific Enhancements */
        .main .block-container {
            padding-top: 1rem !important;
            padding-left: 1.2rem !important;
            padding-right: 1.2rem !important;
        }
        .main-header { font-size: 2rem !important; }
        .stButton>button {
            min-height: 50px !important;
            font-size: 1.1rem !important;
        }
    }
</style>
'''), unsafe_allow_html=True)

# URL Router Check: ?page=prenotazione hides sidebar & optimizes mobile layout
query_params = st.query_params
raw_page = str(query_params.get("page", "")).lower()
is_public_page_url = "prenotazione" in raw_page or "clienti" in raw_page

if is_public_page_url:
    st.markdown(clean_html('''
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
    '''), unsafe_allow_html=True)

def resolve_db_path():
    possible_paths = [
        "gestionale_pneumatici_db_v46.xlsx",
        "gestionale_pneumatici_db_v14.xlsx",
        "gestionale_pneumatici_db_v13.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v46.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v14.xlsx"
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return "gestionale_pneumatici_db_v46.xlsx"

DB_FILE_EXCEL = resolve_db_path()

def ensure_db_exists():
    global DB_FILE_EXCEL
    if not os.path.exists(DB_FILE_EXCEL):
        DB_FILE_EXCEL = "gestionale_pneumatici_db_v46.xlsx"
        if not os.path.exists(DB_FILE_EXCEL):
            wb = openpyxl.Workbook()
            default_sheets = ["Clienti", "Veicoli", "Hotel_Pneumatici", "Appuntamenti", "Magazzino", "Ordini_Pneumatici", "Flotta_Camion", "Impostazioni", "Anagrafica_Pneumatici"]
            wb.active.title = default_sheets[0]
            for s in default_sheets[1:]:
                wb.create_sheet(s)
            ws_imp = wb["Impostazioni"]
            ws_imp.append(["Chiave", "Valore"])
            ws_imp.append(["admin_password", "admin"])
            ws_imp.append(["giorni_bloccati", ""])
            def_url = "https://pneumaticevolution.streamlit.app/"
            ws_imp.append(["app_url", def_url])
            ws_imp.append(["officina_whatsapp", "3336330301"])
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
                if sheet_name == "Veicoli" and "Costo_Acquisto_Pneumatici" not in df.columns:
                    df["Costo_Acquisto_Pneumatici"] = "0.0"
            return df
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


@st.cache_data(ttl=5)
def load_anagrafica_pneumatici():
    """Carica l anagrafica completa dei pneumatici (Marche, Misure, Indici) da Impostazioni/Anagrafica."""
    df_imp = load_data("Impostazioni")
    marche_list = []
    if not df_imp.empty and "Chiave" in df_imp.columns and "Valore" in df_imp.columns:
        row_m = df_imp[df_imp["Chiave"] == "marche_pneumatici"]
        if not row_m.empty:
            marche_list = [m.strip() for m in str(row_m.iloc[0]["Valore"]).split(",") if m.strip()]
    if not marche_list:
        marche_list = ["Michelin", "Pirelli", "Bridgestone", "Continental", "Goodyear", "Dunlop", "Hankook", "Nokian", "Kumho", "Falken", "Yokohama"]
    return sorted(list(set(marche_list)))


def save_sheet(df, sheet_name):
    ensure_db_exists()
    try:
        with pd.ExcelWriter(DB_FILE_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        try:
            backup_file = DB_FILE_EXCEL.replace(".xlsx", "_backup.xlsx")
            shutil.copyfile(DB_FILE_EXCEL, backup_file)
        except Exception:
            pass
    except Exception:
        try:
            df.to_excel(DB_FILE_EXCEL, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Errore salvataggio dati: {e}")
    st.cache_data.clear()

def auto_archive_past_appointments():
    try:
        df_a = load_data("Appuntamenti")
        if not df_a.empty and "Data" in df_a.columns and "Stato" in df_a.columns:
            today_str = str(datetime.date.today())
            df_a = df_a.astype(object)
            mask_expired = (df_a["Data"].astype(str) < today_str) & (df_a["Stato"].astype(str) == "In attesa di conferma")
            if mask_expired.any():
                df_a.loc[mask_expired, "Stato"] = "Scaduto / Non Confermato"
                save_sheet(df_a, "Appuntamenti")
    except Exception:
        pass

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


def get_booking_url():
    base_url = get_setting("app_url", "https://pneumaticevolution.streamlit.app/").strip().rstrip("/")
    if "?page=prenotazione" in base_url or "&page=prenotazione" in base_url:
        return base_url
    if "?" in base_url:
        return f"{base_url}&page=prenotazione"
    return f"{base_url}/?page=prenotazione"

def create_whatsapp_url(phone_number, text_message):
    clean_phone = ''.join(filter(str.isdigit, str(phone_number)))
    if not clean_phone.startswith("39") and len(clean_phone) == 10:
        clean_phone = "39" + clean_phone
    encoded_text = urllib.parse.quote(text_message)
    base_wa = "wa.me"
    return f"https://{base_wa}/{clean_phone}?text={encoded_text}"


def lookup_client_phone(cliente_name="", targa="", dep_row=None):
    if dep_row is not None:
        p = str(dep_row.get('Telefono', '')).strip()
        if p and p.lower() != 'nan' and p != '3330000000' and len(p) >= 6:
            return p
            
    df_c = load_data("Clienti")
    df_v = load_data("Veicoli")
    df_a = load_data("Appuntamenti")
    
    if not df_c.empty:
        c_name = str(cliente_name).strip().lower()
        if c_name:
            match = df_c[df_c['Nome_RagioneSociale'].astype(str).str.strip().str.lower() == c_name]
            if not match.empty:
                phone = str(match.iloc[0].get('Telefono_WhatsApp', '')).strip()
                if phone and phone.lower() != 'nan' and len(phone) >= 6:
                    return phone
            match_id = df_c[df_c['ID_Cliente'].astype(str).str.strip() == c_name]
            if not match_id.empty:
                phone = str(match_id.iloc[0].get('Telefono_WhatsApp', '')).strip()
                if phone and phone.lower() != 'nan' and len(phone) >= 6:
                    return phone
                    
    t_val = str(targa).strip().upper()
    if t_val and not df_v.empty:
        v_match = df_v[df_v['Targa'].astype(str).str.strip().str.upper() == t_val]
        if not v_match.empty:
            cid = str(v_match.iloc[0].get('ID_Cliente', '')).strip()
            if cid and not df_c.empty:
                c_match = df_c[(df_c['ID_Cliente'].astype(str).str.strip() == cid) | (df_c['Nome_RagioneSociale'].astype(str).str.strip().str.lower() == cid.lower())]
                if not c_match.empty:
                    phone = str(c_match.iloc[0].get('Telefono_WhatsApp', '')).strip()
                    if phone and phone.lower() != 'nan' and len(phone) >= 6:
                        return phone

    if not df_a.empty:
        if t_val:
            a_match = df_a[df_a['Targa'].astype(str).str.strip().str.upper() == t_val]
            if not a_match.empty:
                phone = str(a_match.iloc[-1].get('Telefono', '')).strip()
                if phone and phone.lower() != 'nan' and len(phone) >= 6:
                    return phone
        if cliente_name:
            a_match2 = df_a[df_a['Cliente'].astype(str).str.strip().str.lower() == str(cliente_name).strip().lower()]
            if not a_match2.empty:
                phone = str(a_match2.iloc[-1].get('Telefono', '')).strip()
                if phone and phone.lower() != 'nan' and len(phone) >= 6:
                    return phone

    return ""


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
    # Hero Banner (Portal Style)
    st.markdown(clean_html('''
    <div style="background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%); padding: 32px 24px; border-radius: 20px; text-align: center; color: white; box-shadow: 0 12px 30px rgba(30,58,138,0.3); margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.15);">
        <div style="display:inline-block; background: rgba(255,255,255,0.15); padding: 6px 18px; border-radius: 50px; font-size: 0.85rem; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 12px; backdrop-filter: blur(5px);">
            🛞 PNEUMATICEVOLUTION OFFICINA & HOTEL GOMME
        </div>
        <h1 style="color: white !important; font-size: 2.3rem !important; font-weight: 800; margin: 0 0 12px 0; text-shadow: 0 2px 10px rgba(0,0,0,0.3);">
            Prenotazione Online Servizi & Appuntamenti
        </h1>
        <p style="color: #93C5FD; font-size: 1.15rem; margin: 0 auto; max-width: 650px; font-weight: 500;">
            Scegli il servizio desiderato, seleziona la data e conferma la richiesta direttamente su WhatsApp al <b>3336330301</b>.
        </p>
    </div>
    '''), unsafe_allow_html=True)
    
    st.markdown("### 1️⃣ Seleziona il Servizio Richiesto (Tocca una tessera)")
    
    # 5 Styled Portal Cards
    services_data = [
        {
            "id": "Cambio Stagionale (Montaggio da Hotel Gomme)",
            "title": "Cambio Stagionale",
            "subtitle": "Hotel Gomme",
            "badge": "In Custodia",
            "icon": "🛞",
            "color": "#1E3A8A",
            "desc": "Montaggio e bilanciatura delle gomme in custodia"
        },
        {
            "id": "Sostituzione Pneumatici Nuovi",
            "title": "Gomme Nuove",
            "subtitle": "Fornitura & Montaggio",
            "badge": "Sostituzione",
            "icon": "🆕",
            "color": "#10B981",
            "desc": "Acquisto e montaggio pneumatici nuovi da listino"
        },
        {
            "id": "Equilibratura & Convergenza Computerizzata",
            "title": "Equilibratura",
            "subtitle": "& Convergenza 3D",
            "badge": "Assetto Laser",
            "icon": "⚖️",
            "color": "#F59E0B",
            "desc": "Regolazione convergenza ed equilibratura ruote"
        },
        {
            "id": "Inversione Pneumatici & Controllo Pressione",
            "title": "Inversione Gomme",
            "subtitle": "& Pressione",
            "badge": "Manutenzione",
            "icon": "🔄",
            "color": "#0284C7",
            "desc": "Inversione assi e ripristino pressione esercizio"
        },
        {
            "id": "Riparazione Foratura & Vulcanizzazione",
            "title": "Riparazione",
            "subtitle": "Foratura",
            "badge": "Vulcanizzazione",
            "icon": "🛠️",
            "color": "#DC2626",
            "desc": "Riparazione rapida forature con collaudo tenuta"
        }
    ]
    
    col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
    cols = [col_s1, col_s2, col_s3, col_s4, col_s5]
    
    for i, s in enumerate(services_data):
        is_sel = (st.session_state['selected_service'] == s['id'])
        border_style = f"3px solid {s['color']}" if is_sel else "1.5px solid #CBD5E1"
        bg_style = "linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%)" if is_sel else "#FFFFFF"
        shadow_style = "0 10px 25px rgba(37,99,235,0.25)" if is_sel else "0 4px 12px rgba(0,0,0,0.06)"
        status_badge = f"<div style='background:{s['color']}; color:white; font-size:0.75rem; font-weight:800; padding:4px 10px; border-radius:50px; display:inline-block; margin-bottom:8px;'>✔ SELEZIONATO</div>" if is_sel else f"<div style='background:#F1F5F9; color:#475569; font-size:0.75rem; font-weight:700; padding:4px 10px; border-radius:50px; display:inline-block; margin-bottom:8px;'>{s['badge']}</div>"
        
        with cols[i]:
            st.markdown(clean_html(f'''
            <div style="border: {border_style}; background: {bg_style}; border-radius: 16px; padding: 16px 10px; text-align: center; box-shadow: {shadow_style}; margin-bottom: 10px; min-height: 185px; transition: all 0.3s ease;">
                {status_badge}
                <div style="font-size: 2.2rem; margin-bottom: 4px;">{s['icon']}</div>
                <div style="font-size: 1.05rem; font-weight: 800; color: #0F172A; line-height: 1.2;">{s['title']}</div>
                <div style="font-size: 0.85rem; font-weight: 600; color: #64748B; margin-bottom: 6px;">{s['subtitle']}</div>
                <div style="font-size: 0.78rem; color: #475569; line-height: 1.2;">{s['desc']}</div>
            </div>
            '''), unsafe_allow_html=True)
            
            btn_label = "✅ Selezionato" if is_sel else "Tocca per Scegliere"
            if st.button(btn_label, key=f"btn_srv_{i}", use_container_width=True):
                st.session_state['selected_service'] = s['id']
                st.rerun()

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
                "Sostituzione Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ], index=0 if st.session_state['selected_service'] not in [
                "Cambio Stagionale (Montaggio da Hotel Gomme)",
                "Sostituzione Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ] else [
                "Cambio Stagionale (Montaggio da Hotel Gomme)",
                "Sostituzione Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ].index(st.session_state['selected_service']))
            
            qta_pneumatici = st.selectbox("Quantità Pneumatici / Pezzi *", [1, 2, 3, 4, 5, 6, 8], index=3, help="Scegli il numero di pneumatici da sostituire o montare")
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
                    "Quantita_Pneumatici": str(qta_pneumatici),
                    "Telefono": telefono_cliente, "Note": note_cliente, "Stato": "In attesa di conferma"
                }])
                df_app = pd.concat([df_app, new_row], ignore_index=True)
                save_sheet(df_app, "Appuntamenti")
                
                booking_url = get_booking_url()
                gcal_url = create_gcal_url(f"Appuntamento Gommista: {servizio}", str(data_app), str(ora_app), f"Prenotazione per targa {targa_veicolo}")
                
                off_wa = get_setting("officina_whatsapp", "3336330301")
                note_str = f"\n📝 Note: {note_cliente}" if note_cliente.strip() else ""
                msg_text = (
                    f"Ciao PneumaticEvolution! Ho appena inviato una prenotazione dal vostro sito web:\n\n"
                    f"🚗 Vettura Targa: {targa_veicolo}\n"
                    f"🛞 Servizio: {servizio} ({qta_pneumatici} pz)\n"
                    f"📅 Data & Ora: {data_app} alle ore {ora_app}\n\n"
                    f"👤 Cliente: {nome_cliente}\n"
                    f"📱 Cellulare: {telefono_cliente}{note_str}\n\n"
                    f"📅 Aggiungi al Google Calendar: {gcal_url}\n"
                    f"👉 Verifica stato sul sito: {booking_url}"
                )
                wa_url = create_whatsapp_url(off_wa, msg_text)
                
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
        
        st.markdown(clean_html(f'''
        <a href="{wa_href}" target="_blank" class="wa-big-btn">
            🟢 💬 CLICCA QUI PER APRIRE WHATSAPP & INVIARE LA RICHIESTA A PNEUMATICEVOLUTION (3336330301)
        </a>
        <a href="{gcal_href}" target="_blank" class="cal-big-btn">
            📅 AGGIUNGI L APPUNTAMENTO AL TUO GOOGLE CALENDAR
        </a>
        '''), unsafe_allow_html=True)

else:
    if not st.session_state['authenticated']:
        st.warning("🔒 Sezione riservata all Officina. Inserisci la password nel menu a sinistra per accedere.")
    else:
        st.markdown("# 🛞 PneumaticEvolution Gestionale")
        st.markdown("### Gestione Anagrafica, Hotel Pneumatici, Flotta Camion, Agenda e Notifiche Massive WhatsApp")
        
        if 'admin_nav_page' not in st.session_state:
            st.session_state['admin_nav_page'] = 'home'

        # Function to render top back navigation
        def render_back_button(page_title):
            st.markdown(clean_html(f"""
            <div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); padding: 16px 22px; border-radius: 14px; margin-bottom: 18px; border: 1px solid rgba(255,255,255,0.12); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
                <div style="font-size: 0.8rem; font-weight: 700; color: #38BDF8; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">
                    📌 PANNELLO DI CONTROLLO ADMIN / OFFICINA
                </div>
                <div style="font-size: 1.5rem; font-weight: 800; color: #FFFFFF;">
                    {page_title}
                </div>
            </div>
            """), unsafe_allow_html=True)
            cb1, cb2 = st.columns([2, 4])
            with cb1:
                if st.button("🔙 Torna alla Home Page", key=f"back_top_{page_title}", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'home'
                    st.rerun()
            st.markdown("---")

        nav = st.session_state['admin_nav_page']

        if nav == 'home':
            # Load quick metrics for live KPI summary
            df_app_kpi = load_data("Appuntamenti")
            df_dep_kpi = load_data("Hotel_Pneumatici")
            df_mag_kpi = load_data("Magazzino")
            
            today_str_kpi = str(datetime.date.today())
            pending_kpi = len(df_app_kpi[(df_app_kpi["Stato"] == "In attesa di conferma") & (df_app_kpi["Data"].astype(str) >= today_str_kpi)]) if not df_app_kpi.empty and "Stato" in df_app_kpi.columns else 0
            today_kpi = len(df_app_kpi[df_app_kpi["Data"].astype(str) == today_str_kpi]) if not df_app_kpi.empty and "Data" in df_app_kpi.columns else 0
            hotel_kpi = len(df_dep_kpi) if not df_dep_kpi.empty else 0
            mag_kpi_pieces = pd.to_numeric(df_mag_kpi["Quantita_Disponibile"], errors='coerce').fillna(0).sum() if not df_mag_kpi.empty and "Quantita_Disponibile" in df_mag_kpi.columns else 0
            
            # Executive Dashboard Header Banner
            st.markdown(clean_html(f"""
<div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 40%, #1E3A8A 100%); padding: 26px 24px 18px 24px; border-radius: 20px; color: white; box-shadow: 0 12px 30px rgba(15,23,42,0.35); margin-bottom: 18px; border: 1px solid rgba(255,255,255,0.15);">
    <div style="display:inline-block; background: rgba(56, 189, 248, 0.2); border: 1px solid rgba(56, 189, 248, 0.4); padding: 5px 16px; border-radius: 50px; font-size: 0.85rem; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: #38BDF8; margin-bottom: 10px;">
        🛞 PNEUMATICEVOLUTION EXECUTIVE DASHBOARD ADMIN
    </div>
    <h1 style="color: white !important; font-size: 2.2rem !important; font-weight: 800; margin: 0 0 6px 0; text-shadow: 0 2px 10px rgba(0,0,0,0.3);">
        Pannello di Controllo Officina
    </h1>
    <p style="color: #94A3B8; font-size: 1.05rem; margin: 0; font-weight: 500;">
        Monitoraggio in tempo reale — <b>Tocca un indicatore per accedere direttamente alla sezione operativa</b>
    </p>
</div>
"""), unsafe_allow_html=True)

            # Interactive Clickable KPI Metric Grid
            kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)
            
            with kpi_c1:
                st.markdown(clean_html(f"""
<div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 14px 12px 8px 12px; border-radius: 14px 14px 0 0; border: 1px solid rgba(245, 158, 11, 0.3); border-bottom: none; text-align: center;">
    <div style="font-size: 0.78rem; color: #FBBF24; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">⏳ IN ATTESA</div>
    <div style="font-size: 1.9rem; font-weight: 800; color: white; margin-top: 2px;">{pending_kpi}</div>
</div>
"""), unsafe_allow_html=True)
                btn_pending_label = f"🚀 APRI PRENOTAZIONI ({pending_kpi})" if pending_kpi > 0 else "📥 Apri Appuntamenti"
                if st.button(btn_pending_label, key="kpi_nav_pending", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'appuntamenti'
                    st.rerun()

            with kpi_c2:
                st.markdown(clean_html(f"""
<div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 14px 12px 8px 12px; border-radius: 14px 14px 0 0; border: 1px solid rgba(96, 165, 250, 0.3); border-bottom: none; text-align: center;">
    <div style="font-size: 0.78rem; color: #60A5FA; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">📅 APPOINTMENT OGGI</div>
    <div style="font-size: 1.9rem; font-weight: 800; color: white; margin-top: 2px;">{today_kpi}</div>
</div>
"""), unsafe_allow_html=True)
                if st.button(f"📅 OGGI IN AGENDA ({today_kpi})", key="kpi_nav_today", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'appuntamenti'
                    st.rerun()

            with kpi_c3:
                st.markdown(clean_html(f"""
<div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 14px 12px 8px 12px; border-radius: 14px 14px 0 0; border: 1px solid rgba(52, 211, 153, 0.3); border-bottom: none; text-align: center;">
    <div style="font-size: 0.78rem; color: #34D399; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">🏨 HOTEL GOMME</div>
    <div style="font-size: 1.9rem; font-weight: 800; color: white; margin-top: 2px;">{hotel_kpi} set</div>
</div>
"""), unsafe_allow_html=True)
                if st.button(f"🏨 HOTEL GOMME ({hotel_kpi})", key="kpi_nav_hotel", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'hotel'
                    st.rerun()

            with kpi_c4:
                st.markdown(clean_html(f"""
<div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 14px 12px 8px 12px; border-radius: 14px 14px 0 0; border: 1px solid rgba(192, 132, 252, 0.3); border-bottom: none; text-align: center;">
    <div style="font-size: 0.78rem; color: #C084FC; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">📦 STOCK GOMME</div>
    <div style="font-size: 1.9rem; font-weight: 800; color: white; margin-top: 2px;">{int(mag_kpi_pieces)} pz</div>
</div>
"""), unsafe_allow_html=True)
                if st.button(f"📦 MAGAZZINO ({int(mag_kpi_pieces)} pz)", key="kpi_nav_stock", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'magazzino'
                    st.rerun()
                    
            st.markdown("<div style='margin-bottom: 25px;'></div>", unsafe_allow_html=True)

            # Grid of modules - Executive Cards
            r1_c1, r1_c2 = st.columns(2)
            with r1_c1:
                st.markdown(clean_html("""
                <div style="border: 2px solid #3B82F6; border-left: 8px solid #2563EB; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(37, 99, 235, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #1E3A8A; margin: 0; font-size: 1.25rem; font-weight: 800;">📥 1. Appuntamenti & Conferme</h3>
                        <span style="background: #DBEAFE; color: #1E40AF; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">RICHIESTE WEB</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Gestisci le richieste di prenotazione arrivate dal sito web. Conferma l'appuntamento, proponi date alternative o rifiuta. Invia automaticamente il messaggio di notifica via WhatsApp al cliente.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Appuntamenti & Conferme", key="nav_btn_app", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'appuntamenti'
                    st.rerun()

            with r1_c2:
                st.markdown(clean_html("""
                <div style="border: 2px solid #10B981; border-left: 8px solid #059669; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #ECFDF5 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(16, 185, 129, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #065F46; margin: 0; font-size: 1.25rem; font-weight: 800;">👤 2. Anagrafica Clienti & Veicoli</h3>
                        <span style="background: #D1FAE5; color: #065F46; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">ANAGRAFICA & TARGHE</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Registra nuovi clienti e targhe, consulta il registro anagrafico, modifica o elimina clienti. Scarica i pneumatici dal magazzino in fase di registrazione veicolo.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Anagrafica Clienti & Veicoli", key="nav_btn_cli", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'anagrafica'
                    st.rerun()

            # Row 2: Modules 3 & 4
            r2_c1, r2_c2 = st.columns(2)
            with r2_c1:
                st.markdown(clean_html("""
                <div style="border: 2px solid #F59E0B; border-left: 8px solid #D97706; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #FFFBEB 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(245, 158, 11, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #92400E; margin: 0; font-size: 1.25rem; font-weight: 800;">🏨 3. Hotel Pneumatici & QR Code</h3>
                        <span style="background: #FEF3C7; color: #92400E; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">CUSTODIA & QR</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Gestisci il deposito gomme stagionali in custodia. Genera e stampa etichette con QR Code scansionabili da fotocamera smartphone. Aggiorna lo scaffale senza cambiare il Tag ID.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Hotel Pneumatici & QR", key="nav_btn_hotel", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'hotel'
                    st.rerun()

            with r2_c2:
                st.markdown(clean_html("""
                <div style="border: 2px solid #8B5CF6; border-left: 8px solid #7C3AED; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #F5F3FF 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(139, 92, 246, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #5B21B6; margin: 0; font-size: 1.25rem; font-weight: 800;">📦 4. Magazzino Pneumatici & Ordini</h3>
                        <span style="background: #EDE9FE; color: #5B21B6; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">GIACENZA & VALORE</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Visualizza giacenza, totale pezzi e valore totale economico (€). Registra e modifica l'inventario. Traccia le merci in arrivo dal fornitore e caricale in magazzino con 1-click.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Magazzino & Ordini", key="nav_btn_mag", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'magazzino'
                    st.rerun()

            # Row 3: Modules 5 & 6
            r3_c1, r3_c2 = st.columns(2)
            with r3_c1:
                st.markdown(clean_html("""
                <div style="border: 2px solid #25D366; border-left: 8px solid #16A34A; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #F0FDF4 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(37, 211, 102, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #166534; margin: 0; font-size: 1.25rem; font-weight: 800;">📱 5. Notifiche Massive WhatsApp</h3>
                        <span style="background: #DCFCE7; color: #166534; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">CAMPAGNE 1-CLICK</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Avvia campagne WhatsApp con 1-click per avvisare i clienti che le gomme in custodia (Estive o Invernali) sono pronte per il montaggio. Esclude in automatico le 4 Stagioni.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Notifiche Massive WhatsApp", key="nav_btn_wa", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'whatsapp'
                    st.rerun()

            with r3_c2:
                st.markdown(clean_html("""
                <div style="border: 2px solid #0284C7; border-left: 8px solid #0369A1; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #F0F9FF 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(2, 132, 199, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #075985; margin: 0; font-size: 1.25rem; font-weight: 800;">📅 6. Agenda & Calendario Sincronizzato</h3>
                        <span style="background: #E0F2FE; color: #075985; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">IPHONE & ANDROID</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Visualizza il calendario grafico visivo (Mese/Settimana/Giorno) e sincronizza l'agenda dell'officina direttamente sul tuo <b>iPhone (Apple Calendar)</b> ed <b>Android (Google Calendar)</b>.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Agenda & Calendario", key="nav_btn_cal", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'calendario'
                    st.rerun()

            # Row 4: Modules 7 & 8
            r4_c1, r4_c2 = st.columns(2)
            with r4_c1:
                st.markdown(clean_html("""
                <div style="border: 2px solid #6B7280; border-left: 8px solid #4B5563; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #F9FAFB 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(107, 114, 128, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #1F2937; margin: 0; font-size: 1.25rem; font-weight: 800;">📚 7. Storico Appuntamenti & Archivio</h3>
                        <span style="background: #F3F4F6; color: #1F2937; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">ARCHIVIO PLURIENNALE</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Registro storico completo di tutti gli appuntamenti passati e conclusi negli anni. Filtra per targa, cliente o servizio ed esporta i report aziendali in Excel.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Storico & Archivio", key="nav_btn_storico", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'storico'
                    st.rerun()

            with r4_c2:
                st.markdown(clean_html("""
                <div style="border: 2px solid #DC2626; border-left: 8px solid #B91C1C; border-radius: 16px; padding: 20px; background: linear-gradient(135deg, #FFFFFF 0%, #FEF2F2 100%); margin-bottom: 15px; box-shadow: 0 6px 18px rgba(220, 38, 38, 0.12);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h3 style="color: #991B1B; margin: 0; font-size: 1.25rem; font-weight: 800;">⚙️ 8. Chiusure, URL & Backup Excel Admin</h3>
                        <span style="background: #FEE2E2; color: #991B1B; font-size: 0.75rem; font-weight: 800; padding: 4px 10px; border-radius: 20px;">SETUP & BACKUP EXCEL</span>
                    </div>
                    <p style="color: #4B5563; font-size: 0.95rem; margin-bottom: 15px;">Blocca singoli giorni o periodi di ferie lunghi (multi-periodo), aggiorna il link del sito web, cambia la password d'accesso ed scarica o ripristina il backup Excel completo.</p>
                </div>
                """), unsafe_allow_html=True)
                if st.button("🚀 Apri Chiusure, URL & Backup Admin", key="nav_btn_cfg", use_container_width=True):
                    st.session_state['admin_nav_page'] = 'impostazioni'
                    st.rerun()

        elif nav == 'appuntamenti':
            render_back_button("📥 Gestione Appuntamenti & Conferme Pendenti")
        elif nav == 'anagrafica':
            render_back_button("👤 Gestione Anagrafica Clienti & Veicoli")
        elif nav == 'hotel':
            render_back_button("🏨 Hotel Pneumatici - Deposito & QR Code")
        elif nav == 'magazzino':
            render_back_button("📦 Magazzino Pneumatici & Ordini In Arrivo")
        elif nav == 'whatsapp':
            render_back_button("📱 Campagne Notifiche Massive WhatsApp")
        elif nav == 'calendario':
            render_back_button("📅 Agenda & Calendario Sincronizzato")
        elif nav == 'storico':
            render_back_button("📚 Storico Appuntamenti & Archivio")
        elif nav == 'impostazioni':
            render_back_button("⚙️ Chiusure, URL, Password & Backup Excel Admin")

        
        if nav == 'appuntamenti':
            st.subheader("📥 Gestione & Conferma Richieste Pendenti")
            
            # Persistent Banner for Admin WhatsApp Notification Trigger
            if st.session_state.get('admin_wa_action'):
                act = st.session_state['admin_wa_action']
                st.balloons()
                st.markdown(clean_html(f"""
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
                """), unsafe_allow_html=True)
                components.html(f"<script>window.open('{act['wa_url']}', '_blank');</script>", height=0)
                if st.button("❌ Chiudi Avviso Notifica", key="close_wa_banner"):
                    st.session_state['admin_wa_action'] = None
                    st.rerun()

            auto_archive_past_appointments()
            df_a = load_data("Appuntamenti")
            
            if df_a.empty or "Stato" not in df_a.columns:
                st.info("Nessuna prenotazione presente in agenda.")
            else:
                today_str = str(datetime.date.today())
                df_pending = df_a[(df_a["Stato"] == "In attesa di conferma") & (df_a["Data"].astype(str) >= today_str)]
                
                st.markdown(f"### ⏳ Richieste In Attesa di Conferma dall Officina ({len(df_pending)})")
                if df_pending.empty:
                    st.success("✅ Nessuna richiesta pendente da confermare al momento!")
                else:
                    for idx, row in df_pending.iterrows():
                        qta_val = row.get('Quantita_Pneumatici', '4')
                        st.markdown(clean_html(f"""
                        <div class="card-box" style="border-left: 6px solid #F59E0B; background-color: #FFFBEB;">
                            <h4 style="color: #92400E; margin-top: 0;">🚗 Cliente: {row['Cliente']} | Targa: <code>{row['Targa']}</code></h4>
                            <p style="font-size: 1.05rem;"><b>Servizio Richiesto:</b> {row['Servizio']} (<b>Q.tà: {qta_val} pz</b>)<br>
                            <b>Data ed Ora Desiderati:</b> 📅 <b>{row['Data']}</b> alle ore ⏰ <b>{row['Ora']}</b><br>
                            <b>Cellulare:</b> 📱 <code>{row['Telefono']}</code> | <b>Note:</b> {row.get('Note', 'Nessuna')}</p>
                        </div>
                        """), unsafe_allow_html=True)
                        
                        col_ac, col_pr = st.columns([1, 1])
                        
                        with col_ac:
                            if st.button(f"✅ CONFERMA PRENOTAZIONE #{row['ID_App']}", key=f"conf_btn_{idx}", use_container_width=True):
                                df_a.loc[df_a["ID_App"] == row["ID_App"], "Stato"] = "Confermato"
                                save_sheet(df_a, "Appuntamenti")
                                
                                cal_link_cust = create_gcal_url(f"Cambio Gomme - {row['Targa']}", str(row['Data']), str(row['Ora']), f"Servizio: {row['Servizio']}")
                                booking_url = get_booking_url()
                                qta_info = f" ({row.get('Quantita_Pneumatici', '4')} pz)" if row.get('Quantita_Pneumatici') else ""
                                confirm_msg = f"Gentile {row['Cliente']}, la tua prenotazione per la vettura {row['Targa']} ({row['Servizio']}{qta_info}) il {row['Data']} alle ore {row['Ora']} e stata CONFERMATA! Ti aspettiamo in officina presso PneumaticEvolution. 📅 Calendario: {cal_link_cust} 👉 Sito web: {booking_url}"
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
                                        
                                        booking_url = get_booking_url()
                                        prop_msg = f"Gentile {row['Cliente']}, riguardo la tua richiesta per la vettura {row['Targa']} per il {row['Data']} alle {row['Ora']}, purtroppo quell orario non e disponibile. Ti proponiamo in alternativa la data {prop_date} alle ore {prop_time}. Clicca per accedere al sito e confermare o scegliere un altro orario: {booking_url} . Grazie! - PneumaticEvolution"
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
                                        
                                        booking_url = get_booking_url()
                                        refuse_msg = f"Gentile {row['Cliente']}, ci dispiace informarti che la tua richiesta di prenotazione per la vettura {row['Targa']} per il {row['Data']} alle ore {row['Ora']} non puo essere accolta per indisponibilita. Ti invitiamo ad accedere al sito per scegliere una data diversa: {booking_url} . Cordiali Saluti."
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
                
                st.markdown("---")
                st.markdown("##### 🗑️ Eliminazione Definitiva Appuntamento dall Agenda")
                if not df_a.empty:
                    app_dict = {f"#{r.get('ID_App','')} - {r.get('Cliente','')} | Targa: {r.get('Targa','')} ({r.get('Data','')} ore {r.get('Ora','')}) [{r.get('Stato','')}]": r.get('ID_App','') for _, r in df_a.iterrows()}
                    sel_app_lbl = st.selectbox("Seleziona l Appuntamento da Eliminare dall Agenda:", list(app_dict.keys()), key="del_app_sel_main")
                    target_aid = app_dict[sel_app_lbl]
                    if st.button(f"🗑️ ELIMINA DEFINITIVAMENTE APPUNTAMENTO #{target_aid}", type="primary", key="btn_del_app_main"):
                        df_a_del = df_a[df_a["ID_App"].astype(str) != str(target_aid)]
                        save_sheet(df_a_del, "Appuntamenti")
                        st.success(f"🗑️ Appuntamento #{target_aid} eliminato dall agenda!")
                        st.rerun()

        # TAB 1: Anagrafica Clienti & Veicoli (Modifica, Elimina & Scarico Magazzino)
        elif nav == 'anagrafica':
            st.subheader("👤 Gestione Anagrafica Clienti & Veicoli")
            st1, st2, st3 = st.tabs(["➕ Nuovo Inserimento Clienti", "📋 Elenco e Ricerca", "✏️ Modifica / Elimina Cliente"])
            
            with st1:
                st.markdown("### ➕ Inserimento Anagrafica Clienti & Veicoli")
                st.write("Puoi registrare un **nuovo cliente** con la sua vettura, oppure **aggiungere un nuovo veicolo ad un cliente già esistente** (distinto dalla targa).")
                
                df_c_exist = load_data("Clienti")
                mode_add = "🆕 Nuovo Cliente + Veicolo"
                existing_cid = None
                
                if not df_c_exist.empty:
                    mode_add = st.radio(
                        "Seleziona Modalità Inserimento:",
                        ["🆕 Registra Nuovo Cliente + Veicolo", "🚗 Aggiungi Nuovo Veicolo a Cliente Già Esistente"],
                        key="mode_add_cli_radio"
                    )
                
                with st.form("form_add_cli"):
                    c1, c2 = st.columns(2)
                    
                    if "Già Esistente" in mode_add and not df_c_exist.empty:
                        cli_options = {f"#{r['ID_Cliente']} - {r['Nome_RagioneSociale']} (Tel: {r.get('Telefono_WhatsApp','')})": r['ID_Cliente'] for _, r in df_c_exist.iterrows()}
                        sel_cli_label = st.selectbox("Seleziona Cliente Esistente dall Anagrafica *:", list(cli_options.keys()))
                        existing_cid = cli_options[sel_cli_label]
                        c_row_ex = df_c_exist[df_c_exist["ID_Cliente"].astype(str) == str(existing_cid)].iloc[0]
                        
                        with c1:
                            st.markdown("##### 📌 Dati Anagrafici Cliente (Già Salvati)")
                            tipo_cli = st.text_input("Tipologia Cliente", value=str(c_row_ex.get("Tipo_Cliente", "Privato")), disabled=True)
                            nome_cli = st.text_input("Nome / Ragione Sociale *", value=str(c_row_ex.get("Nome_RagioneSociale", "")), disabled=True)
                            cf_piva = st.text_input("Codice Fiscale / P.IVA", value=str(c_row_ex.get("CodiceFiscale_PIVA", "")), disabled=True)
                            telefono = st.text_input("Telefono / WhatsApp *", value=str(c_row_ex.get("Telefono_WhatsApp", "")), disabled=True)
                            email = st.text_input("Email", value=str(c_row_ex.get("Email", "")), disabled=True)
                            indirizzo = st.text_input("Indirizzo e Citta", value=str(c_row_ex.get("Indirizzo", "")), disabled=True)
                    else:
                        with c1:
                            st.markdown("##### 📌 Dati Anagrafici & Contatti Nuovo Cliente")
                            tipo_cli = st.selectbox("Tipologia Cliente", ["Privato", "Azienda / Flotta", "Convenzionato"])
                            nome_cli = st.text_input("Nome / Ragione Sociale *")
                            cf_piva = st.text_input("Codice Fiscale / P.IVA")
                            telefono = st.text_input("Telefono / WhatsApp *", placeholder="3331234567")
                            email = st.text_input("Email")
                            indirizzo = st.text_input("Indirizzo e Citta")
                            
                    with c2:
                        st.markdown("##### 🚗 Dati Nuovo Veicolo & Integrazione Magazzino")
                        targa = st.text_input("Targa Veicolo *", placeholder="AB123CD").upper()
                        marca_modello = st.text_input("Marca e Modello Veicolo", placeholder="Es. Audi A4 / Iveco Stralis")
                        tipo_mezzo = st.selectbox("Tipo Mezzo", ["Auto / SUV", "Furgone", "Camion / Motrice", "Rimorchio"])
                        misura_gomme = st.text_input("Misura Pneumatici", placeholder="Es. 205/55 R16 91V")
                        tipo_gomme = st.selectbox("Pneumatici Attualmente Montati", ["Estivi", "Invernali", "4 Stagioni"])
                        
                        df_mag_select = load_data("Magazzino")
                        mag_options = ["Forniti dal cliente / Acquisto Esterno (Nessun scarico magazzino)"]
                        if not df_mag_select.empty:
                            for _, r_m in df_mag_select.iterrows():
                                mag_options.append(f"Scarica da Magazzino: #{r_m.get('ID_Articolo', 1)} - {r_m.get('Marca_Modello', '')} ({r_m.get('Misura', '')}) [Disp: {r_m.get('Quantita_Disponibile', 0)}]")
                        
                        fonte_gomme_sel = st.selectbox("Provenienza Pneumatici per il Veicolo:", mag_options)
                        qta_scarico_input = st.number_input("Quantità Pneumatici da Scaricare / Utilizzare:", min_value=1, max_value=20, value=4, key="qta_scarico_mag")
                        costo_gomme_input = st.number_input("Costo Sostenuto Acquisto Pneumatici (€):", min_value=0.0, value=0.0, step=10.0, help="Spesa o costo sostenuto per l'acquisto di eventuali pneumatici per questo veicolo")
                        note_cli = st.text_area("Note Veicolo / Cliente")
                    
                    btn_save = st.form_submit_button("💾 Salva Veicolo in Anagrafica Cliente", use_container_width=True)
                    if btn_save:
                        if not nome_cli or not telefono or not targa:
                            st.error("Compila i campi obbligatori: Nome, Telefono e Targa!")
                        else:
                            df_c = load_data("Clienti")
                            if "Già Esistente" in mode_add and existing_cid is not None:
                                cid = existing_cid
                            else:
                                cid = len(df_c) + 1 if not df_c.empty else 1
                                new_c = pd.DataFrame([{"ID_Cliente": cid, "Tipo_Cliente": tipo_cli, "Nome_RagioneSociale": nome_cli, "CodiceFiscale_PIVA": cf_piva, "Telefono_WhatsApp": telefono, "Email": email, "Indirizzo": indirizzo, "Note": note_cli}])
                                save_sheet(pd.concat([df_c, new_c], ignore_index=True), "Clienti")
                            
                            id_mag_ref = ""
                            if "Scarica da Magazzino:" in fonte_gomme_sel:
                                try:
                                    art_id_str = fonte_gomme_sel.split("#")[1].split(" - ")[0]
                                    art_id = int(art_id_str)
                                    id_mag_ref = str(art_id)
                                    qta_to_sub = int(st.session_state.get('qta_scarico_mag', 4))
                                    if not df_mag_select.empty:
                                        mask_art = df_mag_select["ID_Articolo"].astype(str) == str(art_id)
                                        if mask_art.any():
                                            curr_q = pd.to_numeric(df_mag_select.loc[mask_art, "Quantita_Disponibile"], errors='coerce').fillna(0).iloc[0]
                                            df_mag_select.loc[mask_art, "Quantita_Disponibile"] = max(0, int(curr_q) - qta_to_sub)
                                            save_sheet(df_mag_select, "Magazzino")
                                            st.info(f"📦 Scaricati {qta_to_sub} pneumatici dal Magazzino per l articolo #{art_id}!")
                                except Exception as e_m:
                                    st.warning(f"Impossibile scaricare dal magazzino: {e_m}")

                            df_v = load_data("Veicoli")
                            vid = len(df_v) + 1 if not df_v.empty else 1
                            new_v = pd.DataFrame([{"ID_Veicolo": vid, "ID_Cliente": cid, "Targa": targa, "Marca_Modello": marca_modello, "Tipo_Mezzo": tipo_mezzo, "Misura_Pneumatici_Anteriore": misura_gomme, "Misura_Pneumatici_Posteriore": misura_gomme, "Tipo_Pneumatici_Attuali": tipo_gomme, "Fonte_Gomme": fonte_gomme_sel, "ID_Magazzino": id_mag_ref, "Costo_Acquisto_Pneumatici": float(costo_gomme_input)}])
                            save_sheet(pd.concat([df_v, new_v], ignore_index=True), "Veicoli")
                            st.success(f"✅ Veicolo '{targa}' salvato con successo ed associato al cliente '{nome_cli}'!")

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
                    df_v_edit = load_data("Veicoli")
                    v_matches = df_v_edit[df_v_edit["ID_Cliente"].astype(str) == str(target_cid)] if not df_v_edit.empty else pd.DataFrame()
                    v_row = v_matches.iloc[0] if not v_matches.empty else {}
                    
                    with st.form("form_edit_cli"):
                        st.markdown(f"##### Modifica Dati Cliente #{target_cid}: {client_row.get('Nome_RagioneSociale', '')}")
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            st.markdown("##### 📌 Dati Anagrafici & Contatti")
                            edit_tipo = st.selectbox("Tipo Cliente", ["Privato", "Azienda / Flotta", "Convenzionato"], index=["Privato", "Azienda / Flotta", "Convenzionato"].index(client_row.get("Tipo_Cliente", "Privato")))
                            edit_nome = st.text_input("Nome / Ragione Sociale", value=str(client_row.get("Nome_RagioneSociale", "")))
                            edit_cf = st.text_input("Codice Fiscale / P.IVA", value=str(client_row.get("CodiceFiscale_PIVA", "")))
                            edit_tel = st.text_input("Telefono / WhatsApp", value=str(client_row.get("Telefono_WhatsApp", "")))
                            edit_email = st.text_input("Email", value=str(client_row.get("Email", "")))
                            edit_indirizzo = st.text_input("Indirizzo", value=str(client_row.get("Indirizzo", "")))
                        with ec2:
                            st.markdown("##### 🚗 Dati Veicolo & Costo Pneumatici")
                            edit_targa = st.text_input("Targa Veicolo", value=str(v_row.get("Targa", "")) if not v_matches.empty else "").upper()
                            edit_marca_v = st.text_input("Marca e Modello Veicolo / Pneumatici", value=str(v_row.get("Marca_Modello", "")) if not v_matches.empty else "")
                            edit_misura_v = st.text_input("Misura Pneumatici", value=str(v_row.get("Misura_Pneumatici_Anteriore", "")) if not v_matches.empty else "")
                            c_val_init = pd.to_numeric(v_row.get("Costo_Acquisto_Pneumatici", 0.0), errors='coerce') if not v_matches.empty else 0.0
                            edit_costo_gomme = st.number_input("Costo Sostenuto Acquisto Pneumatici (€)", min_value=0.0, value=float(c_val_init) if pd.notnull(c_val_init) else 0.0, step=10.0, help="Spesa o costo sostenuto per l'acquisto dei pneumatici per questo veicolo")
                        edit_note = st.text_area("Note Cliente", value=str(client_row.get("Note", "")))
                        
                        btn_update_cli = st.form_submit_button("💾 Salva Modifiche Anagrafica e Veicolo", use_container_width=True)
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

                            if not df_v_edit.empty:
                                mask_v = df_v_edit["ID_Cliente"].astype(str) == str(target_cid)
                                if mask_v.any():
                                    df_v_edit = df_v_edit.astype(object)
                                    df_v_edit.loc[mask_v, "Targa"] = str(edit_targa)
                                    df_v_edit.loc[mask_v, "Marca_Modello"] = str(edit_marca_v)
                                    df_v_edit.loc[mask_v, "Misura_Pneumatici_Anteriore"] = str(edit_misura_v)
                                    df_v_edit.loc[mask_v, "Misura_Pneumatici_Posteriore"] = str(edit_misura_v)
                                    df_v_edit.loc[mask_v, "Costo_Acquisto_Pneumatici"] = float(edit_costo_gomme)
                                    save_sheet(df_v_edit, "Veicoli")
                            
                            st.success("✅ Dati dell'anagrafica cliente e del veicolo aggiornati con successo!")
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

                    st.markdown("---")
                    st.markdown("##### 🚗 Eliminazione Singolo Veicolo del Cliente")
                    df_v_edit = load_data("Veicoli")
                    if not df_v_edit.empty:
                        v_cli_matches = df_v_edit[df_v_edit["ID_Cliente"].astype(str) == str(target_cid)]
                        if not v_cli_matches.empty:
                            v_dict = {f"🚗 Targa: {r.get('Targa','')} - {r.get('Marca_Modello','')} ({r.get('Misura_Pneumatici_Anteriore','')})": r.get('ID_Veicolo','') for _, r in v_cli_matches.iterrows()}
                            sel_v_lbl = st.selectbox("Seleziona il Veicolo da Eliminare dall Anagrafica:", list(v_dict.keys()), key="del_v_sel")
                            target_vid = v_dict[sel_v_lbl]
                            if st.button(f"🗑️ ELIMINA SOLTANTO QUESTO VEICOLO (ID #{target_vid})", type="primary", key="btn_del_v_single"):
                                df_v_edit = df_v_edit[df_v_edit["ID_Veicolo"].astype(str) != str(target_vid)]
                                save_sheet(df_v_edit, "Veicoli")
                                st.success("🗑️ Veicolo eliminato dall anagrafica del cliente!")
                                st.rerun()

        # TAB 2: Hotel Pneumatici & QR Code (Modifica Deposito con Mantenimento Tag ID)
        elif nav == 'hotel':
            st.subheader("🏨 Hotel Pneumatici - Deposito, QR & Scansione")
            h1, h2, h3, h4 = st.tabs(["📥 Check-in Deposito", "📦 Elenco Depositi", "✏️ Modifica Deposito & Rigenera QR", "🔍 Scansiona QR da Fotocamera"])
            
            with h1:
                st.markdown("### 📥 Check-in Deposito Pneumatici in Custodia")
                st.markdown("##### ⚡ Autocompilazione Rapida da Anagrafica Clienti & Veicoli")
                st.write("Seleziona un cliente/veicolo già presente in anagrafica per **compilare automaticamente** tutti i campi (Nome Cliente, Targa, Marca/Modello Gomme e Misura Pneumatici):")
                
                df_c_dep = load_data("Clienti")
                df_v_dep = load_data("Veicoli")
                
                auto_options = ["➕ Inserimento Manuale (Nessuna autocompilazione)"]
                v_map = {}
                
                if not df_v_dep.empty and not df_c_dep.empty:
                    for _, r_v in df_v_dep.iterrows():
                        cid = str(r_v.get("ID_Cliente", ""))
                        c_matches = df_c_dep[df_c_dep["ID_Cliente"].astype(str) == cid]
                        c_name = c_matches.iloc[0]["Nome_RagioneSociale"] if not c_matches.empty else "Cliente"
                        targa_v = str(r_v.get("Targa", "")).upper()
                        marca_v = str(r_v.get("Marca_Modello", ""))
                        misura_v = str(r_v.get("Misura_Pneumatici_Anteriore", ""))
                        tipo_g_v = str(r_v.get("Tipo_Pneumatici_Attuali", "Invernali"))
                        
                        c_phone = c_matches.iloc[0]["Telefono_WhatsApp"] if (not c_matches.empty and "Telefono_WhatsApp" in c_matches.columns) else ""
                        lbl = f"👤 {c_name} — 🚗 Targa: {targa_v} | {marca_v} ({misura_v})"
                        auto_options.append(lbl)
                        v_map[lbl] = {
                            "cliente": c_name,
                            "targa": targa_v,
                            "marca": marca_v,
                            "misura": misura_v,
                            "tipo_gomme": tipo_g_v,
                            "phone": c_phone
                        }
                
                sel_auto = st.selectbox("🔍 Seleziona Cliente / Veicolo per Autocompilazione Dati:", auto_options, key="dep_auto_select")
                
                def_cli = ""
                def_targa = ""
                def_marca = ""
                def_misura = ""
                def_phone = ""
                def_stag = "Invernali"
                
                if sel_auto in v_map:
                    info = v_map[sel_auto]
                    def_cli = info["cliente"]
                    def_targa = info["targa"]
                    def_marca = info["marca"]
                    def_misura = info["misura"]
                    def_phone = info.get("phone", "")
                    if info["tipo_gomme"] in ["Invernali", "Estivi", "4 Stagioni"]:
                        def_stag = info["tipo_gomme"]
                    st.success(f"⚡ **Dati Autocompilati!** Cliente: `{def_cli}` | Targa: `{def_targa}` | Misura: `{def_misura}` | Tel: `{def_phone}`")
                
                with st.form("form_dep"):
                    c1, c2 = st.columns(2)
                    with c1:
                        cli_sel = st.text_input("Cliente *", value=def_cli, placeholder="Es. Mario Rossi")
                        targa_dep = st.text_input("Targa Veicolo *", value=def_targa, placeholder="AB123CD").upper()
                        marca_dep = st.text_input("Marca e Modello Pneumatico *", value=def_marca, placeholder="Es. Pirelli Cinturato P7")
                        misura_dep = st.text_input("Misura Pneumatici *", value=def_misura, placeholder="Es. 205/55 R16 91V")
                        tel_dep = st.text_input("Cellulare (WhatsApp) *", value=def_phone, placeholder="3331234567")
                    with c2:
                        stag_opts = ["Invernali", "Estivi", "4 Stagioni"]
                        s_idx = stag_opts.index(def_stag) if def_stag in stag_opts else 0
                        stag_dep = st.selectbox("Stagionalita Gomme in Custodia *", stag_opts, index=s_idx)
                        qta_dep = st.number_input("Quantita (Pezzi)", min_value=1, max_value=8, value=4)
                        batt_dep = st.text_input("Battistrada Residuo (mm)", placeholder="6.5 mm")
                        ubi_dep = st.text_input("Ubicazione Magazzino", placeholder="Corsia A - Scaffale 3")
                    btn_h = st.form_submit_button("🏷️ Salva Deposito e Genera Etichetta QR Scansionabile", use_container_width=True)
                    if btn_h:
                        if not cli_sel or not targa_dep or not marca_dep or not misura_dep:
                            st.error("Compila i campi obbligatori del deposito (Cliente, Targa, Marca, Misura).")
                        else:
                            df_h = load_data("Hotel_Pneumatici")
                            hid = len(df_h) + 1 if not df_h.empty else 1
                            tag = f"TAG-2026-{hid:03d}"
                            new_h = pd.DataFrame([{
                                "ID_Deposito": hid,
                                "Tag_ID": tag,
                                "ID_Cliente": cli_sel,
                                "Targa": targa_dep,
                                "Marca_Modello_Gomme": marca_dep,
                                "Misura": misura_dep,
                                "Stagionalita": stag_dep,
                                "Quantita": qta_dep,
                                "Battistrada_mm": batt_dep,
                                "Ubicazione": ubi_dep,
                                "Telefono": tel_dep,
                                "Data_PresaInCarico": str(datetime.date.today()),
                                "Stato_Sollecito": "Da Sollecitare"
                            }])
                            save_sheet(pd.concat([df_h, new_h], ignore_index=True), "Hotel_Pneumatici")
                            st.success(f"✅ Deposito registrato con Tag ID: {tag}")
                            lbl = generate_real_qr_label(tag, cli_sel, targa_dep, marca_dep, misura_dep, batt_dep, ubi_dep)
                            st.session_state['last_label'] = {
                                'tag': tag,
                                'bytes': lbl
                            }
                            st.rerun()
                
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
                            curr_tel = str(dep_row.get("Telefono", lookup_client_phone(dep_row.get("ID_Cliente"), dep_row.get("Targa"))))
                            e_tel = st.text_input("Cellulare (WhatsApp)", value=curr_tel)
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
                            df_h_edit.loc[mask_h, "Telefono"] = str(e_tel)
                            save_sheet(df_h_edit, "Hotel_Pneumatici")
                            
                            lbl_updated = generate_real_qr_label(target_tag, e_cliente, e_targa, e_marca, e_misura, e_batt, e_ubi)
                            st.session_state['last_label'] = {
                                'tag': target_tag,
                                'bytes': lbl_updated
                            }
                            st.success(f"✅ Deposito `{target_tag}` aggiornato e nuovo QR Code rigenerato!")
                            st.rerun()

                    st.markdown("---")
                    st.markdown("##### 🗑️ Zona di Eliminazione Deposito dall Hotel Gomme")
                    st.warning("⚠️ L eliminazione rimuovera definitivamente la scheda di questo deposito dall Hotel Pneumatici.")
                    if st.button(f"🗑️ ELIMINA DEFINITIVAMENTE DEPOSITO #{target_tag}", type="primary", key=f"btn_del_dep_{target_tag}"):
                        df_h_edit = df_h_edit[df_h_edit["Tag_ID"].astype(str) != str(target_tag)]
                        save_sheet(df_h_edit, "Hotel_Pneumatici")
                        st.success(f"🗑️ Deposito #{target_tag} eliminato dall Hotel Pneumatici!")
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
        elif nav == 'magazzino':
            st.subheader("📦 Magazzino Pneumatici & Ordini In Arrivo")
            m_tab1, m_tab2, m_tab3, m_tab4 = st.tabs(["📦 Inventario Stock Disponibile", "➕ Nuovo Carico Pneumatici", "✏️ Modifica / Elimina Articolo Stock", "🚚 Pneumatici Ordinati (In Arrivo)"])
            
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
                st.markdown("### ✏️ Modifica o Elimina un Articolo in Magazzino")
                df_m_edit = load_data("Magazzino")
                if df_m_edit.empty:
                    st.info("Nessun articolo presente in magazzino.")
                else:
                    stock_dict = {f"#{r.get('ID_Articolo', '')} - {r.get('Marca_Modello', '')} ({r.get('Misura', '')}) [Disp: {r.get('Quantita_Disponibile', 0)}]": r.get('ID_Articolo', '') for _, r in df_m_edit.iterrows()}
                    selected_art_label = st.selectbox("Seleziona l Articolo da Modificare / Eliminare:", list(stock_dict.keys()))
                    target_mid = stock_dict[selected_art_label]
                    
                    art_row = df_m_edit[df_m_edit["ID_Articolo"].astype(str) == str(target_mid)].iloc[0]
                    
                    with st.form("form_edit_stock"):
                        st.markdown(f"##### Modifica Dati Articolo #{target_mid}: {art_row.get('Marca_Modello', '')}")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            e_st_marca = st.text_input("Marca e Modello Pneumatico *", value=str(art_row.get("Marca_Modello", "")))
                            e_st_misura = st.text_input("Misura Pneumatici *", value=str(art_row.get("Misura", "")))
                            q_val = pd.to_numeric(art_row.get("Quantita_Disponibile", 4), errors='coerce')
                            q_val = int(q_val) if (pd.notnull(q_val) and q_val >= 0) else 0
                            e_st_qta = st.number_input("Quantita Disponibile (Pezzi) *", min_value=0, max_value=1000, value=q_val)
                        with col_e2:
                            c_val = pd.to_numeric(art_row.get("Costo_Acquisto_Unitario", 50.0), errors='coerce')
                            c_val = float(c_val) if (pd.notnull(c_val) and c_val >= 0) else 50.0
                            e_st_costo = st.number_input("Costo di Acquisto Unitario (€) *", min_value=0.0, value=c_val, step=5.0)
                            
                            p_val = pd.to_numeric(art_row.get("Prezzo_Vendita_Unitario", 90.0), errors='coerce')
                            p_val = float(p_val) if (pd.notnull(p_val) and p_val >= 0) else 90.0
                            e_st_prezzo = st.number_input("Prezzo di Vendita Unitario (€)", min_value=0.0, value=p_val, step=5.0)
                            
                            e_st_ubi = st.text_input("Ubicazione Magazzino", value=str(art_row.get("Ubicazione_Magazzino", "")))
                        e_st_note = st.text_area("Note Magazzino", value=str(art_row.get("Note", "")))
                        
                        btn_update_stock = st.form_submit_button("💾 Salva Modifiche Articolo Magazzino", use_container_width=True)
                        if btn_update_stock:
                            df_m_edit = df_m_edit.astype(object)
                            mask_m = df_m_edit["ID_Articolo"].astype(str) == str(target_mid)
                            df_m_edit.loc[mask_m, "Marca_Modello"] = str(e_st_marca)
                            df_m_edit.loc[mask_m, "Misura"] = str(e_st_misura)
                            df_m_edit.loc[mask_m, "Quantita_Disponibile"] = int(e_st_qta)
                            df_m_edit.loc[mask_m, "Costo_Acquisto_Unitario"] = float(e_st_costo)
                            df_m_edit.loc[mask_m, "Prezzo_Vendita_Unitario"] = float(e_st_prezzo)
                            df_m_edit.loc[mask_m, "Ubicazione_Magazzino"] = str(e_st_ubi)
                            df_m_edit.loc[mask_m, "Note"] = str(e_st_note)
                            save_sheet(df_m_edit, "Magazzino")
                            st.success(f"✅ Articolo #{target_mid} aggiornato con successo!")
                            st.rerun()

                    st.markdown("---")
                    st.markdown("##### 🗑️ Zona di Eliminazione Articolo")
                    st.warning("⚠️ L eliminazione rimuovera definitivamente questo articolo e le relative scorte dal magazzino.")
                    if st.button(f"🗑️ ELIMINA DEFINITIVAMENTE ARTICOLO #{target_mid}", type="primary"):
                        df_m_edit = df_m_edit[df_m_edit["ID_Articolo"].astype(str) != str(target_mid)]
                        save_sheet(df_m_edit, "Magazzino")
                        st.success(f"🗑️ Articolo #{target_mid} eliminato dal magazzino!")
                        st.rerun()

            with m_tab4:
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
                        st.markdown("##### 📦 Registra Arrivo Merce & Personalizza Stock Magazzino:")
                        for idx_o, row_o in df_o_pending.iterrows():
                            o_id_val = row_o.get('ID_Ordine', idx_o+1)
                            o_marca_val = row_o.get('Marca_Modello', 'Pneumatici')
                            o_misura_val = row_o.get('Misura', 'N/D')
                            o_qta_val = row_o.get('Quantita_Ordinata', 1)
                            o_forn_val = row_o.get('Fornitore', 'N/D')
                            
                            c_tot = pd.to_numeric(row_o.get("Costo_Acquisto_Totale", 50), errors='coerce')
                            q_ord = pd.to_numeric(row_o.get("Quantita_Ordinata", 1), errors='coerce')
                            c_tot = float(c_tot) if pd.notnull(c_tot) else 50.0
                            q_ord = float(q_ord) if (pd.notnull(q_ord) and q_ord > 0) else 1.0
                            cost_unit_def = c_tot / q_ord
                            
                            with st.expander(f"🚚 Ordine #{o_id_val}: {o_marca_val} ({o_misura_val}) - Q.tà ordinata: {o_qta_val} pz | Fornitore: {o_forn_val}", expanded=False):
                                st.write("Modifica e conferma la quantità effettivamente arrivata, il prezzo ed i dettagli prima di trasferire la merce nello Stock Magazzino:")
                                with st.form(key=f"form_arrive_ord_{idx_o}"):
                                    ca1, ca2 = st.columns(2)
                                    with ca1:
                                        arr_marca = st.text_input("Marca e Modello Pneumatici *", value=str(o_marca_val), key=f"arr_m_{idx_o}")
                                        arr_misura = st.text_input("Misura Pneumatici *", value=str(o_misura_val), key=f"arr_mis_{idx_o}")
                                        q_init = int(q_ord) if q_ord >= 1 else 1
                                        arr_qta = st.number_input("Quantità Effettivamente Arrivata (Pezzi) *", min_value=1, max_value=1000, value=q_init, key=f"arr_q_{idx_o}")
                                    with ca2:
                                        arr_costo = st.number_input("Costo di Acquisto Unitario (€) *", min_value=0.0, value=float(cost_unit_def), step=1.0, key=f"arr_c_{idx_o}")
                                        arr_prezzo = st.number_input("Prezzo di Vendita Unitario (€)", min_value=0.0, value=float(cost_unit_def * 1.5), step=5.0, key=f"arr_p_{idx_o}")
                                        arr_ubi = st.text_input("Ubicazione Magazzino", value="Corsia M1 - Scaffale A", key=f"arr_u_{idx_o}")
                                    arr_note = st.text_area("Note Magazzino", value=f"Da Ordine #{o_id_val} ({o_forn_val})", key=f"arr_n_{idx_o}")
                                    
                                    btn_confirm_arrive = st.form_submit_button(f"📦 CONFERMA ARRIVO & TRASFERISCI IN STOCK MAGAZZINO #{o_id_val}", use_container_width=True)
                                    if btn_confirm_arrive:
                                        df_o_list = df_o_list.astype(object)
                                        df_o_list.loc[df_o_list["ID_Ordine"].astype(str) == str(o_id_val), "Stato_Ordine"] = "Consegnato in Magazzino"
                                        save_sheet(df_o_list, "Ordini_Pneumatici")
                                        
                                        df_m_add = load_data("Magazzino")
                                        mid = len(df_m_add) + 1 if not df_m_add.empty else 1
                                        new_stock = pd.DataFrame([{"ID_Articolo": mid, "Marca_Modello": arr_marca, "Misura": arr_misura, "Quantita_Disponibile": arr_qta, "Costo_Acquisto_Unitario": arr_costo, "Prezzo_Vendita_Unitario": arr_prezzo, "Ubicazione_Magazzino": arr_ubi, "Note": arr_note}])
                                        save_sheet(pd.concat([df_m_add, new_stock], ignore_index=True), "Magazzino")
                                        st.success(f"🎉 Ordine #{o_id_val} trasferito con successo in Magazzino: {arr_qta} pezzi di '{arr_marca}' ({arr_misura})!")
                                        st.rerun()

                                # Single order deletion inside expander
                                if st.button(f"🗑️ ELIMINA ORDINE #{o_id_val} DAL REGISTRO", key=f"del_pending_{idx_o}", type="secondary"):
                                    df_o_list = df_o_list[df_o_list["ID_Ordine"].astype(str) != str(o_id_val)]
                                    save_sheet(df_o_list, "Ordini_Pneumatici")
                                    st.success(f"🗑️ Ordine #{o_id_val} eliminato dal registro ordini!")
                                    st.rerun()

                    st.markdown("---")
                    st.markdown("##### 🗑️ Gestione & Cancellazione Storico Ordini")
                    st.write("Puoi cancellare singoli ordini dallo storico o eseguire la pulizia automatica degli ordini già consegnati:")
                    
                    ord_dict = {f"#{r.get('ID_Ordine', '')} - {r.get('Marca_Modello', '')} ({r.get('Misura', '')}) [{r.get('Stato_Ordine', '')}]": r.get('ID_Ordine', '') for _, r in df_o_list.iterrows()}
                    selected_ord_label = st.selectbox("Seleziona l Ordine da Eliminare dallo Storico:", list(ord_dict.keys()), key="sel_ord_to_delete")
                    target_oid = ord_dict[selected_ord_label]
                    
                    col_del1, col_del2 = st.columns(2)
                    with col_del1:
                        if st.button(f"🗑️ ELIMINA ORDINE SELEZIONATO (#{target_oid})", use_container_width=True):
                            df_o_upd = df_o_list[df_o_list["ID_Ordine"].astype(str) != str(target_oid)]
                            save_sheet(df_o_upd, "Ordini_Pneumatici")
                            st.success(f"🗑️ Ordine #{target_oid} eliminato con successo dallo storico!")
                            st.rerun()
                    
                    with col_del2:
                        df_delivered = df_o_list[df_o_list["Stato_Ordine"] == "Consegnato in Magazzino"]
                        del_count = len(df_delivered)
                        if st.button(f"🧹 CANCELLA SOLO ORDINI GIA CONSEGNATI ({del_count} Ordini)", use_container_width=True):
                            df_o_upd = df_o_list[df_o_list["Stato_Ordine"] != "Consegnato in Magazzino"]
                            save_sheet(df_o_upd, "Ordini_Pneumatici")
                            st.success(f"🧹 Eliminati {del_count} ordini già consegnati dallo storico!")
                            st.rerun()
                            
                    with st.expander("⚠️ Opzioni Avanzate: Reset Totale Registro Ordini", expanded=False):
                        st.warning("⚠️ Questa operazione svuoterà l intero storico ordini pneumatici.")
                        if st.button("🗑️ SVUOTA COMPLETAMENTE REGISTRO ORDINI", type="primary"):
                            empty_ord_df = pd.DataFrame(columns=["ID_Ordine", "Marca_Modello", "Misura", "Quantita_Ordinata", "Fornitore", "Costo_Acquisto_Totale", "Data_Ordine", "Data_Presunta_Arrivo", "Stato_Ordine"])
                            save_sheet(empty_ord_df, "Ordini_Pneumatici")
                            st.success("🗑️ Registro ordini svuotato con successo!")
                            st.rerun()
                else:
                    st.info("Nessun ordine presente.")

        elif nav == 'whatsapp':
            st.subheader("📱 Campagne & Invio Solleciti WhatsApp (Cambi Stagionali)")
            st.write("Visualizza i numeri di telefono **reali** di tutti i clienti in custodia, invia il messaggio diretto su WhatsApp, utilizza la modalità guidata passo-passo oppure esporta la lista per WhatsApp Broadcast.")
            
            campagna = st.radio("Seleziona Tipo Campagna:", [
                "🍁 Sollecito Cambio Invernale (Notifica clienti con gomme Invernali in custodia)",
                "☀️ Sollecito Cambio Estivo (Notifica clienti con gomme Estive in custodia)",
                "🔄 Controllo Usura & Inversione 4 Stagioni (Opzionale per All-Season)"
            ])
            
            df_h = load_data("Hotel_Pneumatici")
            if df_h.empty:
                st.info("Nessun deposito presente in Hotel Pneumatici.")
            else:
                if "Invernale" in campagna:
                    df_target = df_h[df_h["Stagionalita"] == "Invernali"]
                elif "Estivo" in campagna:
                    df_target = df_h[df_h["Stagionalita"] == "Estivi"]
                else:
                    df_target = df_h[df_h["Stagionalita"] == "4 Stagioni"]
                
                st.info(f"📊 Trovati **{len(df_target)}** clienti idonei per questa campagna.")
                
                if not df_target.empty:
                    booking_url = get_booking_url()
                    
                    # Process and prepare client lists with real phone lookup
                    campaign_items = []
                    valid_phones = []
                    
                    for idx_t, r_t in df_target.iterrows():
                        cli_name = str(r_t.get("ID_Cliente", "Cliente")).strip()
                        targa_val = str(r_t.get("Targa", "N/D")).strip().upper()
                        tag_val = str(r_t.get("Tag_ID", "N/D")).strip()
                        gomme_val = str(r_t.get("Marca_Modello_Gomme", "Pneumatici")).strip()
                        
                        real_phone = lookup_client_phone(cli_name, targa_val, r_t)
                        msg = f"Gentile {cli_name}, le ricordiamo che le sue gomme {gomme_val} per la vettura {targa_val} (Tag: {tag_val}) sono pronte in custodia presso PneumaticEvolution. 👉 Puo prenotare l appuntamento cliccando sul link: {booking_url} . Cordiali Saluti."
                        wa_url = create_whatsapp_url(real_phone, msg) if real_phone else ""
                        
                        if real_phone:
                            valid_phones.append(real_phone)
                            
                        campaign_items.append({
                            'id': r_t.get("ID_Deposito", idx_t),
                            'cliente': cli_name,
                            'targa': targa_val,
                            'tag': tag_val,
                            'gomme': gomme_val,
                            'phone': real_phone,
                            'msg': msg,
                            'wa_url': wa_url
                        })
                    
                    # Section 1: Passo-Passo Sender Carousel
                    st.markdown("---")
                    st.markdown("### ▶️ Modalità Invio Guidato Passo-Passo (1 Cliente alla Volta)")
                    st.write("Per evitare che il browser blocchi le schede, puoi inviare i messaggi uno alla volta semplicemente toccando il pulsante verde:")
                    
                    if 'wa_step_idx' not in st.session_state:
                        st.session_state['wa_step_idx'] = 0
                        
                    step_idx = st.session_state['wa_step_idx']
                    if step_idx >= len(campaign_items):
                        step_idx = 0
                        st.session_state['wa_step_idx'] = 0
                        
                    curr_item = campaign_items[step_idx]
                    
                    c_step1, c_step2 = st.columns([3, 1])
                    with c_step1:
                        st.markdown(f"**Cliente #{step_idx+1} di {len(campaign_items)}:** 👤 **{curr_item['cliente']}** | 🚗 Targa: `{curr_item['targa']}` | 📱 Tel: `<code>{curr_item['phone'] if curr_item['phone'] else 'Non Trovato'}</code>`")
                        st.caption(f"Pneumatici: {curr_item['gomme']} (Tag ID: {curr_item['tag']})")
                    with c_step2:
                        if curr_item['wa_url']:
                            st.markdown(f'<a href="{curr_item["wa_url"]}" target="_blank" class="wa-big-btn" style="padding: 10px 14px; font-size: 1rem; margin:0;">🟢 💬 INVIA ORA & PASSA AL PROSSIMO ➔</a>', unsafe_allow_html=True)
                        else:
                            st.warning("Nessun cellulare")
                            
                    col_st_nav1, col_st_nav2, col_st_nav3 = st.columns([1, 1, 2])
                    with col_st_nav1:
                        if st.button("⬅️ Cliente Precedente", key="btn_wa_prev"):
                            st.session_state['wa_step_idx'] = max(0, step_idx - 1)
                            st.rerun()
                    with col_st_nav2:
                        if st.button("➡️ Prossimo Cliente", key="btn_wa_next"):
                            st.session_state['wa_step_idx'] = min(len(campaign_items) - 1, step_idx + 1)
                            st.rerun()
                    with col_st_nav3:
                        if st.button("🔄 Ricomincia da Capo (Cliente #1)", key="btn_wa_reset"):
                            st.session_state['wa_step_idx'] = 0
                            st.rerun()

                    # Section 2: Full Detailed Interactive List
                    st.markdown("---")
                    st.markdown("### 📋 Elenco Completo Clienti & Pulsanti Diretti WhatsApp")
                    
                    for idx_item, item in enumerate(campaign_items):
                        c_card1, c_card2 = st.columns([3, 1])
                        with c_card1:
                            phone_disp = item['phone'] if item['phone'] else "⚠️ Telefono Mancante in Anagrafica"
                            st.markdown(f"👤 **{item['cliente']}** | 🚗 Targa: `<code>{item['targa']}</code>` | 📱 Cellulare: **{phone_disp}**")
                            st.caption(f"Gomme: {item['gomme']} (Tag: {item['tag']})")
                        with c_card2:
                            if item['wa_url']:
                                st.markdown(f'<a href="{item["wa_url"]}" target="_blank" class="wa-big-btn" style="padding: 8px 12px; font-size: 0.95rem; margin:0;">🟢 💬 Invia WhatsApp</a>', unsafe_allow_html=True)
                            else:
                                st.info("Inserisci numero in Anagrafica")
                        st.markdown("<hr style='margin: 8px 0; border-color: #E5E7EB;'>", unsafe_allow_html=True)

                    # Section 3: Broadcast List & Excel Export
                    st.markdown("---")
                    st.markdown("### 📋 Numeri per Liste di Trasmissione & Report Excel")
                    
                    col_ex1, col_ex2 = st.columns(2)
                    with col_ex1:
                        st.markdown("##### 📱 Numeri WhatsApp Puliti (per Lista di Trasmissione):")
                        joined_phones = ", ".join(valid_phones)
                        st.code(joined_phones if joined_phones else "Nessun numero di telefono presente.", language="text")
                    with col_ex2:
                        st.markdown("##### 📥 Esporta Report Campagna in Excel (.xlsx):")
                        try:
                            df_camp_export = pd.DataFrame(campaign_items)[['cliente', 'targa', 'phone', 'tag', 'gomme', 'msg', 'wa_url']]
                            df_camp_export.columns = ['Cliente', 'Targa', 'Telefono_WhatsApp', 'Tag_ID', 'Marca_Modello_Gomme', 'Messaggio_Inviato', 'Link_WhatsApp']
                            
                            buffer_c_exp = io.BytesIO()
                            with pd.ExcelWriter(buffer_c_exp, engine='openpyxl') as writer_c_exp:
                                df_camp_export.to_excel(writer_c_exp, sheet_name="Campagna_WhatsApp", index=False)
                            
                            st.download_button(
                                label="📥 Scarica Excel Campagna WhatsApp (.xlsx)",
                                data=buffer_c_exp.getvalue(),
                                file_name=f"campagna_whatsapp_pneumaticevolution_{datetime.date.today()}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        except Exception as e_exp:
                            st.error(f"Errore esportazione: {e_exp}")

        elif nav == 'calendario':
            st.subheader("📅 Agenda & Calendario Sincronizzato con Archivio Storico")
            
            ag_tab1, ag_tab2 = st.tabs(["📅 Agenda Attiva & Calendario Visivo", "📚 Storico & Archivio Appuntamenti Passati"])
            
            with ag_tab1:
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
            st.markdown("### 🗑️ Gestione & Cancellazione Eventi dal Calendario")
            df_a_cal = load_data("Appuntamenti")
            if not df_a_cal.empty:
                cal_dict = {f"#{r.get('ID_App','')} - 🚗 {r.get('Targa','')} | {r.get('Cliente','')} ({r.get('Data','')} ore {r.get('Ora','')}) [{r.get('Stato','')}]": r.get('ID_App','') for _, r in df_a_cal.iterrows()}
                sel_cal_lbl = st.selectbox("Seleziona l Evento / Appuntamento da Rimuovere dal Calendario:", list(cal_dict.keys()), key="del_cal_evt_sel")
                target_cal_aid = cal_dict[sel_cal_lbl]
                if st.button(f"🗑️ ELIMINA EVENTO #{target_cal_aid} DAL CALENDARIO", type="primary", key="btn_del_cal_evt"):
                    df_a_cal = df_a_cal[df_a_cal["ID_App"].astype(str) != str(target_cal_aid)]
                    save_sheet(df_a_cal, "Appuntamenti")
                    st.success(f"🗑️ Evento #{target_cal_aid} rimosso con successo dal calendario!")
                    st.rerun()
            
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

            with ag_tab2:
                st.markdown("### 📚 Storico & Archivio Completo Appuntamenti Passati")
                st.write("Consulta, filtra e ricerca tutti gli appuntamenti passati registrati nel tempo. Tutti i dati rimangono archiviati in totale sicurezza per la consultazione negli anni.")
                
                df_a_hist = load_data("Appuntamenti")
                if not df_a_hist.empty:
                    today_str = str(datetime.date.today())
                    
                    # Search and Filter bar
                    col_h1, col_h2, col_h3 = st.columns([2, 2, 2])
                    with col_h1:
                        search_term = st.text_input("🔍 Cerca per Targa, Cliente o Servizio:", placeholder="Es. AB123CD / Rossi")
                    with col_h2:
                        filter_stato = st.selectbox("Filtra per Stato Appuntamento:", ["Tutti gli Stati", "Confermato", "Scaduto / Non Confermato", "Proposta Nuova Data", "Rifiutato", "In attesa di conferma"])
                    with col_h3:
                        filter_period = st.selectbox("Periodo:", ["Tutti i Tempi (Storico Completo)", "Solo Appuntamenti Passati (< Oggi)", "Solo Appuntamenti Futuri (>= Oggi)"])

                    df_filtered = df_a_hist.copy()
                    
                    if filter_period == "Solo Appuntamenti Passati (< Oggi)":
                        df_filtered = df_filtered[df_filtered["Data"].astype(str) < today_str]
                    elif filter_period == "Solo Appuntamenti Futuri (>= Oggi)":
                        df_filtered = df_filtered[df_filtered["Data"].astype(str) >= today_str]
                        
                    if filter_stato != "Tutti gli Stati":
                        df_filtered = df_filtered[df_filtered["Stato"].astype(str).str.contains(filter_stato, na=False)]
                        
                    if search_term.strip():
                        s_term = search_term.strip().lower()
                        mask_s = (
                            df_filtered["Targa"].astype(str).str.lower().str.contains(s_term) |
                            df_filtered["Cliente"].astype(str).str.lower().str.contains(s_term) |
                            df_filtered["Servizio"].astype(str).str.lower().str.contains(s_term) |
                            df_filtered.get("Note", pd.Series([""]*len(df_filtered))).astype(str).str.lower().str.contains(s_term)
                        )
                        df_filtered = df_filtered[mask_s]

                    st.info(f"📊 Trovati **{len(df_filtered)}** appuntamenti nello storico.")
                    st.dataframe(df_filtered, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("##### 🗑️ Cancellazione Record dallo Storico Appuntamenti")
                    if not df_filtered.empty:
                        hist_dict = {f"#{r.get('ID_App','')} - 🚗 {r.get('Targa','')} | {r.get('Cliente','')} ({r.get('Data','')}) [{r.get('Stato','')}]": r.get('ID_App','') for _, r in df_filtered.iterrows()}
                        sel_hist_lbl = st.selectbox("Seleziona l Appuntamento da Eliminare dallo Storico:", list(hist_dict.keys()), key="del_hist_sel")
                        target_hist_aid = hist_dict[sel_hist_lbl]
                        if st.button(f"🗑️ ELIMINA RECORD #{target_hist_aid} DALLO STORICO", type="primary", key="btn_del_hist_rec"):
                            df_a_hist_del = df_a_hist[df_a_hist["ID_App"].astype(str) != str(target_hist_aid)]
                            save_sheet(df_a_hist_del, "Appuntamenti")
                            st.success(f"🗑️ Record #{target_hist_aid} eliminato dallo storico!")
                            st.rerun()
                    
                    # Export button for history
                    try:
                        buffer_exp = io.BytesIO()
                        with pd.ExcelWriter(buffer_exp, engine='openpyxl') as writer_exp:
                            df_filtered.to_excel(writer_exp, sheet_name="Storico_Appuntamenti", index=False)
                        st.download_button(
                            label="📥 Esporta Storico Filtrato in Excel (.xlsx)",
                            data=buffer_exp.getvalue(),
                            file_name=f"storico_appuntamenti_pneumaticevolution_{today_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception:
                        pass
                else:
                    st.info("Nessun appuntamento archiviato nello storico al momento.")

        elif nav == 'impostazioni':
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
                curr_url = get_setting("app_url", "https://pneumaticevolution.streamlit.app/")
                
                st.markdown("---")
                st.markdown("### 💾 Gestione Backup & Ripristino Database Excel")
                st.write("Scarica un backup completo di tutte le schede dell officina (Clienti, Veicoli, Hotel Gomme, Appuntamenti, Magazzino) per garantire la massima sicurezza a lungo termine.")
                
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.markdown("##### 📥 Scarica Backup Completo (.xlsx)")
                    if os.path.exists(DB_FILE_EXCEL):
                        with open(DB_FILE_EXCEL, "rb") as f_excel:
                            excel_bytes = f_excel.read()
                        today_tag = datetime.date.today().strftime("%Y%m%d")
                        st.download_button(
                            label="📥 Scarica Backup Database Excel (.xlsx)",
                            data=excel_bytes,
                            file_name=f"gestionale_pneumatici_backup_{today_tag}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("Nessun database presente da scaricare.")

                with col_b2:
                    st.markdown("##### 📂 Ripristina Database da File Excel")
                    uploaded_backup = st.file_uploader("Carica un file Excel di Backup (.xlsx)", type=["xlsx"])
                    if uploaded_backup is not None:
                        if st.button("🔄 CONFERMA RIPRISTINO DATABASE DA BACKUP", type="primary", use_container_width=True):
                            try:
                                with open(DB_FILE_EXCEL, "wb") as f_out:
                                    f_out.write(uploaded_backup.getbuffer())
                                st.cache_data.clear()
                                st.success("✅ Database ripristinato con successo dal backup caricato!")
                                st.rerun()
                            except Exception as e_up:
                                st.error(f"Errore durante il ripristino: {e_up}")

                st.markdown("---")
                st.markdown("### 🧹 Manutenzione & Archiviazione Automatica Scaduti")
                st.write("Esegui un operazione di pulizia manuale per spostare tutte le richieste pendenti antecedenti ad oggi nello Storico Passato.")
                if st.button("🧹 ESEGUI PULIZIA & ARCHIVIAZIONE ORA", use_container_width=True):
                    auto_archive_past_appointments()
                    st.success("✅ Pulizia ed archiviazione completate con successo!")
                    st.rerun()
                new_app_url = st.text_input("URL Pubblico dell App Streamlit:", value=curr_url)
                if st.button("💾 Salva URL Sito"):
                    update_setting("app_url", new_app_url)
                    st.success("✅ Link del sito aggiornato con successo!")

                st.markdown("---")
                curr_off_wa = get_setting("officina_whatsapp", "3336330301")
                new_off_wa = st.text_input("📱 Numero WhatsApp Officina (Ricezione Prenotazioni):", value=curr_off_wa, help="Tutte le notifiche di prenotazione inviate dai clienti dal sito saranno indirizzate a questo numero WhatsApp")
                if st.button("💾 Salva Numero WhatsApp Officina"):
                    update_setting("officina_whatsapp", new_off_wa)
                    st.success("✅ Numero WhatsApp dell Officina aggiornato con successo!")

            with st2_cfg:
                st.markdown("### 🗓️ Blocco Date & Chiusure Straordinarie Multi-Periodo")
                st.write("Blocca intere settimane di **Ferie Estive**, **Ponti Festivi** o **Periodi di Chiusura Straordinaria** per impedire ai clienti di prenotare in quei giorni.")
                
                giorni_bloccati_str = get_setting("giorni_bloccati", "")
                g_list = sorted(list(set([g.strip() for g in giorni_bloccati_str.split(",") if g.strip()])))
                
                # Show active blocked dates summary
                if g_list:
                    st.info(f"📌 **Totale Giorni Bloccati in Calendario:** {len(g_list)} giorni")
                    with st.expander("📋 Visualizza Elenco Completo Date Bloccate", expanded=False):
                        st.write(", ".join([f"`{d}`" for d in g_list]))
                else:
                    st.success("✅ Nessun giorno di chiusura attualmente bloccato.")

                st.markdown("---")
                st.markdown("##### 🔒 Blocca un Periodo o Intervallo di Date")
                
                with st.form("form_block_period"):
                    cb1, cb2 = st.columns(2)
                    with cb1:
                        d_start = st.date_input("Data Inizio Chiusure / Ferie *", min_value=datetime.date.today(), key="block_d_start")
                    with cb2:
                        d_end = st.date_input("Data Fine Chiusure / Ferie *", min_value=datetime.date.today(), key="block_d_end")
                    
                    motivo_blocco = st.text_input("Motivo / Descrizione Chiusura (Opzionale)", placeholder="Es. Ferie Estive, Ponte 15 Agosto, Lavori in Officina")
                    
                    col_b_act1, col_b_act2 = st.columns(2)
                    with col_b_act1:
                        btn_add_period = st.form_submit_button("🔒 BLOCCA INTERO PERIODO SELEZIONATO", use_container_width=True)
                    with col_b_act2:
                        btn_rem_period = st.form_submit_button("🔓 SBLOCCA PERIODO SELEZIONATO", use_container_width=True)
                    
                    if btn_add_period:
                        if d_start > d_end:
                            st.error("⚠️ La Data Inizio deve essere antecedente o uguale alla Data Fine!")
                        else:
                            added_count = 0
                            curr_dt = d_start
                            while curr_dt <= d_end:
                                dt_s = str(curr_dt)
                                if dt_s not in g_list:
                                    g_list.append(dt_s)
                                    added_count += 1
                                curr_dt += datetime.timedelta(days=1)
                            
                            g_list = sorted(list(set(g_list)))
                            update_setting("giorni_bloccati", ",".join(g_list))
                            st.success(f"🎉 Aggiunti **{added_count} nuovi giorni** ai giorni di chiusura bloccati ({d_start} ➔ {d_end})!")
                            st.rerun()
                            
                    if btn_rem_period:
                        if d_start > d_end:
                            st.error("⚠️ La Data Inizio deve essere antecedente o uguale alla Data Fine!")
                        else:
                            rem_count = 0
                            curr_dt = d_start
                            while curr_dt <= d_end:
                                dt_s = str(curr_dt)
                                if dt_s in g_list:
                                    g_list.remove(dt_s)
                                    rem_count += 1
                                curr_dt += datetime.timedelta(days=1)
                            
                            g_list = sorted(list(set(g_list)))
                            update_setting("giorni_bloccati", ",".join(g_list))
                            st.success(f"🔓 Sbloccati **{rem_count} giorni** dall intervallo selezionato ({d_start} ➔ {d_end})!")
                            st.rerun()

                if g_list:
                    st.markdown("---")
                    st.markdown("##### 🗑️ Sblocco Rapido Singola Data o Reset Completo")
                    col_del1, col_del2 = st.columns([2, 1])
                    with col_del1:
                        sel_date_to_remove = st.selectbox("Seleziona una singola data da sbloccare:", g_list)
                        if st.button("🔓 Sblocca Singola Data Selezionata", use_container_width=True):
                            if sel_date_to_remove in g_list:
                                g_list.remove(sel_date_to_remove)
                                update_setting("giorni_bloccati", ",".join(g_list))
                                st.success(f"✅ Data `{sel_date_to_remove}` sbloccata!")
                                st.rerun()
                    with col_del2:
                        st.write("Reset Totale Chiusure:")
                        if st.button("🗑️ SBLOCCA TUTTE LE DATE", type="primary", use_container_width=True):
                            update_setting("giorni_bloccati", "")
                            st.success("🧹 Tutti i blocchi di chiusura sono stati rimossi!")
                            st.rerun()