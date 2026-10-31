import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
from PIL import Image, ImageDraw
import io
import urllib.parse

st.set_page_config(
    page_title="PneumaticEvolution Gestionale",
    page_icon="🛞",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('''
<style>
    .main-header { font-size: 2.2rem; color: #1E3A8A; font-weight: 700; margin-bottom: 0px; }
    .sub-header { font-size: 1.1rem; color: #4B5563; margin-bottom: 20px; }
    .card-box { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 18px; margin-bottom: 15px; }
    .stButton>button { border-radius: 6px; }
</style>
''', unsafe_allow_html=True)

# Resolution function for Database file path to prevent FileNotFoundError on Cloud
def resolve_db_path():
    possible_paths = [
        "gestionale_pneumatici_db_v9.xlsx",
        "gestionale_pneumatici_db_v8.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v9.xlsx",
        "/workspace/artifacts/gestionale_pneumatici_db_v8.xlsx"
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return "gestionale_pneumatici_db_v9.xlsx"

DB_FILE_EXCEL = resolve_db_path()

def ensure_db_exists():
    global DB_FILE_EXCEL
    if not os.path.exists(DB_FILE_EXCEL):
        DB_FILE_EXCEL = "gestionale_pneumatici_db_v9.xlsx"
        if not os.path.exists(DB_FILE_EXCEL):
            wb = openpyxl.Workbook()
            default_sheets = ["Clienti", "Veicoli", "Hotel_Pneumatici", "Appuntamenti", "Magazzino", "Flotta_Camion"]
            wb.active.title = default_sheets[0]
            for s in default_sheets[1:]:
                wb.create_sheet(s)
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
            st.error(f"Errore durante il salvataggio dei dati: {e}")
    st.cache_data.clear()

if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def generate_tag_label(tag_id, cliente, targa, gomme, misura, battistrada, ubicazione):
    img = Image.new('RGB', (450, 220), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([5, 5, 445, 215], outline=(30, 58, 138), width=3)
    draw.rectangle([5, 5, 445, 40], fill=(30, 58, 138))
    draw.text((15, 12), "PNEUMATICEVOLUTION - HOTEL GOMME", fill=(255, 255, 255))
    draw.text((20, 50), f"TAG ID: {tag_id}", fill=(0, 0, 0))
    draw.text((20, 72), f"CLIENTE: {cliente}", fill=(0, 0, 0))
    draw.text((20, 94), f"TARGA: {targa}", fill=(0, 0, 0))
    draw.text((20, 116), f"PNEUMATICI: {gomme} ({misura})", fill=(0, 0, 0))
    draw.text((20, 138), f"BATTISTRADA: {battistrada}", fill=(0, 0, 0))
    draw.text((20, 160), f"UBICAZIONE: {ubicazione}", fill=(180, 0, 0))
    draw.rectangle([340, 50, 430, 140], fill=(240, 240, 240), outline=(0, 0, 0))
    draw.rectangle([350, 60, 380, 90], fill=(0, 0, 0))
    draw.rectangle([390, 60, 420, 90], fill=(0, 0, 0))
    draw.rectangle([350, 100, 380, 130], fill=(0, 0, 0))
    draw.rectangle([395, 105, 415, 125], fill=(0, 0, 0))
    draw.text((345, 145), f"QR: {tag_id}", fill=(0, 0, 0))
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
            if input_pass in ["admin", "admin123", "pneumatici"]:
                st.session_state['authenticated'] = True
                st.rerun()
            else:
                st.sidebar.error("Password errata! (Password predefinita: admin)")
    else:
        st.sidebar.success("✅ Autenticato come Titolare / Admin")
        if st.sidebar.button("Disconnetti (Logout)"):
            st.session_state['authenticated'] = False
            st.rerun()

if mode == "🔓 Portale Prenotazione Clienti (Pubblico)":
    st.markdown("# 🛞 PneumaticEvolution - Prenotazione Online")
    st.markdown("### Prenota il tuo appuntamento in officina in pochi secondi senza attese!")
    st.markdown("---")
    
    with st.form("public_booking_form"):
        st.subheader("📋 Inserisci i dati della prenotazione")
        col1, col2 = st.columns(2)
        with col1:
            nome_cliente = st.text_input("Nome e Cognome / Ragione Sociale *")
            telefono_cliente = st.text_input("Cellulare (WhatsApp) *", placeholder="3331234567")
            targa_veicolo = st.text_input("Targa Veicolo *", placeholder="AB123CD").upper()
        with col2:
            servizio = st.selectbox("Servizio Richiesto *", [
                "Cambio Stagionale (Montaggio da Hotel Gomme)",
                "Sostituzione 4 Pneumatici Nuovi",
                "Equilibratura & Convergenza Computerizzata",
                "Inversione Pneumatici & Controllo Pressione",
                "Riparazione Foratura & Vulcanizzazione"
            ])
            data_app = st.date_input("Data desiderata *", min_value=datetime.date.today())
            ora_app = st.selectbox("Orario desiderato *", [
                "08:30", "09:15", "10:00", "10:45", "11:30", "14:30", "15:15", "16:00", "16:45", "17:30"
            ])
        note_cliente = st.text_area("Note aggiuntive (es. marca gomme preferita, dettagli vettura)")
        submitted = st.form_submit_button("✅ Conferma e Invia Prenotazione", use_container_width=True)
        
        if submitted:
            if not nome_cliente or not telefono_cliente or not targa_veicolo:
                st.error("Per favore compila tutti i campi obbligatori (*).")
            else:
                df_app = load_data("Appuntamenti")
                new_id = len(df_app) + 1 if not df_app.empty else 1
                new_row = pd.DataFrame([{
                    "ID_App": new_id, "Data": str(data_app), "Ora": str(ora_app),
                    "Cliente": nome_cliente, "Targa": targa_veicolo, "Servizio": servizio,
                    "Telefono": telefono_cliente, "Note": note_cliente, "Stato": "Confermato Online"
                }])
                df_app = pd.concat([df_app, new_row], ignore_index=True)
                save_sheet(df_app, "Appuntamenti")
                st.balloons()
                st.success(f"🎉 Prenotazione confermata per {data_app} alle ore {ora_app}!")
                msg_text = f"Ciao! Ho inviato la prenotazione per la mia vettura {targa_veicolo} per '{servizio}' il {data_app} alle {ora_app}. Grazie!"
                wa_url = create_whatsapp_url(telefono_cliente, msg_text)
                st.markdown(f"💬 [Clicca qui per notificare l Officina via WhatsApp]({wa_url})")

else:
    if not st.session_state['authenticated']:
        st.warning("🔒 Sezione riservata all Officina. Inserisci la password nel menu a sinistra per accedere.")
    else:
        st.markdown("# 🛞 PneumaticEvolution Gestionale")
        st.markdown("### Gestione Anagrafica, Hotel Pneumatici, Flotta Camion, Agenda e Notifiche Massive WhatsApp")
        
        tabs = st.tabs([
            "📊 Dashboard", "👤 Anagrafica Clienti", "🏨 Hotel Pneumatici & QR",
            "📱 Notifiche Massive WhatsApp", "📅 Agenda Appuntamenti", "🚛 Flotta Camion", "📦 Magazzino"
        ])
        
        with tabs[0]:
            st.subheader("📈 Panoramica Operativa Officina")
            df_clienti = load_data("Clienti")
            df_hotel = load_data("Hotel_Pneumatici")
            df_app = load_data("Appuntamenti")
            df_mag = load_data("Magazzino")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Clienti Registrati", len(df_clienti) if not df_clienti.empty else 0)
            col2.metric("Treni Gomme Custodia", len(df_hotel) if not df_hotel.empty else 0)
            col3.metric("Appuntamenti Agenda", len(df_app) if not df_app.empty else 0)
            col4.metric("Articoli Magazzino", len(df_mag) if not df_mag.empty else 0)
            st.markdown("---")
            st.subheader("📅 Prossimi Appuntamenti in Agenda")
            if not df_app.empty:
                st.dataframe(df_app, use_container_width=True)
            else:
                st.info("Nessun appuntamento in agenda al momento.")
                
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
                        telefono = st.text_input("Telefono / WhatsApp *", placeholder="+393331234567")
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
            st.subheader("🏨 Hotel Pneumatici - Deposito & QR")
            h1, h2 = st.tabs(["📥 Check-in Deposito", "📦 Elenco Depositi in Custodia"])
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
                    btn_h = st.form_submit_button("🏷️ Salva Deposito e Genera Etichetta QR", use_container_width=True)
                    if btn_h:
                        df_h = load_data("Hotel_Pneumatici")
                        hid = len(df_h) + 1 if not df_h.empty else 1
                        tag = f"TAG-2026-{hid:03d}"
                        new_h = pd.DataFrame([{"ID_Deposito": hid, "Tag_ID": tag, "ID_Cliente": 1, "Targa": targa_dep, "Marca_Modello_Gomme": marca_dep, "Misura": misura_dep, "Stagionalita": stag_dep, "Quantita": qta_dep, "Battistrada_mm": batt_dep, "Ubicazione": ubi_dep, "Data_PresaInCarico": str(datetime.date.today()), "Stato_Sollecito": "Da Sollecitare"}])
                        save_sheet(pd.concat([df_h, new_h], ignore_index=True), "Hotel_Pneumatici")
                        st.success(f"✅ Deposito registrato con Tag ID: {tag}")
                        lbl = generate_tag_label(tag, cli_sel, targa_dep, marca_dep, misura_dep, batt_dep, ubi_dep)
                        st.image(lbl, caption=f"Etichetta {tag}")
                        st.download_button("📥 Scarica PNG Etichetta da Stampare", data=lbl, file_name=f"Etichetta_{tag}.png", mime="image/png")
            with h2:
                df_h = load_data("Hotel_Pneumatici")
                if not df_h.empty:
                    st.dataframe(df_h, use_container_width=True)
                else:
                    st.info("Nessun deposito in custodia al momento.")

        with tabs[3]:
            st.subheader("📱 Invio Massivo WhatsApp & SMS (Cambi Stagionali)")
            st.write("Filtra automaticamente i clienti ed esclude le gomme **4 Stagioni** dai solleciti obbligatori.")
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
                bulk_list = []
                for _, r in df_target.iterrows():
                    phone = "+393331234567"
                    targa_val = r.get("Targa", "N/D")
                    tag_val = r.get("Tag_ID", "N/D")
                    gomme_val = r.get("Marca_Modello_Gomme", "Pneumatici")
                    msg = f"Gentile Cliente, le ricordiamo che le sue gomme {gomme_val} per la vettura {targa_val} (Tag: {tag_val}) sono pronte in custodia presso PneumaticEvolution. Puo prenotare l appuntamento sul nostro portale. Cordiali Saluti."
                    wa_url = create_whatsapp_url(phone, msg)
                    bulk_list.append(f"Targa: {targa_val} ({phone})\nMessaggio: {msg}\nLink WhatsApp: {wa_url}\n")
                    col_m1, col_m2 = st.columns([3, 1])
                    col_m1.write(f"🚗 **Targa:** {targa_val} | Tag: `{tag_val}` | {gomme_val}")
                    col_m2.markdown(f"[💬 Invia WhatsApp]({wa_url})", unsafe_allow_html=True)
                st.markdown("---")
                st.subheader("📋 Esporta Elenco Completo Messaggi per Invio Collettivo Software")
                st.text_area("Elenco Messaggi Formattati:", value="\n".join(bulk_list), height=200)

        with tabs[4]:
            st.subheader("📅 Agenda Appuntamenti Officina")
            df_a = load_data("Appuntamenti")
            if not df_a.empty:
                st.dataframe(df_a, use_container_width=True)
            else:
                st.info("Agenda vuota.")

        with tabs[5]:
            st.subheader("🚛 Flotta Camion & Mappatura Assi")
            df_flotta = load_data("Flotta_Camion")
            if not df_flotta.empty:
                st.dataframe(df_flotta, use_container_width=True)
            else:
                st.info("Nessun camion registrato.")

        with tabs[6]:
            st.subheader("📦 Magazzino Scorte Pneumatici")
            df_m = load_data("Magazzino")
            if not df_m.empty:
                st.dataframe(df_m, use_container_width=True)
            else:
                st.info("Magazzino vuoto.")
