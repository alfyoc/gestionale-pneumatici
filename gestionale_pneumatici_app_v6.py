import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import io
from PIL import Image, ImageDraw

st.set_page_config(
    page_title="TyreManager Pro v6 - Gestionale Pneumatici & WhatsApp Chat",
    page_icon="🛞",
    layout="wide"
)

def generate_qr_label(tag_id, cliente, targa, marca_misura, ubicazione, mm_usura):
    img = Image.new('RGB', (450, 240), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(5, 5), (445, 235)], outline=(30, 58, 138), width=3)
    draw.rectangle([(5, 5), (445, 40)], fill=(30, 58, 138))
    draw.text((15, 12), f"HOTEL PNEUMATICI - TAG: {tag_id}", fill=(255, 255, 255))
    draw.text((20, 50), f"CLIENTE: {cliente[:30]}", fill=(0, 0, 0))
    draw.text((20, 75), f"TARGA: {targa}", fill=(0, 0, 0))
    draw.text((20, 100), f"GOMME: {marca_misura[:30]}", fill=(0, 0, 0))
    draw.text((20, 125), f"UBICAZIONE: {ubicazione}", fill=(0, 0, 0))
    draw.text((20, 150), f"USURA MEDIA: {mm_usura} mm", fill=(0, 0, 0))
    draw.rectangle([(300, 55), (435, 190)], outline=(0, 0, 0), width=2)
    for i in range(310, 430, 15):
        for j in range(65, 180, 15):
            if (i + j) % 3 == 0:
                draw.rectangle([(i, j), (i+10, j+10)], fill=(0, 0, 0))
    draw.text((310, 195), f"SCAN ID: {tag_id}", fill=(30, 58, 138))
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()

def get_whatsapp_url(phone_number, text=""):
    clean_phone = "".join(filter(str.isdigit, str(phone_number)))
    encoded_text = urllib.parse.quote(text)
    wa_prefix = "https://wa" + ".me/"
    if encoded_text:
        return f"{wa_prefix}{clean_phone}?text={encoded_text}"
    return f"{wa_prefix}{clean_phone}"

if 'clienti' not in st.session_state:
    st.session_state.clienti = [
        {"ID_Cliente": "CLI-001", "Nome": "Mario Rossi", "Telefono": "393331112233", "Email": "mario.rossi@email.com", "Note": "Gomme invernali in custodia"},
        {"ID_Cliente": "CLI-002", "Nome": "Autotrasporti Bianchi SRL", "Telefono": "393404445566", "Email": "info@bianchitrasporti.it", "Note": "Flotta camion - Gomme 4 Stagioni"},
        {"ID_Cliente": "CLI-003", "Nome": "Giuseppe Verdi", "Telefono": "393287778899", "Email": "g.verdi@gmail.com", "Note": "Pneumatici 4 Stagioni (Nessun cambio stagionale necessario)"}
    ]

if 'veicoli' not in st.session_state:
    st.session_state.veicoli = [
        {"Targa": "AB123CD", "ID_Cliente": "CLI-001", "Modello": "Volkswagen Golf", "Tipo": "Auto", "Tipo_Pneumatici": "Estivi Montati / Invernali in Custodia"},
        {"Targa": "EF456GH", "ID_Cliente": "CLI-002", "Modello": "Iveco Stralis 500", "Tipo": "Camion", "Tipo_Pneumatici": "4 Stagioni"},
        {"Targa": "JK789LM", "ID_Cliente": "CLI-003", "Modello": "Fiat Panda", "Tipo": "Auto", "Tipo_Pneumatici": "4 Stagioni"}
    ]

if 'hotel' not in st.session_state:
    st.session_state.hotel = [
        {
            "ID_Tag": "TAG-HOTEL-101",
            "Cliente": "Mario Rossi",
            "Telefono": "393331112233",
            "Targa": "AB123CD",
            "Marca_Gomme": "Pirelli Winter Sottozero",
            "Misura": "205/55 R16",
            "Stagionalita": "Invernale",
            "Richiede_Cambio_Stagionale": "SI",
            "Quantita": 4,
            "Mm_Ant_SX": 6.5, "Mm_Ant_DX": 6.5, "Mm_Post_SX": 5.8, "Mm_Post_DX": 5.8,
            "Ubicazione": "Corsia A - Scaffale 03 - Ripiano 2",
            "Stato_Sollecito": "Non Inviato",
            "Data_Ingresso": "2026-04-15"
        },
        {
            "ID_Tag": "TAG-HOTEL-102",
            "Cliente": "Giuseppe Verdi",
            "Telefono": "393287778899",
            "Targa": "JK789LM",
            "Marca_Gomme": "Michelin CrossClimate 2",
            "Misura": "175/65 R14",
            "Stagionalita": "4 Stagioni",
            "Richiede_Cambio_Stagionale": "NO",
            "Quantita": 4,
            "Mm_Ant_SX": 7.0, "Mm_Ant_DX": 7.0, "Mm_Post_SX": 6.8, "Mm_Post_DX": 6.8,
            "Ubicazione": "Corsia B - Scaffale 01 - Ripiano 1",
            "Stato_Sollecito": "Non Necessario (4 Stagioni)",
            "Data_Ingresso": "2026-06-10"
        }
    ]

if 'prenotazioni' not in st.session_state:
    st.session_state.prenotazioni = [
        {
            "ID_Prenotazione": "PRN-501",
            "Data": str(datetime.date.today() + datetime.timedelta(days=2)),
            "Ora_Inizio": "09:00",
            "Ora_Fine": "09:45",
            "Servizio": "Cambio Stagionale Estivo/Invernale",
            "Nome_Cliente": "Mario Rossi",
            "Telefono": "393331112233",
            "Targa": "AB123CD",
            "Stato": "Confermato",
            "Stato_WhatsApp": "Confermato dal Cliente"
        }
    ]

if 'blocchi_agenda' not in st.session_state:
    st.session_state.blocchi_agenda = []

if 'chat_whatsapp' not in st.session_state:
    st.session_state.chat_whatsapp = [
        {
            "ID_Chat": "CHAT-001",
            "Data_Ora": "2026-10-31 09:15",
            "Cliente": "Mario Rossi",
            "Telefono": "393331112233",
            "Ultimo_Messaggio": "Ciao Mario, ti ricordiamo il cambio pneumatici per il 02/11 ore 09:00.",
            "Stato": "Confermato dal Cliente",
            "Note_Risposta": "Cliente ha confermato via WhatsApp per le 09:00."
        },
        {
            "ID_Chat": "CHAT-002",
            "Data_Ora": "2026-10-31 09:20",
            "Cliente": "Giuseppe Verdi",
            "Telefono": "393287778899",
            "Ultimo_Messaggio": "Ciao Giuseppe, avendo pneumatici 4 Stagioni ti proponiamo un controllo usura/inversione facoltativo.",
            "Stato": "Nuova Risposta / In Gestione",
            "Note_Risposta": "Cliente chiede se l'inversione si può fare di lunedì."
        }
    ]

st.sidebar.title("TyreManager Pro v6")
st.sidebar.markdown("**Gestione Pneumatici & WhatsApp Chat**")

menu = st.sidebar.radio(
    "Navigazione Moduli:",
    [
        "📊 Dashboard Operativa",
        "🏨 Hotel Pneumatici & QR Tag",
        "💬 WhatsApp Chat & Risposte",
        "📅 Agenda Prenotazioni & Blocchi",
        "🍁 Solleciti Stagionali & 4 Stagioni",
        "🚛 Flotta Aziendale Camion",
        "📦 Magazzino Pneumatici",
        "📘 Tabelle Tecniche ETRTO"
    ]
)

if menu == "📊 Dashboard Operativa":
    st.header("📊 Dashboard Operativa Officina")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Totale Clienti", len(st.session_state.clienti))
    col2.metric("Treni in Hotel Pneumatici", len(st.session_state.hotel))
    
    quattro_stagioni_count = sum(1 for h in st.session_state.hotel if h.get("Stagionalita") == "4 Stagioni" or h.get("Richiede_Cambio_Stagionale") == "NO")
    col3.metric("Pneumatici 4 Stagioni", quattro_stagioni_count)
    col4.metric("Prenotazioni in Agenda", len(st.session_state.prenotazioni))
    
    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("💬 Ultimi Messaggi & Risposte WhatsApp")
        df_chat = pd.DataFrame(st.session_state.chat_whatsapp)
        if not df_chat.empty:
            st.dataframe(df_chat[["Data_Ora", "Cliente", "Stato", "Note_Risposta"]], use_container_width=True)
        else:
            st.info("Nessuna conversazione WhatsApp registrata.")
            
    with col_right:
        st.subheader("🏨 Ultimi Depositi in Hotel Pneumatici")
        df_h = pd.DataFrame(st.session_state.hotel)
        if not df_h.empty:
            st.dataframe(df_h[["ID_Tag", "Cliente", "Targa", "Stagionalita", "Richiede_Cambio_Stagionale", "Ubicazione"]], use_container_width=True)
        else:
            st.info("Nessun deposito in custodia registrato.")

elif menu == "🏨 Hotel Pneumatici & QR Tag":
    st.header("🏨 Hotel Pneumatici & Etichettatura QR")
    
    tab1, tab2, tab3 = st.tabs(["➕ Nuova Presa in Carico", "📋 Elenco Gomme in Custodia", "🔍 Scansione / Ristampa Tag QR"])
    
    with tab1:
        st.subheader("Presa in Carico e Deposito Treno Gomme")
        
        with st.form("form_hotel_in"):
            c1, c2 = st.columns(2)
            
            nomi_clienti = [c["Nome"] for c in st.session_state.clienti]
            cliente_sel = c1.selectbox("Seleziona Cliente:", nomi_clienti)
            
            client_obj = next((c for c in st.session_state.clienti if c["Nome"] == cliente_sel), None)
            telefono_cli = client_obj["Telefono"] if client_obj else ""
            
            targhe_cli = [v["Targa"] for v in st.session_state.veicoli if v["ID_Cliente"] == (client_obj["ID_Cliente"] if client_obj else "")]
            if not targhe_cli:
                targa_sel = c2.text_input("Inserisci Nuova Targa Veicolo:", "AB123CD")
            else:
                targa_sel = c2.selectbox("Seleziona Targa Veicolo:", targhe_cli)
                
            col_g1, col_g2, col_g3 = st.columns(3)
            marca_gomme = col_g1.text_input("Marca e Modello Pneumatico:", "Michelin Primacy 4")
            misura_gomme = col_g2.text_input("Misura Complete:", "205/55 R16 91V")
            
            stagionalita = col_g3.selectbox(
                "Stagionalità Pneumatico:",
                ["Invernale", "Estivo", "4 Stagioni (All-Season)"]
            )
            
            if stagionalita == "4 Stagioni (All-Season)":
                richiede_cambio = "NO"
                st.info("ℹ️ **Nota 4 Stagioni:** Per i pneumatici 4 Stagioni il cambio stagionale non è obbligatorio. Il veicolo verrà escluso dai solleciti automatici di cambio di novembre/aprile, ma potrà ricevere suggerimenti per l'inversione periodica del battistrada.")
            else:
                richiede_cambio = "SI"
            
            st.subheader("📏 Misurazione Millimetrica Battistrada Residuo (mm)")
            col_mm1, col_mm2, col_mm3, col_mm4 = st.columns(4)
            mm_ant_sx = col_mm1.number_input("Anteriore SX (mm):", min_value=1.6, max_value=10.0, value=6.5, step=0.1)
            mm_ant_dx = col_mm2.number_input("Anteriore DX (mm):", min_value=1.6, max_value=10.0, value=6.5, step=0.1)
            mm_post_sx = col_mm3.number_input("Posteriore SX (mm):", min_value=1.6, max_value=10.0, value=5.8, step=0.1)
            mm_post_dx = col_mm4.number_input("Posteriore DX (mm):", min_value=1.6, max_value=10.0, value=5.8, step=0.1)
            
            st.subheader("📍 Ubicazione Stoccaggio Magazzino")
            col_u1, col_u2, col_u3 = st.columns(3)
            corsia = col_u1.selectbox("Corsia:", ["Corsia A", "Corsia B", "Corsia C", "Corsia D"])
            scaffale = col_u2.text_input("Numero Scaffale:", "03")
            ripiano = col_u3.text_input("Numero Ripiano / Livello:", "2")
            
            btn_salva_deposito = st.form_submit_button("💾 Salva Deposito & Genera QR Tag")
            
            if btn_salva_deposito:
                new_id_num = len(st.session_state.hotel) + 101
                new_tag = f"TAG-HOTEL-{new_id_num}"
                ubicazione_full = f"{corsia} - Scaffale {scaffale} - Ripiano {ripiano}"
                
                nuovo_deposito = {
                    "ID_Tag": new_tag,
                    "Cliente": cliente_sel,
                    "Telefono": telefono_cli,
                    "Targa": targa_sel,
                    "Marca_Gomme": marca_gomme,
                    "Misura": misura_gomme,
                    "Stagionalita": stagionalita,
                    "Richiede_Cambio_Stagionale": richiede_cambio,
                    "Quantita": 4,
                    "Mm_Ant_SX": mm_ant_sx, "Mm_Ant_DX": mm_ant_dx,
                    "Mm_Post_SX": mm_post_sx, "Mm_Post_DX": mm_post_dx,
                    "Ubicazione": ubicazione_full,
                    "Stato_Sollecito": "Non Inviato" if richiede_cambio == "SI" else "Non Necessario (4 Stagioni)",
                    "Data_Ingresso": str(datetime.date.today())
                }
                
                st.session_state.hotel.append(nuovo_deposito)
                st.success(f"✅ Deposito registrato con successo! Tag Assegnato: **{new_tag}**")
                
                mm_avg = round((mm_ant_sx + mm_ant_dx + mm_post_sx + mm_post_dx) / 4, 1)
                img_bytes = generate_qr_label(new_tag, cliente_sel, targa_sel, f"{marca_gomme} {misura_gomme}", ubicazione_full, mm_avg)
                
                st.image(img_bytes, caption=f"Etichetta QR Generata per {new_tag}", width=400)
                st.download_button(
                    label="📥 Scarica Etichetta da Stampare (PNG)",
                    data=img_bytes,
                    file_name=f"Etichetta_{new_tag}_{targa_sel}.png",
                    mime="image/png"
                )

    with tab2:
        st.subheader("Elenco Pneumatici in Custodia nel Magazzino")
        df_h = pd.DataFrame(st.session_state.hotel)
        if not df_h.empty:
            filtro_stagione = st.multiselect("Filtra per Stagionalità:", ["Tutti", "Invernale", "Estivo", "4 Stagioni"], default=["Tutti"])
            if "Tutti" not in filtro_stagione:
                df_h = df_h[df_h["Stagionalita"].isin(filtro_stagione)]
            st.dataframe(df_h, use_container_width=True)
        else:
            st.info("Nessun dato presente.")

    with tab3:
        st.subheader("Scansione ID Tag & Ristampa Etichetta")
        tag_search = st.text_input("Inserisci o scansiona il Codice Tag ID (es. TAG-HOTEL-101):", "")
        if tag_search:
            found = next((item for item in st.session_state.hotel if item["ID_Tag"] == tag_search), None)
            if found:
                st.success(f"🎯 Deposito Trovato per il Tag **{tag_search}**!")
                st.write(f"**Cliente:** {found['Cliente']} | **Targa:** {found['Targa']}")
                st.write(f"**Gomme:** {found['Marca_Gomme']} ({found['Misura']}) - **Stagionalità:** {found['Stagionalita']}")
                st.write(f"**Ubicazione Magazzino:** {found['Ubicazione']}")
                
                mm_avg = round((found['Mm_Ant_SX'] + found['Mm_Ant_DX'] + found['Mm_Post_SX'] + found['Mm_Post_DX']) / 4, 1)
                img_bytes = generate_qr_label(found['ID_Tag'], found['Cliente'], found['Targa'], f"{found['Marca_Gomme']} {found['Misura']}", found['Ubicazione'], mm_avg)
                st.image(img_bytes, width=350)
                st.download_button(
                    label="📥 Ristampa Etichetta (PNG)",
                    data=img_bytes,
                    file_name=f"Etichetta_{found['ID_Tag']}.png",
                    mime="image/png"
                )
            else:
                st.error("Nessun deposito trovato con questo codice Tag ID.")

elif menu == "💬 WhatsApp Chat & Risposte":
    st.header("💬 WhatsApp Chat & Gestione Risposte Clienti")
    
    col_c1, col_c2 = st.columns([1, 1])
    
    with col_c1:
        st.subheader("📤 Avvia Chat Diretta / Invia Messaggio")
        
        with st.form("form_wa_direct"):
            cliente_sel_chat = st.selectbox("Seleziona Cliente:", [c["Nome"] for c in st.session_state.clienti])
            client_info = next((c for c in st.session_state.clienti if c["Nome"] == cliente_sel_chat), None)
            num_telefono = st.text_input("Numero di Cellulare (con prefisso):", client_info["Telefono"] if client_info else "393330000000")
            
            tipo_messaggio = st.selectbox(
                "Scegli Modello Messaggio:",
                [
                    "Chat Libera (Nessun testo precompilato)",
                    "Conferma Prenotazione Appuntamento",
                    "Promemoria Cambio Stagionale (Hotel Pneumatici)",
                    "Suggerimento Controllo / Inversione 4 Stagioni",
                    "Veicolo Pronto per il Ritiro"
                ]
            )
            
            if tipo_messaggio == "Chat Libera (Nessun testo precompilato)":
                testo_personalizzato = ""
            elif tipo_messaggio == "Conferma Prenotazione Appuntamento":
                testo_personalizzato = f"Gentile {cliente_sel_chat}, Le confermiamo l'appuntamento in officina per il cambio pneumatici. A presto!"
            elif tipo_messaggio == "Promemoria Cambio Stagionale (Hotel Pneumatici)":
                testo_personalizzato = f"Gentile {cliente_sel_chat}, Le ricordiamo il cambio stagionale. Le Sue gomme in custodia sono pronte! Ci risponda a questo messaggio o ci chiami per l'orario."
            elif tipo_messaggio == "Suggerimento Controllo / Inversione 4 Stagioni":
                testo_personalizzato = f"Gentile {cliente_sel_chat}, Le ricordiamo un controllo periodico ed inversione per le sue gomme 4 Stagioni per garantirne il massimo rendimento."
            else:
                testo_personalizzato = f"Gentile {cliente_sel_chat}, Le comunichiamo che il Suo veicolo è pronto per il ritiro in officina."
                
            testo_messaggio_final = st.text_area("Messaggio WhatsApp:", testo_personalizzato, height=120)
            
            btn_wa_send = st.form_submit_button("🟢 Apri WhatsApp Chat")
            
            if btn_wa_send:
                wa_link = get_whatsapp_url(num_telefono, testo_messaggio_final)
                
                st.session_state.chat_whatsapp.append({
                    "ID_Chat": f"CHAT-{len(st.session_state.chat_whatsapp)+101}",
                    "Data_Ora": str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
                    "Cliente": cliente_sel_chat,
                    "Telefono": num_telefono,
                    "Ultimo_Messaggio": testo_messaggio_final if testo_messaggio_final else "(Chat Libera Avviata)",
                    "Stato": "Inviato WA",
                    "Note_Risposta": "In attesa di risposta dal cliente"
                })
                
                st.markdown(f'[📲 Apri WhatsApp Chat con {cliente_sel_chat}]({wa_link})')
                st.success("Conversazione registrata nel log delle chat!")

    with col_c2:
        st.subheader("📥 Registro Risposte & Gestione Conversazioni")
        
        if st.session_state.chat_whatsapp:
            for idx, chat in enumerate(reversed(st.session_state.chat_whatsapp)):
                with st.expander(f"💬 {chat['Cliente']} ({chat['Telefono']}) - Stato: {chat['Stato']}", expanded=(idx==0)):
                    st.write(f"**Data/Ora:** {chat['Data_Ora']}")
                    st.write(f"**Ultimo Messaggio Inviato:** {chat['Ultimo_Messaggio']}")
                    st.write(f"**Note / Risposta Cliente Attuale:** {chat['Note_Risposta']}")
                    
                    with st.form(f"form_reply_{idx}"):
                        nuovo_stato = st.selectbox(
                            "Aggiorna Stato Conversazione:",
                            ["Inviato WA", "Confermato dal Cliente", "Nuova Risposta / In Gestione", "Cliente ha Gomme 4 Stagioni (Nessun Cambio)", "Conclusa"],
                            index=1
                        )
                        nota_risposta_in = st.text_input("Inserisci Testo Risposta Cliente o Nota:", chat['Note_Risposta'])
                        btn_update_chat = st.form_submit_button("💾 Salva Risposta")
                        
                        if btn_update_chat:
                            chat['Stato'] = nuovo_stato
                            chat['Note_Risposta'] = nota_risposta_in
                            st.success("Stato conversazione aggiornato!")
                            st.rerun()
                            
                    direct_wa = get_whatsapp_url(chat['Telefono'])
                    st.markdown(f'[📲 Rispondi a {chat["Cliente"]} su WhatsApp]({direct_wa})')
        else:
            st.info("Nessuna conversazione in storico.")

elif menu == "📅 Agenda Prenotazioni & Blocchi":
    st.header("📅 Agenda Officina & Calcolo Disponibilità Automatica")
    
    tab_p1, tab_p2, tab_p3 = st.tabs(["🌐 Portale Prenotazione Cliente", "🛠️ Calendario Officina", "🛑 Blocchi Orari Manuali (Fermo Ponte/Flotta)"])
    
    with tab_p1:
        st.subheader("Prenotazione Appuntamento Sostituzione & Manutenzione Pneumatici")
        
        col_pr1, col_pr2 = st.columns(2)
        data_pren = col_pr1.date_input("Seleziona Data Appuntamento:", datetime.date.today() + datetime.timedelta(days=1))
        
        servizio_sel = col_pr2.selectbox(
            "Seleziona Tipo di Servizio:",
            [
                "Cambio Stagionale Estivo/Invernale (45 min)",
                "Controllo Usura & Inversione (30 min)",
                "Sostituzione 4 Pneumatici Nuovi (60 min)",
                "Equilibratura & Convergenza (45 min)",
                "Riparazione Foratura (30 min)"
            ]
        )
        
        orari_base = ["08:30", "09:15", "10:00", "10:45", "11:30", "14:30", "15:15", "16:00", "16:45", "17:30"]
        
        prenotati_giorno = [p["Ora_Inizio"] for p in st.session_state.prenotazioni if p["Data"] == str(data_pren)]
        blocchi_giorno = [b["Ora_Inizio"] for b in st.session_state.blocchi_agenda if b["Data"] == str(data_pren)]
        
        orari_disponibili = [o for o in orari_base if o not in prenotati_giorno and o not in blocchi_giorno]
        
        st.subheader("⏱️ Orari Disponibili in Tempo Reale")
        if orari_disponibili:
            ora_scelta = st.radio("Seleziona uno Slot Orario Libero:", orari_disponibili, horizontal=True)
            
            with st.form("form_prenota_online"):
                c_pr_a, c_pr_b = st.columns(2)
                nome_c = c_pr_a.text_input("Nome e Cognome Cliente:", "Mario Rossi")
                tel_c = c_pr_b.text_input("Numero Cellulare:", "393330000000")
                targa_c = c_pr_a.text_input("Targa Veicolo:", "AB123CD")
                note_c = c_pr_b.text_input("Note Aggiuntive:", "Gomme già in custodia Hotel Pneumatici")
                
                btn_conferma_pren = st.form_submit_button("✅ Conferma Prenotazione Appuntamento")
                
                if btn_conferma_pren:
                    nuova_p = {
                        "ID_Prenotazione": f"PRN-{len(st.session_state.prenotazioni)+501}",
                        "Data": str(data_pren),
                        "Ora_Inizio": ora_scelta,
                        "Ora_Fine": "00:00",
                        "Servizio": servizio_sel,
                        "Nome_Cliente": nome_c,
                        "Telefono": tel_c,
                        "Targa": targa_c,
                        "Stato": "Confermato",
                        "Stato_WhatsApp": "Non Inviato"
                    }
                    st.session_state.prenotazioni.append(nuova_p)
                    st.success(f"🎉 Appuntamento confermato per il **{data_pren}** alle **{ora_scelta}**!")
                    
                    wa_text = f"Gentile {nome_c}, confermiamo l'appuntamento per {servizio_sel} il {data_pren} alle ore {ora_scelta} (Targa: {targa_c}). Officina Pneumatici."
                    wa_url = get_whatsapp_url(tel_c, wa_text)
                    st.markdown(f'[📲 Invia Conferma via WhatsApp a {nome_c}]({wa_url})')
        else:
            st.warning("⚠️ Tutti gli slot orari per questa data sono occupati o bloccati dal titolare. Seleziona un'altra data.")

    with tab_p2:
        st.subheader("Vista Calendario Officina")
        df_pr = pd.DataFrame(st.session_state.prenotazioni)
        if not df_pr.empty:
            st.dataframe(df_pr, use_container_width=True)
        else:
            st.info("Nessun dato presente.")

    with tab_p3:
        st.subheader("Inserimento Blocchi Orari Manuali (Chiusure, Manutenzione Flotta Camion, Ferri)")
        with st.form("form_blocco_orario"):
            b_data = st.date_input("Data Blocco:", datetime.date.today())
            b_ora = st.selectbox("Ora da Bloccare:", orari_base)
            b_causale = st.text_input("Motivo del Blocco:", "Manutenzione Camion Flotta / Chiusura straordinaria")
            
            btn_add_block = st.form_submit_button("🛑 Salva Blocco Orario")
            if btn_add_block:
                st.session_state.blocchi_agenda.append({
                    "Data": str(b_data),
                    "Ora_Inizio": b_ora,
                    "Motivo": b_causale
                })
                st.success(f"Slot delle {b_ora} del giorno {b_data} bloccato con successo!")

elif menu == "🍁 Solleciti Stagionali & 4 Stagioni":
    st.header("🍁 Campagne Sollecito Cambio Stagionale")
    
    st.info("📌 **Logica 4 Stagioni:** I clienti che hanno pneumatici 4 Stagioni non devono fare il cambio stagionale. L'app li esclude dai solleciti obbligatori e propone invece un controllo facoltativo di usura/inversione.")
    
    campagna = st.radio(
        "Seleziona la Campagna da Gestire:",
        ["🍁 Cambio Invernale (Scadenza 15 Novembre)", "☀️ Cambio Estivo (Scadenza 15 Aprile)", "🔄 Suggerimento Inversione / Controllo (Pneumatici 4 Stagioni)"],
        horizontal=True
    )
    
    st.markdown("---")
    
    if campagna == "🍁 Cambio Invernale (Scadenza 15 Novembre)":
        st.subheader("📋 Clienti con Gomme Invernali pronte per il Montaggio")
        list_inv = [h for h in st.session_state.hotel if h.get("Stagionalita") == "Invernale" and h.get("Richiede_Cambio_Stagionale") == "SI"]
        
        if list_inv:
            for item in list_inv:
                col_a, col_b = st.columns([3, 1])
                with col_a:
                    st.write(f"**Cliente:** {item['Cliente']} | **Targa:** {item['Targa']} | **Gomme:** {item['Marca_Gomme']} ({item['Misura']})")
                    st.write(f"📍 Ubicazione: *{item['Ubicazione']}* | Tag ID: `{item['ID_Tag']}`")
                with col_b:
                    msg = f"Gentile {item['Cliente']}, Le ricordiamo l'obbligo pneumatici invernali del 15 Novembre. Le Sue gomme in custodia ({item['Marca_Gomme']}) sono pronte! Ci risponda a questo messaggio o chiami per l'appuntamento."
                    url_w = get_whatsapp_url(item['Telefono'], msg)
                    st.markdown(f'[🟢 Sollecita via WhatsApp]({url_w})')
        else:
            st.info("Nessun cliente in elenco per la campagna invernale.")

    elif campagna == "☀️ Cambio Estivo (Scadenza 15 Aprile)":
        st.subheader("📋 Clienti con Gomme Estive pronte per il Montaggio")
        list_est = [h for h in st.session_state.hotel if h.get("Stagionalita") == "Estivo" and h.get("Richiede_Cambio_Stagionale") == "SI"]
        
        if list_est:
            for item in list_est:
                col_a, col_b = st.columns([3, 1])
                with col_a:
                    st.write(f"**Cliente:** {item['Cliente']} | **Targa:** {item['Targa']} | **Gomme:** {item['Marca_Gomme']} ({item['Misura']})")
                    st.write(f"📍 Ubicazione: *{item['Ubicazione']}* | Tag ID: `{item['ID_Tag']}`")
                with col_b:
                    msg = f"Gentile {item['Cliente']}, Le ricordiamo il cambio stagionale estivo del 15 Aprile. Le Sue gomme ({item['Marca_Gomme']}) sono in deposito. Ci risponda per prenotare l'orario!"
                    url_w = get_whatsapp_url(item['Telefono'], msg)
                    st.markdown(f'[🟢 Sollecita via WhatsApp]({url_w})')
        else:
            st.info("Nessun cliente in elenco per la campagna estiva.")

    else:
        st.subheader("🔄 Clienti con Pneumatici 4 Stagioni (All-Season)")
        list_4s = [h for h in st.session_state.hotel if h.get("Stagionalita") == "4 Stagioni" or h.get("Richiede_Cambio_Stagionale") == "NO"]
        
        if list_4s:
            for item in list_4s:
                col_a, col_b = st.columns([3, 1])
                with col_a:
                    st.write(f"**Cliente:** {item['Cliente']} | **Targa:** {item['Targa']} | **Gomme:** {item['Marca_Gomme']} (4 Stagioni)")
                    st.write(f"ℹ️ Status: *Nessun cambio obbligatorio richiesto.* Usura attuale: {item.get('Mm_Ant_SX', 6.0)} mm")
                with col_b:
                    msg = f"Gentile {item['Cliente']}, per le Sue pneumatici 4 Stagioni su targa {item['Targa']} consigliano un controllo ed inversione periodica per prolungarne la durata. Ci risponda per fissare un controllo rapido!"
                    url_w = get_whatsapp_url(item['Telefono'], msg)
                    st.markdown(f'[🟢 Invia Suggerimento WhatsApp]({url_w})')
        else:
            st.info("Nessun cliente con gomme 4 Stagioni registrato.")

elif menu == "🚛 Flotta Aziendale Camion":
    st.header("🚛 Gestione Flotta Aziendale Camion")
    st.dataframe(pd.DataFrame(st.session_state.veicoli)[lambda df: df["Tipo"] == "Camion"], use_container_width=True)

elif menu == "📦 Magazzino Pneumatici":
    st.header("📦 Inventario Magazzino Pneumatici")
    df_mag = pd.read_excel("/workspace/scratch/gestionale_pneumatici_db_v6.xlsx", sheet_name="Magazzino")
    st.dataframe(df_mag, use_container_width=True)

elif menu == "📘 Tabelle Tecniche ETRTO":
    st.header("📘 Tabelle Tecniche & Omologazioni ETRTO")
    df_eq = pd.DataFrame([
        {"Canale Cerchio (pollici)": "5.5 J", "Larghezza Minima (mm)": 165, "Larghezza Ideale (mm)": "175 - 185", "Larghezza Massima (mm)": 195},
        {"Canale Cerchio (pollici)": "6.0 J", "Larghezza Minima (mm)": 175, "Larghezza Ideale (mm)": "185 - 195", "Larghezza Massima (mm)": 205},
        {"Canale Cerchio (pollici)": "6.5 J", "Larghezza Minima (mm)": 185, "Larghezza Ideale (mm)": "195 - 205", "Larghezza Massima (mm)": 215},
        {"Canale Cerchio (pollici)": "7.0 J", "Larghezza Minima (mm)": 195, "Larghezza Ideale (mm)": "205 - 215", "Larghezza Massima (mm)": 225},
        {"Canale Cerchio (pollici)": "7.5 J", "Larghezza Minima (mm)": 205, "Larghezza Ideale (mm)": "215 - 225", "Larghezza Massima (mm)": 235}
    ])
    st.dataframe(df_eq, use_container_width=True)

st.markdown("---")
st.caption("TyreManager Pro v6 - Sviluppato per gestione vendita, officina, flotta camion e comunicazioni WhatsApp.")
