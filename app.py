import os
import io
import re
import math
import random
import psycopg2
import pandas as pd
import streamlit as st
import base64
from datetime import date, datetime, timedelta

# ==========================================
# 1. KONSTANTER & KONFIGURASJON
# ==========================================

# Sett til True når du er klar til å kreve e-post-kode for ALLE brukere.
# Når denne er False, brukes vanlige fellespassord som før, men du kan teste e-post-kode på deg selv.
KREVER_EPOST_INNLOGGING = False

OVERSTYR_NIVAA_KRAV = True

EXCEL_ALIASES = {
    'kampnr': ['Kampnr', 'Kampnummer', 'Match ID'],
    'runde': ['Runde', 'Round'],
    'dato': ['Dato', 'Date'],
    'tid': ['Tid', 'Time'],
    'hjemmelag': ['Hjemmelag', 'Hjemmelag '],
    'bortelag': ['Bortelag', 'Bortelag '],
    'bane': ['Bane', 'Spillested'],
    'arrangor': ['Arrangør', 'Host'],
    'dommer_1': ['Dommer ', 'Dommer 1', 'Dommer'],
    'dommer_2': ['Dommer 2', 'Dommer2'],
    'observator': ['Observatør ', 'Observatør', 'Observer'],
    'turnering': ['Turnering', 'Klasse', 'League']
}

ADMIN_PASSORD = st.secrets["passwords"]["admin"]
LAGLEDER_PASSORD = st.secrets["passwords"]["lagleder"]
DOMMER_PASSORD = st.secrets["passwords"]["dommer"]


# ==========================================
# 2. STYLING (FREIDIG DESIGN)
# ==========================================
def inject_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Play:wght@400;700&display=swap');

        .stApp { background: linear-gradient(135deg, #fdfbfb 0%, #f4f4f4 100%); }
        h1, h2, h3, h4, h5, h6 { color: #1f1f1f !important; font-family: 'Play', sans-serif; letter-spacing: -0.5px; }

        .main-title { 
            text-align: left; 
            font-size: 3rem; 
            font-weight: 800; 
            background: -webkit-linear-gradient(45deg, #E3000F, #9b0b20); 
            -webkit-background-clip: text; 
            -webkit-text-fill-color: transparent; 
            margin-bottom: 0.5rem; 
            padding-top: 1rem; 
        }

        .title-underline {
            height: 5px;
            width: 120px;
            background: linear-gradient(90deg, #FFD700 0%, #FFB300 100%);
            margin-bottom: 2.5rem;
            border-radius: 3px;
        }

        .stButton > button[kind="primary"] { 
            background: linear-gradient(90deg, #E3000F 0%, #C2000B 100%); 
            color: white; 
            border: 2px solid #E3000F; 
            border-radius: 8px; 
            box-shadow: 0 4px 8px rgba(227, 0, 15, 0.25); 
            transition: all 0.3s ease; 
            font-family: 'Arial', sans-serif; 
            font-weight: bold; 
        }
        .stButton > button[kind="primary"]:hover { 
            transform: translateY(-2px); 
            box-shadow: 0 6px 14px rgba(227, 0, 15, 0.4); 
            color: white; 
            border: 2px solid #FFD700;
        }

        .stButton > button[kind="secondary"] { 
            border: 2px solid #1f1f1f; 
            color: #1f1f1f; 
            border-radius: 8px; 
            transition: all 0.3s ease; 
            background-color: white; 
            font-family: 'Arial', sans-serif; 
            font-weight: bold; 
        }
        .stButton > button[kind="secondary"]:hover { 
            border: 2px solid #E3000F; 
            color: #E3000F; 
            background-color: #fff9f9; 
            transform: translateY(-2px); 
            box-shadow: 0 4px 8px rgba(227, 0, 15, 0.1); 
        }

        .stDataFrame { 
            border: none !important; 
            border-radius: 10px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.06); 
            background-color: white; 
            padding: 10px; 
            border-top: 4px solid #FFD700 !important; 
        }

        .streamlit-expanderHeader { 
            font-size: 1.1rem; 
            font-family: 'Play', sans-serif; 
            font-weight: 700; 
            color: #1f1f1f; 
            border-radius: 8px; 
            background-color: white; 
            box-shadow: 0 2px 5px rgba(0,0,0,0.02); 
            margin-bottom: 5px; 
            border-left: 5px solid #E3000F; 
        }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# 3. DATABASE SETUP & HJELPEFUNKSJONER
# ==========================================
@st.cache_resource(ttl=600)
def init_connection():
    conn = psycopg2.connect(st.secrets["supabase"]["db_url"])
    conn.autocommit = True
    return conn


def get_db_connection():
    conn = init_connection()
    if conn.closed != 0:
        st.cache_resource.clear()
        conn = init_connection()
    return conn


def execute_query_to_df(query, conn, params=None):
    c = conn.cursor()
    c.execute(query, params)
    cols = [desc[0] for desc in c.description]
    return pd.DataFrame(c.fetchall(), columns=cols)


def init_db(conn):
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS kamper
                 (
                     kampnr     TEXT PRIMARY KEY,
                     runde      TEXT,
                     dato       TEXT,
                     tid        TEXT,
                     hjemmelag  TEXT,
                     bortelag   TEXT,
                     bane       TEXT,
                     arrangor   TEXT,
                     dommer_1   TEXT,
                     dommer_2   TEXT,
                     observator TEXT,
                     turnering  TEXT,
                     laast      BOOLEAN DEFAULT FALSE,
                     status     TEXT    DEFAULT ''
                 )''')
    c.execute('''CREATE TABLE IF NOT EXISTS dommer_status
                 (
                     navn        TEXT PRIMARY KEY,
                     aktiv       BOOLEAN,
                     type_dommer TEXT
                 )''')
    c.execute('''CREATE TABLE IF NOT EXISTS dommer_ansvar
                 (
                     navn          TEXT PRIMARY KEY,
                     epost         TEXT,
                     telefon       TEXT,
                     lag_liste     TEXT,
                     selvbetjening BOOLEAN DEFAULT TRUE
                 )''')
    c.execute('''CREATE TABLE IF NOT EXISTS dommer_onsker
                 (
                     kampnr      TEXT,
                     dommer_navn TEXT,
                     UNIQUE (kampnr, dommer_navn)
                 )''')
    c.execute('''CREATE TABLE IF NOT EXISTS godkjente_opprykk
                 (
                     dommer_navn TEXT PRIMARY KEY
                 )''')
    c.execute('''CREATE TABLE IF NOT EXISTS engangskoder
                 (
                     epost   TEXT PRIMARY KEY,
                     kode    TEXT,
                     utloper TIMESTAMP
                 )''')

    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS selvbetjening BOOLEAN DEFAULT TRUE")
    c.execute("ALTER TABLE kamper ADD COLUMN IF NOT EXISTS laast BOOLEAN DEFAULT FALSE")
    c.execute("ALTER TABLE kamper ADD COLUMN IF NOT EXISTS status TEXT DEFAULT ''")
    c.execute("ALTER TABLE kamper ADD COLUMN IF NOT EXISTS aktiv BOOLEAN DEFAULT TRUE")

    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS epost TEXT")
    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS siste_innlogging TIMESTAMP")
    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS antall_innlogginger INT DEFAULT 0")
    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS epost_aktivert BOOLEAN DEFAULT FALSE")

    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS siste_innlogging TIMESTAMP")
    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS antall_innlogginger INT DEFAULT 0")
    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS epost_aktivert BOOLEAN DEFAULT FALSE")

    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS epost_bekreftet BOOLEAN DEFAULT FALSE")
    c.execute("ALTER TABLE dommer_status ADD COLUMN IF NOT EXISTS motta_varsler BOOLEAN DEFAULT TRUE")

    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS epost_bekreftet BOOLEAN DEFAULT FALSE")
    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS motta_varsler BOOLEAN DEFAULT TRUE")


def load_data(conn):
    return execute_query_to_df("SELECT * FROM kamper WHERE (aktiv = TRUE OR aktiv IS NULL) ORDER BY dato ASC, tid ASC",
                               conn)


def sync_referees_to_db(db_df, conn, filepath='dommere.xlsx'):
    dommer_info = {}
    if os.path.exists(filepath):
        try:
            df_dommere = pd.read_excel(filepath)
            if 'Fornavn' in df_dommere.columns and 'Etternavn' in df_dommere.columns:
                df_dommere['Fullt Navn'] = df_dommere['Fornavn'].astype(str).str.strip() + " " + df_dommere[
                    'Etternavn'].astype(str).str.strip()
                type_col = 'Dommer' if 'Dommer' in df_dommere.columns else None
                epost_col = 'E-post' if 'E-post' in df_dommere.columns else (
                    'Epost' if 'Epost' in df_dommere.columns else None)
                for _, row in df_dommere.iterrows():
                    navn = str(row['Fullt Navn']).strip()
                    if navn and navn.lower() not in ['nan', 'none']:
                        dtype = str(row[type_col]).strip() if type_col and pd.notna(row[type_col]) else ""
                        epost = str(row[epost_col]).strip() if epost_col and pd.notna(row[epost_col]) else ""
                        dommer_info[navn] = (dtype, epost)
        except Exception:
            pass

    if not db_df.empty:
        for col in ['dommer_1', 'dommer_2', 'observator']:
            if col in db_df.columns:
                for navn in db_df[col].dropna().tolist():
                    navn_clean = str(navn).strip()
                    if navn_clean and navn_clean.lower() not in ['nan', 'none',
                                                                 'nat'] and navn_clean not in dommer_info:
                        dommer_info[navn_clean] = ("", "")
    c = conn.cursor()
    for navn, info in dommer_info.items():
        dtype, epost = info if isinstance(info, tuple) else (info, "")
        c.execute(
            "INSERT INTO dommer_status (navn, type_dommer, epost, aktiv) VALUES (%s, %s, %s, TRUE) ON CONFLICT (navn) DO NOTHING",
            (navn, dtype, epost))
        if dtype or epost:
            c.execute(
                "UPDATE dommer_status SET type_dommer=COALESCE(NULLIF(%s, ''), type_dommer), epost=COALESCE(NULLIF(%s, ''), epost) WHERE navn=%s",
                (dtype, epost, navn))


def get_active_referees_list(conn):
    c = conn.cursor()
    c.execute("SELECT navn FROM dommer_status WHERE aktiv=TRUE ORDER BY navn")
    return [""] + [row[0] for row in c.fetchall()]


def get_match_count(conn, dommer_navn):
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM kamper WHERE dommer_1=%s OR dommer_2=%s", (dommer_navn, dommer_navn))
    return c.fetchone()[0]


def get_age_from_turnering(turnering_str):
    match = re.search(r'(\d{1,2})', str(turnering_str))
    if match: return int(match.group(1))
    return None


def get_allowed_ages(nivå, match_count, is_approved):
    nivå = str(nivå).strip()
    if nivå == 'Nivå 11':
        return [9, 10, 11, 12] if (match_count >= 10 and is_approved) else [9, 10]
    elif nivå == 'Nivå 9':
        return [9, 10, 11, 12]
    elif nivå == 'Nivå 7':
        return [9, 10, 11, 12]
    return [9, 10, 11, 12]


# ==========================================
# 4. E-POST OG INNLOGGINGSKODE
# ==========================================
def generer_og_send_engangskode(conn, mottaker_epost, navn):
    kode = str(random.randint(100000, 999999))
    utloper = datetime.now() + timedelta(minutes=15)

    c = conn.cursor()
    c.execute(
        "INSERT INTO engangskoder (epost, kode, utloper) VALUES (%s, %s, %s) ON CONFLICT (epost) DO UPDATE SET kode = EXCLUDED.kode, utloper = EXCLUDED.utloper",
        (mottaker_epost.lower().strip(), kode, utloper))

    try:
        import smtplib
        from email.mime.text import MIMEText

        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        avsender = st.secrets["email"]["smtp_user"]
        passord = st.secrets["email"]["smtp_password"]

        tekst = f"Hei {navn}!\n\nDin engangskode for innlogging på Dommerberamming Freidig Håndball er: {kode}\n\nKoden er gyldig i 15 minutter."
        msg = MIMEText(tekst)
        msg['From'] = avsender
        msg['To'] = mottaker_epost
        msg['Subject'] = f"Innloggingskode: {kode}"

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(avsender, passord)
        server.send_message(msg)
        server.quit()
        return True, "Koden er sendt til din e-post!"
    except Exception as e:
        return False, f"Feil ved sending av e-post: {e}"


def verifiser_engangskode(conn, mottaker_epost, skrevet_kode):
    c = conn.cursor()
    c.execute("SELECT kode, utloper FROM engangskoder WHERE epost = %s", (mottaker_epost.lower().strip(),))
    rad = c.fetchone()
    if not rad:
        return False, "Ingen kode funnet. Be om ny kode."

    lagret_kode, utloper = rad
    if datetime.now() > utloper:
        return False, "Koden har utløpt. Be om ny kode."

    if skrevet_kode.strip() == lagret_kode:
        c.execute("DELETE FROM engangskoder WHERE epost = %s", (mottaker_epost.lower().strip(),))
        c.execute("UPDATE dommer_status SET epost_bekreftet = TRUE WHERE LOWER(epost) LIKE %s",
                  ('%' + mottaker_epost.lower().strip() + '%',))
        c.execute("UPDATE dommer_ansvar SET epost_bekreftet = TRUE WHERE LOWER(epost) LIKE %s",
                  ('%' + mottaker_epost.lower().strip() + '%',))
        return True, "Vellykket innlogging!"

    return False, "Feil kode."


def registrer_siste_innlogging(conn, tabell, navn):
    c = conn.cursor()
    c.execute(f"""
        UPDATE {tabell} 
        SET siste_innlogging = NOW(), 
            antall_innlogginger = COALESCE(antall_innlogginger, 0) + 1 
        WHERE navn = %s
    """, (navn,))


def send_felles_epost(mottakere, emne, innhold):
    import smtplib
    from email.mime.text import MIMEText

    try:
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        avsender = st.secrets["email"]["smtp_user"]
        passord = st.secrets["email"]["smtp_password"]

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(avsender, passord)

        sendt_antall = 0
        feil_liste = []

        bunntekst = "\n\n---\nDu mottar denne e-posten fordi du er registrert i Dommerberamming Freidig Håndball.\nVil du reservere deg mot e-poster? Svar 'Avmeld' på denne e-posten."
        fullt_innhold = innhold + bunntekst

        for epost_streng, navn in mottakere:
            if not epost_streng or str(epost_streng).strip().lower() in ['nan', 'none', '']:
                continue

            epost_liste = [e.strip() for e in str(epost_streng).replace(',', ';').split(';') if e.strip()]

            for epost in epost_liste:
                try:
                    tekst = f"Hei {navn}!\n\n{fullt_innhold}"
                    msg = MIMEText(tekst)
                    msg['From'] = avsender
                    msg['To'] = epost
                    msg['Subject'] = emne

                    server.send_message(msg)
                    sendt_antall += 1
                except Exception as ex:
                    feil_liste.append(f"{navn} ({epost}): {ex}")

        server.quit()
        return sendt_antall, feil_liste
    except Exception as e:
        return 0, [f"SMTP-feil: {e}"]


# ==========================================
# 5. STREAMLIT APP LOGIKK
# ==========================================
st.set_page_config(page_title="Kampoppsett", layout="wide")
inject_custom_css()

conn = get_db_connection()


@st.cache_resource(ttl=3600)
def oppstart_databasesynk():
    c_conn = get_db_connection()
    init_db(c_conn)
    df_temp = load_data(c_conn)
    sync_referees_to_db(df_temp, c_conn, 'dommere.xlsx')
    return True


oppstart_databasesynk()
df_all = load_data(conn)

turnering_liste = []
if not df_all.empty:
    turneringer = df_all['turnering'].dropna().unique().tolist()
    turnering_liste = sorted(list(set(turneringer)))

df_ansvarlige = execute_query_to_df("SELECT * FROM dommer_ansvar ORDER BY navn", conn)
ansvarlig_navn_liste = df_ansvarlige['navn'].tolist() if not df_ansvarlige.empty else []

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_role = "none"
    st.session_state.user_name = None
    st.session_state.user_teams = []
    st.session_state.venter_paa_kode = False
    st.session_state.epost_for_kode = ""

# -- SIDEBAR INNLOGGING --
st.sidebar.image("FreidigCrest.avif", width=150)
st.sidebar.title("Innlogging")

if not st.session_state.logged_in:
    login_type = st.sidebar.radio("Logg inn som:", ["Administrator", "Dommerkontakt / Lagleder", "Dommer"])

    if login_type == "Administrator":
        pwd = st.sidebar.text_input("Administrator passord", type="password")
        if st.sidebar.button("Logg inn", type="primary"):
            if pwd.strip() == ADMIN_PASSORD:
                st.session_state.logged_in = True
                st.session_state.user_role = "admin"
                st.rerun()
            else:
                st.sidebar.error("Feil passord!")

    elif login_type == "Dommerkontakt / Lagleder":
        if not ansvarlig_navn_liste:
            st.sidebar.info("Ingen dommerkontakter er registrert ennå. Admin må opprette disse.")
        else:
            valgt_navn = st.sidebar.selectbox("Hvem er du?", [""] + ansvarlig_navn_liste)

            c = conn.cursor()
            epost_val = ""
            epost_aktivert = False
            if valgt_navn:
                c.execute("SELECT epost, epost_aktivert FROM dommer_ansvar WHERE navn=%s", (valgt_navn,))
                r = c.fetchone()
                if r:
                    epost_val, epost_aktivert = r[0] or "", bool(r[1])

            bruk_epost_innlogging = KREVER_EPOST_INNLOGGING or epost_aktivert

            if bruk_epost_innlogging and valgt_navn:
                st.sidebar.caption(f"E-postinnlogging aktivert for {valgt_navn}")
                if not epost_val:
                    st.sidebar.error("Ingen e-postadresse registrert på denne brukeren! Kontakt admin.")
                else:
                    if st.sidebar.button("Send meg innloggingskode", type="secondary"):
                        ok, msg = generer_og_send_engangskode(conn, epost_val, valgt_navn)
                        if ok:
                            st.session_state.venter_paa_kode = True
                            st.session_state.epost_for_kode = epost_val
                            st.sidebar.success(msg)
                        else:
                            st.sidebar.error(msg)

                    if st.session_state.venter_paa_kode and st.session_state.epost_for_kode == epost_val:
                        skrevet_kode = st.sidebar.text_input("Skriv inn 6-sifret kode fra e-post")
                        if st.sidebar.button("Bekreft kode og logg inn", type="primary"):
                            suksess, meld = verifiser_engangskode(conn, epost_val, skrevet_kode)
                            if suksess:
                                lag_string = df_ansvarlige[df_ansvarlige['navn'] == valgt_navn]['lag_liste'].iloc[0]
                                mine_lag = [l.strip() for l in lag_string.split("|")] if pd.notna(
                                    lag_string) and lag_string else []

                                reg_siste = "dommer_ansvar"
                                registrer_siste_innlogging(conn, reg_siste, valgt_navn)

                                st.session_state.logged_in = True
                                st.session_state.user_role = "lagleder"
                                st.session_state.user_name = valgt_navn
                                st.session_state.user_teams = mine_lag
                                st.session_state.venter_paa_kode = False
                                st.rerun()
                            else:
                                st.sidebar.error(meld)
            else:
                pwd = st.sidebar.text_input("Passord", type="password")
                if st.sidebar.button("Logg inn", type="primary"):
                    if valgt_navn and pwd.strip() == LAGLEDER_PASSORD:
                        lag_string = df_ansvarlige[df_ansvarlige['navn'] == valgt_navn]['lag_liste'].iloc[0]
                        mine_lag = [l.strip() for l in lag_string.split("|")] if pd.notna(
                            lag_string) and lag_string else []

                        registrer_siste_innlogging(conn, "dommer_ansvar", valgt_navn)

                        st.session_state.logged_in = True
                        st.session_state.user_role = "lagleder"
                        st.session_state.user_name = valgt_navn
                        st.session_state.user_teams = mine_lag
                        st.rerun()
                    else:
                        st.sidebar.error("Feil passord eller manglende navn!")

    elif login_type == "Dommer":
        aktiv_dommer_liste = get_active_referees_list(conn)
        if len(aktiv_dommer_liste) <= 1:
            st.sidebar.info("Ingen aktive dommere registrert ennå.")
        else:
            valgt_dommer = st.sidebar.selectbox("Hvem er du?", aktiv_dommer_liste)

            c = conn.cursor()
            epost_val = ""
            epost_aktivert = False
            if valgt_dommer:
                c.execute("SELECT epost, epost_aktivert FROM dommer_status WHERE navn=%s", (valgt_dommer,))
                r = c.fetchone()
                if r:
                    epost_val, epost_aktivert = r[0] or "", bool(r[1])

            bruk_epost_innlogging = KREVER_EPOST_INNLOGGING or epost_aktivert

            if bruk_epost_innlogging and valgt_dommer:
                st.sidebar.caption(f"E-postinnlogging aktivert for {valgt_dommer}")
                if not epost_val:
                    st.sidebar.error("Ingen e-postadresse registrert på denne dommeren! Kontakt admin.")
                else:
                    if st.sidebar.button("Send meg innloggingskode", type="secondary"):
                        ok, msg = generer_og_send_engangskode(conn, epost_val, valgt_dommer)
                        if ok:
                            st.session_state.venter_paa_kode = True
                            st.session_state.epost_for_kode = epost_val
                            st.sidebar.success(msg)
                        else:
                            st.sidebar.error(msg)

                    if st.session_state.venter_paa_kode and st.session_state.epost_for_kode == epost_val:
                        skrevet_kode = st.sidebar.text_input("Skriv inn 6-sifret kode fra e-post")
                        if st.sidebar.button("Bekreft kode og logg inn", type="primary"):
                            suksess, meld = verifiser_engangskode(conn, epost_val, skrevet_kode)
                            if suksess:
                                registrer_siste_innlogging(conn, "dommer_status", valgt_dommer)

                                st.session_state.logged_in = True
                                st.session_state.user_role = "dommer"
                                st.session_state.user_name = valgt_dommer
                                st.session_state.user_teams = []
                                st.session_state.venter_paa_kode = False
                                st.rerun()
                            else:
                                st.sidebar.error(meld)
            else:
                pwd = st.sidebar.text_input("Passord", type="password")
                if st.sidebar.button("Logg inn", type="primary"):
                    if valgt_dommer and pwd.strip() == DOMMER_PASSORD:
                        registrer_siste_innlogging(conn, "dommer_status", valgt_dommer)

                        st.session_state.logged_in = True
                        st.session_state.user_role = "dommer"
                        st.session_state.user_name = valgt_dommer
                        st.session_state.user_teams = []
                        st.rerun()
                    else:
                        st.sidebar.error("Feil passord eller manglende navn!")
else:
    rolle_navn = "Administrator" if st.session_state.user_role == "admin" else st.session_state.user_name
    st.sidebar.success(f"Logget inn som: {rolle_navn}")
    if st.session_state.user_role == "lagleder":
        st.sidebar.caption(f"Ansvar for: {', '.join(st.session_state.user_teams)}")
    if st.sidebar.button("Logg ut", type="secondary"):
        st.session_state.logged_in = False
        st.session_state.user_role = "none"
        st.session_state.user_name = None
        st.session_state.user_teams = []
        st.session_state.venter_paa_kode = False
        st.rerun()

user_role = st.session_state.user_role

st.markdown('''
    <h1 class="main-title">Dommerberamming Freidig Håndball DBH og D1</h1>
    <div class="title-underline"></div>
''', unsafe_allow_html=True)

if user_role == "none":
    st.info("Vennligst logg inn via menyen til venstre for å få tilgang til systemet.")
    st.stop()

# ==========================================
# DOMMER-VISNING (SELVBETJENING)
# ==========================================
if user_role == "dommer":
    c = conn.cursor()
    c.execute("SELECT type_dommer FROM dommer_status WHERE navn=%s", (st.session_state.user_name,))
    nivaa_row = c.fetchone()
    nivaa = nivaa_row[0] if nivaa_row else "Ukjent"

    antall_kamper = get_match_count(conn, st.session_state.user_name)

    c.execute("SELECT 1 FROM godkjente_opprykk WHERE dommer_navn=%s", (st.session_state.user_name,))
    er_godkjent = bool(c.fetchone())

    tillatte_aldre = get_allowed_ages(nivaa, antall_kamper, er_godkjent)

    st.subheader(f"Velkommen, {st.session_state.user_name} 🟨")
    st.markdown(f"**Ditt nivå:** {nivaa} | **Kamper dømt:** {antall_kamper}")

    if OVERSTYR_NIVAA_KRAV:
        st.warning(
            "⚠️ **Midlertidig unntakstilstand:** Alle nivå- og alderssperrer er midlertidig deaktivert. Du kan nå se og melde interesse på alle åpne kamper. **Vennligst meld kun interesse på kamper du er trygg på at du kan dømme.**")
    elif nivaa == 'Nivå 11' and antall_kamper >= 10:
        if er_godkjent:
            st.success("Du er godkjent for å dømme 11-årskamper!")
        else:
            st.info("Du har passert 10 kamper og avventer godkjenning fra lagleder for å dømme 11-årskamper.")

    c.execute("""
              SELECT k.kampnr,
                     k.dato,
                     k.tid,
                     k.hjemmelag,
                     k.bortelag,
                     k.bane,
                     k.turnering,
                     k.arrangor
              FROM kamper k
              WHERE (k.laast = FALSE OR k.laast IS NULL)
                AND (k.aktiv = TRUE OR k.aktiv IS NULL)
              ORDER BY k.dato, k.tid
              """)
    alle_ulåste = c.fetchall()

    c.execute("SELECT lag_liste FROM dommer_ansvar WHERE selvbetjening=TRUE")
    selvbetjening_lag = []
    for row in c.fetchall():
        if row[0]:
            selvbetjening_lag.extend([re.sub(r'\s+', '', str(l).lower()) for l in row[0].split("|")])

    tilgjengelige_kamper = []
    for kamp in alle_ulåste:
        alder = get_age_from_turnering(kamp[6])
        renset_turnering = re.sub(r'\s+', '', str(kamp[6]).lower())

        har_alders_tilgang = OVERSTYR_NIVAA_KRAV or (alder in tillatte_aldre)

        if har_alders_tilgang and renset_turnering in selvbetjening_lag:
            tilgjengelige_kamper.append(kamp)

    if tilgjengelige_kamper:
        kamper_df = pd.DataFrame(tilgjengelige_kamper,
                                 columns=['kampnr', 'dato', 'tid', 'hjemmelag', 'bortelag', 'bane', 'turnering',
                                          'arrangor'])

        c.execute("SELECT COUNT(*) FROM dommer_status WHERE type_dommer=%s AND aktiv=TRUE", (nivaa,))
        antall_dommere_samme_nivaa = c.fetchone()[0] or 1
        antall_plasser = len(tilgjengelige_kamper) * 2
        maks_kvote = 99

        st.info(
            f"Basert på kapasitet og rettferdig fordeling kan du melde interesse på inntil **{maks_kvote}** av disse kampene.")

        c.execute("SELECT kampnr FROM dommer_onsker WHERE dommer_navn=%s", (st.session_state.user_name,))
        mine_onsker = [row[0] for row in c.fetchall()]
        kamper_df['Ønsker å dømme'] = kamper_df['kampnr'].isin(mine_onsker)

        edited_onsker = st.data_editor(
            kamper_df, hide_index=True,
            disabled=["dato", "tid", "hjemmelag", "bortelag", "bane", "turnering"],
            width="stretch",
            column_config={
                "Ønsker å dømme": st.column_config.CheckboxColumn("Meld interesse", default=False),
                "arrangor": None,
                "kampnr": None,
                "runde": None
            }
        )

        if st.button("Lagre mine ønsker", type="primary"):
            valgte_kampnr = edited_onsker[edited_onsker['Ønsker å dømme'] == True]['kampnr'].tolist()
            if len(valgte_kampnr) > maks_kvote:
                st.error(
                    f"Du har valgt {len(valgte_kampnr)} kamper, men kvoten din er {maks_kvote}. Vennligst fjern noen ønsker før du lagrer.")
            else:
                c.execute("DELETE FROM dommer_onsker WHERE dommer_navn=%s", (st.session_state.user_name,))
                for knr in valgte_kampnr:
                    c.execute("INSERT INTO dommer_onsker (kampnr, dommer_navn) VALUES (%s, %s)",
                              (knr, st.session_state.user_name))
                st.success("Dine ønsker er lagret!")
    else:
        st.info(
            "Det er for øyeblikket ingen ledige kamper som passer ditt nivå (eller årskullene har stengt for selvbetjening).")
    st.stop()

# ==========================================
# ADMIN-VISNING (IMPORT, REGISTER OG STATISTIKK)
# ==========================================
if user_role == "admin":
    with st.expander("📈 Brukerstatistikk & E-post-styring"):
        st.subheader("Send e-post og administrer e-post-innlogging")
        st.info(
            "Huk av for 'Velg' på personene du vil sende e-post til, eller huk av for 'E-post Innlogging' for å aktivere kode-innlogging.")

        tab_lagledere, tab_dommere = st.tabs(["👥 Dommerkontakter / Lagledere", "🏃‍♂️ Dommere"])
        mottakere_valgt = []

        with tab_lagledere:
            df_ansvar_stat = execute_query_to_df(
                "SELECT navn, epost, epost_bekreftet, motta_varsler, epost_aktivert, siste_innlogging, antall_innlogginger FROM dommer_ansvar ORDER BY navn",
                conn)
            if not df_ansvar_stat.empty:
                df_ansvar_stat['epost_aktivert'] = df_ansvar_stat['epost_aktivert'].astype(bool)
                df_ansvar_stat['Velg'] = False

                df_ansvar_stat['epost_bekreftet'] = df_ansvar_stat['epost_bekreftet'].astype(bool)
                df_ansvar_stat['motta_varsler'] = df_ansvar_stat['motta_varsler'].astype(bool)

                edited_ansvar = st.data_editor(
                    df_ansvar_stat, hide_index=True, width="stretch",
                    column_config={
                        "Velg": st.column_config.CheckboxColumn("Velg 📩", default=False),
                        "navn": st.column_config.TextColumn("Navn", disabled=True),
                        "epost": st.column_config.TextColumn("E-post (Sema-separert)"),
                        "epost_bekreftet": st.column_config.CheckboxColumn("Bekreftet ✅", disabled=True),
                        "motta_varsler": st.column_config.CheckboxColumn("Motta Varsler 🔔", default=True),
                        "epost_aktivert": st.column_config.CheckboxColumn("E-post Innlogging"),
                        "siste_innlogging": st.column_config.DatetimeColumn("Sist Logget Inn", disabled=True),
                        "antall_innlogginger": st.column_config.NumberColumn("Antall", disabled=True)
                    }
                )
                if st.button("Lagre Endringer for Kontakter", type="secondary"):
                    c = conn.cursor()
                    for _, row in edited_ansvar.iterrows():
                        c.execute(
                            "UPDATE dommer_ansvar SET epost=%s, epost_aktivert=%s, motta_varsler=%s WHERE navn=%s",
                            (row['epost'], bool(row['epost_aktivert']), bool(row['motta_varsler']), row['navn']))
                    st.success("Oppdatert!")
                    st.rerun()

                for _, r in edited_ansvar[edited_ansvar['Velg'] == True].iterrows():
                    if r['epost']:
                        mottakere_valgt.append((r['epost'], r['navn']))

        with tab_dommere:
            df_dommer_stat = execute_query_to_df(
                "SELECT navn, epost, epost_bekreftet, motta_varsler, epost_aktivert, siste_innlogging, antall_innlogginger FROM dommer_status WHERE aktiv=TRUE ORDER BY navn",
                conn)
            if not df_dommer_stat.empty:
                df_dommer_stat['epost_aktivert'] = df_dommer_stat['epost_aktivert'].astype(bool)
                df_dommer_stat['Velg'] = False

                df_dommer_stat['epost_bekreftet'] = df_dommer_stat['epost_bekreftet'].astype(bool)
                df_dommer_stat['motta_varsler'] = df_dommer_stat['motta_varsler'].astype(bool)

                edited_dommer_stat = st.data_editor(
                    df_dommer_stat, hide_index=True, width="stretch",
                    column_config={
                        "Velg": st.column_config.CheckboxColumn("Velg 📩", default=False),
                        "navn": st.column_config.TextColumn("Navn", disabled=True),
                        "epost": st.column_config.TextColumn("E-post (Sema-separert)"),
                        "epost_bekreftet": st.column_config.CheckboxColumn("Bekreftet ✅", disabled=True),
                        "motta_varsler": st.column_config.CheckboxColumn("Motta Varsler 🔔", default=True),
                        "epost_aktivert": st.column_config.CheckboxColumn("E-post Innlogging"),
                        "siste_innlogging": st.column_config.DatetimeColumn("Sist Logget Inn", disabled=True),
                        "antall_innlogginger": st.column_config.NumberColumn("Antall", disabled=True)
                    }
                )
                if st.button("Lagre Endringer for Dommere", type="secondary"):
                    c = conn.cursor()
                    for _, row in edited_dommer_stat.iterrows():
                        c.execute(
                            "UPDATE dommer_status SET epost=%s, epost_aktivert=%s, motta_varsler=%s WHERE navn=%s",
                            (row['epost'], bool(row['epost_aktivert']), bool(row['motta_varsler']), row['navn']))
                    st.success("Oppdatert!")
                    st.rerun()

                for _, r in edited_dommer_stat[edited_dommer_stat['Velg'] == True].iterrows():
                    if r['epost']:
                        mottakere_valgt.append((r['epost'], r['navn']))

        st.divider()
        st.markdown(f"### 📧 Send e-post til valgte brukere ({len(mottakere_valgt)} valgt)")

        if mottakere_valgt:
            st.caption("Valgte mottakere: " + ", ".join([f"{n} ({e})" for e, n in mottakere_valgt]))
            epost_emne = st.text_input("E-post emne", value="Viktig melding fra Dommerberamming Freidig Håndball")
            epost_innhold = st.text_area("Meldingstekst",
                                         value="Hei!\n\nHusk å gå inn på https://dommere.streamlit.app og sjekke dine kamper.\n\nMed vennlig hilsen\nAdministrator",
                                         height=150)

            if st.button("🚀 Send e-post til valgte brukere nå", type="primary"):
                with st.spinner(f"Sender e-post til {len(mottakere_valgt)} mottaker(e)..."):
                    sendt_ant, feil = send_felles_epost(mottakere_valgt, epost_emne, epost_innhold)
                    if sendt_ant > 0:
                        st.success(f"Vellykket! Sendte {sendt_ant} e-post(er).")
                    if feil:
                        for f_msg in feil:
                            st.error(f_msg)
        else:
            st.info("Huk av i boksene for 'Velg 📩' i tabellene ovenfor for å velge hvem som skal motta e-post.")

    with st.expander("Importér kamper fra Excel"):
        uploaded_file = st.file_uploader("Last opp Kamper (Excel)", type=["xlsx", "xls", "csv"])
        if uploaded_file and st.button("Importér Kamper", type="primary"):
            try:
                df_import = pd.read_csv(uploaded_file, sep=None, engine='python') if uploaded_file.name.endswith(
                    '.csv') else pd.read_excel(uploaded_file)
                rename_dict = {col: db_col for col in df_import.columns for db_col, aliases in EXCEL_ALIASES.items() if
                               str(col).strip().lower() in [a.lower() for a in aliases]}
                df_final = df_import.rename(columns=rename_dict)[
                    [col for col in df_import.rename(columns=rename_dict).columns if col in EXCEL_ALIASES.keys()]]
                if 'dato' in df_final.columns: df_final['dato'] = pd.to_datetime(df_final['dato'],
                                                                                 errors='coerce').dt.strftime(
                    '%Y-%m-%d')
                if 'tid' in df_final.columns: df_final['tid'] = df_final['tid'].astype(str).str[:5]
                if 'kampnr' in df_final.columns: df_final['kampnr'] = df_final['kampnr'].astype(str).str.replace(".0",
                                                                                                                 "",
                                                                                                                 regex=False)

                if 'turnering' in df_final.columns: df_final['turnering'] = df_final['turnering'].astype(str).apply(
                    lambda x: x.split(',')[0].strip())

                df_final = df_final.fillna('')

                imported_count = 0
                c = conn.cursor()

                importerte_kampnr = [str(k).replace(".0", "").strip() for k in df_final['kampnr'].dropna().tolist() if
                                     str(k).strip()]
                if importerte_kampnr:
                    c.execute("""
                              UPDATE kamper
                              SET aktiv  = FALSE,
                                  status = '❌ Kansellert/Flyttet'
                              WHERE kampnr NOT IN %s
                                AND (aktiv = TRUE OR aktiv IS NULL)
                              """, (tuple(importerte_kampnr),))

                for _, row in df_final.iterrows():
                    knr = str(row.get('kampnr', '')).strip()
                    if not knr: continue

                    c.execute("SELECT dato, tid, bane, status FROM kamper WHERE kampnr=%s", (knr,))
                    eksisterende = c.fetchone()

                    if eksisterende:
                        ex_dato, ex_tid, ex_bane, ex_status = eksisterende
                        ny_dato, ny_tid, ny_bane = row.get('dato', ''), row.get('tid', ''), row.get('bane', '')
                        ny_status = ex_status
                        if ex_dato != ny_dato or ex_tid != ny_tid or ex_bane != ny_bane:
                            ny_status = "🔄 Endret"

                        c.execute('''UPDATE kamper
                                     SET runde=%s,
                                         dato=%s,
                                         tid=%s,
                                         hjemmelag=%s,
                                         bortelag=%s,
                                         bane=%s,
                                         arrangor=%s,
                                         turnering=%s,
                                         status=%s,
                                         aktiv= TRUE
                                     WHERE kampnr = %s''',
                                  (row.get('runde'), ny_dato, ny_tid, row.get('hjemmelag'), row.get('bortelag'),
                                   ny_bane, row.get('arrangor'), row.get('turnering'), ny_status, knr))
                    else:
                        c.execute('''INSERT INTO kamper (kampnr, runde, dato, tid, hjemmelag, bortelag, bane, arrangor,
                                                         dommer_1, dommer_2, observator, turnering, status, aktiv)
                                     VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, TRUE)''',
                                  (knr, row.get('runde'), row.get('dato'), row.get('tid'),
                                   row.get('hjemmelag'), row.get('bortelag'), row.get('bane'), row.get('arrangor'),
                                   row.get('dommer_1'), row.get('dommer_2'), row.get('observator'),
                                   row.get('turnering'), '🆕 Ny'))
                    imported_count += 1

                df_all = load_data(conn)
                st.success(f"Suksess! Synkroniserte {imported_count} kamper.")
            except Exception as e:
                st.error(f"Feil ved import: {e}")

    with st.expander("Importér Dommerregister fra Excel"):
        uploaded_dommere = st.file_uploader("Last opp Dommere (Excel)", type=["xlsx", "xls"])
        if uploaded_dommere and st.button("Importér Dommere", type="primary"):
            try:
                temp_filename = "temp_dommere_upload.xlsx"
                with open(temp_filename, "wb") as f:
                    f.write(uploaded_dommere.getbuffer())
                sync_referees_to_db(df_all, conn, temp_filename)
                os.remove(temp_filename)
                st.success("Suksess! Dommerregisteret ble oppdatert.")
                st.rerun()
            except Exception as e:
                st.error(f"Feil ved import av dommere: {e}")

    st.divider()
    st.subheader("Administrer Dommere")
    with st.expander("Aktiver / Deaktiver Dommere"):
        df_dommere_status = execute_query_to_df("SELECT * FROM dommer_status ORDER BY navn", conn)
        df_dommere_status['aktiv'] = df_dommere_status['aktiv'].astype(bool)
        edited_dommere = st.data_editor(
            df_dommere_status, hide_index=True, width="stretch",
            column_config={
                "navn": st.column_config.TextColumn("Dommernavn", disabled=True),
                "type_dommer": st.column_config.TextColumn("Type / Nivå", disabled=True),
                "aktiv": st.column_config.CheckboxColumn("Aktiv", default=True)
            }
        )
        if st.button("Lagre Dommerstatus", type="secondary"):
            c = conn.cursor()
            for _, row in edited_dommere.iterrows():
                c.execute("UPDATE dommer_status SET aktiv=%s WHERE navn=%s",
                          (bool(row['aktiv']), row['navn']))
            st.success("Dommerstatus oppdatert!")

    st.divider()
    st.subheader("Register for Dommerkontakter")
    with st.expander("Opprett eller rediger dommerkontakt"):
        valgt_handling = st.selectbox("Hvem vil du redigere?", ["-- Opprett ny kontakt --"] + ansvarlig_navn_liste)

        with st.form("kontakt_form", clear_on_submit=True):
            col1, col2 = st.columns(2)

            if valgt_handling == "-- Opprett ny kontakt --":
                ny_navn = col1.text_input("Navn på dommerkontakt")
            else:
                ny_navn = valgt_handling
                col1.text_input("Navn på dommerkontakt", value=ny_navn, disabled=True)

            ny_epost = col1.text_input("E-post (La stå tom for å beholde eksisterende)")
            ny_tlf = col2.text_input("Telefon (La stå tom for å beholde eksisterende)")
            valgte_lag = col2.multiselect("Velg årskull/turnering de har ansvar for:", turnering_liste)

            send_epost = st.checkbox("Send velkomst-epost med innlogging ved lagring")

            if st.form_submit_button("Lagre Kontakt"):
                if ny_navn and valgte_lag:
                    c = conn.cursor()
                    lag_str = "|".join(valgte_lag)

                    if valgt_handling == "-- Opprett ny kontakt --":
                        c.execute('''INSERT INTO dommer_ansvar (navn, epost, telefon, lag_liste)
                                     VALUES (%s, %s, %s, %s)''', (ny_navn, ny_epost, ny_tlf, lag_str))
                        epost_som_brukes = ny_epost
                    else:
                        epost_som_brukes = ny_epost if ny_epost else \
                            df_ansvarlige[df_ansvarlige['navn'] == ny_navn]['epost'].values[0]
                        ny_tlf_val = ny_tlf if ny_tlf else \
                            df_ansvarlige[df_ansvarlige['navn'] == ny_navn]['telefon'].values[0]
                        c.execute('''UPDATE dommer_ansvar
                                     SET epost=%s,
                                         telefon=%s,
                                         lag_liste=%s
                                     WHERE navn = %s''', (epost_som_brukes, ny_tlf_val, lag_str, ny_navn))

                    if send_epost and epost_som_brukes:
                        if send_velkomst_epost(epost_som_brukes, ny_navn, ", ".join(valgte_lag)):
                            st.toast("Velkomst-epost sendt!", icon="📧")
                        else:
                            st.warning("Klarte ikke å sende e-post. Sjekk secrets.toml.")

                    st.success(f"Oppdatert ansvar for {ny_navn}!")
                    st.rerun()
                else:
                    st.warning("Du må fylle inn navn og velge minst ett årskull.")

        if not df_ansvarlige.empty:
            st.markdown("**Oversikt over eksisterende kontakter:**")
            st.info("Bruk nedtrekksmenyen over for å redigere årskullene, så formatet lagres riktig.")
            edited_kontakter = st.data_editor(
                df_ansvarlige.drop(columns=['selvbetjening'], errors='ignore'),
                hide_index=True, width="stretch", disabled=["lag_liste"]
            )
            if st.button("Oppdater endringer gjort i tabellen", type="secondary"):
                c = conn.cursor()
                for _, row in edited_kontakter.iterrows():
                    c.execute("UPDATE dommer_ansvar SET epost=%s, telefon=%s WHERE navn=%s",
                              (row['epost'], row['telefon'], row['navn']))
                st.success("Kontaktliste oppdatert.")
                st.rerun()

# ==========================================
# FELLES FOR ADMIN OG LAGLEDER: BERAMMING
# ==========================================
st.divider()
c = conn.cursor()

if user_role == "lagleder":
    st.subheader(f"Lagleder Dashboard: {st.session_state.user_name}")

    # 1. SELVBETJENING TOGGLE
    c.execute("SELECT selvbetjening FROM dommer_ansvar WHERE navn=%s", (st.session_state.user_name,))
    sb_status = c.fetchone()
    current_sb = bool(sb_status[0]) if sb_status else True
    ny_sb = st.toggle("La dommere melde interesse for mine kamper selv (Selvbetjening)", value=current_sb)
    if ny_sb != current_sb:
        c.execute("UPDATE dommer_ansvar SET selvbetjening=%s WHERE navn=%s", (bool(ny_sb), st.session_state.user_name))
        st.success("Selvbetjeningsstatus oppdatert!")

    # 2. GODKJENNING AV NIVÅ 11
    c.execute("""
              SELECT ds.navn
              FROM dommer_status ds
              WHERE ds.type_dommer = 'Nivå 11'
                AND ds.navn NOT IN (SELECT dommer_navn FROM godkjente_opprykk)
              """)
    potensielle_opprykk = c.fetchall()
    kandidater_til_godkjenning = [row[0] for row in potensielle_opprykk if get_match_count(conn, row[0]) >= 10]

    if kandidater_til_godkjenning:
        st.warning("🔔 Følgende Nivå 11-dommere har dømt 10+ kamper og avventer godkjenning for 11-årskamper:")
        for kand in kandidater_til_godkjenning:
            if st.button(f"Godkjenn {kand} for 11-årskamper", key=kand):
                c.execute("INSERT INTO godkjente_opprykk (dommer_navn) VALUES (%s)", (kand,))
                st.success(f"{kand} er nå godkjent!")
                st.rerun()
    st.divider()

if user_role == "admin":
    st.subheader("Beramming av dommere (Alle kamper)")
    view_df = df_all.copy()
else:
    st.subheader("Mine kamper")
    renset_ansvar = [re.sub(r'\s+', '', str(t).lower()) for t in st.session_state.user_teams]
    mask = df_all['turnering'].apply(lambda x: re.sub(r'\s+', '', str(x).lower()) in renset_ansvar)
    view_df = df_all[mask].copy()

    vis_laaste = st.checkbox("Vis også låste kamper (TA)", value=False)
    if not vis_laaste:
        view_df = view_df[view_df['laast'] == False]

if not view_df.empty:
    with st.expander("📊 Vis dommerstatistikk (Antall tildelte kamper)"):
        c.execute("""
                  SELECT dommer, COUNT(*) as antall
                  FROM (SELECT dommer_1 as dommer
                        FROM kamper
                        WHERE dommer_1 != ''
                          AND dommer_1 IS NOT NULL
                        UNION ALL
                        SELECT dommer_2 as dommer
                        FROM kamper
                        WHERE dommer_2 != ''
                          AND dommer_2 IS NOT NULL) as subquery
                  GROUP BY dommer
                  ORDER BY antall DESC
                  """)
        stats = c.fetchall()
        if stats:
            st.dataframe(pd.DataFrame(stats, columns=["Dommer", "Antall kamper"]), hide_index=True)
        else:
            st.info("Ingen kamper er tildelt ennå.")

    c.execute("""
              SELECT dommer, COUNT(*)
              FROM (SELECT dommer_1 as dommer
                    FROM kamper
                    WHERE dommer_1 != ''
                      AND dommer_1 IS NOT NULL
                    UNION ALL
                    SELECT dommer_2 as dommer
                    FROM kamper
                    WHERE dommer_2 != ''
                      AND dommer_2 IS NOT NULL) sub
              GROUP BY dommer
              """)
    match_count_dict = dict(c.fetchall())

    c.execute("SELECT navn, type_dommer FROM dommer_status")
    nivaa_dict = {row[0]: row[1] for row in c.fetchall()}

    c.execute("SELECT kampnr, dommer_navn FROM dommer_onsker")
    onsker_data = c.fetchall()

    onsker_dict = {}
    for knr, dnavn in onsker_data:
        nivaa = nivaa_dict.get(dnavn, "Ukjent")
        kamper_dømt = match_count_dict.get(dnavn, 0)
        formatert_navn = f"{dnavn} ({nivaa} | {kamper_dømt} kamper)"
        onsker_dict.setdefault(knr, []).append(formatert_navn)

    view_df['Interesserte dommere'] = view_df['kampnr'].map(lambda x: onsker_dict.get(x, []))
    if 'laast' not in view_df.columns: view_df['laast'] = False
    view_df['laast'] = view_df['laast'].astype(bool)
    if 'status' not in view_df.columns: view_df['status'] = ''

    aktiv_dommer_liste = get_active_referees_list(conn)
    filter_options = ["Alle"] + sorted(view_df['dato'].dropna().unique().tolist())
    filter_date = st.selectbox("Filtrer på dato:", filter_options)
    display_df = view_df if filter_date == "Alle" else view_df[view_df['dato'] == filter_date]

    if user_role == "admin":
        st.markdown("Fyll inn dommere. **Huk av for 'Låst (TA)'** når kampen er ferdig berammet i MinIdrett/TA.")
    else:
        st.markdown(
            "Sett opp dommere på dine kamper. **Admin vil låse kampen** når den er ferdig berammet i Turneringsadmin (TA).")

    kampnr_konfig = (
        st.column_config.TextColumn("Kamp-ID", disabled=True)
        if user_role == "admin"
        else None
    )

    edited_df = st.data_editor(
        display_df, num_rows="dynamic" if user_role == "admin" else "fixed", width="stretch",
        disabled=["kampnr", "dato", "hjemmelag", "bortelag", "turnering", "Interesserte dommere",
                  "status"], hide_index=True,
        column_config={
            "kampnr": kampnr_konfig,
            "dommer_1": st.column_config.SelectboxColumn("Dommer 1 🔽",
                                                         help="Klikk to ganger for å velge fra listen over aktive dommere.",
                                                         options=aktiv_dommer_liste),
            "dommer_2": st.column_config.SelectboxColumn("Dommer 2 🔽",
                                                         help="Klikk to ganger for å velge fra listen over aktive dommere.",
                                                         options=aktiv_dommer_liste),
            "Interesserte dommere": st.column_config.ListColumn("🙋‍♂️ Ønsker (Nivå | Kamper)", width="large"),
            "laast": st.column_config.CheckboxColumn("Låst (TA) 🔒"),
            "status": st.column_config.TextColumn("Status"),
            "arrangor": None,
            "observator": None,
            "runde": None
        }
    )
    col_save, col_clear, col_export = st.columns([1.5, 1.5, 2])
    with col_save:
        if st.button("Lagre Beramming", type="primary"):
            endret_antall = 0
            for idx, row in edited_df.iterrows():
                orig_row = display_df.loc[idx]

                har_endring = (
                        str(row.get('dommer_1', '')) != str(orig_row.get('dommer_1', '')) or
                        str(row.get('dommer_2', '')) != str(orig_row.get('dommer_2', '')) or
                        str(row.get('observator', '')) != str(orig_row.get('observator', '')) or
                        str(row.get('bane', '')) != str(orig_row.get('bane', '')) or
                        str(row.get('tid', '')) != str(orig_row.get('tid', '')) or
                        bool(row.get('laast', False)) != bool(orig_row.get('laast', False))
                )

                if har_endring:
                    c.execute('''UPDATE kamper
                                 SET dommer_1=%s,
                                     dommer_2=%s,
                                     observator=%s,
                                     bane=%s,
                                     tid=%s,
                                     laast=%s
                                 WHERE kampnr = %s''',
                              (str(row.get('dommer_1', '')), str(row.get('dommer_2', '')),
                               str(row.get('observator', '')),
                               str(row.get('bane', '')), str(row.get('tid', '')), bool(row.get('laast', False)),
                               str(row['kampnr'])))
                    endret_antall += 1

            if endret_antall > 0:
                st.success(f"Oppdatert {endret_antall} kamp(er)!")
                st.rerun()
            else:
                st.info("Ingen endringer å lagre.")

    with col_clear:
        if st.button("Fjern 'Ny/Endret'-varsler", type="secondary"):
            for kampnr in view_df['kampnr']:
                c.execute("UPDATE kamper SET status='' WHERE kampnr=%s", (kampnr,))
            st.rerun()

    with col_export:
        if user_role == "admin":
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer: df_all.to_excel(writer, index=False,
                                                                                      sheet_name='Kamper')
            st.download_button(label="⬇️ Eksporter hele databasen til Excel", data=buffer.getvalue(),
                               file_name=f"dommeroppsett_eksport_{date.today().strftime('%Y%m%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="secondary")
else:
    st.info("Fant ingen kamper.")