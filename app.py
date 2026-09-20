import os
import io
import re
import math
import psycopg2
import pandas as pd
import streamlit as st
import base64
from datetime import date
from weasyprint import HTML
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 1. KONSTANTER & ALIASER
# ==========================================
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

ADMIN_PASSORD = "admin123"
LAGLEDER_PASSORD = "dommer123"
DOMMER_PASSORD = "fløyte"


# ==========================================
# 2. STYLING (MODERNE DESIGN)
# ==========================================
def inject_custom_css():
    st.markdown('''
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Play:wght@400;700&display=swap');
        .stApp { background: linear-gradient(135deg, #fcfcfc 0%, #f0f0f0 100%); }
        h1, h2, h3, h4, h5, h6, .st-emotion-cache-10trblm { color: #1f1f1f !important; font-family: 'Play', sans-serif; letter-spacing: -0.5px; }
        .main-title { text-align: left; font-size: 3rem; font-weight: 800; background: -webkit-linear-gradient(45deg, #FF5800, #FF8C00); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 2rem; padding-top: 1rem; }
        .st-emotion-cache-10trblm { border-bottom: 2px solid #eaeaea; padding-bottom: 10px; margin-bottom: 20px; }
        .stButton > button[kind="primary"] { background: linear-gradient(90deg, #FF5800 0%, #FF8C00 100%); color: white; border: none; border-radius: 8px; box-shadow: 0 4px 8px rgba(255, 88, 0, 0.25); transition: all 0.3s ease; font-family: 'Arial', sans-serif; font-weight: bold; }
        .stButton > button[kind="primary"]:hover { transform: translateY(-2px); box-shadow: 0 6px 14px rgba(255, 88, 0, 0.4); color: white; }
        .stButton > button[kind="secondary"] { border: 2px solid #1f1f1f; color: #1f1f1f; border-radius: 8px; transition: all 0.3s ease; background-color: white; font-family: 'Arial', sans-serif; font-weight: bold; }
        .stButton > button[kind="secondary"]:hover { border: 2px solid #FF5800; color: #FF5800; background-color: #fffaf6; transform: translateY(-2px); box-shadow: 0 4px 8px rgba(255, 88, 0, 0.1); }
        .stDataFrame { border: none !important; border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); background-color: white; padding: 10px; }
        .streamlit-expanderHeader { font-size: 1.1rem; font-family: 'Play', sans-serif; font-weight: 700; color: #1f1f1f; border-radius: 8px; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.02); margin-bottom: 5px; }
    </style>
    ''', unsafe_allow_html=True)


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

    # Legger til nye kolonner hvis de ikke finnes (støttet i Postgres)
    c.execute("ALTER TABLE dommer_ansvar ADD COLUMN IF NOT EXISTS selvbetjening BOOLEAN DEFAULT TRUE")
    c.execute("ALTER TABLE kamper ADD COLUMN IF NOT EXISTS laast BOOLEAN DEFAULT FALSE")
    c.execute("ALTER TABLE kamper ADD COLUMN IF NOT EXISTS status TEXT DEFAULT ''")


def load_data(conn):
    return execute_query_to_df("SELECT * FROM kamper", conn)


def sync_referees_to_db(db_df, conn, filepath='dommere.xlsx'):
    dommer_info = {}
    if os.path.exists(filepath):
        try:
            df_dommere = pd.read_excel(filepath)
            if 'Fornavn' in df_dommere.columns and 'Etternavn' in df_dommere.columns:
                df_dommere['Fullt Navn'] = df_dommere['Fornavn'].astype(str).str.strip() + " " + df_dommere[
                    'Etternavn'].astype(str).str.strip()
                type_col = 'Dommer' if 'Dommer' in df_dommere.columns else None
                for _, row in df_dommere.iterrows():
                    navn = str(row['Fullt Navn']).strip()
                    if navn and navn.lower() not in ['nan', 'none']:
                        dtype = str(row[type_col]).strip() if type_col and pd.notna(row[type_col]) else ""
                        dommer_info[navn] = dtype
        except Exception:
            pass

    if not db_df.empty:
        for col in ['dommer_1', 'dommer_2', 'observator']:
            if col in db_df.columns:
                for navn in db_df[col].dropna().tolist():
                    navn_clean = str(navn).strip()
                    if navn_clean and navn_clean.lower() not in ['nan', 'none',
                                                                 'nat'] and navn_clean not in dommer_info:
                        dommer_info[navn_clean] = ""
    c = conn.cursor()
    for navn, dtype in dommer_info.items():
        c.execute(
            "INSERT INTO dommer_status (navn, type_dommer, aktiv) VALUES (%s, %s, TRUE) ON CONFLICT (navn) DO NOTHING",
            (navn, dtype))
        if dtype:
            c.execute(
                "UPDATE dommer_status SET type_dommer=%s WHERE navn=%s AND (type_dommer IS NULL OR type_dommer='')",
                (dtype, navn))


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
        return [9, 10, 11] if (match_count >= 10 and is_approved) else [9, 10]
    elif nivå == 'Nivå 9':
        return [11, 12]
    elif nivå == 'Nivå 7':
        return [12]
    return [9, 10, 11, 12]


# ==========================================
# 4. PDF GENERATOR & EPOST
# ==========================================
def generate_schedule_pdf(df, title="Dommeroppsett"):
    df = df.fillna('')
    if 'dato' in df.columns and 'tid' in df.columns:
        df = df.sort_values(by=["dato", "tid"])
    table_html = ""
    current_date = ""
    for index, row in df.iterrows():
        if row["dato"] != current_date:
            table_html += f'<tr class="date-group"><td colspan="6">{row["dato"]}</td></tr>'
            current_date = row["dato"]
        kamp_info = f'<strong>{row.get("hjemmelag", "")} - {row.get("bortelag", "")}</strong><br><span class="desc">{row.get("turnering", "")}</span>'
        dommere = []
        if str(row.get("dommer_1", "")).strip(): dommere.append(f'D1: {row.get("dommer_1", "")}')
        if str(row.get("dommer_2", "")).strip(): dommere.append(f'D2: {row.get("dommer_2", "")}')
        if str(row.get("observator", "")).strip(): dommere.append(f'Obs: {row.get("observator", "")}')
        dommer_str = "<br>".join(dommere) if dommere else "<i>Ikke berammet</i>"
        table_html += f'''<tr><td style="white-space: nowrap;">{row.get("tid", "")}</td><td style="font-size: 8pt; color: #555;">{row.get("kampnr", "")}</td><td>{kamp_info}</td><td>{row.get("bane", "")}</td><td>{row.get("arrangor", "")}</td><td>{dommer_str}</td></tr>'''

    html_template = f'''
    <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    @page {{ size: A4 landscape; margin-top: 18mm; margin-bottom: 15mm; margin-left: 15.4mm; margin-right: 15.4mm; background-color: #ffffff; @bottom-right {{ content: "Side " counter(page) " av " counter(pages); font-family: 'Arial', sans-serif; font-size: 10pt; color: #1f1f1f; }} @bottom-left {{ content: "Generert fra Kamp- og Dommersystem"; font-family: 'Arial', sans-serif; font-size: 10pt; color: #1f1f1f; }} }}
    body {{ font-family: 'Arial', sans-serif; margin: 0; padding: 0; color: #000000; }}
    .header {{ padding-bottom: 10px; margin-bottom: 20px; width: 100%; border-bottom: 2px solid #FF5800; }}
    .header h1 {{ margin: 0; color: #FF5800; font-family: 'Play', sans-serif; font-size: 24pt; padding-bottom: 5px; font-weight: normal; }}
    .header .meta {{ font-size: 11pt; color: #1f1f1f; margin-top: 10px; font-family: 'Arial', sans-serif; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 10.5pt; margin-top: 15px; page-break-inside: avoid; }}
    th {{ background-color: #f2f2f2; color: #000000; text-align: left; padding: 10px 8px; font-weight: bold; border: 1pt solid #000000; text-transform: uppercase; font-size: 9pt; letter-spacing: 0.5px; }}
    td {{ padding: 10px 8px; border: 1pt solid #000000; vertical-align: top; line-height: 1.3; }}
    .date-group td {{ background-color: #e2e8f0; color: #000000; font-weight: bold; font-size: 12pt; padding: 10px 8px; font-family: 'Play', sans-serif; border-bottom: 1pt solid #000000; }}
    .desc {{ font-size: 9pt; color: #444; display: block; margin-top: 2px; font-style: italic; }}
    tr {{ page-break-inside: avoid; }}
    </style></head><body>
    <div class="header"><h1>{title}</h1><div class="meta">Utskriftsdato: <strong>{date.today().strftime('%d.%m.%Y')}</strong></div></div>
    <table><thead><tr><th style="width: 5%;">Tid</th><th style="width: 10%;">Kampnr</th><th style="width: 30%;">Kamp & Turnering</th><th style="width: 20%;">Bane</th><th style="width: 15%;">Arrangør</th><th style="width: 20%;">Dommere / Obs</th></tr></thead><tbody>{table_html}</tbody></table>
    </body></html>'''
    return HTML(string=html_template, base_url=os.path.abspath(os.getcwd())).write_pdf()


def send_velkomst_epost(mottaker_epost, navn, lag_liste_str):
    try:
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        avsender = st.secrets["email"]["smtp_user"]
        passord = st.secrets["email"]["smtp_password"]
        admin_epost = st.secrets["email"]["admin_epost"]

        msg = MIMEMultipart()
        msg['From'] = avsender
        msg['To'] = mottaker_epost
        msg['Reply-To'] = admin_epost
        msg['Subject'] = "Tilgang til nytt dommeroppsett"

        body = f"""Hei {navn}!\n\nDu har nå fått tilgang som dommerkontakt i det nye systemet for dommeroppsett.\n\nDine ansvarsområder (turnering/årsklasse): {lag_liste_str}\n\nSlik logger du inn:\n1. Gå til: https://dommere.streamlit.app\n2. Velg "Dommerkontakt / Lagleder" i menyen.\n3. Velg ditt navn i listen.\n4. Passordet er: {LAGLEDER_PASSORD}\n\nVi setter utrolig stor pris på alle tilbakemeldinger. Bare svar direkte på denne e-posten!\n\nVennlig hilsen,\nAdministrator"""
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(avsender, passord)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"E-POST FEIL: {e}")
        return False


# ==========================================
# 5. STREAMLIT APP LOGIKK
# ==========================================
st.set_page_config(page_title="Kampoppsett", layout="wide")
inject_custom_css()

conn = get_db_connection()
init_db(conn)

df_all = load_data(conn)
sync_referees_to_db(df_all, conn, 'dommere.xlsx')

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

# -- SIDEBAR INNLOGGING --
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
            pwd = st.sidebar.text_input("Passord", type="password")
            if st.sidebar.button("Logg inn", type="primary"):
                if valgt_navn and pwd.strip() == LAGLEDER_PASSORD:
                    lag_string = df_ansvarlige[df_ansvarlige['navn'] == valgt_navn]['lag_liste'].iloc[0]
                    if pd.notna(lag_string) and lag_string:
                        # Splitter KUN på loddrett strek
                        mine_lag = [l.strip() for l in lag_string.split("|")]
                    else:
                        mine_lag = []

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
            pwd = st.sidebar.text_input("Passord", type="password")
            if st.sidebar.button("Logg inn", type="primary"):
                if valgt_dommer and pwd.strip() == DOMMER_PASSORD:
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
        st.rerun()

user_role = st.session_state.user_role
st.markdown('<h1 class="main-title">Kamp- og Dommeroppsett</h1>', unsafe_allow_html=True)

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
    if nivaa == 'Nivå 11' and antall_kamper >= 10:
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
              WHERE k.laast = FALSE
                 OR k.laast IS NULL
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
        if alder in tillatte_aldre and renset_turnering in selvbetjening_lag:
            tilgjengelige_kamper.append(kamp)

    if tilgjengelige_kamper:
        kamper_df = pd.DataFrame(tilgjengelige_kamper,
                                 columns=['kampnr', 'dato', 'tid', 'hjemmelag', 'bortelag', 'bane', 'turnering',
                                          'arrangor'])

        c.execute("SELECT COUNT(*) FROM dommer_status WHERE type_dommer=%s AND aktiv=TRUE", (nivaa,))
        antall_dommere_samme_nivaa = c.fetchone()[0] or 1
        antall_plasser = len(tilgjengelige_kamper) * 2
        maks_kvote = max(2, math.ceil(antall_plasser / antall_dommere_samme_nivaa))

        st.info(
            f"Basert på kapasitet og rettferdig fordeling kan du melde interesse på inntil **{maks_kvote}** av disse kampene.")

        c.execute("SELECT kampnr FROM dommer_onsker WHERE dommer_navn=%s", (st.session_state.user_name,))
        mine_onsker = [row[0] for row in c.fetchall()]
        kamper_df['Ønsker å dømme'] = kamper_df['kampnr'].isin(mine_onsker)

        edited_onsker = st.data_editor(
            kamper_df, hide_index=True,
            disabled=["kampnr", "dato", "tid", "hjemmelag", "bortelag", "bane", "turnering", "arrangor"],
            use_container_width=True,
            column_config={"Ønsker å dømme": st.column_config.CheckboxColumn("Meld interesse", default=False)}
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
# ADMIN-VISNING (IMPORT OG REGISTER)
# ==========================================
if user_role == "admin":
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

                # Klipper bort alt etter første komma
                if 'turnering' in df_final.columns: df_final['turnering'] = df_final['turnering'].astype(str).apply(
                    lambda x: x.split(',')[0].strip())

                df_final = df_final.fillna('')

                imported_count = 0
                c = conn.cursor()
                for _, row in df_final.iterrows():
                    if not row.get('kampnr'): continue

                    c.execute("SELECT dato, tid, bane, status FROM kamper WHERE kampnr=%s", (row['kampnr'],))
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
                                         status=%s
                                     WHERE kampnr = %s''',
                                  (row.get('runde'), ny_dato, ny_tid, row.get('hjemmelag'), row.get('bortelag'),
                                   ny_bane, row.get('arrangor'), row.get('turnering'), ny_status, row['kampnr']))
                    else:
                        c.execute('''INSERT INTO kamper (kampnr, runde, dato, tid, hjemmelag, bortelag, bane, arrangor,
                                                         dommer_1, dommer_2, observator, turnering, status)
                                     VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)''',
                                  (row['kampnr'], row.get('runde'), row.get('dato'), row.get('tid'),
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
        edited_dommere = st.data_editor(df_dommere_status, hide_index=True, use_container_width=True,
                                        column_config={"navn": st.column_config.TextColumn("Dommernavn", disabled=True),
                                                       "type_dommer": st.column_config.TextColumn("Type / Nivå",
                                                                                                  disabled=True),
                                                       "aktiv": st.column_config.CheckboxColumn("Aktiv", default=True)})
        if st.button("Lagre Dommerstatus", type="secondary"):
            c = conn.cursor()
            for _, row in edited_dommere.iterrows(): c.execute("UPDATE dommer_status SET aktiv=%s WHERE navn=%s",
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
                        epost_som_brukes = ny_epost if ny_epost else df_ansvarlige[df_ansvarlige['navn'] == ny_navn]['epost'].values[0]
                        ny_tlf_val = ny_tlf if ny_tlf else df_ansvarlige[df_ansvarlige['navn'] == ny_navn]['telefon'].values[0]
                        c.execute('''UPDATE dommer_ansvar
                                     SET epost=%s,telefon=%s,lag_liste=%s
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
            edited_kontakter = st.data_editor(df_ansvarlige.drop(columns=['selvbetjening'], errors='ignore'),
                                              hide_index=True, use_container_width=True, disabled=["lag_liste"])
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

    c.execute("SELECT kampnr, dommer_navn FROM dommer_onsker")
    onsker_data = c.fetchall()

    c.execute("SELECT navn, type_dommer FROM dommer_status")
    nivaa_dict = {row[0]: row[1] for row in c.fetchall()}

    onsker_dict = {}
    for knr, dnavn in onsker_data:
        nivaa = nivaa_dict.get(dnavn, "Ukjent")
        kamper_dømt = get_match_count(conn, dnavn)
        formatert_navn = f"{dnavn} ({nivaa} | {kamper_dømt} kamper)"
        onsker_dict.setdefault(knr, []).append(formatert_navn)

    view_df['Interesserte dommere'] = view_df['kampnr'].map(lambda x: " \n ".join(onsker_dict.get(x, [])))
    if 'laast' not in view_df.columns: view_df['laast'] = False
    view_df['laast'] = view_df['laast'].astype(bool)
    if 'status' not in view_df.columns: view_df['status'] = ''

    aktiv_dommer_liste = get_active_referees_list(conn)
    filter_options = ["Alle"] + sorted(view_df['dato'].dropna().unique().tolist())
    filter_date = st.selectbox("Filtrer på dato:", filter_options)
    display_df = view_df if filter_date == "Alle" else view_df[view_df['dato'] == filter_date]

    st.markdown("Fyll inn dommere. **Huk av for 'Låst (TA)'** når kampen er ferdig berammet i MinIdrett.")
    edited_df = st.data_editor(
        display_df, num_rows="dynamic" if user_role == "admin" else "fixed", use_container_width=True,
        disabled=["kampnr", "runde", "dato", "hjemmelag", "bortelag", "turnering", "arrangor", "Interesserte dommere",
                  "status"], hide_index=True,
        column_config={
            "dommer_1": st.column_config.SelectboxColumn("Dommer 1", options=aktiv_dommer_liste),
            "dommer_2": st.column_config.SelectboxColumn("Dommer 2", options=aktiv_dommer_liste),
            "observator": st.column_config.SelectboxColumn("Observatør", options=aktiv_dommer_liste),
            "Interesserte dommere": st.column_config.TextColumn("🙋‍♂️ Ønsker (Nivå | Kamper)"),
            "laast": st.column_config.CheckboxColumn("Låst (TA) 🔒"),
            "status": st.column_config.TextColumn("Status")
        }
    )

    col_save, col_clear, col_export = st.columns([1.5, 1.5, 2])
    with col_save:
        if st.button("Lagre Beramming", type="primary"):
            for _, row in edited_df.iterrows():
                c.execute('''UPDATE kamper
                             SET dommer_1=%s,
                                 dommer_2=%s,
                                 observator=%s,
                                 bane=%s,
                                 tid=%s,
                                 laast=%s
                             WHERE kampnr = %s''',
                          (str(row.get('dommer_1', '')), str(row.get('dommer_2', '')), str(row.get('observator', '')),
                           str(row.get('bane', '')), str(row.get('tid', '')), bool(row.get('laast', False)),
                           str(row['kampnr'])))
            st.success("Oppdatert!")

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

# ==========================================
# GENERER PDF
# ==========================================
st.divider()
if not view_df.empty:
    col1, col2 = st.columns(2)
    with col1:
        pdf_title = st.text_input("Tittel på dokumentet", value="Dommeroppsett")
        pdf_filter = st.selectbox("Hvilke kamper skal inkluderes i PDF?",
                                  ["Alle dine viste kamper"] + sorted(view_df['dato'].dropna().unique().tolist()))

    if st.button("Generer PDF"):
        with st.spinner("Lager PDF..."):
            pdf_data_source = view_df if pdf_filter == "Alle dine viste kamper" else view_df[
                view_df['dato'] == pdf_filter]
            pdf_bytes = generate_schedule_pdf(pdf_data_source, title=pdf_title)
            os.makedirs("output", exist_ok=True)
            file_path = os.path.join("output", f"oppsett_{date.today().strftime('%Y%m%d')}.pdf")
            with open(file_path, "wb") as f: f.write(pdf_bytes)
            b64 = base64.b64encode(pdf_bytes).decode()
            st.markdown(
                f'''<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}" style="text-decoration: none;"><button style="border: 2px solid #1f1f1f; color: #1f1f1f; border-radius: 8px; background-color: white; font-family: 'Arial', sans-serif; font-weight: bold; padding: 0.5rem 1rem; cursor: pointer; transition: all 0.2s ease;">⬇️ Last ned PDF</button></a>''',
                unsafe_allow_html=True)