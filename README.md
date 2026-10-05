# 🟨 Dommerberamming Freidig Håndball

En Streamlit-basert webapplikasjon bygget for Freidig Håndball (DBH og D1) for å forenkle og automatisere prosessen med dommerberamming. Applikasjonen lar administratorer importere kamper fra MinIdrett, mens dommere selv kan melde interesse for ledige kamper basert på deres nivå og erfaring.

## 🌟 Nøkkelfunksjoner

Applikasjonen har rollebasert tilgangskontroll med tre hovedvisninger:

### 1. Administrator (`admin`)

* **Import/Eksport:** Last opp kampoppsett og dommerregister direkte fra Excel-filer (MinIdrett-eksport). Appen oppdaterer kun endringer (f.eks. flyttede kamper) og bevarer eksisterende beramming.
* **Dommerhåndtering:** Aktiver/deaktiver dommere, og administrer dommerkontakter (lagledere).
* **Full oversikt:** Beram dommere på tvers av alle årskull, merk kamper som "Låst i TA", og hold oversikt over dommerstatistikk.
* **Kommunikasjon:** Send automatiserte velkomst-e-poster til nye lagledere med innloggingsdetaljer.
* **PDF-generering:** Lag og last ned skreddersydde, utskriftsvennlige PDF-rapporter (med Freidig-profil) av dommeroppsettet.

### 2. Dommerkontakt / Lagleder (`lagleder`)

* **Årskull-fokus:** Får kun opp kamper for de lagene/årskullene de har fått tildelt ansvar for.
* **Selvbetjening:** Kan skru av og på muligheten for at dommere kan melde interesse for sine kamper.
* **Godkjenning (Opprykk):** Får automatisk varsel når en Nivå 11-dommer har dømt 10 kamper, og kan godkjenne dem for å dømme eldre klasser.

### 3. Dommer (`dommer`)

* **Selvbetjening:** Logg inn og se ledige kamper som matcher dommerens godkjente nivå (f.eks. Nivå 11, Nivå 9).
* **Meld interesse:** Velg hvilke kamper som passer. Systemet viser disse ønskene direkte i berammingstabellen til admin/lagleder.

---

## 🛠 Teknologi og Avhengigheter

* **Frontend:** [Streamlit](https://streamlit.io/) (Python)
* **Database:** PostgreSQL via [Supabase](https://supabase.com/) (`psycopg2`)
* **Datahåndtering:** `pandas`, `openpyxl`
* **PDF-generering:** `weasyprint`
* **E-post:** `smtplib` (Innebygd Python-bibliotek)

### Nødvendige Python-pakker (`requirements.txt`)

```text
streamlit
pandas
psycopg2-binary
weasyprint
openpyxl

```

---

## 🚀 Oppsett og Installasjon

**1. Klon repository og installer avhengigheter:**

```bash
git clone <din-repo-url>
cd Dommere
pip install -r requirements.txt

```

**2. Konfigurer Secrets:**
Opprett en fil kalt `secrets.toml` inne i en `.streamlit/`-mappe i rotkatalogen. Denne må inneholde databasetilkobling, passord og e-postinnstillinger:

```toml
# .streamlit/secrets.toml

[passwords]
admin = "admin123"
lagleder = "dommer123"
dommer = "fløyte"

[supabase]
db_url = "postgresql://bruker:passord@db.supabase.co:5432/postgres"

[email]
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_user = "din.epost@gmail.com"
smtp_password = "app-spesifikt-passord"
admin_epost = "admin.epost@freidig.no"

```

**3. Kjør applikasjonen lokalt:**

```bash
streamlit run app_2.py

```

---

## 🗄 Databasestruktur (PostgreSQL)

Appen genererer og vedlikeholder følgende tabeller automatisk ved første oppstart:

* `kamper`: Lagrer alle kampdetaljer (kampnr, lag, tid, bane, tildelte dommere, status).
* `dommer_status`: Register over alle dommere, deres aktive status og dommernivå.
* `dommer_ansvar`: Håndterer lagledere/dommerkontakter og hvilke årskull de styrer.
* `dommer_onsker`: Koblingstabell for dommere som har meldt interesse for spesifikke kamper.
* `godkjente_opprykk`: Holder oversikt over Nivå 11-dommere som er godkjent for å dømme eldre klasser.

---

## 💡 Brukerveiledning for Excel-import

1. Eksporter kamper eller dommerliste fra MinIdrett/TurneringsAdmin i Excel-format (`.xlsx`).
2. Appen bruker fleksible aliaser for å kjenne igjen kolonner (f.eks. `Kampnr`, `Dato`, `Tid`, `Dommer 1`, `Klasse`).
3. Logg inn som Administrator i appen.
4. Åpne **"Importér kamper fra Excel"** eller **"Importér Dommerregister fra Excel"**, velg filen og trykk import. Alt synkroniseres til databasen på få sekunder.