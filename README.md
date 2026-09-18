# 🏆 Kamp- og Dommeroppsett

Dette er en skreddersydd webapplikasjon bygget med [Streamlit](https://streamlit.io/) for å administrere kamper, beramme dommere og generere utskriftsvennlige dommeroppsett (PDF). Systemet er spesielt tilpasset idrettsklubber med behov for rollebasert tilgang og enkel import/eksport til Excel.

## ✨ Nøkkelfunksjoner

- **Rollebasert tilgang:** Egen innlogging for Administrator og Dommerkontakter/Lagledere.
- **Smart Excel-import:** Last opp `kamper.xlsx` for å oppdatere tider, baner og lag, *uten* å overskrive dommere som allerede er berammet.
- **Dommerregister:** Henter automatisk dommere fra `dommere.xlsx`. Admin kan deaktivere dommere som er midlertidig utilgjengelige.
- **Dommerkontakter (Lagledere):** Admin kan knytte kontaktpersoner til ett eller flere lag. Når de logger inn, ser og administrerer de *kun* sine egne kamper.
- **PDF-generering:** Genererer proffe dommeroppsett (A4 landskap) for valgte datoer direkte fra nettleseren.
- **Eksport:** Eksporter det oppdaterte oppsettet tilbake til Excel med ett klikk.

## 🛠 Installasjon (Lokalt)

For å kjøre denne appen lokalt på din egen datamaskin, trenger du Python installert.

1. **Last ned/klon prosjektet** til din maskin.
2. **Åpne en terminal** i prosjektmappen og installer avhengighetene:
   ```bash
   pip install -r requirements.txt