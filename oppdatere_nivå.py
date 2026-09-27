import pandas as pd
from io import StringIO

# 1. Legg inn svarene fra Google Forms
forms_data = """Tidsmerke	Hva er navnet ditt?	Ønsker du å dømme kamper denne sesongen?	Ønsker du å være medlem av Dommergruppa i Spond?	Hvor mange kamper ser du for deg å dømme?	Hvilke e-postadresser skal brukes for varsler fremover? Skriv kun inn din egen hvis du styrer alt selv og ikke trenger foresatte på kopi
20.09.2026 kl. 20.11.30	Julie Lorentsen	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	Julie.r.lorentsen@gmail.com
20.09.2026 kl. 20.16.53	Christina Marie Gripp Jensen	Ja	Ja	Kan bidra så mye som mulig så lenge det ikke kræsjer med oppsett fra regionen	christina.mg.jensen@gmail.com
20.09.2026 kl. 20.22.43	Lene Westby	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	lenewestby23@gmail.com
20.09.2026 kl. 20.30.36	Sofia Dahl	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	sofia23dahl@gmail.com
20.09.2026 kl. 20.35.28	Linnea Tessem Størseth	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	linnea.storseth@gmail.com
20.09.2026 kl. 21.11.56	Tyra Ormbostad	Ja	Ja	Så mye som mulig	tyraormb@gmail.com
20.09.2026 kl. 21.29.56	Julie Westeren	Ja	Ja	Så mye som mulig	Julielawe@icloud.com eller hege.langeland@hotmail.com
20.09.2026 kl. 21.40.34	Ingrid S. Gabrielsen	Ja	Ja	Så mye som mulig	ingsangab@gmail.com
20.09.2026 kl. 22.12.31	Mari Elshaug-Tørstad	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	marielshaugtorstad@gmail.com
20.09.2026 kl. 22.14.17	Helena Bakken Nakstad	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	helena.bakken.nakstad@gmail.com
20.09.2026 kl. 22.16.37	Ragne Fougner Nordhagen	Ja	Ja	Så mye som mulig så lenge det ikke kræsjer med egen aktivitet	ragnefn@gmail.com
21.09.2026 kl. 03.41.33	Vilma Dalheim	Nei	Nei		
21.09.2026 kl. 08.22.06	Viola Margrethe Arnesen Aandal	Ja	Ja	Så mye som mulig	Fotoanette@gmail.com"""

df_forms = pd.read_csv(StringIO(forms_data), sep='\t')

# Hent ut navnene på de som har svart 'Ja' og gjort dem om til små bokstaver
aktive_dommere = df_forms[df_forms['Ønsker du å dømme kamper denne sesongen?'] == 'Ja']['Hva er navnet ditt?'].str.strip().str.lower().tolist()

# 2. Last inn Excel-filen
# NB: Pass på at du har lukket filen i Microsoft Excel før du kjører scriptet!
filnavn = 'dommere.xlsx'
try:
    df_excel = pd.read_excel(filnavn)

    # Lag en midlertidig kolonne med fullt navn (i små bokstaver) for å treffe riktig person
    df_excel['Fullt Navn'] = (df_excel['Fornavn'].astype(str).str.strip() + ' ' + df_excel['Etternavn'].astype(str).str.strip()).str.lower()

    # 3. Oppdater nivået til "Nivå 9" for alle som har et treff
    mask = df_excel['Fullt Navn'].isin(aktive_dommere)
    df_excel.loc[mask, 'Dommer'] = 'Nivå 9'

    # Rydder opp ved å slette den midlertidige navnekolonnen
    df_excel = df_excel.drop(columns=['Fullt Navn'])

    # 4. Lagre endringene tilbake til originalfilen
    df_excel.to_excel(filnavn, index=False)
    print(f"✅ Suksess! Oppdaterte nivået til 'Nivå 9' for {mask.sum()} dommere.")

except Exception as e:
    print(f"❌ Det oppstod en feil (er filen åpen i Excel?): {e}")