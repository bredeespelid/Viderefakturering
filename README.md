Dette skriptet automatiserer prosessen med å ekstrahere data fra PDF-filer og lagre resultatet i en Excel-fil. Skriptet bruker PyPDF2 for å lese PDF-filer, pandas for datahåndtering, tkinter for filvalg og openpyxl for å skrive data til en Excel-fil.

Forutsetninger
For å kjøre dette skriptet trenger du:

Python installert på systemet ditt.
Følgende Python-pakker: PyPDF2, pandas, tkinter, openpyxl, re.
Funksjonalitet
Skriptet utfører følgende operasjoner:

Ekstrahere data fra PDF-filer:

Leser tekst fra hver side i PDF-filen.
Ekstraherer relevante data fra tekstlinjer som inneholder spesifikke mønstre og tekststrenger.
Henter rekvisisjonsnummer hvis det er tilstede.
Formatere og rense data:

Deler opp tekstlinjer i flere kolonner.
Omstrukturerer og omdøper kolonner til spesifiserte navn.
Konverterer tekstverdier til numeriske verdier der det er nødvendig.
Filtrerer ut rader hvor "Beløp" er 0.
Beregner rabatt basert på antall og enhetspris.
Legger til rekvisisjonsnummer som en ny kolonne.
Velge PDF-filer: Bruker en filvelger-dialog for å la brukeren velge en eller flere PDF-filer.

Lagre data til Excel:

Bruker en filvelger-dialog for å la brukeren velge lagringssted for den nye Excel-filen.
Skriver hver DataFrame til et eget ark i Excel-filen.
