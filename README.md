ReadMe

#1-28
Hente Excel-fil med PDF-links: 
python "C:\Users\Michala Hammer\Documents\Michala\Programmering\PDF downloader\Programmering\Kode\Main.py"

#29-
Download af nye PDF-filer:
Programmet henter links fra kolonne D og bruger filnavnet fra kolonne C.
PDF-filer gemmes i undermappen PDFs. BR-numre, der allerede findes i
PDF_links_log.json, springes over og bliver ikke hentet igen. Links, der
allerede findes i JSON-loggen, springes også over, selv hvis BR-nummeret er nyt.
Hvis et link i kolonne D fejler (for eksempel HTTP 404 eller timeout), gemmes
fejlen i error_log.json, og programmet prøver reserve-linket i kolonne E.
Hvis reserve-linket virker, gemmes PDF'en i PDFs, og reserve-linket registreres
i PDF_links_log.json.
Fejlende links registreres kun i error_log.json, ikke i PDF_links_log.json.

# Backup og gendannelse
PDF_links_backup.json opdateres, når PDF_links_log.json opdateres. Den indeholder
de hentede PDF-links og filoplysninger. Backupfejl registreres i
error_backup.json, uden at stoppe de øvrige downloads.

Gendan loggen og hent manglende PDF-filer fra backup med:
python "C:\Users\Michala Hammer\Documents\Michala\Programmering\PDF downloader\Programmering\Kode\Main.py" --reset
Resetresultater, herunder forsøgte links og årsagen til filer, der springes over
eller fejler, registreres i reset_log.json.
