import csv, datetime

# Das aus Evento exportierte Excel-Sheet muss in Excel per Export in ein CSV 
# umgewandelt werden (==> Exportieren, Dateityp, CSV). Die erwartete
# Spaltensortierung und -beschriftung findet sich in KOPFZEILE_CSV_EVENTO.

print("\n-------------------------------------------------------------------")
IMPORT_FILENAME = input("Name der Evento-CSV-Datei (mit Endung): ")

# -----------------------------------------------------------------------
#IMPORT_FILENAME = "11-02-21-Klassenlisten alle Stufen.csv"
TESTUSER_FILENAME = "testuser.csv"
EXPORT_FILENAME = str(datetime.date.today()) + "_Import_Absenzensystem.csv"
# Klassenname für Testuser und Austauschschüler*innen
AUSTAUSCH_SuS = "21v"
CODIERUNG_CSV_EVENTO = "ISO-8859-1"
# erwartete Kopfzeile Evento-Export
KOPFZEILE_CSV_EVENTO = ['ID Person', 'Anlassnummer', 'Geschlecht', 'Geburtsdatum', 'Nachname', 'Vornamen', 'Adresszeile 1', 'PLZ', 'Ort', 'Telefon', 'E-Mail 2', 'E-Mail 3', 'AD10015: SF', 'AD10010: S3', 'AD10020: KU', 'AD10025: EF', 'AD10030: (diverse)', 'AD10031: S+', 'AD10013: BI', 'AD10027: MINT', 'AD1382: Musikinstrument', 'Heimatort', 'GV: Vollständiger Name', 'Promotionsstatus (aktuell)', 'Status (Anmeldung)']
CODIERUNG_CSV_ABSENZENSYSTEM = "ISO-8859-1"
# Kopfzeile Absenzensystem-Import
KOPFZEILE_CSV_ABSENZENSYSTEM = ['ID Person', 'Anlassnummer', 'Geschlecht', 'Nachname', 'Vornamen', 'Adresszeile 1', 'PLZ', 'Ort', 'Telefon privat', 'SF', 'S3', 'KU', 'EF', 'PF5', 'S+', 'BI', 'MINT', 'AD1382: Musikinstrument', 'Staatsangehoerigkeit', 'GV: Vollstaendiger Name', 'Promotionsstatus', 'Geburtsdatum', 'UPN', 'MailNickname', 'GV: E-Mail', 'GV: E-Mail 2']
# Sonderzeichen mit Ersatzzeichen
ERSATZ = [["ä","ae"],
          ["ö","oe"],
          ["ü","ue"],
          ["å","a"],
          ["ã","a"],
          ["â","a"],
          ["á","a"],
          ["à","a"],
          ["ç","c"],
          ["é","e"],
          ["è","e"],
          ["ë","e"],
          ["ê","e"],
          ["î","i"],
          ["ì","i"],
          ["í","i"],
          ["ï","i"],
          ["ñ","n"],
          ["ó","o"],
          ["ô","o"],
          ["ò","o"],
          ["õ","o"],
          ["ø","o"],
          ["ß","ss"],
          ["š","s"],
          ["ś","s"],
          ["ŝ","s"],
          ["ý","y"],
          ["ž","z"],
          ["'",""]]
# -----------------------------------------------------------------------

print("")

# Liste zur Aufnahme des Evento-CSV
import_liste = []
with open(IMPORT_FILENAME, "r", encoding=CODIERUNG_CSV_EVENTO) as csvdatei:
    csv_reader_object = csv.reader(csvdatei, delimiter=';')
    for row in csv_reader_object:
        import_liste.append(row)

# Ist die Kopfzeile des Evento-Files gleich geblieben? Vergleich mit Vorlage.
# Skript bricht nicht ab, gibt aber eine Warnung aus.
if import_liste[0] != KOPFZEILE_CSV_EVENTO:
    print("Die Spaltentitel im Importfile haben sich verändert!")
    print("Das generierte csv könnte nicht den Anforderungen entsprechen!")

# Liste zur Aufnahme des CSV für den Absenzensystem-Import
export_liste = []
export_liste.append(KOPFZEILE_CSV_ABSENZENSYSTEM)

# Evento-Daten werden ohne Kopfzeile zeilenweise eingelesen. Jeder Datensatz muss 
# umsortiert werden, damit er für den Absenzensystem-Import taugt.
for i in range(1, len(import_liste)):
    # Leerzeilen in den Evento-Daten werden übergangen
    if import_liste[i][0] == "":
        print("Zeile",i+1,"des Importfiles ist eine Leerzeile und wurde ignoriert.")
        continue
    datensatz = [import_liste[i][0],   # 00: ID Person
                 import_liste[i][1],   # 01: Anlassnumer
                 import_liste[i][2],   # 02: Geschlecht
                 import_liste[i][4],   # 03: Nachname
                 import_liste[i][5],   # 04: Vorname
                 import_liste[i][6],   # 05: Adresszeile 1
                 import_liste[i][7],   # 06: PLZ
                 import_liste[i][8],   # 07: Ort
                 import_liste[i][9],   # 08: Telefon
                 import_liste[i][12],  # 09: SF
                 import_liste[i][13],  # 10: S3
                 import_liste[i][14],  # 11: KU
                 import_liste[i][15],  # 12: EF
                 import_liste[i][16],  # 13: PF5
                 import_liste[i][17],  # 14: S+
                 import_liste[i][18],  # 15: BI
                 import_liste[i][19],  # 16: MINT
                 import_liste[i][20],  # 17: Musikinstrument
                 import_liste[i][21],  # 18: Heimatort
                 import_liste[i][22],  # 19: GV: Vollständiger Name
                 import_liste[i][23],  # 20: Promotionsstatus (aktuell)
                 import_liste[i][3],   # 21: Geburtsdatum
                 "",                   # 22: UPN
                 "",                   # 23: MailNickname
                 import_liste[i][10],  # 24: GV: E-Mail
                 import_liste[i][11]]  # 25: GV: E-Mail 2

    # SuS im Austausch kommen in die Klasse AUSTAUSCH_SuS
    if import_liste[i][24] != "jA.Aufgenommen":
        datensatz[1] = AUSTAUSCH_SuS
  

    # Im Namen Gross- durch Kleinbuchstaben ersetzen
    nachname_neu = datensatz[3].lower()
    vorname_neu = datensatz[4].lower()

    # Sonderzeichen in Vor- und Nachname werden gemäss ERSATZ ersetzt
    for e in ERSATZ:
        nachname_neu = nachname_neu.replace(e[0], e[1])
        vorname_neu = vorname_neu.replace(e[0], e[1])

    # Nur Nachname: Lücke auf den ersten vier Positionen löschen.
    # ("von gunten" => "vongunten", "de weck" => "deweck" aber "hugi meier" => "hugi meier")
    x = nachname_neu.find(" ", 0, 4)
    if x != -1: # Nachname enthält (mind.) ein Leerzeichen auf Position 1 bis 4
        nachname_neu = nachname_neu[:x] +  nachname_neu[x+1:]

    # Vor- und Nachname: Name nach der ersten Lücke kappen ("hugi meier" => "hugi", "lou maria" => "lou")
    y = nachname_neu.find(" ")
    z = vorname_neu.find(" ")
    if y != -1:  # Nachname enthält (mind.) ein Leerzeichen
        nachname_neu = nachname_neu[:y]
    if z != -1: # Vorname enthält (mind.) ein Leerzeichen
        vorname_neu = vorname_neu[:z]
        
    # MailNickname und Mailadresse (=UPN) erstellen
    nickname = vorname_neu + "." + nachname_neu
    email = nickname + "@students.lerbermatt.ch"
    datensatz[23] = nickname
    datensatz[22] = email

    # SuS-Datensatz in die Export-Liste einfügen
    export_liste.append(datensatz)

# Testuser hinzufügen (diese liegen schon im Absenzensystem-Format vor und
# können der Export-Liste einfach angefügt werden).
with open(TESTUSER_FILENAME, "r", encoding=CODIERUNG_CSV_ABSENZENSYSTEM) as csvdatei:
    csv_reader_object = csv.reader(csvdatei, delimiter=';')
    for row in csv_reader_object:
        export_liste.append(row)

# CSV-Datei für den Absenzensystem-Import
with open(EXPORT_FILENAME, "w", newline="", encoding=CODIERUNG_CSV_ABSENZENSYSTEM) as csvdatei:
    writer = csv.writer(csvdatei, delimiter=";")
    for row in export_liste:
        writer.writerow(row)

print("CSV mit",len(export_liste),"Zeilen erstellt (davon eine Kopfzeile).")
print("-------------------------------------------------------------------\n")
