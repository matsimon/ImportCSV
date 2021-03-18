# ImportCSV
Evento-Export-CSV ==> Absenzensystem-Import-CSV

Das Skript kann über die Konsole (CMD) gestartet werden. Im selben Ordner müssen sich das Evento-Export-CSV (11-02-21-Klassenlisten alle Stufen.CSV als Beispiel in der Anlage) und die Datei mit den Testusern (testuser.csv in der Anlage) befinden.
Das Evento-Export-CSV erhält man, indem man das xlsx aus Evento in Excel als CSV exportiert (==> Datei, Exportieren, Dateityp, CSV). Ändert sich die Spaltenreihenfolge oder -beschriftung im Evento-File, so muss das Skript geringfügig überarbeitet werden. Das Skript erkennt das aber und promptet eine entsprechende Meldung.

Das Skript kann in der CMD gemäss dem Screenshot angekickt werden. Das zweimalige "dir" im Screenshot dient nur dazu, dass man sieht, dass sich danach eine zusätzliche Datei im Verzeichnis befindet (das Absenzensystem-Import-CSV).

Das Skript (import.Skript.py in der Anlage) habe ich ausführlich dokumentiert. Die wichtigen Grössen können im Kopf der Datei verändert werden. Die Sonderzeichenliste füllt die Zeilen 23-52, die Logik zum Erstellen der SuS-Mailadressen die Zeilen 113-140 (inkl. erklärendem Kommentar + Beispiele).

Damit python von der Konsole aus läuft, muss es a) installiert sein (Version 3!) und b) die Systemvariable PATH richtig gesetzt sein. Dazu gibt's Anleitungen en masse auf dem Internet.
Z.B: https://praxistipps.chip.de/python-in-cmd-nutzen-so-gehts_96172#:~:text=W%C3%A4hlen%20Sie%20bei%20den%20Benutzervariablen,Python%5CPython36%22)%20an.
