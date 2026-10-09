# Jahresarchiv lokal anlegen

Das Dashboard enthält eine vorbereitete, zunächst deaktivierte CSV-Verbindung
`StatReaktJahresarchiv`. Der aktuelle Datenbestand liefert keine vollständigen
Vorjahre. Leere Archivdaten werden deshalb nicht als Nullleistung dargestellt.

1. In `Jahresarchiv_Quellen.json` für jedes gewünschte Aufnahme-Kalenderjahr
   **genau eine** endgültige Jahresquelle auswählen. Dateipfad und tatsächlichen
   Datenstand `stand_date` anpassen und `enabled` auf `true` setzen. Mehrere
   Zwischenstände desselben Jahres dürfen nicht zusammengeführt werden.
2. Das Windows-Skript lokal ausführen:
   `powershell.exe -NoProfile -File .\Jahresarchiv_erstellen.ps1 -ManifestPath .\Jahresarchiv_Quellen.json`
   Für XLSX/XLSM ist Excel Desktop erforderlich; CSV-Quellen werden ohne Excel
   gelesen. Falls Windows das heruntergeladene Skript blockiert, dessen
   Dateieigenschaften öffnen und bei Bedarf die Download-Blockierung aufheben.
3. Die erzeugte `Jahresarchiv.csv.pruefung.json` mit den Jahresfallzahlen,
   Monatsumfängen sowie gültigen, fehlenden und negativen Zeitpaaren prüfen.
4. Um das Archiv im Dashboard zu aktivieren, das Skript mit
   `-DashboardPath 'S:\...\StatBoardReakt_2026_Dashboard.xlsm'` ausführen.
   Es erstellt eine neue Datei `..._mit_Jahresarchiv.xlsm`, aktiviert deren
   Archivabfrage und liest die CSV ein. Die ursprüngliche Dashboard-Datei
   wird nicht geändert. Für einen eigenen neuen Dateinamen kann zusätzlich
   `-UpdatedDashboardPath 'S:\...\Dashboard_mit_Archiv.xlsm'` angegeben werden.
5. Künftig zuerst die CSV mit dem Skript erneuern, dann die neue Dashboard-Datei
   in Excel mit **Daten > Alle aktualisieren** aktualisieren. Excel liest aus
   der Archiv-CSV; das Exportskript übernimmt die ausgewählten Jahresdateien.

Kennwortgeschützte XLSX/XLSM werden mit einer verdeckten Kennworteingabe geöffnet.
Ein Kennwort wird weder im Manifest noch in Skript, CSV oder Prüfbericht abgelegt.
Das Skript öffnet Jahresquellen nur lesend mit deaktivierten Makros/Ereignissen
und ohne Aktualisierung externer Arbeitsmappenlinks. Es führt keinen VBA-Code aus.

Die Archiv-CSV enthält zwölf Spalten: `AufnahmeDatum` als ISO-Datum, `Jahr`,
`Monat`, `FA`, fünf Dauern `Dauer1_s` bis `Dauer5_s` in ganzen Sekunden,
`QuelleJahr`, `QuelleDatei` als Dateiname und `QuellstandDate` als ISO-Datum.
Fehlende oder negative Dauern bleiben leer. Negative Werte werden nicht auf
Null gesetzt. Die Kennzahlen sind Aufnahme→Anforderung,
Anforderung→Sichtung, Anforderung→Durchführung, Durchführung→Vidierung und
Anforderung→Vidierung.

Pro Jahresquelle bleibt die **erste vollständige Quellzeile je Fallnummer**
erhalten. Spätere Zeilen derselben Fallnummer werden verworfen, auch wenn sie
besser ausgefüllte Zeitstempel haben. Danach zählt nur das Aufnahme-Kalenderjahr
der ausgewählten Quelle. Identische ausgegebene Statistikzeilen bleiben erhalten:
verschiedene Fälle dürfen dieselben Daten und Dauern besitzen.

Das Skript erwartet ein Rohdatenimport-Blatt mit Überschriften in Zeile 12 und
den ursprünglichen A:P-Quellspalten ab Zeile 13. Für ältere, anders angeordnete
Dateien lassen sich `sheet`, `header_row`, `first_data_row` und `columns` im
Manifest ausdrücklich zuordnen. `columns` ordnet den ursprünglichen Feldnamen
Excel-Spaltenbuchstaben zu; bei CSV-Quellen ordnet es abweichende CSV-Feldnamen
zu. Alte XLS-Dateien vorher lokal als XLSX speichern. Ein passendes altes
Rohdatenschema ist erforderlich; beliebige Auswertungsblätter werden nicht als
Rohfälle rekonstruiert.

Patienten-/Mitarbeiternamen und Fallnummern werden nicht ausgegeben. Die
fallbezogenen Datums-, Fachabteilungs- und Dauerwerte bleiben sensible
Statistikdaten und werden lokal auf dem Abteilungslaufwerk gespeichert.
Das Skript lehnt eine Archiv-CSV innerhalb eines Git-Verzeichnisses ab.

Die anfänglich deaktivierte Abfrage lässt sich durch **Alle aktualisieren**
allein nicht aktivieren. Der optionale Dashboard-Schritt setzt nativ
`QueryTable.EnableRefresh=True`. Die Linux-Prüfung kann Excel Desktop und den
tatsächlichen Zugriff auf das S:-Laufwerk nicht ausführen.
