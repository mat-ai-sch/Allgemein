Die aktuelle Arbeitsmappe ist **StatBoardReakt_2026_Dashboard_v15_ohne_Makros_geschuetzt.xlsx**.

In GitHub die Datei öffnen und **Download raw file** wählen. Das bisherige Öffnungskennwort gilt weiter. Es steht ausschließlich im privaten Chat.

Die Datei im Jahresordner 2026 neben **StatReaktImport.csv** speichern. Zum Einlesen der CSV in Excel Desktop **Daten > Alle aktualisieren** wählen. Die Aktualisierung und die anschließende Berechnung vollständig abwarten. Unter **Formeln > Berechnungsoptionen** muss **Automatisch** gewählt sein.

Version 15 bleibt eine **makrofreie XLSX-Datei**. Sie enthält keinen VBA-Code, keine Makro-Schaltflächen und keine zusätzliche Maussteuerung. Makros müssen nicht aktiviert werden. Filter, Formeln, Diagramme und die CSV-Verbindung bleiben erhalten. Die CSV wird weiterhin über **Daten > Alle aktualisieren** eingelesen.

Im Diagramm **Durchschnittliche Dauer nach Monat** stehen die Tageswerte direkt an der Linie, jeweils mit zwei Nachkommastellen und **T**, zum Beispiel **2,34 T**. Die Grafik ist flacher. Die Tagesachse beginnt bei null; Obergrenze und Schrittweite passen sich weiterhin der Auswahl an. Ihre Zahlen werden mit der nötigen Genauigkeit angezeigt, damit auch kleine Zeitwerte unterscheidbar bleiben. Fehlende Monate erhalten keinen Nullwert.

Die Jahresübersicht behält die Aufteilung auf zwei Seiten: Der Zeitverlauf steht auf Seite 1; Fallzahlen und die fünf Zeitspannen stehen auf Seite 2. „Fälle zuordnen nach“ bestimmt die Zuordnung zum Jahr und zum Monat. Die Erklärung nennt die Zeit, innerhalb derer die Hälfte der Fälle bearbeitet wurde. Für längere Bearbeitungszeiten wird die tatsächliche Fallzahl angegeben. Das vermeidet pauschale Prozentangaben bei kleinen Fallzahlen oder gleichen Zeitwerten. Vidierung und tatsächliche Entlassung sowie die Dezemberfälle im Januarbericht sind kürzer erklärt. Der kurze Datenhinweis steht ab Zeile 88. Rohdaten, Berechnungsmodelle, Verbindungen und Druckbereiche bleiben unverändert.

Die Diagramme behalten sichtbare Achsen, ungeglättete Linien und Lücken bei fehlenden Monatsdaten. Fallbalken zeigen Zahlen direkt an. Zum Ablesen einzelner Datenpunkte dienen die normalen Excel-Hinweise; alle Monatswerte stehen außerdem im Blatt **Monatsvergleich**.

Die **Jahresübersicht** enthält die gemeinsamen Filter. Blaue Felder mit ▾ enthalten Auswahlmenüs. Die Vergleichsgrenze wird in Stunden eingegeben; der Tageswert steht daneben. **Monatsvergleich** und **Fachabteilungen** übernehmen die Filter. Über **Filter ändern** gelangt man zurück zur Jahresübersicht.

**Jahresvergleich** verwendet die gewählte Fachabteilung und Zeitspanne, ordnet Fälle aber stets nach dem Aufnahmedatum zu. Vergleichsjahr, Monatsfenster und Bestätigung der vollständigen Quellen werden dort getrennt gewählt. Ohne freigegebenen Vergleich erscheint keine Vergleichsreihe.

Die zwölf **Monatsberichte** verwenden das CSV-Quelljahr, das Aufnahmedatum und alle Fachabteilungen. Ihre feste Reaktionszeit bedeutet **Anforderung bis Durchführung**. Der Januarbericht enthält zusätzlich die im Dezember des Vorjahres aufgenommenen Fälle mit Vidierung im Januar. Bereits im Dezember abgeschlossene Fälle zählen dort nicht mit.

**Datenqualität** zeigt den Weg von CSV-Zeilen über eindeutige Fallnummern und die gewählte Auswahl bis zu den ausgewerteten Zeitwerten. Pro Fallnummer zählt die erste CSV-Zeile. Unterschiedliche Fallzahlen ergeben sich aus Mehrfachzeilen, Jahres- und Fachabteilungsfiltern sowie fehlenden oder negativen Zeitangaben. Ein unverändertes Ergebnis allein belegt weder eine erfolgreiche noch eine fehlgeschlagene Aktualisierung. Die Prüfschritte stehen im Blatt **Hilfe**.

Die sechs Ansichten haben jeweils **zwei DIN-A4-Seiten im Hochformat**. Jeder Monatsbericht hat **drei Seiten**: Monatsübersicht, Zeitspannen nach Fachabteilung und Verteilung der Reaktionszeiten. Manuelle Seitenumbrüche trennen diese Inhalte. Das vollständige Neuwerk-Logo steht in den Kopfzeilen und wird auf Folgeseiten wiederholt.

Zeitwerte stehen in **Stunden und Tagen**. Ein Tag bedeutet 24 Stunden, einschließlich Nächten, Wochenenden und Feiertagen. Tageswerte werden aus den ungerundeten Zeitwerten berechnet. Ein Strich bedeutet, dass kein auswertbarer Wert vorliegt. Der Abschluss bedeutet Vidierung; ein tatsächliches Entlassdatum liegt nicht vor.

Die Berechnung behält die Optimierungen aus Version 9. Die wiederholte vollständige Neuberechnung bei jeder Änderung bleibt ausgeschaltet; beim Öffnen wird vollständig berechnet. Nach einem Filterwechsel Werte erst ablesen, wenn Excel nicht mehr **Berechnen** anzeigt.

Die optionalen [Hilfsdateien für das Jahresarchiv](Jahresarchiv/Jahresarchiv_Bedienung.md) dienen zum Vergleich abgeschlossener Jahresquellen. Es sind zunächst keine historischen Daten geladen. Die Archivverbindung bleibt bis zur Einrichtung deaktiviert.

Für eine neue Jahresdatei die CSV-Verbindung auf den neuen Jahresordner einstellen, **CSV-Quelljahr** im Blatt **Hilfe** anpassen und das Auswertungsjahr in der Jahresübersicht wählen.

Im Repository liegt ausschließlich die verschlüsselte Arbeitsmappe mit allgemeinen Hilfsdateien. Die unverschlüsselte Arbeitsmappe, Rohdaten und erzeugte Archiv-CSV bleiben auf dem privaten Abteilungslaufwerk.
