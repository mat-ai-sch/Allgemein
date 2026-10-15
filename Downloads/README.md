Die aktuelle Arbeitsmappe ist **StatBoardReakt_2026_Dashboard_v13_geschuetzt.xlsm**.

In GitHub die Datei öffnen und **Download raw file** wählen. Das bisherige Öffnungskennwort gilt weiter. Es steht ausschließlich im privaten Chat.

Die Datei im Jahresordner 2026 neben **StatReaktImport.csv** speichern. Zum Einlesen der CSV in Excel Desktop **Daten > Alle aktualisieren** wählen. Die Aktualisierung und die anschließende Berechnung vollständig abwarten. Unter **Formeln > Berechnungsoptionen** muss **Automatisch** gewählt sein.

Version 13 ergänzt die Maussteuerung in den beiden Diagrammen der Jahresübersicht und im Diagramm des Monatsvergleichs. Dafür die Datei in **Excel Desktop** öffnen und die Makros aktivieren. Neben der Diagrammüberschrift **Monatswerte ansehen** wählen oder das Diagramm einmal anklicken. Anschließend die Maus innerhalb der Zeichenfläche seitlich bewegen. Eine senkrechte Linie und eine Wertekarte ordnen die Mausposition dem nächsten Monat zu. Die Karte enthält den ausgeschriebenen Monat, das Jahr, die gewählte Zeitspanne, Stunden und Tage sowie die ausgewerteten und gesamten Fälle. Während einer laufenden Berechnung erscheinen keine alten Zahlen als neues Ergebnis.

Die Maussteuerung liest ausschließlich vorhandene Monatsergebnisse. Sie schreibt keine Datenzellen und fordert keine Neuberechnung an. Fehlende Monatsdaten, eine Auswahl ohne Fälle und fehlende Zeitwerte werden getrennt benannt. Die Startschaltflächen sind als nicht druckend eingestellt; temporäre Karten und Führungslinien werden vor dem Drucken und Speichern entfernt. Änderungen an diesen Anzeigeobjekten können eine Speichernachfrage auslösen.

Die Makroquellen, ihre Einbindung und die unveränderte Datenbasis wurden geprüft. **Kompilierung und Mausbedienung in Microsoft Excel Desktop wurden hier nicht ausgeführt.** Die zugehörige Desktop-Prüfung steht noch aus. Es ist ausschließlich neuer Anzeigecode eingebunden; der frühere automatische Importcode ist nicht enthalten. Die CSV wird weiterhin über **Daten > Alle aktualisieren** eingelesen.

Die Jahresübersicht behält die Aufteilung aus Version 12: Der größere Zeitverlauf steht auf Seite 1; Fallzahlen und die fünf Zeitspannen stehen auf Seite 2. „Fälle zuordnen nach“ bestimmt die Zuordnung zum Jahr und zum Monat. Median und 90-Prozent-Wert werden gemeinsam erklärt; die ersten 50 Prozent sind in den rund 90 Prozent bereits enthalten. Der kurze Datenhinweis steht ab Zeile 88. Rohdaten, Berechnungsmodelle, Verbindungen und Druckbereiche bleiben unverändert.

Die Diagramme behalten sichtbare Achsen, ungeglättete Linien und Lücken bei fehlenden Monatsdaten. Fallbalken zeigen Zahlen direkt an. Ohne aktivierte Makros bleiben die Diagramme und normalen Excel-Datenpunkthinweise verfügbar; die zusätzliche Monatsanzeige benötigt den neuen Ereigniscode.

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
