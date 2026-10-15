Die aktuelle Arbeitsmappe ist **StatBoardReakt_2026_Dashboard_v19_ohne_Makros_geschuetzt.xlsx**.

In GitHub die Datei öffnen und **Download raw file** wählen. Das bisherige Öffnungskennwort gilt weiter. Es steht ausschließlich im privaten Chat.

Die Datei im Jahresordner 2026 neben **StatReaktImport.csv** speichern. Zum Einlesen der CSV in Excel Desktop **Daten > Alle aktualisieren** wählen. Die Aktualisierung und die anschließende Berechnung vollständig abwarten. Unter **Formeln > Berechnungsoptionen** muss **Automatisch** gewählt sein.

Version 19 bleibt eine **makrofreie XLSX-Datei**. Sie enthält keinen VBA-Code, keine Makro-Schaltflächen und keine zusätzliche Maussteuerung. Makros müssen nicht aktiviert werden. Filter, Formeln, Diagramme und die CSV-Verbindung bleiben erhalten. Die CSV wird weiterhin über **Daten > Alle aktualisieren** eingelesen.

**Monatsvergleich**, **Fachabteilungen** und **Jahresvergleich** folgen dem Raster und der Gestaltung der Jahresübersicht: gleiche Schriften und Bezeichnungen, ruhigere Diagramme, klare Filterfelder und feine Trennlinien zwischen Tabellenzeilen. Zweizeilige Zeitwerte stehen in 34 Punkt hohen Datensätzen. Monatsnamen werden ausgeschrieben; die Jahre sind eindeutig benannt. Der Monatsverlauf verwendet dieselbe Grafik wie die Jahresübersicht. Für Fachabteilungen bleiben die horizontalen Balken mit Tageswerten am Balkenende erhalten.

Der **Jahresvergleich** hat eine eigene Monatswertetabelle für beide Jahre mit Stunden, Tagen und deren Unterschied. Sie folgt denselben freigegebenen Werten wie die beiden Jahreslinien. Fehlende oder gesperrte Vergleichswerte erscheinen als Strich. Drei Druckseiten trennen Diagramm und Kennzahlen, Monatswerte sowie Jahreskennzahlen und Quellen.

Auf Seite 2 der **Jahresübersicht** haben die Tabellenzeilen jetzt 28 statt 20 Punkt Höhe. Größere Abstände trennen Diagramm, Tabelle und Erklärungen. Die beiden Erklärungstexte enthalten jeweils eine Leerzeile; der Januarhinweis erhält mehr Platz. Der kurze Hinweis am Seitenende nutzt weniger Leerraum. Die zusätzliche Achsenbeschriftung im Fallzahlendiagramm entfällt, damit sie die Skalenwerte nicht überdeckt. Die erste Seite behält ihre bisherige Darstellung.

Im Diagramm **Durchschnittliche Dauer nach Monat** steht der Tageswert direkt unter dem jeweiligen Monatsnamen, jeweils mit zwei Nachkommastellen und **T**, zum Beispiel **2,34 T**. Die Werte liegen damit außerhalb der Zeichenfläche. Die Einheit steht oben im Titel; der senkrechte Achsentitel entfällt. Die Tagesachse beginnt bei null; Obergrenze und Schrittweite passen sich weiterhin der Auswahl an. Fehlende Zeitwerte erscheinen als Strich und als Lücke in der Linie. Eine tatsächliche Dauer von null wird als **0,00 T** angezeigt.

Die Jahresübersicht behält die Aufteilung auf zwei Seiten: Der Zeitverlauf steht auf Seite 1; Fallzahlen und die fünf Zeitspannen stehen auf Seite 2. „Fälle zuordnen nach“ bestimmt die Zuordnung zum Jahr und zum Monat. Ein kurzer Hinweis erklärt die Durchschnittsdauer und die Zeitrechnung einschließlich Nächten, Wochenenden und Feiertagen. Die missverständliche Verteilungsangabe und ihre zusätzliche Zählformel wurden entfernt. Vidierung und tatsächliche Entlassung sowie die Dezemberfälle im Januarbericht sind kurz erklärt. Der kurze Datenhinweis steht ab Zeile 88. Rohdaten, statistische Berechnungen und CSV-Verbindungen bleiben erhalten. Die Jahresübersicht behält ihren Druckbereich. Zwölf Hilfszellen liefern ausschließlich die Monatsbeschriftungen.

Die Diagramme behalten sichtbare Achsen, ungeglättete Linien und Lücken bei fehlenden Monatsdaten. Fallbalken zeigen Zahlen direkt an. Zum Ablesen einzelner Datenpunkte dienen die normalen Excel-Hinweise; alle Monatswerte stehen außerdem im Blatt **Monatsvergleich**.

**Jahresübersicht**, **Monatsvergleich**, **Fachabteilungen** und **Jahresvergleich** haben denselben bedienbaren Auswahlbereich. Jahr, Datumszuordnung, Fachabteilung, Zeitspanne und Vergleichsgrenze lassen sich direkt im jeweiligen Blatt ändern. Jedes Blatt speichert seine eigene Auswahl. Blaue Felder mit ▾ enthalten Auswahlmenüs. Die Grenze wird in Stunden eingegeben; der Tageswert steht daneben. Der Monatsvergleich behält seine zwei Monatsfelder; Fachabteilungen zeigt alle Abteilungen und hebt die ausgewählte hervor.

Die sechs Navigationsreiter sind auf allen 18 sichtbaren Blättern über ihre gesamte Fläche anklickbar. Sie führen sowohl zu vorherigen als auch zu nachfolgenden Kategorien.

**Jahresvergleich** verwendet seine eigene Auswahl und dieselbe Datumszuordnung für beide Jahre. Die aktuelle CSV unterstützt die angebotenen Datumsgrundlagen. Historische Jahresarchive enthalten ausschließlich Aufnahmemonate. Wird bei einem Archiv eine andere Datumszuordnung gewählt, bleiben die Vergleichswerte leer; ein Hinweis fordert zur Auswahl von **Aufnahme** auf. Vergleichsjahr, Monatsfenster und Bestätigung der vollständigen Quellen werden weiterhin im Jahresvergleich gewählt. Ohne freigegebenen Vergleich erscheint keine Vergleichsreihe.

Die zwölf **Monatsberichte** verwenden das CSV-Quelljahr, das Aufnahmedatum und alle Fachabteilungen. Ihre feste Reaktionszeit bedeutet **Anforderung bis Durchführung**. Der Januarbericht enthält zusätzlich die im Dezember des Vorjahres aufgenommenen Fälle mit Vidierung im Januar. Bereits im Dezember abgeschlossene Fälle zählen dort nicht mit.

**Datenqualität** bezieht sich weiterhin auf die Auswahl der Jahresübersicht und zeigt den Weg von CSV-Zeilen über eindeutige Fallnummern und die gewählte Auswahl bis zu den ausgewerteten Zeitwerten. Pro Fallnummer zählt die erste CSV-Zeile. Unterschiedliche Fallzahlen ergeben sich aus Mehrfachzeilen, Jahres- und Fachabteilungsfiltern sowie fehlenden oder negativen Zeitangaben. Ein unverändertes Ergebnis allein belegt weder eine erfolgreiche noch eine fehlgeschlagene Aktualisierung. Die Prüfschritte stehen im Blatt **Hilfe**.

**Jahresübersicht**, **Monatsvergleich**, **Fachabteilungen**, **Datenqualität** und **Hilfe** haben jeweils **zwei DIN-A4-Seiten im Hochformat**. Der **Jahresvergleich** hat drei Seiten. Jeder Monatsbericht hat ebenfalls **drei Seiten**: Monatsübersicht, Zeitspannen nach Fachabteilung und Verteilung der Reaktionszeiten. Manuelle Seitenumbrüche trennen diese Inhalte. Das vollständige Neuwerk-Logo steht in den Kopfzeilen und wird auf Folgeseiten wiederholt.

Zeitwerte stehen in **Stunden und Tagen**. Ein Tag bedeutet 24 Stunden, einschließlich Nächten, Wochenenden und Feiertagen. Tageswerte werden aus den ungerundeten Zeitwerten berechnet. Ein Strich bedeutet, dass kein auswertbarer Wert vorliegt. Der Abschluss bedeutet Vidierung; ein tatsächliches Entlassdatum liegt nicht vor.

Die Berechnung behält die Optimierungen aus Version 9. Die drei zusätzlichen Filterbereiche verwenden nur die jeweils benötigten Berechnungen; der vollständige Berechnungsbereich wird dafür nicht dreimal kopiert. Die wiederholte vollständige Neuberechnung bei jeder Änderung bleibt ausgeschaltet; beim Öffnen wird vollständig berechnet. Nach einem Filterwechsel Werte erst ablesen, wenn Excel nicht mehr **Berechnen** anzeigt.

Die optionalen [Hilfsdateien für das Jahresarchiv](Jahresarchiv/Jahresarchiv_Bedienung.md) dienen zum Vergleich abgeschlossener Jahresquellen. Es sind zunächst keine historischen Daten geladen. Die Archivverbindung bleibt bis zur Einrichtung deaktiviert.

Für eine neue Jahresdatei die CSV-Verbindung auf den neuen Jahresordner einstellen, **CSV-Quelljahr** im Blatt **Hilfe** anpassen und das Auswertungsjahr in der Jahresübersicht wählen.

Im Repository liegt ausschließlich die verschlüsselte Arbeitsmappe mit allgemeinen Hilfsdateien. Die unverschlüsselte Arbeitsmappe, Rohdaten und erzeugte Archiv-CSV bleiben auf dem privaten Abteilungslaufwerk.
