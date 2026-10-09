# Excel-Makro: Datum der Maßnahmenstände ändern

`FMEA_Massnahmenstand_Datum.bas` setzt in einem APIS IQ-FMEA XML-Export das Datum **aller Maßnahmenstände** auf ein eingegebenes Datum.

- Auf Wunsch (Ja/Nein-Abfrage) setzt es auch die **Termine der Maßnahmen in den Folgeständen** auf dieses Datum. Folgestände sind alle Stände nach dem ersten Stand einer Ursache.
- Die **Historie** der Maßnahmen wird nie verändert.
- Das Datum wird so gespeichert wie von APIS: 23:59:59 Uhr, Zeitzone Europe/Berlin. Sommer- und Winterzeit stellt das Makro automatisch ein.
- Außer den Datumswerten bleibt die Datei Byte für Byte gleich.
- Das Ergebnis wird als neue Datei neben dem Original gespeichert: `<Name>_JJJJ-MM-TT.xml`. Das Original bleibt unverändert.
- Ist die Datei anders aufgebaut als erwartet, bricht das Makro mit einer Meldung ab und schreibt keine Datei.

## Einrichten (einmalig)

1. Excel öffnen und eine neue, leere Arbeitsmappe anlegen.
2. Mit `Alt + F11` den VBA-Editor öffnen.
3. **Datei → Datei importieren…** wählen und `FMEA_Massnahmenstand_Datum.bas` auswählen.
4. Den VBA-Editor schließen und die Mappe als **Excel-Arbeitsmappe mit Makros (\*.xlsm)** speichern, z. B. `FMEA-Werkzeuge.xlsm`.
5. Optional eine Schaltfläche anlegen: **Entwicklertools → Einfügen → Schaltfläche**, dann das Makro `Massnahmenstaende_Datum_aendern` zuweisen.

## Benutzen

1. Die Mappe öffnen und Makros zulassen.
2. Mit `Alt + F8` das Makro **Massnahmenstaende_Datum_aendern** auswählen und **Ausführen** klicken, oder die Schaltfläche verwenden.
3. Den XML-Export auswählen.
4. Das neue Datum eingeben, z. B. `30.04.2027`.
5. Die Frage beantworten, ob auch die Termine der Maßnahmen in den Folgeständen gesetzt werden sollen.
6. Am Ende zeigt eine Zusammenfassung, was gesetzt wurde. Bei `dispensen_medien.xml` sind das 431 Maßnahmenstände (215 Anfang, 216 Folgestände) und, wenn gewählt, 484 Termine.

Bei großen Dateien kann das ein paar Sekunden dauern.

## Getestet

Den Kern des Makros habe ich mit LibreOffice Basic im VBA-Modus geprüft, mit `dispensen_medien.xml` und dem ursprünglichen Export. Dabei habe ich Sommer- und Winterzeit getestet, mit und ohne Termine, eine fehlerhafte Datei, eine fehlende Datei und Eingabedatei gleich Ausgabedatei. Jede Ergebnisdatei habe ich zusätzlich unabhängig gegen das Original verglichen.

Die Excel-Dialoge (Dateiauswahl, Eingabe, Meldungen) ließen sich dort nicht ausführen. Sie sind in Excel noch nicht getestet.
