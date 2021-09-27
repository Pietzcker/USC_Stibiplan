# USC_Stibiplan

Diese Skripte erzeugen einen "hübschen" Stimmbildungsplan aus ComMusic-Daten.

 * Das Skript "Stimmbildungsplan-Generator" erzeugt eine Terminübersicht, geordnet nach Wochentagen (Seiten), Stimmbildner:innen (Spalten) und Uhrzeiten (Zeilen).
 * Das Skript "Stibi-nach-Stimmbildner" erzeugt 
   * eine Übersichtsliste aller Schülerinnen und Schüler, geordnet nach Stimmbildner:innen (Seiten), Schüler:innen (Zeilen) und Kontaktdaten inkl. Eltern (Spalten)
   * je eine Liste aller Mailadressen der Eltern und Kinder pro Stimmbildner:in

### Wie funktioniert's? ###

* Starten der Abfrage "Gesamtliste Stimmbildung (Vorlage für Übersichtsplan)"
* Ziel der Abfrage: Zwischenablage
* Starten des Skripts "Stimmbildungsplan-Generator.py" bzw. "Stibi-nach-Stimmbildner.py"

### Was brauche ich? ###

* Installation von ComMusic
* Import der Abfrage in den ComMusic-Reporter
* Installation von Python 3.7 oder höher
* Installation der Module xlsxwriter und win32clipboard
