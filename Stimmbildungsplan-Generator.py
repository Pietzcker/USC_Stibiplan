# Input: Reporter-Abfrage "Gesamtliste Stimmbildung (Vorlage für Übersichtsplan)"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import xlsxwriter
import datetime

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d")

abk = {"Sopran 1": "S1", "Sopran 2": "S2", "Alt": "A",
       "Vorchor I": "V1", "Vorchor II": "V2", "Kinderchor": "K", "Jugendchor": "J",
       "Schnupperer Vorchor II": "sV2", "Schnupperer Kinderchor": "sK",
       "Schnupperer Jugendchor": "sJ"}

form_header = {"bold": True, "align": "center", "valign": "top", "size": 12, "border": 2}
form_zelle_weiss = {"align": "left", "valign": "top", "size": 10, "text_wrap": True, "border": 1, "bg_color": "#FFFFFF"}
form_zelle_grau =  {"align": "left", "valign": "top", "size": 10, "text_wrap": True, "border": 1, "bg_color": "#DDDDDD"}
form_li_rand = {"left": 2, "bottom": 1, "align": "left", "valign": "vcenter", "size": 10}
form_re_rand_weiss = {"right": 2, "bottom": 1, "align": "left", "valign": "top", "size": 10, "text_wrap": True, "bg_color": "#FFFFFF"}
form_re_rand_grau = {"right": 2, "bottom": 1, "align": "left", "valign": "top", "size": 10, "text_wrap": True, "bg_color": "#DDDDDD"}
form_ob_rand = {"top": 2}

print("Bitte Reporter-Abfrage 'Gesamtliste Stimmbildung (Vorlage für Übersichtsplan)'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("lfd. Nr.\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

with io.StringIO(data) as infile:
    daten = list(csv.DictReader(infile, delimiter="\t"))

spatzen = {}
stibis = {}
namen = set()

for eintrag in daten:
    if "." in eintrag["Uhrzeit"]: print(eintrag)
    spatzen.setdefault(eintrag["Wochentag"], {}).setdefault(
                       eintrag["Uhrzeit"].replace(".", ":"), {}).setdefault(
                       eintrag["Stimmbildner(in)"],[]).append(
                       f'{eintrag["Vorname"]} {eintrag["Name"]} ({abk.get(eintrag["Stimme"])} {abk.get(eintrag["Bereich"])})')
    stibis.setdefault(eintrag["Stimmbildner(in)"],{}).setdefault(
                      eintrag["Wochentag"],set()).add(eintrag["Raum"])
    namen.add(eintrag["Stimmbildner(in)"])

namen = sorted(list(namen-{""}))

max_breite = {name: 0 for name in namen}
for wochentag in spatzen:
    for uhrzeit in spatzen[wochentag]:
        for stibi in spatzen[wochentag][uhrzeit]:
            if stibi:
                for eintrag in spatzen[wochentag][uhrzeit][stibi]:
                    if len(eintrag) > max_breite[stibi]:
                        max_breite[stibi] = len(eintrag)

with xlsxwriter.Workbook(f"stibi_{heute}.xlsx") as outxlsx:
    excel = outxlsx.add_worksheet("Stibi-Plan")
    excel.set_paper(9) # DIN A4
    excel.set_landscape()
    excel.set_margins(0.3, 0.3, 0.6, 0.6) # Seitenränder in Zoll
    for spalte, eintrag in enumerate(max_breite.items()):
        excel.set_column(spalte+1, spalte+1, eintrag[1]*0.85)
    excel.set_column(0, 0, 10) # Spalte Datum/Uhrzeit
    excel.fit_to_pages(1,0) # An Seitenbreite anpassen
    excel.set_header("&CStimmbildungs-Plan Ulmer Spatzen Chor, Stand: &D")
    format_header = outxlsx.add_format(form_header)
    format_zelle_weiss = outxlsx.add_format(form_zelle_weiss)
    format_zelle_grau = outxlsx.add_format(form_zelle_grau)
    format_li_rand = outxlsx.add_format(form_li_rand)
    format_re_rand_weiss = outxlsx.add_format(form_re_rand_weiss)
    format_re_rand_grau = outxlsx.add_format(form_re_rand_grau)
    format_ob_rand = outxlsx.add_format(form_ob_rand)
    zeile = 0
    umbrüche = []
    for wochentag in spatzen:
        if not wochentag: 
            continue
        datum = wochentag
# Überschriften
        header = {datum: datum}
        for name in namen:
            raum = stibis[name].get(wochentag,[])
            if len(raum) > 1: 
                print(f"\nWarnung: Mehrere Räume am {wochentag} für {name}: {raum}")
                raum = [input("Bitte eingeben, welcher Raum verwendet werden soll: ")] 
            header[name] = f"{name} ({list(raum)[0]})" if raum else name
        for spalte, item in enumerate(header.items()):
            excel.write_string(zeile, spalte, item[1], format_header)
        zeile += 1
# Termine
        zeiten = sorted(spatzen[wochentag].keys())
        for uhrzeit in zeiten:
            for stimmbildner in spatzen[wochentag][uhrzeit]:
                spatzen[wochentag][uhrzeit][stimmbildner] = "\n".join(spatzen[wochentag][uhrzeit][stimmbildner]) + "\n"
            d = {datum: uhrzeit}
            d.update(spatzen[wochentag][uhrzeit])
            letzte_spalte = len(namen)
            for spalte, item in enumerate([datum]+namen):
                contents = d.get(item,"")
                if spalte == 0:
                    format = format_li_rand
                elif spalte == letzte_spalte:
                    if contents:
                        format = format_re_rand_weiss
                    else:
                        format = format_re_rand_grau
                else:
                    if contents:
                        format = format_zelle_weiss
                    else:
                        format = format_zelle_grau
                excel.write_string(zeile, spalte, contents, format)
            zeile += 1
        for spalte in range(len(namen)+1):
            excel.write(zeile, spalte, None, format_ob_rand) 
        zeile += 1
        umbrüche.append(zeile)
    excel.write_string(zeile, 0, "Keine Stibi:")
    zeile += 1
    for uhrzeit in spatzen[""]:
        for stibi in spatzen[""][uhrzeit]:
            if uhrzeit or stibi: 
                excel.write_string(zeile, 0, f"{uhrzeit} {stibi}")
                zeile += 1
            for spatz in spatzen[""][uhrzeit][stibi]:
                if spatz != "  (None None)":
                    excel.write_string(zeile, 0, f"{spatz}")
                    zeile += 1
    excel.set_h_pagebreaks(umbrüche)
    
print(f"Fertig! Die Datei stibi_{heute}.xlsx wurde im aktuellen Ordner abgelegt.")
