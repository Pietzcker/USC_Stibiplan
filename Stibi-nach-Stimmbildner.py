# Input: Reporter-Abfrage "Gesamtliste Stimmbildung (Vorlage für Übersichtsplan)"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import xlsxwriter
import datetime

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d")

form_header = {"bold": True, "align": "center", "valign": "top", "size": 12, "border": 2}
form_zelle_weiss = {"align": "left", "valign": "top", "size": 10, "text_wrap": True, "border": 1, "bg_color": "#FFFFFF"}
form_zelle_grau =  {"align": "left", "valign": "top", "size": 10, "text_wrap": True, "border": 1, "bg_color": "#DDDDDD"}
form_li_rand = {"left": 2, "bottom": 1, "align": "left", "valign": "vcenter", "size": 10}
form_re_rand_weiss = {"right": 2, "bottom": 1, "align": "left", "valign": "top", "size": 10, "text_wrap": True, "bg_color": "#FFFFFF"}
form_re_rand_grau = {"right": 2, "bottom": 1, "align": "left", "valign": "top", "size": 10, "text_wrap": True, "bg_color": "#DDDDDD"}
form_ob_rand = {"top": 2}

print("Bitte Reporter-Abfrage ")
print("'Liste Stimmbildung mit allen Telefonnummern und Mailadressen (Vorlage für Skript)'")
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

stibis = {}
namen = set()

for eintrag in daten:
    stibi = eintrag["Ausbilder"]
    spatz = "{}, {}".format(eintrag["Name"], eintrag["Vorname"])
    tel_spatz = eintrag["Telefon"]
    tel_eltern = eintrag["Eltern Telefon"]
    stibis.setdefault(stibi, {})
    stibis[stibi].setdefault(spatz, {})
    if "@" in tel_spatz:
        stibis[stibi][spatz].setdefault("Mail", set()).add(tel_spatz)
    elif tel_spatz.startswith("01"):
        stibis[stibi][spatz].setdefault("Mobil", set()).add(tel_spatz)
    else:
        stibis[stibi][spatz].setdefault("Telefon", set()).add(tel_spatz)

    if "@" in tel_eltern:
        stibis[stibi][spatz].setdefault("Mail Eltern", set()).add(tel_eltern)
    elif tel_eltern.startswith("01"):
        stibis[stibi][spatz].setdefault("Mobil Eltern", set()).add(tel_eltern)
    else:
        stibis[stibi][spatz].setdefault("Telefon Eltern", set()).add(tel_eltern)
    
    namen.add(eintrag["Ausbilder"])

namen = sorted(list(namen-{""}))

#input(stibis)
#sys.exit()

#max_breite = {name: 0 for name in namen}
#for wochentag in spatzen:
#    for uhrzeit in spatzen[wochentag]:
#        for stibi in spatzen[wochentag][uhrzeit]:
#            if stibi:
#                for eintrag in spatzen[wochentag][uhrzeit][stibi]:
#                    if len(eintrag) > max_breite[stibi]:
#                        max_breite[stibi] = len(eintrag)
breiten = (20, 16, 16, 35, 16, 16, 35)

with xlsxwriter.Workbook(f"stibi_tel_mail_{heute}.xlsx") as outxlsx:
    excel = outxlsx.add_worksheet("Stibi-Liste")
    excel.set_paper(9) # DIN A4
    excel.set_margins(0.3, 0.3, 0.6, 0.6) # Seitenränder in Zoll
    for spalte, eintrag in enumerate(breiten):
        excel.set_column(spalte, spalte, eintrag*0.85)
    #excel.set_column(0, 6, 20) # Spalte Datum/Uhrzeit
    excel.set_landscape()
    excel.fit_to_pages(1,0) # An Seitenbreite anpassen
    excel.set_header("&CStimmbildungs-Kontaktliste Ulmer Spatzen Chor, Stand: &D")
    format_header = outxlsx.add_format(form_header)
    format_zelle_weiss = outxlsx.add_format(form_zelle_weiss)
    format_zelle_grau = outxlsx.add_format(form_zelle_grau)
    format_li_rand = outxlsx.add_format(form_li_rand)
    format_re_rand_weiss = outxlsx.add_format(form_re_rand_weiss)
    format_re_rand_grau = outxlsx.add_format(form_re_rand_grau)
    format_ob_rand = outxlsx.add_format(form_ob_rand)
    zeile = 0
    umbrüche = []
    for stibi in stibis:
    # Überschriften
        for spalte, item in enumerate([stibi, "Telefon", "Mobil", "Mail", "Telefon Eltern", "Mobil Eltern", "Mail Eltern"]):
            excel.write_string(zeile, spalte, item, format_header)
        zeile += 1
        for spatz in stibis[stibi]:
            felder = [spatz]
            for eintrag in ("Telefon", "Mobil", "Mail", "Telefon Eltern", "Mobil Eltern", "Mail Eltern"):
                nummern = list(stibis[stibi][spatz].get(eintrag, [""]))
                felder.append("\n".join(nummern))
            for spalte, item in enumerate(felder):
                excel.write_string(zeile, spalte, item, format_zelle_weiss)
            zeile += 1
        umbrüche.append(zeile)
    excel.set_h_pagebreaks(umbrüche)
    
    # Arbeitsblatt nur mit Mailadressen (Eltern & Kinder)
    
    for stibi in stibis:
        excel = outxlsx.add_worksheet(stibi[:stibi.index(",")])  # Nur Nachname der Stibi als Worksheet-Name
        excel.set_paper(9) # DIN A4
        excel.set_margins(0.3, 0.3, 0.6, 0.6) # Seitenränder in Zoll
        excel.set_column(0, 0, 35)
        excel.set_header("&CStimmbildungs-Mailliste (Eltern und Kinder) Ulmer Spatzen Chor, Stand: &D")

        excel.write_string(0, 0, stibi, format_header)

        # Eine überspringen, damit nachher in Excel Auswahl mit Strg-A möglich wird
        excel.write_string(2, 0, "spatzenchor@ulm.de", format_zelle_weiss)

        zeile = 3 
        mails = set()
        for spatz in stibis[stibi]:
            adressen = list(stibis[stibi][spatz].get("Mail", [])) + list(stibis[stibi][spatz].get("Mail Eltern", []))
            for mail in adressen:
                mails.add(mail)
        for item in mails:
            excel.write_string(zeile, 0, item, format_zelle_weiss)
            zeile += 1

    
print(f"Fertig! Die Datei stibi_tel_mail_{heute}.xlsx wurde im aktuellen Ordner abgelegt.")
