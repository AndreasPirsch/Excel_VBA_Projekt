# Excel_VBA_Projekt
1. Use-Case
Das vorliegende Tool dient der digitalen Verwaltung von Urlaubsanträgen in Excel. Ziel ist es, Prozesse zur Antragstellung, Konfliktprüfung und Verwaltung effizient zu gestalten und in einer nutzerfreundlichen Umgebung darzustellen.

2. Funktionen im Überblick
•	Automatische Konfliktprüfung (s. Code in Modul_UrlaubPruefen)
prüft bei jedem Antrag auf Überschneidungen mit bereits genehmigten Anträgen.

•	Farbliche Statusanzeige (s. Code in Modul_UrlaubPruefen)

o	Grün: Kein Konflikt

o	Rot: Überschneidung

o	Status basiert auf „offen“, „genehmigt“, „abgelehnt“


 ![grafik](https://github.com/user-attachments/assets/edc40549-3342-435b-b521-49a9d52e5eaa)

Abbildung 1: Überschneidungsprüfung mit farblicher Markierung 

•	Button „Neuer Antrag“ (s. Code in Modul_NeueAnfrageHinzufuegen)
erstellt automatisch eine neue Eingabezeile mit vorbereiteten Dropdowns.
•	Dropdown-Felder für Statusauswahl
Einheitliche Einträge zur Reduzierung von Fehlern.
•	Sortierung nach Startdatum 
Automatisch bei jeder Änderung: Übersicht bleibt erhalten.


![grafik](https://github.com/user-attachments/assets/78b99efa-58ac-418b-b649-34c6aee6a636)

Abbildung 2: Code aus Tabelle1(Urlaubsanträge)

•	Export als PDF (s. Code in Modul_ExportiereAlsPDF)
ermöglicht den einfachen Export der aktuellen Liste.
•	Echtzeitprüfung über Worksheet-Events
Jeder Eintrag wird direkt auf Gültigkeit geprüft – ohne zusätzliches Auslösen eines Makros.


![grafik](https://github.com/user-attachments/assets/1ddd0163-90cf-432f-890a-357cea7350a3)

Abbildung 3: Code aus Tabelle1(Urlaubsanträge)

•	Automatische E-Mail bei Statusänderung (s. Code in Modul_SendeStatusMail)
Sobald sich der Status eines Urlaubsantrags ändert (z. B. von offen zu genehmigt oder abgelehnt), wird automatisch eine E-Mail über Microsoft Outlook generiert. Diese informiert über den Antragsteller, den Zeitraum und den neuen Status.
Die Funktion ermöglicht eine direkte Kommunikation mit z. B. der Personalabteilung oder Teamleitung und reduziert manuellen Aufwand.

3. Fazit und Ausblick:
Das Urlaubsantrags-Tool bietet eine simple, automatisierte Lösung zur Verwaltung von Urlaubsanträgen innerhalb von Excel. Durch einfache Bedienung, Echtzeitprüfung und integrierten Export ist es sowohl für kleine Teams als auch für wachsende Strukturen flexibel einsetzbar.
Eine zukünftige Erweiterung könnte z. B. die Integration einer Übersicht aller Urlaubstage je Mitarbeiter oder ein Maximum an Urlaubern je Abteilung beinhalten.

4. Sonstiges:

Das Modul „Modul_UrlaubKonflikteMitHinweis“ wurde während der Entwicklung zu Testzwecken verwendet. Es bietet zusätzlich eine interaktive Rückmeldung per MsgBox bei erkannten Konflikten, wird jedoch im aktuellen Stand des Tools nicht verwendet. Eine Integration zur Unterstützung weniger erfahrener Nutzer wäre denkbar.
