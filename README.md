# wpContacts-Session-Outlook-Add-In
Projekt wpContacts Session Outlook-Add-In

## Zweck
Das Add-In für Outlook 2016/2019 syncronisiert die Daten der Mandatare und Mitarbeiter in einen Outlook-Ordner im Standard-Postfach.

Damit stehen diese Daten nicht nur am Arbeitsplatz, sondern auch auf dem Laptop und Mobiltelefon zur Verfügung, auch wenn man keine Verbindung zur Session hat.

## Funktion/Ablauf
Die Daten werden erst aus dem Kontakteordner gelöscht um dann neu angelegt zu werden.

Für alle Gremien werden Verteilerlisten angelegt.

## Abhängigkeiten

1. log4net 
https://logging.apache.org/log4net/
via NUGET: reinstall log4net

2. WiX Toolset
https://github.com/wixtoolset/wix3/releases
Homepage: https://wixtoolset.org/
THE MOST POWERFUL SET OF TOOLS AVAILABLE TO CREATE YOUR WINDOWS INSTALLATION EXPERIENCE.

3. WIX-Plugin
https://marketplace.visualstudio.com/items?itemName=WixToolset.WixToolsetVisualStudio2019Extension

4. WAX (optional)
https://marketplace.visualstudio.com/items?itemName=TomEnglert.Wax

## Installation
Mit WiX Toolset wird ein MSI-Paket erzeugt und kann als Administrator installiert werden.
Den Pfad zur zentralen Konfiguration setzt man mit der Umgebungsvariable "wpconfig".

## Update und Deinstallation
Das MSI prüft ob eine Version installiert ist und aktualisiert diese.
Deinstalliert wird das Add-In über die Systemsteuerung.

## Anmerkungen

* wpconfig
Es ist möglich eine Konfiguration auf eine zentrale Share abzulegen. 
Der Name der Datei muss wpContacts-Session-Outlook-Addin.xml sein.
Mit einer **Umgebungsvariable** kann eingestellt werden auf welchem Pfad diese Datei zu finden ist.

* wpdebug
Wird die Umgebungsvariable wpdebug=1 ist das Logging etwas ausführlicher.

* Das Löschen der Kontakte dauert länger als notwendig, da es keine Möglichkeit gibt alle Kontakte auf einmal zu löschen.
* Der Kontakteordner wird als "Outlook-Adressbuch" angelegt.

Im EventLog findet man Infos zum Add-In und Ladezeiten.
https://support.microsoft.com/en-us/topic/outlook-application-event-log-entries-for-add-in-load-time-c3d39bb6-1488-250c-0b5a-ba3cf4b6f15e
https://docs.microsoft.com/de-at/archive/blogs/vsod/resolving-performance-issues-with-loading-office-add-ins-vsto-add-ins-or-shared-add-ins

## Änderungen

17.01.2021 Die 'Microsoft Exchange Web Services Managed API' wird nicht mehr benötigt.

## Bilder
![Hauptmenu](./Bilder/wpContacts_Hauptmenu.png)

Mehr Bilder: [hier](/Bilder)

## License
Copyright 2020-2021 Alexander Waller

Licensed under the AGPLv3: https://www.gnu.org/licenses/agpl-3.0.html

