# wpContacts-Session-Outlook-Add-In
Projekte wpContacts Session Outlook-Add-In

## Zweck
Das Add-In für Outlook 2016/2019 syncronisiert die Daten der Mandatare und Mitarbeiter in einen Outlook-Ordner im Standard-Postfach.
Damit stehen diese Daten nicht nur am Arbeitsplatz, sondern auch auf dem Laptop und Mobiltelefon zur Verfügung, auch wenn man keine Verbindung zur Session hat.

## Funktion/Ablauf
Die Daten werden erst aus dem Kontakteordner gelöscht um dann neu angelegt zu werden.
Für alle Gremien werden Verteilerlisten angelegt.

## Abhängigkeiten

1. Microsoft Exchange Web Services Managed API 2.2
https://www.microsoft.com/en-us/download/details.aspx?id=42951

2. log4net 
https://logging.apache.org/log4net/
via NUGET: reinstall log4net

3. WiX Toolset
https://github.com/wixtoolset/wix3/releases
Homepage: https://wixtoolset.org/
THE MOST POWERFUL SET OF TOOLS AVAILABLE TO CREATE YOUR WINDOWS INSTALLATION EXPERIENCE.

4. WIX-Plugin
https://marketplace.visualstudio.com/items?itemName=WixToolset.WixToolsetVisualStudio2019Extension

5. WAX (optional)
https://marketplace.visualstudio.com/items?itemName=TomEnglert.Wax

## Installation
Mit WiX Toolset wird ein MSI-Paket erzeugt und kann als Administrator installiert werden.

## Update und Deinstallation
Das MSI prüft ob eine Version installiert ist und aktualisiert diese.
Deinstalliert wird das Add-In über die Systemsteuerung.

## Anmerkungen

* wpconfig
Es ist möglich eine Konfiguration auf eine zentrale Share abzulegen. 
Der Name der Datei muss wpContacts-Session-Outlook-Addin.xml sein.
Mit einer **Umgebungsvariable** kann eingestellt werden auf welchem Pfad diese Datei zu finden ist.

* Das Löschen der Kontakte dauert länger als notwendig, da es keine Möglichkeit gibt alle Kontakte auf einmal zu löschen.

* Der Kontakteordner wird als "Outlook-Adressbuch" angelegt.


![Hauptmenu](./Bilder/wpContacts_Hauptmenu.png)
Mehr Bilder: [hier](/Bilder)
