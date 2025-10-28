# Parced ZUGFeRD Dateien
Läd ZUGFeRD dateien und parced xml dateien und läd einige Beträge in Excel. Läuft auf LinuX und Windows 11
Graphische Oberfläche wurde mit Qt aus Python PyQt6 erstellt.

## Funktion
Die Ausgabedatei im xltx format als Template wird aus einem base64 Binärstring erstellt, welche das Skrit get_binär_str.py erstellt.

## Nuitka
Die Ausführbaren Dateien wurden mit nuitka erstellt und stehen zum Download bereit

## Vorraussetzungen
Stelle sicher, dass Python 3 und alle Requirements installiert ist:
pip install -r requirements.txt

## Windowns Install
Die einlesen.zip herunterladen und mit dem [wininstaller](https://github.com/rootloewe/wininstaller) installieren auf Windows.

## Debian
mit rootrechten
dpkg -i einlesen.deb oder
apt-get install einlesen.deb
Alle anderen Disties gibts ein tar.gz zum Download

## Hinweis
Das Programm wurde speziell für Kundenanforderungen angepaßt, nicht garantiert, ob es auch für Ihre Zwecke dienlich ist. 
Kann gerne auf Anfrage angepaßt werden.

## Download
noch in bearbeitung
| Windows                                                                                                                                  | MacOS                                                                                                                                        | LinuX                                                                                                                                     |Debian                                                                                                                                    |SuSe                                                                                                                                      |Arch                                                                                                                                      |
|------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| <img src="./logos/windows.png" width="100" alt="Windows">                                                                                | <img src="./logos/apple.png" width="90" alt="MacOS">                                                                                         | <img src="./logos/linux.png" width="100" alt="Linux">                                                                                     |<img src="./logos/debian.png" width="100" alt="Linux">                                                                                    |<img src="./logos/suse.png" width="100" alt="Linux">                                                                                      |<img src="./logos/arch.png" width="100" alt="Linux">                                                                                      |
| [ZIP-Archiv](https://github.com/rootloewe/einlesen/releases/download/2.5.4/einlesen.zip)<br>(x86_64)                                     | [DMG-Archiv] <br>(Intel)                                                                                                                     | [TAR.GZ-Archiv](https://github.com/rootloewe/einlesen/releases/download/2.5.4/einlesen.tar.gz)<br>(x86_64)                                |[deb-Installer](https://github.com/rootloewe/einlesen/releases/download/2.5.4/einlesen.deb)<br>(x86_64)                                   |[rpm-Installer](https://github.com/rootloewe/einlesen/releases/download/2.5.4/einlesen-2.5.4-1.x86_64.rpm)<br>(x86_64)                    |[zst-Installer](https://github.com/rootloewe/einlesen/releases/download/2.5.4/einlesen-2.4.5-1-x86_64.pkg.tar.zst)<br>(x86_64)            |


## Arch Linux
erst die Abhängigkeiten installieren:
pacman -Syu qt6-base libxcb xcb-util xcb-util-wm xcb-util-image python-pyqt5 python-pyqt6-sip qt6-xcb-private-headers
pacman -U einlesen-2.5.4.pkg.tar.zst


## Suse
mit rootrechten
zypper install libgthread-2_0-0
rpm -i einlesen-2.5.4-1.x86_64.rpm oder
zypper install einlesen-2.5.4-1.x86_64.rpm

Download GPG Schlüssel: [GPG-KEY](https://github.com/rootloewe/einlesen/blob/master/GPG-KEY/RPM-GPG-KEY-einlesen.asc) 
rpm --import RPM-GPG-KEY-einlesen.asc


## Lizenz
GNU GENERAL PUBLIC LICENSE
