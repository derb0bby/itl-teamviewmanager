# Team Views Manager Excel Add-in

Team Views Manager ist ein Excel Office Add-in, das Teams dabei hilft, effizient zwischen verschiedenen Spaltenansichten in Excel-Arbeitsmappen zu wechseln. Es ermöglicht verschiedenen Teams, ihre bevorzugten Spaltenkonfigurationen zu speichern und schnell zwischen ihnen zu wechseln, wodurch sie sich besser auf die für ihre spezifischen Teamanforderungen relevanten Daten konzentrieren können.

## Funktionen

- **Vorkonfigurierte Team-Ansichten**: Schnellzugriffsschaltflächen für verschiedene Teams (Lager, SM, Technik)
- **Benutzerdefinierte Ansichtskonfiguration**: Jedes Team kann mehrere Arbeitsblattansichten mit spezifischen sichtbaren Spalten konfigurieren
- **Dauerhafte Einstellungen**: Konfigurationen werden direkt in der Arbeitsmappe gespeichert
- **Mehrblatt-Unterstützung**: Konfiguration von Ansichten für mehrere Arbeitsblätter innerhalb derselben Arbeitsmappe
- **Eingabevalidierung**: Integrierte Validierung für Excel-Spaltenreferenzen
- **Benutzerfreundliche Oberfläche**: Einfaches dialogbasiertes Konfigurationssystem

## Systemvoraussetzungen

- Excel (Office 365, Excel 2016 oder neuer)
- Arbeitsmappe muss auf SharePoint oder OneDrive gespeichert sein für die Funktionalität benannter Ansichten
- Internetverbindung für das Laden des Add-ins

## Installation

1. Laden Sie die Add-in-Manifestdatei herunter
2. Öffnen Sie in Excel `Einfügen > Office Add-ins`
3. Wählen Sie `Mein Add-in hochladen` und wählen Sie die Manifestdatei aus
4. Der Team Views Manager erscheint in Ihrem Excel-Menüband

## Verwendung

### Grundlegende Nutzung

1. Klicken Sie auf eine der Team-Schaltflächen (Team Lager, Team SM, Team Technik), um die jeweilige Team-Ansicht anzuwenden
2. Das Add-in blendet alle Spalten aus, außer denen, die für das ausgewählte Team konfiguriert wurden
3. Die Ansicht jedes Teams wird als benannte Ansicht in Excel gespeichert, um einfaches Umschalten zu ermöglichen

### Ansichten Konfigurieren

1. Klicken Sie auf die Schaltfläche "Einstellungen"
2. Wählen Sie das Team aus dem Dropdown-Menü
3. Für jedes Arbeitsblatt, das Sie konfigurieren möchten:
   - Wählen Sie das Arbeitsblatt aus dem Dropdown-Menü
   - Geben Sie die sichtbaren Spalten mit Excel-Spaltenbuchstaben ein (z.B. "A,C,E")
   - Fügen Sie weitere Arbeitsblatt-Konfigurationen über "Arbeitsblattansicht hinzufügen" hinzu
4. Klicken Sie auf "Speichern", um die Konfiguration zu sichern

### Spaltenformat

- Verwenden Sie kommagetrennte Spaltenbuchstaben (z.B. "A,B,C" oder "A,AA,BC")
- Unterstützt ein- und mehrbuchstabige Spaltenreferenzen (bis zu Excels Maximum XFD)
- Groß-/Kleinschreibung wird nicht beachtet (wird automatisch in Großbuchstaben umgewandelt)
- Leerzeichen werden automatisch entfernt

## Technische Details

### Speicherung

- Konfigurationen werden in den benutzerdefinierten Eigenschaften der Arbeitsmappe gespeichert
- Maximale Konfigurationsgröße beträgt 5MB
- Jedes Team kann mehrere Arbeitsblatt-Konfigurationen haben
- Konfigurationen bleiben mit der Arbeitsmappe erhalten

### Ansichtsverwaltung

- Erstellt oder aktualisiert benannte Blattansichten für jedes Team
- Ansichten beginnen mit "TVM_" gefolgt vom Teamnamen
- Ausgeblendeter/sichtbarer Status wird pro Spalte verwaltet
- Änderungen werden beim Wechsel der Ansichten sofort angewendet

## Einschränkungen

- Benötigt SharePoint- oder OneDrive-Speicherung für die Funktionalität benannter Ansichten
- Maximale Spaltenreferenz ist XFD (Excel-Limit)
- Konfigurationsspeicher ist auf 5MB pro Arbeitsmappe begrenzt

## Fehlerbehandlung

Das Add-in enthält umfassende Fehlerbehandlung für:

- Ungültige Spaltenreferenzen
- Speicherbegrenzungen
- Netzwerkverbindungsprobleme
- Fehlende Konfigurationen
- Ungültige Arbeitsblatt-Referenzen

## Support

Bei Problemen oder Fragen öffnen Sie bitte ein Issue im GitHub-Repository.

## Mitwirken

Beiträge sind willkommen! Reichen Sie gerne einen Pull Request ein.

## Lizenz

Dieses Projekt ist unter der MIT Lizenz lizenziert - siehe die [LICENSE.md](LICENSE.md) Datei für Details.