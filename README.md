# Traueranzeigen Downloader

Ein Desktop-Tool (PySide6), das Traueranzeigen von mehreren Portalen durchsucht und die Ergebnisse als **XLSX / JSON / CSV / XML** exportiert. Optional können Anzeigenbilder heruntergeladen und per **OCR (Tesseract)** ausgelesen werden.

## Features

- **Input:** URL *oder* Name (ohne `http`)
  - **URL-Modus:** Crawlt ab der angegebenen URL
  - **Name-Modus:** Sucht auf allen unterstützten Portalen über `/traueranzeigen-suche/<slug>`
- **Modi:** Bilder / Daten / Bilder + Daten
- **Export:** XLSX / JSON / CSV / XML (URL immer letzte Spalte / Key)
- **Lokale Suche/Filter** in der Tabelle
- **Max. Personen:** stoppt automatisch
- **OCR (manuell)** für markierte Zeilen → schreibt Text in **„Zusatzinformationen“**
- **Shortcuts:** `Entf`, `Ctrl+S`, `Ctrl+Q`, `Ctrl+A`, `Ctrl+O`
- **Sortierung:** Klick auf Spaltenkopf (A–Z / Z–A), inkl. Sort-Key (PLZ/Datum sauber)
- **Kein Duplikat-Filter** (es kann mehrere Anzeigen pro Person geben)

## Unterstützte Portale

- `trauer-anzeigen.de`
- `abschied-nehmen.de` / `www.abschied-nehmen.de`
- `ok-trauer.de` / `www.ok-trauer.de`
- `gedenken.freiepresse.de`
- `vrm-trauer.de` / `www.vrm-trauer.de`

> Hinweis: Portale können HTML/CSS ändern. Der Parser arbeitet heuristisch und filtert „Junk“ (Livechat, Vorsorge, Aufgeben etc.) möglichst zuverlässig.

---

## Voraussetzungen

- **Python 3.10+** empfohlen
- Internetverbindung
- Optional für OCR:
  - **Tesseract OCR** (siehe Abschnitt „OCR“)

---

## Installation

### 1) Repository klonen

```bash
git clone <REPO-URL>
cd <REPO-ORDNER>
