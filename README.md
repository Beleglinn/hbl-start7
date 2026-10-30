# HBL Start7 Scout

Holt alle Spielerstatistiken der **Opel HBL** inklusive **Handball Performance Index (HPI)** und dessen
Verlauf pro Spieltag und baut daraus

- `index.html` – handytaugliche Übersicht (Suche, Filter, Sortierung, Detailansicht mit HPI-Kurve)
- `hbl_start7.xlsx` – Excel mit den Blättern *Spieler*, *HPI-Verlauf*, *Torhüter*, *Tabelle & Gegner*, *Info*
- `players.json` – alle Daten maschinenlesbar

## Einrichtung (einmalig, ca. 5 Minuten)

1. Neues Repository auf GitHub anlegen (z. B. `hbl-start7`) und den Inhalt dieses Ordners hochladen
   (inkl. `.github/workflows/build.yml`).
2. **Settings → Pages → Build and deployment → Source: „GitHub Actions“** auswählen.
3. **Actions → „HBL Start7 Scout“ → Run workflow** einmal starten.
4. Nach 1–3 Minuten ist die Seite unter `https://<dein-github-name>.github.io/hbl-start7/` erreichbar.
   Am Handy als Lesezeichen bzw. „Zum Home-Bildschirm“ hinzufügen.

> GitHub Pages ist bei kostenlosen Konten nur für **öffentliche** Repos verfügbar (die Daten sind ohnehin
> öffentlich). Mit privatem Repo ohne Pro-Plan: Excel direkt am Workflow-Lauf unter *Artifacts* herunterladen.

## Aktualisieren vom Handy

- Automatisch: täglich ca. 7 Uhr und 23 Uhr.
- Sofort: GitHub-App → Repo → **Actions** → „HBL Start7 Scout“ → **Run workflow**. Danach Seite neu laden.

## Lokal ausführen

```bash
pip install -r requirements.txt
python hbl_start7.py            # Ausgabe nach ./site
python hbl_start7.py --dump raw # zusätzlich Rohdaten der APIs sichern
```

Weitere Optionen: `--hpi-season 247` (ältere HPI-Saison, z. B. 2025/26), `--season-id <uuid>` (Statistik-Saison
fest vorgeben), `--from-dump raw` (offline aus Rohdaten bauen).

## Kennzahlen

| Spalte | Bedeutung |
|---|---|
| Ø HPI | Saisonschnitt des HPI (0–100) – Grundlage der Start7-Punkte |
| Form | Ø HPI der letzten 3 Spieltage mit Wert |
| Δ Form | Form minus Saisonschnitt (positiv = im Aufwind) |
| Letzter ▲▼ | HPI des letzten Spiels (nach Datum, verlegte Spieltage berücksichtigt) und Trend zum vorherigen Spiel |
| Letztes Spiel | Datum, Gegner und Ergebnis des jüngsten Spiels mit HPI-Wert (Detailansicht & Excel) |
| Spieltag-Filter | Zeigt nur Spieler mit HPI-Wert an diesem Spieltag und sortiert nach diesem Wert |
| σ | Streuung des HPI – klein = konstant, groß = Wundertüte |
| HPI-Sp | Spiele mit HPI-Wert (HBL-Rangliste verlangt mindestens 5) |
| Ø o.7m | HPI ohne Siebenmeter-Aktionen |
| Ø Vorsaison | HPI-Schnitt der Vorsaison zum Vergleich |
| Crunch | Crunchtime-Wert der HBL (Schlussphase enger Spiele) |
| Nächster / Geg. GT/Sp | Nächster Gegner (H/A) und dessen Gegentore pro Spiel – hoch = gutes Matchup für Angreifer |

Ansichten: *Start7*, *HPI*, *Aufsteiger* (größtes Δ Form), *Angriff*, *Abwehr*, *Torhüter*, *Tracking*, *Gegner*.

## Datenquellen

- `https://hpi.handball-bundesliga.de/api/…` – HPI-Saisonindex, HPI je Spieltag, Crunchtime
- `https://www.opel-hbl.de/api/synergy/…` – Feldspieler-/Torwartstatistik, Kader, Spielplan

Das sind die (inoffiziellen) Schnittstellen, die auch die HBL-Webseite nutzt. Ändert die HBL ihre Seite,
kann das Skript angepasst werden müssen. Die Server antworten gelegentlich langsam oder mit 504 – das Skript
wiederholt Anfragen automatisch.
