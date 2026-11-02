# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Projekt

Ein einziges Skript (`hbl_start7.py`) holt öffentliche Daten der Opel HBL (Handball-Bundesliga) und baut daraus
eine Scouting-Seite für das Fantasy-Managerspiel **Start7**: `site/index.html` (handytauglich), `site/hbl_start7.xlsx`
und `site/players.json`. GitHub Actions (`.github/workflows/build.yml`) baut zweimal täglich und deployt auf GitHub
Pages. UI-Texte, Kommentare, README und Commit-Messages sind auf Deutsch.

## Befehle

```bash
pip install -r requirements.txt                       # requests, openpyxl
python3 hbl_start7.py                                 # live laden, Ausgabe nach ./site
python3 hbl_start7.py --dump raw                      # zusätzlich Rohantworten nach ./raw speichern
python3 hbl_start7.py --from-dump raw --out site      # offline aus Rohdaten bauen (schnell, reproduzierbar)
python3 hbl_start7.py --now 2026-10-30T12:00:00+00:00 # Referenzzeitpunkt (für "nächste Spiele", Tage seit Spiel)
```

Es gibt keine Tests und keinen Linter. Prüfen heißt: einmal mit `--dump raw` laden, dann iterativ mit
`--from-dump raw` bauen und `site/players.json` auswerten. Das Frontend-JS steckt im Python-String `HTML_TEMPLATE`;
Syntaxfehler fallen erst im Browser auf – nach Änderungen das `<script>` der erzeugten `index.html` mindestens mit
`node -e "new Function(...)"` parsen oder die Seite headless öffnen.

## Architektur (Pipeline in `main`)

1. `Fetcher` – GET mit Retries; `--dump` schreibt jede Antwort als Datei, `--from-dump` liest nur noch daraus
   (Dateiname aus URL + Parametern, siehe `Fetcher._key`).
2. `fetch_all` – zwei Quellen:
   - **HPI-API** (`hpi.handball-bundesliga.de/api`): Saisonindex, Vorsaison, Crunchtime und HPI **pro Spieltag**
     (ein Request je Spieltag). Spieler-ID ist `dc_id`.
   - **Synergy/Sportradar** über `opel-hbl.de/api/synergy`: Saisonstatistik Feldspieler + Torwart, Kader, Spielplan.
     IDs kommen mit Präfix (`h1s44:<uuid>`) → `strip_ns`. Die Synergy-Personen-UUID entspricht dem HPI-`dc_id`.
3. `build_teams` – Tabelle, Torverhältnis, letzte 5, nächste Gegner aus dem Spielplan.
4. `build_players` – führt alles pro (Spieler, Team) zusammen; ein Spieler mit Vereinswechsel kann mehrfach vorkommen.
5. `write_html` (JSON wird als `__DATA__` in `HTML_TEMPLATE` eingesetzt), `write_excel` (`COLUMNS` steuert das
   Blatt *Spieler*), `players.json`.

Frontend: Spalten sind in `COLS` definiert (`h` = Überschrift, `t` = Beschreibung für Tooltip und Erklärzeile über
der Tabelle – Pflicht, ohne doppelte Anführungszeichen, `v` = Sortierwert, `r` = Rendering), Reiter in `VIEWS`. Neue Kennzahl =
Feld in `build_players` + Eintrag in `COLS`/`VIEWS` + ggf. `COLUMNS` (Excel) + Zeile in der README-Tabelle.

## Daten-Fallstricke (alle schon einmal Ursache von Fehlern)

- **Spieltage nach Datum sortieren, nicht nach Nummer.** Spieltage werden verlegt (z. B. ST 16 am 30.09. gespielt).
  `md_numbers` ist deshalb nach Datum sortiert, Form/„Letzter“/letztes Spiel nutzen das Spieldatum des Vereins.
- `roundNumber` im Spielplan ist ein **String**.
- `index.last` und `index.trend` der HPI-API sind unbrauchbar (zeigen meist auf Spieltag 1) → werden selbst berechnet.
- `index.max_unweighted` sind **keine** Start7-Punkte.
- Die Torwart-Statistik (`statisticType=goalkeeper`) enthält alle Spieler, auch Feldspieler (mit Nullen).
- In der Saisonstatistik ist nur ein Teil der Tore einer Wurfart zugeordnet; Würfe pro Wurfart nur über die Quote
  ableitbar (`shots_from`). Quote 0 bei 0 Toren = Wurfzahl unbekannt.
- Ein Spieltag ohne HPI-Wert heißt nicht zwingend „nicht gespielt“ (zu wenige Aktionen für einen HPI).

## Start7-Punkte (geschätzt)

Die echten Punkte gibt es nur in der App (keine öffentliche API). `start7_points` schätzt sie aus der
Saisonstatistik nach der offiziellen Tabelle (<https://start7.de/spielregeln/punkteberechnung/>), hinterlegt als
`S7_SHOT`, `S7_SAVE`, `S7_EVENT`, `S7_OFF_TARGET`. Ligaweite Hilfswerte liefert `s7_league` (Ø Assist-Punkte,
Fehlwürfe neben das Tor je Parade). Kalibriert mit 14 App-Werten nach ST 11/2026-27 (Tore ohne Wurfart = 8 Punkte);
Abweichung meist ±10–20 %. Abwehrspezialisten werden unterschätzt (Offensivfouls, Stopfouls, 7 m verursacht fehlen
in den Daten). Neue Kalibrierungen nur übernehmen, wenn sie auch ohne den jeweiligen Spieler (Leave-one-out) halten.
