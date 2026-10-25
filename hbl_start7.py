#!/usr/bin/env python3
"""
HBL Start7 Scout
================
Holt alle Spielerstatistiken der Opel HBL inkl. Handball Performance Index (HPI)
und dessen Verlauf pro Spieltag und erzeugt daraus

  * eine interaktive, handytaugliche HTML-Seite (index.html)
  * eine Excel-Datei (hbl_start7.xlsx)

Datenquellen (oeffentlich, dieselben wie auf opel-hbl.de):
  * https://hpi.handball-bundesliga.de/api/...      (HPI, HPI pro Spieltag)
  * https://www.opel-hbl.de/api/synergy/...         (Feldspieler-/Torwartstatistik,
                                                      Kader, Spielplan)

Aufruf:
  python hbl_start7.py                 # aktuelle Saison, Ausgabe nach ./site
  python hbl_start7.py --out public    # anderes Ausgabeverzeichnis
  python hbl_start7.py --dump raw      # Rohantworten zusaetzlich speichern
  python hbl_start7.py --from-dump raw # offline aus gespeicherten Rohdaten bauen
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
import statistics
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

HPI_API = "https://hpi.handball-bundesliga.de/api"
HBL_SITE = "https://www.opel-hbl.de"
SYN_API = HBL_SITE + "/api/synergy"
# Wettbewerbs-IDs der Opel HBL im Sportradar-Datacore (aus dem Frontend der HBL-Seite)
HBL_COMPETITION_IDS = ["4c445e5c-3956-11ef-9d0e-b74f5c057367", "4c5a3af0-3956-11ef-9b6c-b74f5c057367"]
# Fallback, falls die Saison nicht automatisch ermittelt werden kann (Saison 2026/27)
DEFAULT_SYNERGY_SEASON = "c4a3125f-79f2-11f1-9a19-5f7c8c2ed877"
TZ = ZoneInfo("Europe/Berlin")

POS_DE = {
    "Goalkeeper": ("TW", "Torwart"),
    "Left Wing": ("LA", "Linksaußen"),
    "Right Wing": ("RA", "Rechtsaußen"),
    "Left Back": ("RL", "Rückraum Links"),
    "Right Back": ("RR", "Rückraum Rechts"),
    "Centre Back": ("RM", "Rückraum Mitte"),
    "Center Back": ("RM", "Rückraum Mitte"),
    "Pivot": ("KM", "Kreismitte"),
}
ROSTER_POS = {"G": "TW", "GK": "TW", "LW": "LA", "RW": "RA", "LB": "RL", "RB": "RR",
              "CB": "RM", "P": "KM", "PV": "KM", "LP": "KM", "PIV": "KM"}
POS_ORDER = ["TW", "LA", "RL", "RM", "RR", "RA", "KM", "?"]


def log(msg: str) -> None:
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", file=sys.stderr, flush=True)


# --------------------------------------------------------------------------- #
# HTTP / Rohdaten
# --------------------------------------------------------------------------- #
class Fetcher:
    """Laedt JSON von den APIs (mit Retries) oder aus einem Dump-Verzeichnis."""

    def __init__(self, dump_dir: Path | None = None, from_dump: Path | None = None, timeout: int = 90):
        self.dump_dir = dump_dir
        self.from_dump = from_dump
        self.timeout = timeout
        self.session = requests.Session()
        retry = Retry(total=6, connect=6, read=6, status=6, backoff_factor=2,
                      status_forcelist=(429, 500, 502, 503, 504), allowed_methods=("GET",))
        self.session.mount("https://", HTTPAdapter(max_retries=retry, pool_maxsize=8))
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/128.0 Safari/537.36 hbl-start7-scout",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "de-DE,de;q=0.9",
            "Referer": HBL_SITE + "/de/hbl/statistiken",
        })

    @staticmethod
    def _key(url: str, params: dict | None) -> str:
        raw = url.split("://", 1)[-1] + ("?" + "&".join(f"{k}={v}" for k, v in sorted(params.items())) if params else "")
        return re.sub(r"[^A-Za-z0-9._-]+", "_", raw)[:180] + ".json"

    def get(self, url: str, params: dict | None = None, optional: bool = False):
        key = self._key(url, params)
        if self.from_dump:
            p = self.from_dump / key
            if p.exists():
                return json.loads(p.read_text(encoding="utf-8"))
            if optional:
                return None
            raise FileNotFoundError(f"Dump fehlt: {p}")
        try:
            r = self.session.get(url, params=params, timeout=self.timeout)
            r.raise_for_status()
            data = r.json()
        except Exception as e:  # noqa: BLE001
            if optional:
                log(f"  (optional) {url} fehlgeschlagen: {e}")
                return None
            raise
        if self.dump_dir:
            self.dump_dir.mkdir(parents=True, exist_ok=True)
            (self.dump_dir / key).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return data

    def get_text(self, url: str) -> str | None:
        if self.from_dump:
            return None
        try:
            r = self.session.get(url, timeout=self.timeout)
            r.raise_for_status()
            return r.text
        except Exception:  # noqa: BLE001
            return None


# --------------------------------------------------------------------------- #
# Helfer
# --------------------------------------------------------------------------- #
def iso_minutes(s: str | None) -> float | None:
    """'PT528M34S' -> 528.57"""
    if not s:
        return None
    m = re.fullmatch(r"P(?:T)?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?", s)
    if not m:
        return None
    h, mi, se = (float(x) if x else 0.0 for x in m.groups())
    return round(h * 60 + mi + se / 60, 1)


def rnd(x, n=1):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    return round(float(x), n)


def per_game(v, g, n=2):
    if v is None or not g:
        return None
    return round(v / g, n)


def strip_ns(i: str | None) -> str | None:
    """'h1s44:uuid' -> 'uuid'"""
    return i.split(":", 1)[-1] if i else i


def age_from(dob: str | None, today: dt.date) -> int | None:
    if not dob:
        return None
    try:
        d = dt.date.fromisoformat(dob[:10])
    except ValueError:
        return None
    return today.year - d.year - ((today.month, today.day) < (d.month, d.day))


# --------------------------------------------------------------------------- #
# Datenbeschaffung
# --------------------------------------------------------------------------- #
def find_synergy_season(f: Fetcher, override: str | None) -> str:
    if override:
        return override
    for cid in HBL_COMPETITION_IDS:
        data = f.get(f"{SYN_API}/seasons", {"competitionId": cid}, optional=True)
        if not data:
            continue
        seasons = data.get("data", data) if isinstance(data, dict) else data
        seasons = [s for s in seasons or [] if s.get("status") not in ("DRAFT", "PENDING")]
        if seasons:
            active = [s for s in seasons if s.get("status") == "ACTIVE"] or seasons
            active.sort(key=lambda s: s.get("startDate") or "", reverse=True)
            return active[0]["seasonId"]
    # Fallback: Saison-ID aus dem Nuxt-Payload der Statistikseite lesen
    html = f.get_text(HBL_SITE + "/de/hbl/statistiken/feldspieler/feldspieler")
    if html:
        m = re.search(r'"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",\{"id":\d+,"resourceType":\d+\},'
                      r'"\d{4}-\d{2}-\d{2}","ACTIVE"', html)
        if m:
            return m.group(1)
    log(f"  Saison-ID nicht ermittelbar, nutze Fallback {DEFAULT_SYNERGY_SEASON}")
    return DEFAULT_SYNERGY_SEASON


def fetch_all(f: Fetcher, synergy_season: str | None, hpi_season: int | None) -> dict:
    log("HPI: Turnier/Saisons laden")
    tour = f.get(f"{HPI_API}/tournament/1")["data"]
    seasons = sorted(tour["seasons"], key=lambda s: s["id"])
    cur = next((s for s in seasons if s["id"] == hpi_season), None) if hpi_season else seasons[-1]
    if cur is None:
        raise SystemExit(f"HPI-Saison {hpi_season} nicht gefunden")
    prev = next((s for s in reversed(seasons) if s["id"] < cur["id"]), None)

    log(f"HPI: Saisonindex {cur['name']} (id {cur['id']})")
    hpi_season_idx = f.get(f"{HPI_API}/index/season/{cur['id']}")["data"]
    hpi_prev_idx = f.get(f"{HPI_API}/index/season/{prev['id']}", optional=True) if prev else None
    hpi_prev_idx = (hpi_prev_idx or {}).get("data", [])
    crunch = f.get(f"{HPI_API}/crunchtime/season/{cur['id']}", optional=True) if cur.get("hasCrunchtime") else None
    crunch = (crunch or {}).get("data", [])

    mds = sorted([m for m in cur["matchdays"] if m.get("started")], key=lambda m: m["number"])
    log(f"HPI: {len(mds)} Spieltage einzeln laden")

    def load_md(m):
        d = f.get(f"{HPI_API}/index/season/{cur['id']}/matchday/{m['id']}", optional=True)
        return m, (d or {}).get("data", [])

    with ThreadPoolExecutor(max_workers=4) as ex:
        md_data = list(ex.map(load_md, mds))

    sid = find_synergy_season(f, synergy_season)
    log(f"Statistik: Saison {sid}")
    fp = f.get(f"{SYN_API}/season-statistic/player-overview", {"seasonId": sid, "statisticType": "fieldplayer"})
    gk = f.get(f"{SYN_API}/season-statistic/player-overview", {"seasonId": sid, "statisticType": "goalkeeper"})
    log("Statistik: Kader & Spielplan")
    roster, off = [], 0
    while True:
        r = f.get(f"{SYN_API}/season/roster", {"seasonId": sid, "offset": off}, optional=True)
        if not r or not r.get("data"):
            break
        roster += r["data"]
        if not (r.get("links") or {}).get("next"):
            break
        off += 1000
    fixtures = f.get(f"{SYN_API}/season/fixtures", {"seasonId": sid}, optional=True) or {}

    return dict(hpi_tour=tour, hpi_cur=cur, hpi_prev=prev, hpi_idx=hpi_season_idx, hpi_prev_idx=hpi_prev_idx,
                crunch=crunch, md_data=md_data, synergy_season=sid, fp=fp, gk=gk, roster=roster, fixtures=fixtures)


# --------------------------------------------------------------------------- #
# Aufbereitung
# --------------------------------------------------------------------------- #
def build_teams(raw: dict, now: dt.datetime) -> dict:
    """Tabelle, Torverhältnis und naechste Gegner aus dem Spielplan."""
    fx = raw["fixtures"] or {}
    ents = ((fx.get("includes") or {}).get("resources") or {}).get("entities") or {}
    for res in (raw["fp"], raw["gk"]):
        ents = {**(((res.get("includes") or {}).get("resources") or {}).get("entities") or {}), **ents}
    teams: dict[str, dict] = {}

    def team(eid):
        if eid not in teams:
            e = ents.get(f"h1s44:{eid}") or ents.get(eid) or {}
            teams[eid] = dict(id=eid, name=e.get("nameFullLocal") or e.get("nameFullLatin") or eid[:8],
                              abbr=e.get("codeLocal") or e.get("codeLatin") or "", games=0, w=0, d=0, l=0,
                              gf=0, ga=0, pts=0, fixtures=[])
        return teams[eid]

    for e in ents.values():
        if e.get("entityId"):
            team(e["entityId"])

    for m in fx.get("data", []):
        comps = m.get("competitors") or []
        if len(comps) != 2:
            continue
        a, b = comps
        start = m.get("startTimeUTC")
        try:
            start_dt = dt.datetime.fromisoformat(start).replace(tzinfo=dt.timezone.utc) if start else None
        except ValueError:
            start_dt = None
        for me, op in ((a, b), (b, a)):
            t = team(me["entityId"])
            t["fixtures"].append(dict(opp=op["entityId"], home=bool(me.get("isHome")), start=start_dt,
                                      round=m.get("roundNumber"), status=m.get("status"),
                                      gf=me.get("score"), ga=op.get("score")))
        if m.get("status") == "CONFIRMED":
            try:
                sa, sb = int(a.get("score")), int(b.get("score"))
            except (TypeError, ValueError):
                continue
            for me, gf, ga in ((a, sa, sb), (b, sb, sa)):
                t = team(me["entityId"])
                t["games"] += 1
                t["gf"] += gf
                t["ga"] += ga
                if gf > ga:
                    t["w"] += 1
                    t["pts"] += 2
                elif gf == ga:
                    t["d"] += 1
                    t["pts"] += 1
                else:
                    t["l"] += 1

    ranked = sorted([t for t in teams.values() if t["games"] or t["fixtures"]],
                    key=lambda t: (-t["pts"], -(t["gf"] - t["ga"]), -t["gf"]))
    for i, t in enumerate(ranked, 1):
        t["rank"] = i
        t["ga_pg"] = per_game(t["ga"], t["games"], 1)
        t["gf_pg"] = per_game(t["gf"], t["games"], 1)
        upcoming = sorted([x for x in t["fixtures"] if x["status"] in ("SCHEDULED", "IN_PROGRESS")
                           and x["start"] and x["start"] > now - dt.timedelta(hours=3)], key=lambda x: x["start"])
        t["next"] = upcoming[:3]
        played = sorted([x for x in t["fixtures"] if x["status"] == "CONFIRMED" and x["start"]],
                        key=lambda x: x["start"])
        t["last5"] = "".join("S" if int(x["gf"]) > int(x["ga"]) else "U" if x["gf"] == x["ga"] else "N"
                             for x in played[-5:] if str(x["gf"]).isdigit() and str(x["ga"]).isdigit())
    return {t["id"]: t for t in ranked}


def build_players(raw: dict, teams: dict, today: dt.date) -> tuple[list[dict], list[int]]:
    md_numbers = [m["number"] for m, _ in raw["md_data"]]
    # HPI pro Spieltag: dc_id -> {number: hpi}
    md_hpi: dict[str, dict[int, float]] = {}
    md_hpi_team: dict[str, str] = {}
    for m, rows in raw["md_data"]:
        for r in rows:
            v = (r.get("index") or {}).get("avg")
            if v is not None:
                md_hpi.setdefault(r["dc_id"], {})[m["number"]] = round(v, 1)
                md_hpi_team[r["dc_id"]] = (r.get("team") or {}).get("dc_id")

    hpi = {r["dc_id"]: r for r in raw["hpi_idx"]}
    hpi_prev = {r["dc_id"]: r for r in raw["hpi_prev_idx"]}
    crunch = {r["dc_id"]: r.get("crunchtime") or {} for r in raw["crunch"]}

    persons = {}
    entities = {}
    for res in (raw["fp"], raw["gk"]):
        rr = (res.get("includes") or {}).get("resources") or {}
        persons.update({strip_ns(k): v for k, v in (rr.get("persons") or {}).items()})
        entities.update({strip_ns(k): v for k, v in (rr.get("entities") or {}).items()})

    roster_pos = {}
    for r in raw["roster"]:
        roster_pos[(r.get("personId"), r.get("entityId"))] = r.get("position")
        roster_pos.setdefault((r.get("personId"), None), r.get("position"))

    stats_fp = {(strip_ns(d["person"]["id"]), strip_ns(d["entity"]["id"])): d.get("statistics") or {}
                for d in raw["fp"].get("data", [])}
    stats_gk = {(strip_ns(d["person"]["id"]), strip_ns(d["entity"]["id"])): d.get("statistics") or {}
                for d in raw["gk"].get("data", [])}

    keys = set(stats_fp) | set(stats_gk)
    # Spieler mit HPI aber (noch) ohne Statistik ebenfalls aufnehmen
    for pid, r in hpi.items():
        tid = (r.get("team") or {}).get("dc_id")
        if not any(k[0] == pid for k in keys):
            keys.add((pid, tid))

    players = []
    for pid, tid in keys:
        sf = stats_fp.get((pid, tid), {})
        sg = stats_gk.get((pid, tid), {})
        p = persons.get(pid, {})
        h = hpi.get(pid)
        # HPI nur zuordnen, wenn das Team passt (Vereinswechsel) – sonst trotzdem, wenn Spieler nur einmal vorkommt
        if h and (h.get("team") or {}).get("dc_id") not in (tid, None):
            if sum(1 for k in keys if k[0] == pid) > 1:
                h = None
        hi = (h or {}).get("index") or {}
        hi7 = (h or {}).get("index_ex7m") or {}
        hp = ((hpi_prev.get(pid) or {}).get("index") or {})
        team = teams.get(tid) or {}
        ent = entities.get(tid, {})

        pos_en = (h or {}).get("position")
        if pos_en in POS_DE:
            pos = POS_DE[pos_en][0]
        else:
            rp = roster_pos.get((pid, tid)) or roster_pos.get((pid, None)) or ""
            pos = ROSTER_POS.get(rp.upper(), "?")
        if pos == "?" and sg.get("goalKeeperShotsSaved"):
            pos = "TW"

        series = md_hpi.get(pid, {})
        played_vals = [series[n] for n in md_numbers if n in series]
        last3 = played_vals[-3:]
        form3 = rnd(statistics.fmean(last3)) if last3 else None
        avg = rnd(hi.get("avg"))
        games = sf.get("games") or sg.get("games") or 0
        minutes = iso_minutes(sf.get("timeOnPlayingField") or sg.get("timeOnPlayingField"))
        goals = sf.get("goalsScored") if sf else sg.get("goalsScored")
        name = (p.get("nameFullLocal") or
                (f"{(h or {}).get('firstname', '')} {(h or {}).get('lastname', '')}".strip()) or pid[:8])
        last_name = p.get("nameFamilyLocal") or (h or {}).get("lastname") or name.split(" ")[-1]

        # naechste Gegner
        nxt = []
        for x in (team.get("next") or [])[:3]:
            o = teams.get(x["opp"], {})
            nxt.append(dict(opp=o.get("abbr") or o.get("name", "?"), opp_name=o.get("name", "?"),
                            home=x["home"], date=x["start"].astimezone(TZ).strftime("%a %d.%m. %H:%M")
                            .replace("Mon", "Mo").replace("Tue", "Di").replace("Wed", "Mi").replace("Thu", "Do")
                            .replace("Fri", "Fr").replace("Sat", "Sa").replace("Sun", "So"),
                            opp_rank=o.get("rank"), opp_ga_pg=o.get("ga_pg"), opp_gf_pg=o.get("gf_pg"),
                            round=x["round"]))

        speed = sf.get("speedMax") if sf else sg.get("speedMax") if sg else None
        players.append(dict(
            id=pid, name=name, last=last_name, team=team.get("name") or ent.get("nameFullLocal") or "",
            team_abbr=team.get("abbr") or ent.get("codeLocal") or "", team_id=tid, team_rank=team.get("rank"),
            pos=pos, age=age_from(p.get("dob"), today), nat=p.get("nationality"),
            height=(p.get("additionalDetails") or {}).get("height"),
            url=(h or {}).get("url") or f"{HBL_SITE}/de/player/{pid}",
            # HPI
            hpi_avg=avg, hpi_avg_ex7m=rnd(hi7.get("avg")), hpi_max=rnd(hi.get("max")), hpi_last=rnd(hi.get("last")),
            hpi_trend=hi.get("trend"), hpi_std=rnd(hi.get("stddev")), hpi_md=hi.get("matchdays"),
            hpi_md_needed=hi.get("matchdays_needed"), hpi_events=hi.get("events"),
            hpi_form3=form3, hpi_form_delta=rnd(form3 - avg) if (form3 is not None and avg is not None) else None,
            hpi_prev=rnd(hp.get("avg")), hpi_prev_md=hp.get("matchdays"),
            hpi_series=[series.get(n) for n in md_numbers],
            crunch_avg=rnd((crunch.get(pid) or {}).get("avg")),
            # Basis
            games=games, minutes=minutes, min_pg=per_game(minutes, games, 1),
            goals=goals, goals_pg=per_game(goals, games),
            shots=sf.get("shots"), acc=rnd(sf.get("shootingAccuracy")),
            g7=sf.get("sevenMetreGoalsScored"), acc7=rnd(sf.get("sevenMetreShootingAccuracy")),
            g_field=sf.get("fieldGoalsScored"), g_wing=sf.get("wingGoalsScored"), g_back=sf.get("backCourtGoalsScored"),
            g_9m=sf.get("nineMetreGoalsScored"), g_pivot=sf.get("pivotGoalsScored"), g_fb=sf.get("fastBreakGoalsScored"),
            assists=sf.get("assists") if sf else sg.get("assists"),
            assists_pg=per_game(sf.get("assists") if sf else sg.get("assists"), games),
            steals=sf.get("steals") if sf else sg.get("steals"), blocks=sf.get("blocks"),
            tf=sf.get("technicalFaults"), tf_pg=per_game(sf.get("technicalFaults"), games),
            two_min=sf.get("twoMinuteSuspensions") if sf else sg.get("twoMinuteSuspensions"),
            yellow=sf.get("yellowCards") if sf else sg.get("yellowCards"),
            red=sf.get("redCards") if sf else sg.get("redCards"),
            dist_km=rnd(((sf or sg).get("distanceTotal") or 0) / 1000, 1) if (sf or sg).get("distanceTotal") else None,
            speed_kmh=rnd(speed * 3.6, 1) if speed else None,
            # Torwart
            saves=sg.get("goalKeeperShotsSaved"), save_pct=rnd(sg.get("goalKeeperSaveAccuracy")),
            saves_pg=per_game(sg.get("goalKeeperShotsSaved"), games, 1),
            ga=sg.get("goalKeeperGoalsAgainst"), saves7=sg.get("goalKeeperSevenMetreShotsSaved"),
            save7_pct=rnd(sg.get("goalKeeperSevenMetreSaveAccuracy")),
            save_wing_pct=rnd(sg.get("goalKeeperWingSaveAccuracy")),
            save_back_pct=rnd(sg.get("goalKeeperBackCourtSaveAccuracy")),
            save_fb_pct=rnd(sg.get("goalKeeperFastBreakSaveAccuracy")),
            # Spielplan
            next=nxt,
        ))
    # Nur echte Spieler (Statistik oder HPI vorhanden)
    players = [p for p in players if p["games"] or p["hpi_avg"] is not None]
    players.sort(key=lambda p: (-(p["hpi_avg"] or 0), -(p["games"] or 0)))
    return players, md_numbers


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
COLUMNS = [  # key, header, width, group
    ("name", "Spieler", 24, "Basis"), ("team_abbr", "Team", 7, "Basis"), ("pos", "Pos", 5, "Basis"),
    ("age", "Alter", 6, "Basis"), ("nat", "Nat", 6, "Basis"),
    ("hpi_avg", "Ø HPI", 8, "HPI"), ("hpi_form3", "Form (Ø letzte 3)", 9, "HPI"),
    ("hpi_form_delta", "Form vs Ø", 8, "HPI"), ("hpi_last", "HPI letzter ST", 9, "HPI"),
    ("hpi_max", "Max HPI", 8, "HPI"), ("hpi_std", "Streuung (σ)", 9, "HPI"),
    ("hpi_md", "HPI-Spiele", 8, "HPI"), ("hpi_avg_ex7m", "Ø HPI ohne 7m", 9, "HPI"),
    ("hpi_prev", "Ø HPI Vorsaison", 9, "HPI"), ("crunch_avg", "Crunchtime Ø", 9, "HPI"),
    ("games", "Spiele", 7, "Basis"), ("min_pg", "Min/Spiel", 8, "Basis"),
    ("goals", "Tore", 6, "Angriff"), ("goals_pg", "Tore/Spiel", 8, "Angriff"), ("shots", "Würfe", 6, "Angriff"),
    ("acc", "Quote %", 7, "Angriff"), ("g7", "7m-Tore", 7, "Angriff"), ("acc7", "7m %", 6, "Angriff"),
    ("g_fb", "Gegenstoß", 8, "Angriff"), ("g_wing", "Außen", 6, "Angriff"), ("g_back", "Rückraum", 8, "Angriff"),
    ("g_9m", "9m", 5, "Angriff"), ("g_pivot", "Kreis", 6, "Angriff"),
    ("assists", "Assists", 7, "Angriff"), ("assists_pg", "Assists/Spiel", 9, "Angriff"),
    ("tf", "Techn. Fehler", 8, "Disziplin"), ("tf_pg", "TF/Spiel", 7, "Disziplin"),
    ("steals", "Steals", 7, "Abwehr"), ("blocks", "Blocks", 7, "Abwehr"),
    ("two_min", "2 Min", 6, "Disziplin"), ("yellow", "Gelb", 5, "Disziplin"), ("red", "Rot", 5, "Disziplin"),
    ("saves", "Paraden", 8, "Torwart"), ("save_pct", "Paraden %", 9, "Torwart"), ("saves_pg", "Paraden/Spiel", 9, "Torwart"),
    ("ga", "Gegentore", 9, "Torwart"), ("saves7", "7m gehalten", 9, "Torwart"), ("save7_pct", "7m %", 6, "Torwart"),
    ("dist_km", "Laufweg km", 9, "Tracking"), ("speed_kmh", "Max km/h", 8, "Tracking"),
    ("next1", "Nächster Gegner", 16, "Spielplan"), ("next1_rank", "Gegner Platz", 8, "Spielplan"),
    ("next1_ga", "Gegner Gegentore/Sp", 10, "Spielplan"), ("team", "Verein", 24, "Basis"),
]


def next_label(p) -> str:
    if not p["next"]:
        return ""
    n = p["next"][0]
    return f"{'H' if n['home'] else 'A'} vs {n['opp']} ({n['date']})"


def write_excel(path: Path, players: list[dict], md_numbers: list[int], teams: dict, meta: dict) -> None:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="1F2A44")
    head_font = Font(bold=True, color="FFFFFF")
    scale = ColorScaleRule(start_type="num", start_value=50, start_color="F8696B",
                           mid_type="num", mid_value=70, mid_color="FFEB84",
                           end_type="num", end_value=90, end_color="63BE7B")

    def header(ws, cols):
        for i, (h, w) in enumerate(cols, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "B2"

    # Spieler
    ws = wb.active
    ws.title = "Spieler"
    header(ws, [(h, w) for _, h, w, _ in COLUMNS])
    for r, p in enumerate(players, 2):
        n = p["next"][0] if p["next"] else {}
        row = dict(p, next1=next_label(p), next1_rank=n.get("opp_rank"), next1_ga=n.get("opp_ga_pg"))
        for c, (k, *_rest) in enumerate(COLUMNS, 1):
            cell = ws.cell(row=r, column=c, value=row.get(k))
            if k == "name":
                cell.hyperlink = p["url"]
                cell.font = Font(color="1F4E99", underline="single")
    last = len(players) + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"
    for k in ("hpi_avg", "hpi_form3", "hpi_last", "hpi_avg_ex7m", "hpi_prev"):
        col = get_column_letter([c[0] for c in COLUMNS].index(k) + 1)
        ws.conditional_formatting.add(f"{col}2:{col}{last}", scale)
    col = get_column_letter([c[0] for c in COLUMNS].index("hpi_form_delta") + 1)
    ws.conditional_formatting.add(f"{col}2:{col}{last}", ColorScaleRule(
        start_type="num", start_value=-10, start_color="F8696B", mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="num", end_value=10, end_color="63BE7B"))

    # HPI-Verlauf
    ws2 = wb.create_sheet("HPI-Verlauf")
    cols = [("Spieler", 24), ("Team", 7), ("Pos", 5), ("Ø HPI", 7), ("Form", 7)] + [(f"ST {n}", 6) for n in md_numbers]
    header(ws2, cols)
    for r, p in enumerate(players, 2):
        vals = [p["name"], p["team_abbr"], p["pos"], p["hpi_avg"], p["hpi_form3"]] + p["hpi_series"]
        for c, v in enumerate(vals, 1):
            ws2.cell(row=r, column=c, value=v)
    last2 = len(players) + 1
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last2}"
    ws2.conditional_formatting.add(f"D2:{get_column_letter(len(cols))}{last2}", scale)

    # Torhüter
    ws3 = wb.create_sheet("Torhüter")
    gcols = [("name", "Spieler", 24), ("team_abbr", "Team", 7), ("hpi_avg", "Ø HPI", 7), ("hpi_form3", "Form", 7),
             ("games", "Spiele", 7), ("min_pg", "Min/Spiel", 8), ("saves", "Paraden", 8), ("saves_pg", "Paraden/Spiel", 9),
             ("save_pct", "Paraden %", 9), ("ga", "Gegentore", 9), ("saves7", "7m gehalten", 9), ("save7_pct", "7m %", 7),
             ("save_wing_pct", "Außen %", 8), ("save_back_pct", "Rückraum %", 9), ("save_fb_pct", "Gegenstoß %", 9),
             ("goals", "Tore", 6), ("assists", "Assists", 7)]
    header(ws3, [(h, w) for _, h, w in gcols])
    gks = sorted([p for p in players if p["pos"] == "TW"], key=lambda p: -(p["hpi_avg"] or 0))
    for r, p in enumerate(gks, 2):
        for c, (k, *_r) in enumerate(gcols, 1):
            ws3.cell(row=r, column=c, value=p.get(k))
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(gcols))}{len(gks) + 1}"

    # Tabelle & Spielplan
    ws4 = wb.create_sheet("Tabelle & Gegner")
    tcols = [("Pl.", 5), ("Verein", 26), ("Sp", 5), ("S", 4), ("U", 4), ("N", 4), ("Tore", 9), ("Diff", 6),
             ("Pkt", 5), ("Tore/Sp", 8), ("Gegentore/Sp", 11), ("Form (letzte 5)", 12),
             ("Nächstes Spiel", 30), ("Danach", 30)]
    header(ws4, tcols)
    for r, t in enumerate(sorted(teams.values(), key=lambda t: t["rank"]), 2):
        def fx(i):
            if len(t["next"]) <= i:
                return ""
            x = t["next"][i]
            o = teams.get(x["opp"], {})
            return f"{'H' if x['home'] else 'A'} vs {o.get('name', '?')} – {x['start'].astimezone(TZ):%d.%m. %H:%M}"
        vals = [t["rank"], t["name"], t["games"], t["w"], t["d"], t["l"], f"{t['gf']}:{t['ga']}", t["gf"] - t["ga"],
                t["pts"], t["gf_pg"], t["ga_pg"], t["last5"], fx(0), fx(1)]
        for c, v in enumerate(vals, 1):
            ws4.cell(row=r, column=c, value=v)

    # Info
    ws5 = wb.create_sheet("Info")
    info = [("Erstellt", meta["generated"]), ("Saison", meta["season"]),
            ("Spieltage mit HPI", ", ".join(map(str, md_numbers))),
            ("Quelle HPI", "hpi.handball-bundesliga.de"), ("Quelle Statistik", "opel-hbl.de / Sportradar"),
            ("", ""),
            ("Ø HPI", "Saisondurchschnitt des Handball Performance Index (0–100)."),
            ("Form (Ø letzte 3)", "Durchschnitt der letzten 3 Spieltage mit HPI-Wert."),
            ("Form vs Ø", "Form minus Saisonschnitt: positiv = Spieler im Aufwind."),
            ("Streuung (σ)", "Standardabweichung des HPI: klein = konstant, groß = schwankend."),
            ("HPI-Spiele", "Spieltage mit genug Aktionen für einen HPI-Wert (HBL-Ranking braucht 5)."),
            ("Crunchtime Ø", "HPI-Beitrag in den Schlussminuten enger Spiele (HBL-Wert)."),
            ("Gegner Gegentore/Sp", "Wie viele Tore der nächste Gegner im Schnitt kassiert – hoch = gutes Matchup für Angreifer.")]
    for r, (a, b) in enumerate(info, 1):
        ws5.cell(row=r, column=1, value=a).font = Font(bold=True)
        ws5.cell(row=r, column=2, value=b)
    ws5.column_dimensions["A"].width = 22
    ws5.column_dimensions["B"].width = 100
    wb.save(path)


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
HTML_TEMPLATE = r"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0f172a">
<title>HBL Start7 Scout</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--fg:#111827;--mut:#6b7280;--line:#e5e7eb;--acc:#2563eb;--up:#16a34a;--down:#dc2626;--chip:#eef2ff;--head:#0f172a}
@media (prefers-color-scheme:dark){:root{--bg:#0b1020;--card:#121a2e;--fg:#e5e7eb;--mut:#94a3b8;--line:#22304d;--acc:#60a5fa;--up:#4ade80;--down:#f87171;--chip:#1e293b;--head:#020617}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;-webkit-text-size-adjust:100%}
header{background:var(--head);color:#fff;padding:14px 16px 10px;position:sticky;top:0;z-index:5}
header h1{font-size:18px;margin:0}
header .sub{font-size:12px;opacity:.75;margin-top:2px}
.bar{display:flex;flex-wrap:wrap;gap:8px;padding:10px 16px;background:var(--card);border-bottom:1px solid var(--line);position:sticky;top:58px;z-index:4}
.bar input,.bar select{font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg);min-width:0}
.bar input[type=search]{flex:1 1 160px}
.tabs{display:flex;gap:6px;overflow-x:auto;padding:8px 16px;-webkit-overflow-scrolling:touch}
.tabs button{flex:0 0 auto;font:inherit;font-size:13px;padding:6px 12px;border-radius:999px;border:1px solid var(--line);background:var(--card);color:var(--fg)}
.tabs button.on{background:var(--acc);border-color:var(--acc);color:#fff}
.wrap{overflow-x:auto;margin:0 0 24px;-webkit-overflow-scrolling:touch}
table{border-collapse:separate;border-spacing:0;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:7px 8px;border-bottom:1px solid var(--line);white-space:nowrap;text-align:right}
th{position:sticky;top:0;background:var(--card);font-size:12px;color:var(--mut);cursor:pointer;user-select:none;z-index:2}
th.sorted{color:var(--acc)}
td:first-child,th:first-child{position:sticky;left:0;background:var(--card);text-align:left;z-index:1;max-width:46vw;overflow:hidden;text-overflow:ellipsis}
th:first-child{z-index:3}
tbody tr{background:var(--card);cursor:pointer}
tbody tr:hover td{background:var(--chip)}
.tl{text-align:left}
.pos{display:inline-block;min-width:26px;text-align:center;font-size:11px;padding:1px 4px;border-radius:5px;background:var(--chip);color:var(--mut)}
.hpi{display:inline-block;min-width:34px;padding:2px 6px;border-radius:6px;font-weight:600;color:#111}
.up{color:var(--up)}.down{color:var(--down)}.mut{color:var(--mut)}
.small{font-size:12px}
svg.spark{vertical-align:middle}
.empty{padding:30px;text-align:center;color:var(--mut)}
dialog{border:none;border-radius:14px;padding:0;width:min(640px,96vw);max-height:92vh;background:var(--card);color:var(--fg)}
dialog::backdrop{background:rgba(0,0,0,.55)}
.dlg{padding:16px}
.dlg h2{margin:0 0 2px;font-size:19px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px;margin-top:12px}
.kpi{background:var(--bg);border-radius:10px;padding:8px 10px}
.kpi b{display:block;font-size:17px}
.kpi span{font-size:11px;color:var(--mut)}
.close{float:right;font-size:22px;background:none;border:none;color:var(--fg);cursor:pointer}
.sect{margin-top:14px;font-weight:600;font-size:13px;color:var(--mut);text-transform:uppercase;letter-spacing:.04em}
footer{padding:16px;color:var(--mut);font-size:12px}
a{color:var(--acc)}
</style>
</head>
<body>
<header><h1>HBL Start7 Scout</h1><div class="sub" id="sub"></div></header>
<div class="bar">
  <input type="search" id="q" placeholder="Spieler oder Verein suchen …">
  <select id="team"><option value="">Alle Vereine</option></select>
  <select id="pos"><option value="">Alle Positionen</option></select>
  <select id="minmd" title="Mindestanzahl Spiele">
    <option value="0">alle Spieler</option><option value="1">≥1 Spiel</option><option value="3" selected>≥3 Spiele</option>
    <option value="5">≥5 Spiele</option>
  </select>
</div>
<div class="tabs" id="tabs"></div>
<div class="wrap"><table id="tbl"><thead></thead><tbody></tbody></table></div>
<div class="tabs"><span class="small mut" id="count"></span></div>
<footer>
  Daten: <a href="https://www.opel-hbl.de/de/hbl/statistiken">opel-hbl.de</a> &amp; <a href="https://hpi.handball-bundesliga.de">HPI der HBL</a>.
  Spalten antippen zum Sortieren, Zeile antippen für Details &amp; HPI-Verlauf.
  <a href="hbl_start7.xlsx">Excel herunterladen</a>.
  <br>Form = Ø HPI der letzten 3 Spieltage · Δ = Form minus Saisonschnitt · σ = Streuung (Konstanz) · Gegner GT/Sp = Gegentore pro Spiel des nächsten Gegners.
</footer>
<dialog id="dlg"><div class="dlg" id="dlgc"></div></dialog>
<script>
const DATA = __DATA__;
const P = DATA.players, MD = DATA.matchdays;
const POSN = {TW:"Torwart",LA:"Linksaußen",RA:"Rechtsaußen",RL:"Rückraum links",RR:"Rückraum rechts",RM:"Rückraum Mitte",KM:"Kreis","?":"unbekannt"};
const f1 = v => v==null ? "–" : (Math.round(v*10)/10).toLocaleString("de-DE");
const f0 = v => v==null ? "–" : String(v);
function hpiColor(v){ if(v==null) return "transparent"; const t=Math.max(0,Math.min(1,(v-50)/45));
  const h=t*125; return `hsl(${h} 70% ${62 - t*8}%)`; }
const hpiCell = v => v==null ? '<span class="mut">–</span>' : `<span class="hpi" style="background:${hpiColor(v)}">${Math.round(v)}</span>`;
const delta = v => v==null ? '<span class="mut">–</span>' : `<span class="${v>0.5?'up':v<-0.5?'down':'mut'}">${v>0?'+':''}${f1(v)}</span>`;
const trend = t => t>0 ? '<span class="up">▲</span>' : t<0 ? '<span class="down">▼</span>' : '<span class="mut">●</span>';
function path(s,x,y){let d='',pen=false; s.forEach((v,i)=>{if(v==null){pen=false;return;} d+=(pen?'L':'M')+x(i).toFixed(1)+' '+y(v).toFixed(1)+' '; pen=true;}); return d;}
function spark(s, w=90, h=24){ const pts=s.map((v,i)=>[i,v]).filter(p=>p[1]!=null); if(pts.length<1) return '<span class="mut">–</span>';
  const n=Math.max(s.length-1,1), y=v=>h-2-(Math.max(30,Math.min(100,v))-30)/70*(h-4), x=i=>2+i/n*(w-4);
  const d=path(s,x,y); const l=pts[pts.length-1];
  return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><line x1="0" x2="${w}" y1="${y(70)}" y2="${y(70)}" stroke="currentColor" stroke-opacity=".15"/><path d="${d}" fill="none" stroke="var(--acc)" stroke-width="1.6"/><circle cx="${x(l[0])}" cy="${y(l[1])}" r="2.4" fill="var(--acc)"/></svg>`; }
const nextTxt = p => p.next.length ? `${p.next[0].home?'H':'A'} ${p.next[0].opp}` : '–';

const COLS = {
  name:{h:"Spieler",cls:"tl",v:p=>p.last,r:p=>`${p.name}<div class="small mut">${p.team_abbr} · <span class="pos">${p.pos}</span></div>`},
  pos:{h:"Pos",v:p=>"TWLARLRMRRRAKM?".indexOf(p.pos),r:p=>`<span class="pos">${p.pos}</span>`},
  team_abbr:{h:"Team",cls:"tl",v:p=>p.team_abbr},
  hpi_avg:{h:"Ø HPI",v:p=>p.hpi_avg,r:p=>hpiCell(p.hpi_avg)},
  hpi_form3:{h:"Form",v:p=>p.hpi_form3,r:p=>hpiCell(p.hpi_form3)},
  hpi_form_delta:{h:"Δ Form",v:p=>p.hpi_form_delta,r:p=>delta(p.hpi_form_delta)},
  hpi_last:{h:"Letzter",v:p=>p.hpi_last,r:p=>hpiCell(p.hpi_last)+" "+trend(p.hpi_trend)},
  spark:{h:"Verlauf",v:p=>p.hpi_form3,r:p=>spark(p.hpi_series)},
  hpi_max:{h:"Max",v:p=>p.hpi_max,r:p=>f1(p.hpi_max)},
  hpi_std:{h:"σ",v:p=>p.hpi_std,r:p=>f1(p.hpi_std)},
  hpi_md:{h:"HPI-Sp",v:p=>p.hpi_md,r:p=>f0(p.hpi_md)},
  hpi_avg_ex7m:{h:"Ø o.7m",v:p=>p.hpi_avg_ex7m,r:p=>f1(p.hpi_avg_ex7m)},
  hpi_prev:{h:"Ø Vorsaison",v:p=>p.hpi_prev,r:p=>hpiCell(p.hpi_prev)},
  crunch_avg:{h:"Crunch",v:p=>p.crunch_avg,r:p=>f1(p.crunch_avg)},
  games:{h:"Sp",v:p=>p.games,r:p=>f0(p.games)},
  min_pg:{h:"Min/Sp",v:p=>p.min_pg,r:p=>f1(p.min_pg)},
  goals:{h:"Tore",v:p=>p.goals,r:p=>f0(p.goals)},
  goals_pg:{h:"T/Sp",v:p=>p.goals_pg,r:p=>f1(p.goals_pg)},
  shots:{h:"Würfe",v:p=>p.shots,r:p=>f0(p.shots)},
  acc:{h:"Quote",v:p=>p.acc,r:p=>p.acc==null?"–":f1(p.acc)+"%"},
  g7:{h:"7m",v:p=>p.g7,r:p=>f0(p.g7)},
  acc7:{h:"7m %",v:p=>p.acc7,r:p=>p.acc7==null?"–":f1(p.acc7)+"%"},
  g_fb:{h:"Gegenst.",v:p=>p.g_fb,r:p=>f0(p.g_fb)},
  assists:{h:"Assists",v:p=>p.assists,r:p=>f0(p.assists)},
  assists_pg:{h:"A/Sp",v:p=>p.assists_pg,r:p=>f1(p.assists_pg)},
  steals:{h:"Steals",v:p=>p.steals,r:p=>f0(p.steals)},
  blocks:{h:"Blocks",v:p=>p.blocks,r:p=>f0(p.blocks)},
  tf:{h:"TF",v:p=>p.tf,r:p=>f0(p.tf)},
  tf_pg:{h:"TF/Sp",v:p=>p.tf_pg,r:p=>f1(p.tf_pg)},
  two_min:{h:"2'",v:p=>p.two_min,r:p=>f0(p.two_min)},
  red:{h:"Rot",v:p=>p.red,r:p=>f0(p.red)},
  saves:{h:"Paraden",v:p=>p.saves,r:p=>f0(p.saves)},
  saves_pg:{h:"Par/Sp",v:p=>p.saves_pg,r:p=>f1(p.saves_pg)},
  save_pct:{h:"Par %",v:p=>p.save_pct,r:p=>p.save_pct==null?"–":f1(p.save_pct)+"%"},
  ga:{h:"Gegentore",v:p=>p.ga,r:p=>f0(p.ga)},
  save7_pct:{h:"7m %",v:p=>p.save7_pct,r:p=>p.save7_pct==null?"–":f1(p.save7_pct)+"%"},
  dist_km:{h:"km",v:p=>p.dist_km,r:p=>f1(p.dist_km)},
  speed_kmh:{h:"km/h",v:p=>p.speed_kmh,r:p=>f1(p.speed_kmh)},
  next:{h:"Nächster",cls:"tl",v:p=>p.next.length?p.next[0].opp_rank:null,r:p=>nextTxt(p)},
  next_rank:{h:"Geg.-Pl.",v:p=>p.next.length?p.next[0].opp_rank:null,r:p=>p.next.length?f0(p.next[0].opp_rank):"–"},
  next_ga:{h:"Geg. GT/Sp",v:p=>p.next.length?p.next[0].opp_ga_pg:null,r:p=>p.next.length?f1(p.next[0].opp_ga_pg):"–"},
};
const VIEWS = {
  "Start7":    {cols:["name","hpi_avg","hpi_form3","hpi_form_delta","hpi_last","spark","hpi_std","hpi_md","min_pg","next","next_ga"], sort:"hpi_avg"},
  "HPI":       {cols:["name","hpi_avg","hpi_form3","hpi_last","hpi_max","hpi_std","hpi_md","hpi_avg_ex7m","hpi_prev","crunch_avg","spark"], sort:"hpi_avg"},
  "Aufsteiger":{cols:["name","hpi_form_delta","hpi_form3","hpi_avg","hpi_last","spark","games","next"], sort:"hpi_form_delta", minmd:3},
  "Angriff":   {cols:["name","goals","goals_pg","shots","acc","g7","acc7","g_fb","assists","assists_pg","tf","hpi_avg"], sort:"goals"},
  "Abwehr":    {cols:["name","steals","blocks","two_min","red","tf_pg","min_pg","hpi_avg"], sort:"steals"},
  "Torhüter":  {cols:["name","hpi_avg","hpi_form3","saves","saves_pg","save_pct","ga","save7_pct","games","min_pg","next","next_ga"], sort:"hpi_avg", pos:"TW"},
  "Tracking":  {cols:["name","dist_km","speed_kmh","min_pg","games","hpi_avg"], sort:"speed_kmh"},
  "Gegner":    {cols:["name","next","next_rank","next_ga","hpi_avg","hpi_form3"], sort:"next_ga"},
};
let view="Start7", sortKey="hpi_avg", sortDir=-1;
const $=s=>document.querySelector(s);
$("#sub").textContent = `${DATA.season} · Stand ${DATA.generated} · ${P.length} Spieler · HPI bis Spieltag ${MD[MD.length-1]??"–"}`;
[...new Set(P.map(p=>p.team))].filter(Boolean).sort().forEach(t=>$("#team").insertAdjacentHTML("beforeend",`<option>${t}</option>`));
["TW","LA","RL","RM","RR","RA","KM"].forEach(p=>$("#pos").insertAdjacentHTML("beforeend",`<option value="${p}">${POSN[p]}</option>`));
Object.keys(VIEWS).forEach(v=>$("#tabs").insertAdjacentHTML("beforeend",`<button data-v="${v}">${v}</button>`));
$("#tabs").onclick=e=>{const b=e.target.closest("button"); if(!b) return; view=b.dataset.v; sortKey=VIEWS[view].sort; sortDir=-1;
  if(VIEWS[view].pos) $("#pos").value=VIEWS[view].pos; else if($("#pos").value==="TW" && view!=="Torhüter") $("#pos").value="";
  if(VIEWS[view].minmd!=null && +$("#minmd").value<VIEWS[view].minmd) $("#minmd").value=String(VIEWS[view].minmd); render();};
["q","team","pos","minmd"].forEach(id=>$("#"+id).addEventListener("input",render));
function render(){
  document.querySelectorAll("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.v===view));
  const V=VIEWS[view], q=$("#q").value.trim().toLowerCase(), tm=$("#team").value, ps=$("#pos").value, mn=+$("#minmd").value;
  let rows=P.filter(p=>(!q||(p.name+" "+p.team).toLowerCase().includes(q))&&(!tm||p.team===tm)&&(!ps||p.pos===ps)&&((p.games||0)>=mn||(p.hpi_md||0)>=mn));
  const c=COLS[sortKey]; rows.sort((a,b)=>{const x=c.v(a),y=c.v(b); if(x==null&&y==null) return 0; if(x==null) return 1; if(y==null) return -1; return (x<y?-1:x>y?1:0)*sortDir;});
  $("#tbl thead").innerHTML="<tr>"+V.cols.map(k=>`<th data-k="${k}" class="${(COLS[k].cls||"")+(k===sortKey?" sorted":"")}">${COLS[k].h}${k===sortKey?(sortDir<0?" ↓":" ↑"):""}</th>`).join("")+"</tr>";
  $("#tbl tbody").innerHTML = rows.length ? rows.map(p=>`<tr data-id="${p.id}|${p.team_id}">`+V.cols.map(k=>`<td class="${COLS[k].cls||""}">${(COLS[k].r||(p=>f0(COLS[k].v(p))))(p)}</td>`).join("")+"</tr>").join("")
    : `<tr><td colspan="${V.cols.length}" class="empty">Keine Spieler gefunden</td></tr>`;
  $("#count").textContent=`${rows.length} Spieler`;
}
$("#tbl thead").onclick=e=>{const th=e.target.closest("th"); if(!th) return; const k=th.dataset.k; if(k===sortKey) sortDir*=-1; else {sortKey=k; sortDir=k==="name"||k==="team_abbr"||k==="next_rank"?1:-1;} render();};
$("#tbl tbody").onclick=e=>{const tr=e.target.closest("tr[data-id]"); if(!tr) return; const [id,tid]=tr.dataset.id.split("|"); detail(P.find(p=>p.id===id&&String(p.team_id)===tid));};
function bigChart(s){ const w=600,h=180,l=28,b=22; const pts=s.map((v,i)=>[i,v]).filter(p=>p[1]!=null); if(!pts.length) return '<p class="mut">Noch kein HPI-Wert.</p>';
  const n=Math.max(s.length-1,1), x=i=>l+i/n*(w-l-10), y=v=>8+(100-Math.max(30,Math.min(100,v)))/70*(h-b-8);
  let g=""; [40,50,60,70,80,90,100].forEach(v=>g+=`<line x1="${l}" x2="${w-6}" y1="${y(v)}" y2="${y(v)}" stroke="currentColor" stroke-opacity=".1"/><text x="${l-4}" y="${y(v)+4}" font-size="10" text-anchor="end" fill="currentColor" opacity=".55">${v}</text>`);
  MD.forEach((m,i)=>g+=`<text x="${x(i)}" y="${h-6}" font-size="10" text-anchor="middle" fill="currentColor" opacity=".55">${m}</text>`);
  const d=path(s,x,y);
  const dots=pts.map(p=>`<circle cx="${x(p[0])}" cy="${y(p[1])}" r="4" fill="${hpiColor(p[1])}" stroke="var(--card)" stroke-width="1.5"><title>ST ${MD[p[0]]}: ${f1(p[1])}</title></circle><text x="${x(p[0])}" y="${y(p[1])-8}" font-size="10" text-anchor="middle" fill="currentColor">${Math.round(p[1])}</text>`).join("");
  return `<svg viewBox="0 0 ${w} ${h}" width="100%" role="img" aria-label="HPI-Verlauf">${g}<path d="${d}" fill="none" stroke="var(--acc)" stroke-width="2"/>${dots}</svg><div class="small mut">x: Spieltag · y: HPI (Lücke = kein Einsatz/kein Wert)</div>`; }
function kpi(label,val){return `<div class="kpi"><b>${val}</b><span>${label}</span></div>`;}
function detail(p){ if(!p) return; const gk=p.pos==="TW";
  const nx=p.next.map(n=>`<li>${n.date}: ${n.home?"Heim":"Auswärts"} gegen <b>${n.opp_name}</b> (Pl. ${f0(n.opp_rank)}, ${f1(n.opp_ga_pg)} Gegentore/Sp)</li>`).join("")||"<li>–</li>";
  $("#dlgc").innerHTML=`<button class="close" onclick="dlg.close()" aria-label="Schließen">×</button>
  <h2>${p.name}</h2><div class="mut">${p.team} · ${POSN[p.pos]||p.pos}${p.age?` · ${p.age} J.`:""}${p.nat?` · ${p.nat}`:""}${p.height?` · ${p.height} cm`:""}</div>
  <div class="sect">HPI-Verlauf</div>${bigChart(p.hpi_series)}
  <div class="grid">${kpi("Ø HPI",hpiCell(p.hpi_avg))}${kpi("Form (letzte 3)",hpiCell(p.hpi_form3))}${kpi("Δ Form",delta(p.hpi_form_delta))}${kpi("Letzter "+trend(p.hpi_trend),hpiCell(p.hpi_last))}
  ${kpi("Max",f1(p.hpi_max))}${kpi("Streuung σ",f1(p.hpi_std))}${kpi("HPI-Spiele",f0(p.hpi_md))}${kpi("Ø ohne 7m",f1(p.hpi_avg_ex7m))}${kpi("Ø Vorsaison",hpiCell(p.hpi_prev))}${kpi("Crunchtime Ø",f1(p.crunch_avg))}</div>
  <div class="sect">Statistik</div><div class="grid">
  ${kpi("Spiele",f0(p.games))}${kpi("Min/Spiel",f1(p.min_pg))}
  ${gk?kpi("Paraden",f0(p.saves))+kpi("Paraden %",f1(p.save_pct))+kpi("Paraden/Sp",f1(p.saves_pg))+kpi("Gegentore",f0(p.ga))+kpi("7m gehalten",`${f0(p.saves7)} (${f1(p.save7_pct)}%)`)+kpi("Außen %",f1(p.save_wing_pct))+kpi("Rückraum %",f1(p.save_back_pct)):""}
  ${kpi("Tore",f0(p.goals))}${gk?"":kpi("Tore/Sp",f1(p.goals_pg))+kpi("Würfe",f0(p.shots))+kpi("Quote",p.acc==null?"–":f1(p.acc)+"%")+kpi("7m",`${f0(p.g7)} (${p.acc7==null?"–":f1(p.acc7)+"%"})`)+kpi("Gegenstoß",f0(p.g_fb))+kpi("Außen / Rückraum",`${f0(p.g_wing)} / ${f0(p.g_back)}`)+kpi("9m / Kreis",`${f0(p.g_9m)} / ${f0(p.g_pivot)}`)}
  ${kpi("Assists",f0(p.assists))}${kpi("Steals",f0(p.steals))}${gk?"":kpi("Blocks",f0(p.blocks))+kpi("Techn. Fehler",f0(p.tf))}
  ${kpi("2 Min / Rot",`${f0(p.two_min)} / ${f0(p.red)}`)}${kpi("Laufweg",p.dist_km==null?"–":f1(p.dist_km)+" km")}${kpi("Max-Speed",p.speed_kmh==null?"–":f1(p.speed_kmh)+" km/h")}</div>
  <div class="sect">Nächste Spiele</div><ul class="small">${nx}</ul>
  <a href="${p.url}" target="_blank" rel="noopener">Spielerprofil auf opel-hbl.de ↗</a>`;
  dlg.showModal(); }
const dlg=$("#dlg"); dlg.addEventListener("click",e=>{if(e.target===dlg) dlg.close();});
render();
</script>
</body>
</html>
"""


def write_html(path: Path, players: list[dict], md_numbers: list[int], meta: dict) -> None:
    payload = dict(players=players, matchdays=md_numbers, **meta)
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), default=str).replace("</", "<\\/")
    path.write_text(HTML_TEMPLATE.replace("__DATA__", data), encoding="utf-8")


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Opel HBL Spielerstatistiken + HPI für Start7")
    ap.add_argument("--out", default="site", help="Ausgabeverzeichnis (Standard: site)")
    ap.add_argument("--dump", help="Rohantworten der APIs in dieses Verzeichnis speichern")
    ap.add_argument("--from-dump", help="Statt APIs gespeicherte Rohdaten aus diesem Verzeichnis nutzen")
    ap.add_argument("--season-id", help="Synergy-Saison-UUID manuell setzen")
    ap.add_argument("--hpi-season", type=int, help="HPI-Saison-ID manuell setzen (z. B. 260 = 2026/27)")
    ap.add_argument("--now", help="Referenzzeitpunkt ISO (nur für Tests)")
    a = ap.parse_args(argv)

    t0 = time.time()
    f = Fetcher(Path(a.dump) if a.dump else None, Path(a.from_dump) if a.from_dump else None)
    raw = fetch_all(f, a.season_id, a.hpi_season)
    now = dt.datetime.fromisoformat(a.now).astimezone(dt.timezone.utc) if a.now else dt.datetime.now(dt.timezone.utc)
    teams = build_teams(raw, now)
    players, md_numbers = build_players(raw, teams, now.date())
    meta = dict(season=raw["hpi_cur"]["name"], generated=now.astimezone(TZ).strftime("%d.%m.%Y %H:%M"),
                synergy_season=raw["synergy_season"])

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    write_html(out / "index.html", players, md_numbers, meta)
    write_excel(out / "hbl_start7.xlsx", players, md_numbers, teams, meta)
    (out / "players.json").write_text(json.dumps(players, ensure_ascii=False, default=str, indent=1), encoding="utf-8")
    log(f"Fertig: {len(players)} Spieler, {len(teams)} Teams, Spieltage {md_numbers} -> {out}/ "
        f"({time.time() - t0:.0f}s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
