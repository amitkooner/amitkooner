#!/usr/bin/env python3
"""
Export the full history of an ESPN fantasy football league.

Pulls every season from --start to --end and writes CSVs (plus an optional
Excel workbook) covering:

  teams.csv          one row per team per season: owner, record, PF/PA, final rank, move counts
  matchups.csv       every matchup score, regular season and playoffs
  draft.csv          every draft pick (snake or auction), with player names
  lineups.csv        weekly starters/bench with player points (where ESPN still has it)
  transactions.csv   adds, drops, waivers and trades (where ESPN still has it)
  owners_all_time.csv   career record per owner across all seasons
  head_to_head.csv      owner-vs-owner all-time records

Raw API responses are cached under <out>/raw/ so re-runs are fast and nothing is lost.

Private leagues need two cookies from a logged-in browser session on espn.com:
    export ESPN_S2='AEB...'        # the espn_s2 cookie
    export ESPN_SWID='{ABCD-...}'  # the SWID cookie, including braces

Usage:
    python espn_history.py --league-id 123456 --start 2010
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import sys
import time
from collections import defaultdict
from pathlib import Path

import requests

BASE = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl"
HISTORY_CUTOFF = 2018  # seasons before this live under /leagueHistory/

LINEUP_SLOTS = {
    0: "QB", 1: "TQB", 2: "RB", 3: "RB/WR", 4: "WR", 5: "WR/TE", 6: "TE", 7: "OP",
    8: "DT", 9: "DE", 10: "LB", 11: "DL", 12: "CB", 13: "S", 14: "DB", 15: "DP",
    16: "D/ST", 17: "K", 18: "P", 19: "HC", 20: "BE", 21: "IR", 23: "FLEX", 24: "ER",
}
POSITIONS = {1: "QB", 2: "RB", 3: "WR", 4: "TE", 5: "K", 7: "P", 9: "DT", 10: "DE",
             11: "LB", 12: "CB", 13: "S", 14: "HC", 16: "D/ST"}
STARTER_EXCLUDE = {20, 21}  # bench, IR

# Message types in the league activity feed (communication endpoint)
ACTIVITY_TYPES = {178: "FA ADDED", 179: "DROPPED", 180: "WAIVER ADDED",
                  181: "DROPPED", 239: "DROPPED", 244: "TRADED"}


class ESPN:
    def __init__(self, league_id: int, out: Path, espn_s2: str | None, swid: str | None,
                 refresh: bool = False, delay: float = 0.4):
        self.league_id = league_id
        self.raw_dir = out / "raw"
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.refresh = refresh
        self.delay = delay
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "espn-league-history-export"
        if espn_s2 and swid:
            self.s.cookies.set("espn_s2", espn_s2, domain=".espn.com")
            self.s.cookies.set("SWID", swid, domain=".espn.com")

    def _url(self, season: int, suffix: str = "") -> tuple[str, dict]:
        if season < HISTORY_CUTOFF:
            return f"{BASE}/leagueHistory/{self.league_id}{suffix}", {"seasonId": season}
        return f"{BASE}/seasons/{season}/segments/0/leagues/{self.league_id}{suffix}", {}

    def get(self, season: int, cache_key: str, views: list[str], suffix: str = "",
            extra: dict | None = None, fantasy_filter: dict | None = None):
        cache = self.raw_dir / f"{season}_{cache_key}.json"
        if cache.exists() and not self.refresh:
            return json.loads(cache.read_text())

        url, params = self._url(season, suffix)
        params = {**params, **(extra or {}), "view": views}
        headers = {"x-fantasy-filter": json.dumps(fantasy_filter)} if fantasy_filter else {}
        for attempt in range(4):
            r = self.s.get(url, params=params, headers=headers, timeout=30)
            if r.status_code == 429 or r.status_code >= 500:
                time.sleep(2 ** attempt)
                continue
            break
        time.sleep(self.delay)
        if r.status_code in (401, 403):
            raise PermissionError(
                f"{season}: ESPN returned {r.status_code}. The league is private (or the cookies "
                "are wrong/expired): set ESPN_S2 and ESPN_SWID.")
        if r.status_code == 404:
            return None
        r.raise_for_status()
        data = r.json()
        if season < HISTORY_CUTOFF and isinstance(data, list):
            data = data[0] if data else None
        cache.write_text(json.dumps(data))
        return data

    def players(self, season: int) -> dict[int, dict]:
        """Season-wide player id -> {name, position} map."""
        cache = self.raw_dir / f"{season}_players.json"
        if cache.exists() and not self.refresh:
            data = json.loads(cache.read_text())
        else:
            r = self.s.get(f"{BASE}/seasons/{season}/players",
                           params={"scoringPeriodId": 0, "view": "players_wl"},
                           headers={"x-fantasy-filter": json.dumps({"filterActive": None})},
                           timeout=60)
            time.sleep(self.delay)
            data = r.json() if r.ok else []
            cache.write_text(json.dumps(data))
        return {p["id"]: {"name": p.get("fullName"),
                          "position": POSITIONS.get(p.get("defaultPositionId"), "")}
                for p in data if isinstance(p, dict) and "id" in p}


# ---------------------------------------------------------------------------
# helpers

def team_name(t: dict) -> str:
    if t.get("name"):
        return t["name"].strip()
    return f"{t.get('location', '')} {t.get('nickname', '')}".strip()


def member_names(league: dict) -> dict[str, str]:
    out = {}
    for m in league.get("members", []):
        full = f"{m.get('firstName', '')} {m.get('lastName', '')}".strip()
        out[m["id"]] = full or m.get("displayName") or m["id"]
    return out


def owner_of(t: dict, members: dict[str, str]) -> tuple[str, str]:
    """Returns (owner_id, owner_name). Uses the primary owner when there are co-owners."""
    ids = t.get("owners") or ([t["primaryOwner"]] if t.get("primaryOwner") else [])
    primary = t.get("primaryOwner") or (ids[0] if ids else "")
    names = " & ".join(members.get(i, i) for i in ids)
    return primary, names


def scoring_periods(league: dict) -> dict[int, list[int]]:
    """matchup period -> scoring periods (weeks)."""
    mp = league.get("settings", {}).get("scheduleSettings", {}).get("matchupPeriods", {})
    if mp:
        return {int(k): v for k, v in mp.items()}
    weeks = {m.get("matchupPeriodId") for m in league.get("schedule", [])}
    return {w: [w] for w in sorted(x for x in weeks if x)}


def write_csv(path: Path, rows: list[dict]):
    if not rows:
        path.write_text("")
        return
    fields = list(dict.fromkeys(k for r in rows for k in r))
    with path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def ts(ms) -> str:
    return dt.datetime.fromtimestamp(ms / 1000).isoformat(sep=" ", timespec="minutes") if ms else ""


# ---------------------------------------------------------------------------
# per-season extraction

def export_season(api: ESPN, season: int, tables: dict[str, list], skip_weekly: bool):
    league = api.get(season, "league",
                     ["mTeam", "mSettings", "mStandings", "mMatchup", "mMatchupScore", "mDraftDetail"])
    if not league:
        print(f"  {season}: no league found, skipping")
        return False

    members = member_names(league)
    players = api.players(season)
    teams = {t["id"]: t for t in league.get("teams", [])}
    tname = {tid: team_name(t) for tid, t in teams.items()}
    towner = {tid: owner_of(t, members) for tid, t in teams.items()}

    def pname(pid):
        return players.get(pid, {}).get("name") or ""

    # --- teams / standings
    for tid, t in teams.items():
        rec = t.get("record", {}).get("overall", {})
        tc = t.get("transactionCounter", {})
        oid, oname = towner[tid]
        tables["teams"].append({
            "season": season, "team_id": tid, "team_name": tname[tid],
            "abbrev": t.get("abbrev"), "owner_id": oid, "owner": oname,
            "wins": rec.get("wins"), "losses": rec.get("losses"), "ties": rec.get("ties"),
            "points_for": round(rec.get("pointsFor", t.get("points", 0)) or 0, 2),
            "points_against": round(rec.get("pointsAgainst", 0) or 0, 2),
            "playoff_seed": t.get("playoffSeed"),
            "final_rank": t.get("rankCalculatedFinal") or t.get("rankFinal"),
            "draft_day_projected_rank": t.get("draftDayProjectedRank"),
            "acquisitions": tc.get("acquisitions"), "drops": tc.get("drops"),
            "trades": tc.get("trades"), "faab_spent": tc.get("acquisitionBudgetSpent"),
            "lineup_moves": tc.get("moveToActive"),
        })

    # --- matchups
    for m in league.get("schedule", []):
        home, away = m.get("home", {}), m.get("away")
        h_id = home.get("teamId")
        a_id = away.get("teamId") if away else None
        tables["matchups"].append({
            "season": season, "matchup_period": m.get("matchupPeriodId"),
            "playoff_tier": m.get("playoffTierType", "NONE"),
            "home_team_id": h_id, "home_team": tname.get(h_id),
            "home_owner_id": towner.get(h_id, ("", ""))[0], "home_owner": towner.get(h_id, ("", ""))[1],
            "home_score": round(home.get("totalPoints", 0) or 0, 2),
            "away_team_id": a_id, "away_team": tname.get(a_id) if a_id else "BYE",
            "away_owner_id": towner.get(a_id, ("", ""))[0] if a_id else "",
            "away_owner": towner.get(a_id, ("", ""))[1] if a_id else "",
            "away_score": round(away.get("totalPoints", 0) or 0, 2) if away else None,
            "winner": m.get("winner"),
        })

    # --- draft
    dd = league.get("draftDetail", {})
    auction = league.get("settings", {}).get("draftSettings", {}).get("type") == "AUCTION"
    for p in dd.get("picks", []):
        tid = p.get("teamId")
        pid = p.get("playerId")
        tables["draft"].append({
            "season": season, "overall_pick": p.get("overallPickNumber") or p.get("id"),
            "round": p.get("roundId"), "round_pick": p.get("roundPickNumber"),
            "team_id": tid, "team_name": tname.get(tid), "owner": towner.get(tid, ("", ""))[1],
            "player_id": pid, "player": pname(pid),
            "position": players.get(pid, {}).get("position", ""),
            "bid_amount": p.get("bidAmount") if auction else "",
            "keeper": p.get("keeper"),
            "nominating_team_id": p.get("nominatingTeamId") if auction else "",
        })

    if skip_weekly:
        return True

    # --- weekly lineups + transactions
    periods = scoring_periods(league)
    final_week = league.get("status", {}).get("finalScoringPeriod") or max(
        (w for ws in periods.values() for w in ws), default=0)
    current = league.get("scoringPeriodId") or final_week
    seen_tx = set()
    n_lineups = n_tx = 0
    for mp, weeks in sorted(periods.items()):
        for week in weeks:
            if week > min(final_week, current):
                continue
            box = api.get(season, f"box_w{week:02d}", ["mMatchupScore", "mScoreboard"],
                          extra={"scoringPeriodId": week},
                          fantasy_filter={"schedule": {"filterMatchupPeriodIds": {"value": [mp]}}})
            for m in (box or {}).get("schedule", []):
                if m.get("matchupPeriodId") != mp:
                    continue
                for side in ("home", "away"):
                    s = m.get(side) or {}
                    roster = s.get("rosterForCurrentScoringPeriod") or {}
                    for e in roster.get("entries", []):
                        ppe = e.get("playerPoolEntry", {})
                        pl = ppe.get("player", {})
                        slot = e.get("lineupSlotId")
                        tables["lineups"].append({
                            "season": season, "matchup_period": mp, "week": week,
                            "team_id": s.get("teamId"), "team_name": tname.get(s.get("teamId")),
                            "owner": towner.get(s.get("teamId"), ("", ""))[1],
                            "player_id": e.get("playerId"),
                            "player": pl.get("fullName") or pname(e.get("playerId")),
                            "position": POSITIONS.get(pl.get("defaultPositionId"), ""),
                            "slot": LINEUP_SLOTS.get(slot, slot),
                            "started": slot not in STARTER_EXCLUDE,
                            "points": round(ppe.get("appliedStatTotal", 0) or 0, 2),
                        })
                        n_lineups += 1

            tx = api.get(season, f"tx_w{week:02d}", ["mTransactions2"], extra={"scoringPeriodId": week})
            for t in (tx or {}).get("transactions", []):
                if t.get("status") != "EXECUTED" or t.get("type") in ("ROSTER", "FUTURE_ROSTER", "DRAFT"):
                    continue
                if t["id"] in seen_tx:
                    continue
                seen_tx.add(t["id"])
                for it in t.get("items", []):
                    if it.get("type") == "LINEUP":
                        continue
                    pid = it.get("playerId")
                    tables["transactions"].append({
                        "season": season, "week": week, "date": ts(t.get("processDate") or t.get("proposedDate")),
                        "source": "transactions", "transaction_id": t["id"], "type": t.get("type"),
                        "action": it.get("type"),
                        "team_id": t.get("teamId"), "team_name": tname.get(t.get("teamId")),
                        "from_team": tname.get(it.get("fromTeamId"), "") if it.get("fromTeamId", -1) >= 0 else "",
                        "to_team": tname.get(it.get("toTeamId"), "") if it.get("toTeamId", -1) >= 0 else "",
                        "player_id": pid, "player": pname(pid), "bid_amount": t.get("bidAmount"),
                    })
                    n_tx += 1

    # Older seasons rarely have mTransactions2 data; try the league activity feed too.
    if n_tx == 0:
        n_tx = export_activity_feed(api, season, tables, tname, pname)

    print(f"  {season}: {len(teams)} teams, {len(league.get('schedule', []))} matchups, "
          f"{len(dd.get('picks', []))} picks, {n_lineups} lineup rows, {n_tx} transaction rows")
    return True


def export_activity_feed(api: ESPN, season: int, tables, tname, pname) -> int:
    flt = {"topics": {"filterType": {"value": ["ACTIVITY_TRANSACTIONS"]},
                      "limit": 2000, "limitPerMessageSet": {"value": 25}, "offset": 0,
                      "sortMessageDate": {"sortPriority": 1, "sortAsc": False},
                      "sortFor": {"sortPriority": 2, "sortAsc": False},
                      "filterIncludeMessageTypeIds": {"value": list(ACTIVITY_TYPES)}}}
    try:
        data = api.get(season, "activity", ["kona_league_communication"], suffix="/communication/",
                       fantasy_filter=flt)
    except requests.HTTPError:
        return 0
    n = 0
    for topic in (data or {}).get("topics", []):
        for msg in topic.get("messages", []):
            mt = msg.get("messageTypeId")
            team_id = msg.get("from") if mt == 244 else msg.get("for") if mt == 239 else msg.get("to")
            pid = msg.get("targetId")
            tables["transactions"].append({
                "season": season, "week": "", "date": ts(topic.get("date")),
                "source": "activity_feed", "transaction_id": topic.get("id"),
                "type": ACTIVITY_TYPES.get(mt, mt), "action": ACTIVITY_TYPES.get(mt, mt),
                "team_id": team_id, "team_name": tname.get(team_id, ""),
                "from_team": tname.get(msg.get("from"), "") if mt == 244 else "",
                "to_team": tname.get(msg.get("to"), "") if mt == 244 else "",
                "player_id": pid, "player": pname(pid), "bid_amount": "",
            })
            n += 1
    return n


# ---------------------------------------------------------------------------
# all-time aggregates

def all_time(tables: dict[str, list]):
    # Group by ESPN member id (stable even if someone renames); display their latest name.
    owners = defaultdict(lambda: defaultdict(float))
    season_owner = {(r["season"], r["team_id"]): r["owner_id"] for r in tables["teams"]}
    display = {}
    for r in sorted(tables["teams"], key=lambda r: r["season"]):
        display[r["owner_id"]] = r["owner"]

    for r in tables["teams"]:
        o = owners[r["owner_id"]]
        o["seasons"] += 1
        for k in ("wins", "losses", "ties"):
            o[k] += r[k] or 0
        o["points_for"] += r["points_for"] or 0
        o["points_against"] += r["points_against"] or 0
        o["championships"] += 1 if r["final_rank"] == 1 else 0
        o["runner_ups"] += 1 if r["final_rank"] == 2 else 0
        o["last_place"] += 1 if r["final_rank"] and r["final_rank"] == max(
            x["final_rank"] or 0 for x in tables["teams"] if x["season"] == r["season"]) else 0
        o["trades"] += r["trades"] or 0
        o["acquisitions"] += r["acquisitions"] or 0

    # playoff appearances: any winners-bracket matchup
    playoff = {(m["season"], tid) for m in tables["matchups"] if m["playoff_tier"] == "WINNERS_BRACKET"
               for tid in (m["home_team_id"], m["away_team_id"]) if tid is not None}
    for season, tid in playoff:
        if (season, tid) in season_owner:
            owners[season_owner[(season, tid)]]["playoff_appearances"] += 1

    high = defaultdict(float)
    h2h = defaultdict(lambda: defaultdict(float))
    for m in tables["matchups"]:
        if m["away_team_id"] is None or m["winner"] not in ("HOME", "AWAY", "TIE"):
            continue
        a, b = m["home_owner_id"], m["away_owner_id"]
        high[a] = max(high[a], m["home_score"])
        high[b] = max(high[b], m["away_score"])
        kind = "regular" if m["playoff_tier"] == "NONE" else "playoff"
        for me, opp, mine, theirs in ((a, b, m["home_score"], m["away_score"]),
                                      (b, a, m["away_score"], m["home_score"])):
            rec = h2h[(me, opp)]
            res = "ties" if m["winner"] == "TIE" else ("wins" if mine > theirs else "losses")
            rec[f"{kind}_{res}"] += 1
            rec["points_for"] += mine
            rec["points_against"] += theirs

    owner_rows = []
    for oid, o in owners.items():
        g = o["wins"] + o["losses"] + o["ties"]
        owner_rows.append({
            "owner": display.get(oid, oid), "seasons": int(o["seasons"]),
            "wins": int(o["wins"]), "losses": int(o["losses"]), "ties": int(o["ties"]),
            "win_pct": round((o["wins"] + 0.5 * o["ties"]) / g, 3) if g else None,
            "points_for": round(o["points_for"], 2), "points_against": round(o["points_against"], 2),
            "avg_pf_per_season": round(o["points_for"] / o["seasons"], 2),
            "championships": int(o["championships"]), "runner_ups": int(o["runner_ups"]),
            "last_place_finishes": int(o["last_place"]),
            "playoff_appearances": int(o["playoff_appearances"]),
            "highest_week": round(high.get(oid, 0), 2),
            "trades": int(o["trades"]), "acquisitions": int(o["acquisitions"]),
        })
    owner_rows.sort(key=lambda r: (-r["championships"], -(r["win_pct"] or 0)))

    h2h_rows = []
    for (me, opp), r in sorted(h2h.items(), key=lambda kv: (display.get(kv[0][0], ""), display.get(kv[0][1], ""))):
        h2h_rows.append({
            "owner": display.get(me, me), "opponent": display.get(opp, opp),
            "reg_wins": int(r["regular_wins"]), "reg_losses": int(r["regular_losses"]),
            "reg_ties": int(r["regular_ties"]),
            "playoff_wins": int(r["playoff_wins"]), "playoff_losses": int(r["playoff_losses"]),
            "points_for": round(r["points_for"], 2), "points_against": round(r["points_against"], 2),
        })
    return owner_rows, h2h_rows


def write_xlsx(path: Path, sheets: dict[str, list[dict]]):
    try:
        from openpyxl import Workbook
    except ImportError:
        print("openpyxl not installed; skipping Excel workbook")
        return
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name[:31])
        if not rows:
            continue
        fields = list(dict.fromkeys(k for r in rows for k in r))
        ws.append(fields)
        for r in rows:
            ws.append([r.get(f) for f in fields])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main():
    today = dt.date.today()
    default_end = today.year if today.month >= 8 else today.year - 1
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--league-id", type=int, required=True)
    ap.add_argument("--start", type=int, default=2010)
    ap.add_argument("--end", type=int, default=default_end)
    ap.add_argument("--out", type=Path, default=Path(__file__).parent / "output")
    ap.add_argument("--skip-weekly", action="store_true",
                    help="skip per-week lineups/transactions (much faster: ~1 request per season)")
    ap.add_argument("--refresh", action="store_true", help="ignore cached raw responses")
    ap.add_argument("--no-xlsx", action="store_true")
    args = ap.parse_args()

    s2, swid = os.environ.get("ESPN_S2"), os.environ.get("ESPN_SWID")
    if not (s2 and swid):
        print("ESPN_S2 / ESPN_SWID not set: this only works if the league is public.")
    args.out.mkdir(parents=True, exist_ok=True)
    api = ESPN(args.league_id, args.out, s2, swid, refresh=args.refresh)

    tables = {k: [] for k in ("teams", "matchups", "draft", "lineups", "transactions")}
    for season in range(args.start, args.end + 1):
        try:
            export_season(api, season, tables, args.skip_weekly)
        except PermissionError as e:
            sys.exit(str(e))
        except requests.RequestException as e:
            print(f"  {season}: request failed ({e}); continuing")

    owner_rows, h2h_rows = all_time(tables)
    sheets = {**tables, "owners_all_time": owner_rows, "head_to_head": h2h_rows}
    for name, rows in sheets.items():
        write_csv(args.out / f"{name}.csv", rows)
    if not args.no_xlsx:
        write_xlsx(args.out / "league_history.xlsx", sheets)
    print(f"\nWrote {', '.join(f'{k} ({len(v)})' for k, v in sheets.items())} to {args.out}/")


if __name__ == "__main__":
    main()
