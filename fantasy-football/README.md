# ESPN fantasy league history export

Pulls every season of an ESPN fantasy football league (2010 onward) into CSVs and one Excel workbook.

## Run it

```bash
pip install requests openpyxl

# Only needed for private leagues. Copy both cookies from espn.com in your browser:
# DevTools → Application → Cookies → https://www.espn.com
export ESPN_S2='AEB...'
export ESPN_SWID='{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'

python espn_history.py --league-id 123456 --start 2010
```

Your league ID is the `leagueId=` number in your league's URL.

Options:
- `--skip-weekly` skips weekly lineups and transactions (about 1 request per season instead of about 35).
- `--refresh` re-downloads instead of using cached data.
- `--end YEAR` sets the last season to pull.

## Output (`output/`)

| File | Contents |
|---|---|
| `teams.csv` | team per season: owner, W-L-T, PF/PA, seed, final rank, trade/add counts |
| `matchups.csv` | every game's score, regular season and playoffs |
| `draft.csv` | every pick with player, position, auction price, keeper flag |
| `lineups.csv` | weekly roster per team: slot, started or benched, player points |
| `transactions.csv` | adds, drops, waivers and trades |
| `owners_all_time.csv` | career record, titles, playoff trips, best week per owner |
| `head_to_head.csv` | owner-vs-owner records, regular season and playoffs |
| `league_history.xlsx` | all of the above as sheets |
| `raw/` | untouched ESPN JSON, cached so re-runs are fast |

## What ESPN still has

Seasons before 2018 come from ESPN's `leagueHistory` endpoint, which keeps less detail:
- **Records, scores and drafts:** generally complete back to 2010.
- **Weekly lineups and player points:** sparse or missing before about 2018–2019.
- **Transactions:** mostly gone before about 2018. The per-team counts of trades and adds in `teams.csv` often survive even when the individual moves don't.
