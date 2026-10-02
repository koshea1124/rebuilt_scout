# Project status (Oct 1, 2026)

Handoff notes for whoever picks this up next, including a future Claude chat:
"Read PROJECT_STATUS.md in github.com/koshea1124/rebuilt_scout and continue."

## What exists

- **Scouting app** (GitHub Pages): https://koshea1124.github.io/rebuilt_scout/
  - Match tab: won-auto toggle locks inactive shifts, fuel pads (−5 −1 +1 +5 +10 +25 +50) for auto,
    transition, shifts 1 to 4 and end game; inactive-hub role per shift; climb, defense, driver, fouls, notes.
  - Robot position picker auto-fills team # from the Blue Alliance schedule.
  - Pit tab and Teams tab (sortable rankings, team detail, schedule loader, missing-report list).
  - Offline-first (Firestore cache + service worker), team passcode, installs to home screen.
  - Cardinal/gold theme, Russo One headings, Roboto body, team logo icons.
- **Firebase project** `rebuilt-scouting-f33c1`: Firestore collections `matches`, `pit`, `schedules`,
  `members`, `config/secret` (passcode), `settings/tba` (Blue Alliance key). Rules in `firestore.rules`.
- **Google Sheet sync** (`sheets/`): Team Averages (dropdown sort), Fuel Accuracy vs The Blue Alliance,
  Match Data, Pit Data. Setup steps in `sheets/README.md`.

## Open items

- [ ] Google Sheet: `setup` failed with "Firebase refused access". Check `appsscript.json` has the
      datastore scope and the script runs as the Firebase project owner; re-copy the latest `Code.gs`
      (it now prints Google's exact reason).
- [ ] Confirm the Firestore rules in the console include the `schedules` and `settings` blocks.
- [ ] Save the Blue Alliance key in the app (Teams tab > Blue Alliance key) if not done.
- [ ] Find Tidal Tumble's Blue Alliance event key (Beach Blitz is `2026cabl`).

## Ideas discussed, not built

- Per-scout accuracy (tower climb per robot, won-auto check) in the Sheet.
- Admin-protected delete button for bad match reports.
- CSV export, Statbotics EPA next to teams.

## Rules of thumb

- After changing any app file, bump `VERSION` in `sw.js` so phones update.
- Point values live in `index.html` (`CLIMB_PTS`, `pts`) and `sheets/Code.gs` (`CLIMB_PTS`, `AUTO_CLIMB_PTS`).
