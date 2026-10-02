# RoboWarriors REBUILT Scout

Match and pit scouting app for the FRC 2026 game **REBUILT**, built for Team 2659 RoboWarriors.
It runs in any phone browser, installs to the home screen like an app, and keeps working with no signal in the stands.

**App:** https://koshea1124.github.io/rebuilt_scout/

## What it does

- **Match tab**: fuel counters for Auto, Teleop shifts and End Game, tower climb, inactive-shift role, defense, driver rating, robot status, fouls, notes.
- **Pit tab**: drivetrain, intake, shooter, fuel capacity, climb level, auto claims, weight, notes.
- **Teams tab**: sortable rankings (estimated points, fuel by period, climb rate, foul points, driver rating) with a per-team detail view.

All data is shared live between every scout's phone through Firebase Firestore.

## For scouts

1. Open the app link on your phone and enter the team passcode (once per phone).
2. Add it to your home screen:
   - **iPhone (Safari):** Share button, then **Add to Home Screen**.
   - **Android (Chrome):** menu, then **Install app** or **Add to Home screen**.
3. Open it once at home on Wi‑Fi before each event so everything is downloaded.
4. Set the event code at the top (for example `TIDAL26`). Everyone at an event should use the same code.

No signal? Keep scouting. Reports save on the phone and upload automatically when it reconnects. The status line at the top shows how many are waiting.

## Match schedule (The Blue Alliance)

1. Set the event code at the top to the event's Blue Alliance key, for example `2026CABL` for Beach Blitz.
2. Teams tab > **Load schedule** (needs internet; usually posted the morning of the event).
3. Each scout picks their robot position (Red 1 to Blue 3) on the Match tab. Typing a match number fills in the team and alliance.
4. The schedule card shows which played matches are missing reports.

The Blue Alliance read key is saved once in the app (Teams tab > **Blue Alliance key**) and stored in Firestore `settings/tba`, never in this repo.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | The whole app: layout, styles and logic |
| `config.js` | Firebase project settings (not secret) |
| `sw.js` | Offline support. **Bump `VERSION` after any change** so phones pick up the update |
| `manifest.webmanifest`, `icons/` | Home screen install name and icons |
| `firestore.rules` | Database security rules. Paste into Firebase console > Firestore > Rules |

## Admin setup (one time)

1. **Firebase**: Firestore database created, **Authentication > Sign-in method > Anonymous** enabled.
2. **Authorized domain**: Authentication > Settings > Authorized domains > add `koshea1124.github.io`.
3. **Passcode**: Firestore > Data > collection `config`, document `secret`, string field `code`.
4. **Rules**: copy `firestore.rules` into Firestore > Rules and publish.
5. **GitHub Pages**: repo Settings > Pages > Source: Deploy from a branch > `main` / root.

To change the passcode, edit `config/secret`. New phones need the new code; phones that already joined keep working.
To remove all phones' access, delete documents in the `members` collection.

## Scoring assumptions

Estimated points per match = fuel scored (1 each) + auto Tower L1 (15) + end game climb (L1 10, L2 20, L3 30).
Check these against the latest REBUILT Team Update and adjust `CLIMB_PTS` and the `pts` function in `index.html` if they change.

## Next season

Copy this repo, change the scouting fields in `index.html` for the new game, update the point values, and bump `VERSION` in `sw.js`.
