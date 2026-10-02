# Scouting data in Google Sheets

A Google Sheet that pulls every match and pit report from the team's Firebase database
every 5 minutes and keeps a sortable **Team Averages** tab.

| Tab | What's in it |
| --- | --- |
| Team Averages | One row per team: matches, avg/best estimated points, fuel by period, climb rates, top climb, defense %, foul points, driver rating, robot issues, plus drivetrain/shooter/capacity from pit scouting |
| Fuel Accuracy | For each played qualification match and alliance: scouted vs official fuel (auto, teleop, end game, total) from The Blue Alliance, with % error |
| Match Data | Every match report, including per-shift fuel and inactive-hub roles |
| Pit Data | Every pit report |

At the top of Team Averages, three dropdowns choose the **event**, the **column to sort by** and the **order**.
The sheet re-sorts as soon as you change one. Use **Scouting > Sync now** for an instant refresh.

## Setup (about 5 minutes, once)

Use the Google account that created the Firebase project (it must be an owner or editor of it).

1. Go to [sheets.new](https://sheets.new) to make a blank sheet. Name it something like "2659 Scouting".
2. **Extensions > Apps Script**.
3. Delete the starter code in `Code.gs` and paste in [`Code.gs`](Code.gs) from this folder.
4. Click the **gear** (Project Settings) and tick **Show "appsscript.json" manifest file in editor**.
   Back in the editor, open `appsscript.json` and replace it with [`appsscript.json`](appsscript.json).
5. Click **Save**, choose **setup** in the function dropdown, then **Run**.
6. Google asks for permission. Click **Review permissions**, pick your account, then **Advanced > Go to (project name)** and **Allow**.
   The warning appears because this is your own unpublished script.
7. Return to the sheet. The Team Averages tab fills in, and a sync runs every 5 minutes from now on.

## Fuel Accuracy

- Works for events whose event code in the app is the Blue Alliance key (like `2026CABL`), using the Blue Alliance key saved in the app.
- Official fuel = hub points from The Blue Alliance (1 point per FUEL scored in an active hub), the same thing scouts count.
  Teleop = transition + shifts 1 to 4.
- % error = (scouted − official) ÷ official. Positive means scouts counted too many.
- Fuel is only reported per alliance, so the fair comparison needs all 3 robots scouted ("Complete"). Rows missing robots will read low.
- Green within 10%, yellow within 25%, red over 25%. The summary at the top averages the complete alliances.

## Notes

- The sheet only reads from Firebase; editing the sheet never changes scouting data.
- Manual sorting with Data > Sort is undone at the next sync. Use the dropdowns instead.
- If point values change, update `CLIMB_PTS` and `AUTO_CLIMB_PTS` at the top of `Code.gs` (and in `index.html`).
- Error "Firebase refused access": the account running the script isn't on the Firebase project.
  Add it in Firebase console > Project settings > Users and permissions, or run setup from the owner account.
