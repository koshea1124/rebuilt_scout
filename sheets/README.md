# Scouting data in Google Sheets

A Google Sheet that pulls every match and pit report from the team's Firebase database
every 5 minutes and keeps a sortable **Team Averages** tab.

| Tab | What's in it |
| --- | --- |
| Team Averages | One row per team: matches, avg/best estimated points, fuel by period, climb rates, top climb, defense %, foul points, driver rating, robot issues, plus drivetrain/shooter/capacity from pit scouting |
| Match Preview | Pick an event and match number to see the six robots side by side, Red vs Blue, with alliance totals and a projected score |
| Fuel Accuracy | For each played qualification match and alliance: scouted vs official fuel (auto, teleop, end game, total) from The Blue Alliance, with % error |
| Scouter Accuracy | One row per scouter: average absolute fuel % error (total, auto, teleop, end game) across the complete alliances they helped scout, most accurate first |
| Match Data | Every match report, including per-shift fuel and inactive-hub roles |
| Pit Data | Every pit report |
| Schedule | Qualification schedule per event (from the app's Load schedule and from The Blue Alliance); feeds Match Preview |

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

## Match Preview

- Choose the **Event** and **Match #** at the top. The six team numbers fill in from the Schedule tab.
- No schedule yet (or planning a playoff alliance)? Type team numbers straight into the team row; the table updates.
- **Stats from** switches between this event only and all events (handy early in an event).
- Per robot: matches scouted, avg/best estimated points, avg/best fuel for auto, teleop, end game and total,
  auto and end game climb %, top climb, avg tower points, defense %, foul points, driver rating, robot issues,
  usual role when their hub is inactive, and drivetrain/shooter/capacity from pit scouting.
- Alliance columns add up the three robots (driver rating is averaged). Gold marks the alliance with the edge.
- The projected score is the sum of each robot's average estimated points (fuel + tower). It ignores fouls and defense.
- RP rows use the regional thresholds (100 and 360 fuel, 50 tower points); edit `RP_*` in `Code.gs` if they differ.

## Fuel Accuracy

- Works for events whose event code in the app is the Blue Alliance key (like `2026CABL`), using the Blue Alliance key saved in the app.
- Official fuel = hub points from The Blue Alliance (1 point per FUEL scored in an active hub), the same thing scouts count.
  Teleop = transition + shifts 1 to 4.
- % error = (scouted − official) ÷ official. Positive means scouts counted too many.
- Fuel is only reported per alliance, so the fair comparison needs all 3 robots scouted ("Complete"). Rows missing robots will read low.
- Green within 10%, yellow within 25%, red over 25%. The summary at the top averages the complete alliances.

## Scouter Accuracy

- Uses only "Complete" rows from Fuel Accuracy (all 3 robots scouted). Each scouter on that alliance gets the alliance's error.
- Errors are absolute values before averaging, so +11% and −11% average to 11%.
- Names are matched ignoring capitals and extra spaces ("Ana" = "ana ").
- Each number includes the other two scouters' mistakes. Rotate who scouts together so it evens out, and treat rows flagged "Fewer than 5 checks" as rough (change `MIN_CHECKS` in `Code.gs`).
- "Tends to count" shows whether their alliances usually came out high or low; it is a coaching hint, not part of the average.

## Notes

- The sheet only reads from Firebase; editing the sheet never changes scouting data.
- Manual sorting with Data > Sort is undone at the next sync. Use the dropdowns instead.
- If point values change, update `CLIMB_PTS` and `AUTO_CLIMB_PTS` at the top of `Code.gs` (and in `index.html`).
- Error "Firebase refused access": the account running the script isn't on the Firebase project.
  Add it in Firebase console > Project settings > Users and permissions, or run setup from the owner account.
