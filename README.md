# BizMatch Scheduler

A web app that builds optimal, conflict-free 1:1 business matching schedules (startup × investor / corporate speed meetings). You set up the event and participants, click **Generate schedule**, and get a schedule you can download as Excel. It runs on Streamlit, and the scheduling is done by Google OR-Tools CP-SAT.

```
bizmatch/
├── app.py                 # the whole app: parsing, CP-SAT model, views, Excel export
├── requirements.txt
├── .streamlit/config.toml # theme colour and upload limit
└── sample_data/           # example startups.csv and hosts.csv
```

## What it does

**Tab 1. Event & Participants**
- Multi-day events (one row per day, with dates), day start and end, meeting length (default 20 min), optional gap between meetings, lunch break, number of tables, time zone.
- **Sessions** such as `Session A 10:00-12:00; Session B 13:00-15:00; Session C 15:00-17:00`. They group the timeline exactly like the team's sheet.
- **Pitch polishing** that runs alongside the meetings. Each startup gets one polishing slot per selected day, and never has a meeting in that slot. *Coaches* sets how many startups are polished at the same time.
- Editable tables for startups and hosts. You can also import from CSV or Excel, and download a template.
- **Paste straight from Google Sheets or Excel.** Copy the cells including your header row and paste them into *Paste startups / hosts*. Your own column names are fine: the app guesses what each column means (`Startup`, `Sector`, `Attending`, `Pitch practice` ...) and you can correct any guess before importing. A host column written like `Raymond - Focus Tech` can be split into contact and organisation. Uploaded CSV / Excel files go through the same check.
- Meeting length in 1-minute steps, and the gap between meetings in minutes plus seconds (e.g. 1 min 10 s).
- **Column names** used on the board and in Excel can be renamed (e.g. Table → Booth).
- **Save project / Open project** stores everything, including the generated schedule and live statuses, in a single `.json` file. Anyone on the team can open it later.

**Tab 2. Generate**
- Hard rules, always enforced
  - A startup and a host meet **at most once** across the whole event
  - Nobody is double booked in any slot
  - Concurrent meetings in a slot never exceed the number of tables
  - Attendance days, available windows, cutoffs, blocked times and max meetings per participant are respected
  - A host never meets a startup from one of its **Excluded Domains**
  - A startup in pitch polishing has no meeting in that slot. Polishing is scheduled before meetings, so it is only skipped when it truly cannot fit
- Host packing (pick one)
  - **Front-packed**: each host's meetings start at the first slot they are available and run without gaps. Once a host has no valid meetings left, its later slots stay empty (`-`), meaning it finishes early. The lunch break is not counted as a gap.
  - **Contiguous block**: one block with no gaps, which may start later. Use this when there are more hosts than tables, because then hosts cannot all start in the first slot.
  - **Soft**: gaps are allowed but penalised.
- Objective, in this order: total meetings, then priority startups, then fairness (concave reward, so every startup gets its 2nd meeting before anyone gets a 5th), then earlier slots.
- **Re-plan mode**: in the middle of an event, keeps every meeting that is already Ongoing, Done or No-Show, and reschedules the rest.

**Tab 3. Results & Live Board**
- **Drag-and-drop board.** Drag any meeting to another table or time. It snaps onto the cell under the pointer, and a preview tells you what will happen before you let go:
  - **Green**: it fits.
  - **Amber**: meetings in the way (same table, same startup or same host at that time) get pushed to the nearest free time where both sides are available. If one cannot fit anywhere, it goes to *Unplaced*.
  - **Red**: a clash, for example the startup or host is not available then. You can still drop it, and it stays marked with a red `!` until it is fixed.
- Pitch polishing slots can be dragged the same way. A startup's meeting at the new polishing time gets pushed away.
- Drop a meeting on **Unplaced** to take it out of the schedule, then drag it back later. Click a meeting to set its status, pin it or unschedule it. Ctrl+Z or ↶ Undo reverts a change.
- **Pitching practice** table under the board: Time, Startup, Presenter, Round (1st/2nd), Status and Focus / Mentor Note for each polishing session. It is the same data as the board. Dragging a polishing slot on the board changes its time here, and a status or note set here shows on the board.
- Meetings you move are pinned. When you generate again with *Keep pinned meetings* ticked, they stay where you put them.
- NOT HERE / ONLY notes stay hidden until you press **Show attendance notes**. In the Excel file they sit in a folded row group, opened with the `+` in the left margin.
- A timeline for each day in the team's sheet layout: Session, Time Slot, Table 1 / Status 1 … Table N / Status N, **Polishing Schedule** and **Waiting List** (startups that are neither in a meeting nor in polishing, or `None (All in Sessions)`). The live board in the app also shows resting hosts.
- Each host (organisation) keeps one colour everywhere, so you can follow a host across tables and sessions. Pink and yellow are reserved for polishing: pink is a startup's first polishing session, yellow its second.
- `NOT HERE` and `ONLY 27 NOV` style notes at the top are generated from the Days column.
- Status chips: `Scheduled`, `Ongoing`, `Done`, `No-Show`. On the event day, statuses follow the clock automatically and the current slot is marked NOW. The board refreshes every 30 s.
- The live status board lets you override any meeting, for example to mark a No-Show. A manual status always wins over the clock.
- An independent verifier re-checks every hard rule on the final schedule and shows the result.
- Startup summary (name, total meetings, partners matched on each day) and host summary.
- **Download Excel**. Sheets: `Matching Schedule`, `Pitching Practice Schedule`, `Startup Summary`, `Host Summary`, `All Meetings`. The POLISHING SCHEDULE column of the main sheet is a formula that reads the Pitching Practice sheet, so changing a startup there in Excel or Google Sheets updates the main schedule too. Every meeting has a Status dropdown, and its colour changes with the status (conditional formatting). Each meeting cell reads `Startup` with `(Contact - Organisation)` below it, filled with the host's colour, and a No-Show greys it out and strikes it through. Upload the file to Google Drive and open it with Google Sheets: the dropdowns and colours keep working there, so the sheet can be the live board on the day. The sheet prints landscape, fitted to page width.

## Participant columns

| Column | Startups | Hosts | Format |
|---|---|---|---|
| Name | ✓ | ✓ | Unique (for hosts, the organisation) |
| Contact | | ✓ | Person attending, shown as `Raymond · Focus Tech` |
| Presenter | ✓ | | Who pitches, e.g. `Jay & Serena`. Shown in the pitching practice schedule |
| Domain | ✓ | | `Fintech` or `Climate, Hardware` |
| Type | | ✓ | Investor, Corporate, Agency, Mentor, Other |
| Days | ✓ | ✓ | Blank or `All` = every day. Otherwise `Day 2`, `1,2`, or your day label. `Not here` (or `None`, `Absent`) = not attending, listed under NOT HERE |
| Available From | ✓ | ✓ | `HH:MM`. Blank = day start |
| Available Until | ✓ | ✓ | `HH:MM` cutoff. No meeting may end after it (e.g. `16:00` means "cannot meet after 16:00") |
| Blocked Times | ✓ | ✓ | Restricted windows, e.g. `Day 1 14:00-15:00; 16:20-16:40`. Without a day, the window applies to every day |
| Priority | ✓ | | TRUE / yes / 1 |
| Pitch Polishing | ✓ | | Ticked by default. Untick to skip polishing for that startup |
| Excluded Domains | | ✓ | `K-Beauty, Hardware` |
| Max Meetings | ✓ | ✓ | Blank = no limit |

Common header names such as `Startup Name`, `Company`, `Sector`, `Cutoff`, `From`, `Until` and `Excludes` are recognised on import. Excel time cells work too.

## Run it locally

You need Python 3.10 to 3.13.

```bash
cd bizmatch
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Your browser will open at http://localhost:8501. The app starts with a sample event, so you can press **Generate schedule** straight away.

## Deploy free on Streamlit Community Cloud

1. The code lives in **github.com/jayjunhyoung-crypto/bizmatching** (branch `main`, file `app.py`).
2. Go to **https://share.streamlit.io** and sign in with GitHub.
3. Click **Create app**, choose **Deploy a public app from GitHub**, and select `jayjunhyoung-crypto/bizmatching`, branch `main` and main file `app.py`. Pick an app URL such as `ksc-bizmatch`.
4. Open **Advanced settings** and set **Python version** to 3.12. The default may be older.
5. Click **Deploy**. The first build installs OR-Tools and takes a few minutes. You get a URL like `https://<name>.streamlit.app`.
6. **Keep it inside the team.** In the app, click **Share** at the top right (or open App settings and go to **Sharing**), make the app private and invite your colleagues by email. Viewers sign in to open it. A free account can have one private app.
7. To update the app later, push to `main` and it redeploys automatically.

Notes
- Participant data only lives in the browser session while you use the app. Nothing is stored on the server, so use **Save project** to keep your work.
- Free apps go to sleep after a period of inactivity. The first visit after that takes a short while to wake the app up.
- Light or dark mode: open ⋮ at the top right, then Settings, then Theme.
- Set the time zone in Tab 1. Streamlit Cloud servers run on UTC, and the live statuses use the time zone you choose.

## Tuning

- **Solver time limit**: 20 s is plenty for about 40 startups × 15 hosts × 2 days. For bigger events, raise it. The app reports whether the schedule is *Optimal* or a *good schedule within the time limit*.
- **Priority weight**: 1 means no preference. At 3, a priority startup's marginal meeting is worth about three times as much as a normal startup's once slots get scarce.
- The weights are constants at the top of `app.py` (`W_BASE`, `W_FAIR_STARTUP`, and so on).
