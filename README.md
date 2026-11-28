# BizMatch

A web app for running accelerator programs: each program's stages and deadlines, its startups, intake
interviews, results and KPIs, and the proposal, interim and final reports as 한글 (.hwpx) files. It also builds
optimal, conflict-free 1:1 business matching schedules (startup × investor / corporate speed meetings) with
Google OR-Tools CP-SAT. It runs on Streamlit.

```
bizmatch/
├── app.py                 # pages, company directory, business matching (CP-SAT model, board, Excel export)
├── programs.py            # Programs page: stages, startups, intake, KPIs and outcomes, documents
├── docgen.py              # document sections built from the data, Markdown -> .hwpx (한글)
├── ai.py                  # optional AI drafts (Claude API)
├── store.py               # team workspace database (events, directory, program documents)
├── i18n.py                # English and Korean text
├── board/ summary/ local/ # browser components: matching board, summary, browser backup
├── requirements.txt
└── .streamlit/config.toml # theme colour and upload limit
```

## Programs

**Programs** is the first page. Every accelerator program has its own structure, so a program starts from a
template and everything in it can be changed.

* **Templates**: global acceleration, exhibition or CES delegation, business matching event, or blank. A template
  fills in the stages (proposal, contract, selection, kick-off, intake, workshops, matching, interim report, demo
  day, final report, settlement and so on) spread over the program dates, each with a checklist, plus default
  KPIs and intake questions. **Create an example program** shows a filled one.
* The page lists **deadlines in the next two weeks** across all programs and a card per program with its
  progress.
* Inside a program:
  * **Overview**: name, client, dates, PM, team, objective, budget, and the business matching events that
    belong to the program (their meetings count toward the KPIs and reports).
  * **Stages & schedule**: a table of stages with status, dates, owner and notes, a timeline with today marked,
    and a checklist per stage.
  * **Startups**: pick the program's startups from Companies and see intake status, meetings and outcomes for
    each.
  * **Intake interview**: the team records each interview while talking (same questions for every startup,
    editable under **Questions**). **Summarize with AI** turns the notes into a summary and suggests tags,
    stage and a one line intro, which you can apply to Companies. Download one startup or all of them as 한글.
  * **Results & KPIs**: meetings from the linked events (done, no-shows, upcoming, partners), indicators with
    target and actual (counted automatically from meetings, interviews and the outcome log, or typed), and an
    outcome log (follow-up meetings, NDA, MOU, PoC, investment, contract).
  * **Reports**: a proposal with a quote (line items, VAT, totals), an interim report and a final report. Every
    section starts filled from the program data and keeps following it until someone edits that section.
    Each section can be refilled from the data or drafted by AI, then edited freely. **Download 한글 file** gives
    a .hwpx in the usual Korean government report layout (title band, numbered section boxes, □ and - bullets,
    ruled tables, page numbers). It opens in 한글 2014 or later; if the client wants .hwp, open it in 한글 and
    use Save As. **Preview pages** shows an approximate layout in the browser.

Everything is saved as you type. In the team workspace programs are shared (table `bm_docs`); when two people
change the same program at the same moment, both changes are kept unless they touched the same field, in which
case the first one stays and the second person sees a note.

### AI drafts (optional)

AI drafts use the Claude API. Without a key the app works the same and the AI buttons are hidden. To turn them
on, add this to the secrets (Streamlit Cloud: **⋮ → Settings → Secrets**; locally `.streamlit/secrets.toml`):

```toml
[anthropic]
api_key = "sk-ant-..."
model = "claude-sonnet-5-5"   # optional
```

Only what a section needs is sent (program details, stages, indicators, meeting counts, outcomes, intake notes).
Drafts are written in the usual Korean report style and mark anything they cannot confirm with (확인 필요);
check numbers and names before sending a report.

## Companies and business matching

The other two pages in the top bar.

1. **Companies** is the team's directory: one row per startup and per partner (investor, corporate, agency).
   Startups keep Company (EN), Company (KR), CEO and CEO (EN), email, phone, website, tags, a one-liner,
   founded year, stage, employees, presenter, booth, deck link and notes. Partners keep organisation, contact,
   title, email, phone, type, country, the domains they prefer, only meet or exclude, website and notes.
   * Add your own fields under **Fields** (for example Investment ask or Patent).
   * **Paste or upload** takes cells copied from Google Sheets or Excel, or a CSV/XLSX file, with Korean or
     English headers. Each column is matched to a field automatically and you can change the match. "Add new
     and update existing" matches companies by name, so pasting an updated sheet only fills in what changed.
   * Workbooks with several tabs: the app reads every tab, skips cover or notes tabs, finds the header row even
     when a title sits above it, and loads the tab that fits best (by tab name and column names). A **Tabs to
     import** box lets you pick a different tab, or several tabs with the same columns to combine them. When a
     second tab clearly holds the other list (for example "투자자 리스트"), it is prepared on that side too. The
     participant uploads in Business matching work the same way.
   * Search looks in every field. **Download Excel** gives both lists as one workbook.
   * **Company profile** shows one company on a card, plus every meeting it has in business matching
     (the open event, and in the team workspace every saved event).
2. **Business matching** is the scheduler below. In an event's participant tables, **📇 From companies**
   adds companies from the directory (English name, tags, presenter or CEO; for partners the contact, type and
   domain rules), and **Refresh details** updates the ones already in the event after you edit the directory.

In the team workspace the directory is shared and saved automatically (table `bm_directory`). In personal
mode the app keeps a copy of the event and the directory in your browser after every change, so a refresh or a
closed tab loses nothing (it is per browser, so use **Save project** or the team workspace to move work to
another computer). Paste imports never drop a column: a header the app does not know becomes a new field.

## What it does

**Language.** The whole app (setup, results and the drag-and-drop board) is available in English and Korean (한국어). Switch at the top right. The choice is kept in the link (`?lang=ko`), so you can share a Korean link. The Excel file stays in English so it can go to participants, and Korean column headings (스타트업, 발표자, 분야, 참석일 ...) and inputs such as `불참` or `1일차` are understood when you paste.

**Tab 1. Event & Participants**
- Multi-day events (one row per day, with dates), day start and end, meeting length (default 20 min), optional gap between meetings, lunch break, number of tables, time zone.
- **Sessions** such as `Session A 10:00-12:00; Session B 13:00-15:00; Session C 15:00-17:00`. They group the timeline exactly like the team's sheet.
- **Pitch polishing** that runs alongside the meetings. Each startup gets one polishing slot per selected day, and never has a meeting in that slot. *Coaches* sets how many startups are polished at the same time.
- Editable tables for startups and hosts. You can also import from CSV or Excel, and download a template.
- **Paste straight from Google Sheets or Excel.** Copy the cells including your header row and paste them into *Paste startups / hosts*. Your own column names are fine: the app guesses what each column means (`Startup`, `Sector`, `Attending`, `Pitch practice` ...) and you can correct any guess before importing. A host column written like `Raymond - Focus Tech` can be split into contact and organisation. Uploaded CSV / Excel files go through the same check.
- Meeting length in 1-minute steps, and the gap between meetings in minutes plus seconds (e.g. 1 min 10 s).
- **Column names** used on the board and in Excel can be renamed (e.g. Table → Booth).
- The headings of the Startups and Hosts tables can be renamed with **Rename columns** (e.g. Name → Startup, Domain → Sector). Only the display changes, and pasted sheets with those headings are matched automatically.
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
- **Pin** a meeting from its menu to lock it: a pinned meeting cannot be dragged or pushed by other moves, and it always stays where it is when you generate again. ★ marks a priority startup.
- NOT HERE / ONLY notes stay hidden until you press **Show attendance notes**. In the Excel file they sit in a folded row group, opened with the `+` in the left margin.
- A timeline for each day in the team's sheet layout: Session, Time Slot, Table 1 / Status 1 … Table N / Status N, **Polishing Schedule** and **Waiting List** (startups that are neither in a meeting nor in polishing, or `None (All in Sessions)`). The live board in the app also shows resting hosts.
- Each host (organisation) keeps one colour everywhere, so you can follow a host across tables and sessions. Pitch polishing is shown in neutral slate grey with a dark outline (1st light grey, 2nd deeper grey), so it never looks like a partner. In the waiting list every startup also gets its own colour.
- `NOT HERE` and `ONLY 27 NOV` style notes at the top are generated from the Days column.
- Status chips: `Scheduled`, `Ongoing`, `Done`, `No-Show`. On the event day, statuses follow the clock automatically and the current slot is marked NOW. The board refreshes every 30 s.
- The live status board lets you override any meeting, for example to mark a No-Show. A manual status always wins over the clock.
- An independent verifier re-checks every hard rule on the final schedule and shows the result.
- Startup summary (name, total meetings, partners matched on each day) and host summary.
- **The board is the page**: it is as tall as the day's schedule, so you scroll the page and the header row follows you. **Full screen** in the board toolbar puts the board over the whole browser window for the event day (Esc to leave).
- **Find and highlight** on the live board: type or pick a startup or partner and everything else fades, with a line showing their meeting count, polishing time and next meeting. Clicking a waiting startup or a free partner tag does the same.
- **Cancellations (who steps in)**: mark a meeting No-Show, or pick "Cancelled? Find who steps in" in its menu, and a panel lists who can take the empty seat right now. If the startup cancelled, it lists startups that are free at that time and have never met that partner (or have a later meeting with them today that can be pulled forward), best fit first. If the partner cancelled, it lists free partners the startup has not met. One click puts them in; the cancelled meeting goes to Unplaced as a no-show so you can rebook it. A no-show chip shows ↺ and how many substitutes there are.
- **Waiting list** only shows startups who have not met at least one partner present in that time slot; the small number on each tag is how many of those partners they have not met (hover to see who).
- **Partner card**: click a free partner's name tag (or "Recommend startups" in a meeting's menu, or search a partner) to see how many startups they have met and how many are left, plus a ranked list of startups they have not met yet. Ranking puts preferred domain matches, ★ priority startups and startups with fewer meetings first, and only lists startups that pass the partner's Only and Excluded rules. Each row has a + button with the earliest time today when both are free and a table is open (next to the partner's other meetings when possible); one click adds the meeting.
- **Free partners** column on the live board: partners who are available but not meeting in that slot appear as dashed name tags. Drag one onto an empty table and pick a startup from the list (only startups that are free, have not met that partner yet and are not in an excluded domain are offered), or drop it straight onto a startup in the waiting list.
- **Summary tab**: a live dashboard for the event day. Totals with a completion bar (done, ongoing, upcoming, no-show, pitch practice), a "Now" strip listing which table is meeting right now and when the next round starts, and one row per startup (or per partner) with a progress bar and every meeting as a chip coloured by status. Upcoming and ongoing meetings come first and finished ones move to the end, and by default rows are sorted so whoever is up next is on top. Search, filter (priority, has a no-show, in a meeting now, part days) and sort. Click any chip to mark it done or a no-show; the board and Excel follow.
- **Download Excel**. Sheets: `Matching Schedule`, `Pitching Practice Schedule`, `Startup Summary`, `Host Summary`, `All Meetings`. The POLISHING SCHEDULE column of the main sheet is a formula that reads the Pitching Practice sheet, so changing a startup there in Excel or Google Sheets updates the main schedule too. Every meeting has a Status dropdown, and its colour changes with the status (conditional formatting). Each meeting cell reads `Startup` with `(Contact - Organisation)` below it, filled with the host's colour, and a No-Show greys it out and strikes it through. Upload the file to Google Drive and open it with Google Sheets: the dropdowns and colours keep working there, so the sheet can be the live board on the day. The sheet prints landscape, fitted to page width.

## Participant columns

| Column | Startups | Hosts | Format |
|---|---|---|---|
| Name | ✓ | ✓ | Unique (for hosts, the organisation) |
| Contact | | ✓ | Person attending, shown as `Raymond · Focus Tech` |
| Presenter | ✓ | | Who pitches, e.g. `Jay & Serena`. Shown in the pitching practice schedule |
| Domain | ✓ | | Tags, comma separated. Put the sector and the business model together, e.g. `Fintech, AI, B2C` |
| Type | | ✓ | Investor, Corporate, Agency, Mentor, Other |
| Days | ✓ | ✓ | Blank or `All` = every day. Otherwise `Day 2`, `1,2`, or your day label. `Not here` (or `None`, `Absent`) = not attending, listed under NOT HERE |
| Available From | ✓ | ✓ | `HH:MM`. Blank = day start |
| Available Until | ✓ | ✓ | `HH:MM` cutoff. No meeting may end after it (e.g. `16:00` means "cannot meet after 16:00") |
| Blocked Times | ✓ | ✓ | Restricted windows, e.g. `Day 1 14:00-15:00; 16:20-16:40`. Without a day, the window applies to every day |
| Priority | ✓ | | TRUE / yes / 1 |
| Pitch Polishing | ✓ | | Ticked by default. Untick to skip polishing for that startup |
| Only Domains (Only Meets) | | ✓ | Hard rule. Meets only startups with one of these tags, e.g. `AI` or `AI, B2C`. Blank = anyone |
| Preferred Domains (Prefers) | | ✓ | Soft wish, e.g. `B2C`. These startups are matched with them first; free time can still go to others |
| Excluded Domains | | ✓ | Hard rule. Never these tags, e.g. `K-Beauty, Hardware` |
| Max Meetings | ✓ | ✓ | Blank = no limit |

Tags match by whole words, so `AI` matches `AI`, `AI/ML` and `Gen AI` but not `Retail`, and `B2C` matches `B2C Commerce`. If a partner's Only Meets tags match no startup, the Generate tab warns before you run it.

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
3. Click **Create app**, choose **Deploy a public app from GitHub**, and select `jayjunhyoung-crypto/bizmatching`, branch `main` and main file `app.py`. Pick an app URL such as `bizmatching`.
4. Open **Advanced settings** and set **Python version** to 3.12. The default may be older.
5. Click **Deploy**. The first build installs OR-Tools and takes a few minutes. You get a URL like `https://<name>.streamlit.app`.
6. **Keep it inside the team.** In the app, click **Share** at the top right (or open App settings and go to **Sharing**), make the app private and invite your colleagues by email. Viewers sign in to open it. A free account can have one private app.
7. To update the app later, push to `main` and it redeploys automatically.

Notes
- Without a team database (below), participant data only lives in the browser tab while you use the app, so use **Save project** to keep your work.
- Free apps go to sleep after a period of inactivity. The first visit after that takes a short while to wake the app up.
- Light or dark mode: open ⋮ at the top right, then Settings, then Theme.
- Set the time zone in Tab 1. Streamlit Cloud servers run on UTC, and the live statuses use the time zone you choose.

## Team workspace (shared events and login)

With a database connected, everyone on the team opens the same events. Edits save automatically,
the board shows who saved last, and on the event day several people can update statuses on the live board
at once. Each browser checks for other people's changes every 10 seconds.

How saving works
- The setup and schedule of an event are saved as one record with a version number. If two people save at
  the same moment, the second person gets the newer version loaded and a message to redo their last change.
- Meeting statuses and polishing notes are saved one by one, so two people marking different meetings never
  overwrite each other.
- **Manage → Delete** hides an event for the whole team.

### 1. Database (Supabase, free)

1. Create a project at **https://supabase.com** (region Singapore). Write down the database password.
2. Click **Connect** at the top of the project, choose the **Session pooler** connection string
   (Streamlit Cloud needs the pooler because it has no IPv6), and replace `[YOUR-PASSWORD]` with your password.
3. That's it. The app creates its two tables (`bm_events`, `bm_event_items`) on first start and turns on
   row level security so they are not readable through Supabase's public API.

Free Supabase projects pause after a week with no use. Open the Supabase dashboard and click **Restore** if
that happens. Any Postgres works, for example Neon.

### 2. Google login (recommended)

1. Open **https://console.cloud.google.com**, create a project, then **APIs & Services → OAuth consent screen**.
   If your company uses Google Workspace, choose **Internal** so only company accounts can sign in.
2. **Credentials → Create credentials → OAuth client ID → Web application**.
   Add the authorised redirect URI `https://bizmatching.streamlit.app/oauth2callback`
   (and `http://localhost:8501/oauth2callback` for local testing).
3. Copy the client ID and client secret.

### 3. Paste the secrets into Streamlit Cloud

Open the app on **https://share.streamlit.io**, then **⋮ → Settings → Secrets**, and paste:

```toml
[database]
url = "postgresql://postgres.abcdefgh:YOUR_PASSWORD@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"

[auth]
redirect_uri = "https://bizmatching.streamlit.app/oauth2callback"
cookie_secret = "any long random text, e.g. from: python -c 'import secrets; print(secrets.token_hex(32))'"
client_id = "1234-abc.apps.googleusercontent.com"
client_secret = "GOCSPX-..."
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"

[access]
allowed_domains = ["yourcompany.com"]     # anyone with an email at these domains may sign in
allowed_emails = ["partner@gmail.com"]    # plus these individual addresses
```

Save, and the app restarts with the team workspace. Leave out `[access]` to let any Google account in.

**Quicker start without Google**: leave out `[auth]` and set a shared passcode instead. Each person types the
passcode and their name. This is fine for trying it out, but Google login is safer for real participant data.

```toml
[database]
url = "postgresql://..."

[access]
passcode = "choose-a-passcode"
```

For local runs put the same content in `.streamlit/secrets.toml` (it is git ignored). Never commit secrets.

## Big events

Tested with 33 partners, 25 to 40 startups, and a CES style run (33 startups, 50 to 80 partners, 4 days of 30 minute slots in Las Vegas time: 735 to 1,010 meetings in 25 to 35 seconds, no clashes), 2 days of 20 minute slots and pitch polishing: about 25 seconds to
a full schedule with no clashes. The app first builds a quick schedule slot by slot and hands it to the
solver as a starting point, then gives big events extra solving time automatically (up to 2 minutes). If the
solver cannot beat the quick schedule in time, the quick one is used. With more than 16 partners every partner
still gets its own pastel colour. On the board, more than 10 tables switches to slimmer cells, the Session and
Time columns stay put while you scroll sideways, and Full screen is the easiest way to see a 30 table board.

## Tuning

- **Solver time limit**: 20 s is plenty for about 40 startups × 15 hosts × 2 days. For bigger events, raise it. The app reports whether the schedule is *Optimal* or a *good schedule within the time limit*.
- **Priority weight**: 1 means no preference. At 3, a priority startup's marginal meeting is worth about three times as much as a normal startup's once slots get scarce.
- The weights are constants at the top of `app.py` (`W_BASE`, `W_FAIR_STARTUP`, and so on).
