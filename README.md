# Book Work Clock

A small, privacy-friendly Windows 11 desktop application that tracks time spent on individual books in Microsoft Word.

## First version (MVP)

- Add, edit, and archive books, each linked to a local Word document (.doc, .docx, or .docm).
- Automatically start/resume tracking when the matching Word document is in front and the computer is active.
- Automatically switch between registered books.
- Prompt after **15 seconds** without mouse **or keyboard** activity.
- Prompt when switching away from Word; choose **Keep tracking**, **Pause**, **Stop**, or **Pause + don't ask again**.
- Optional automatic-tracking checkbox; manually restart with **Start / Resume**.
- Track sessions in a local SQLite database, with dates, start/end times, and duration.
- Keep historical time when a book is removed (archived), and show recent time entries.
- Run in the Windows notification area when the main window is closed.

## Install on Windows 11

1. Download/clone the repository and open PowerShell in its folder.
2. Install Python 3.11+ from python.org, if it is not already installed.
3. Create and activate a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```

4. Launch:

   ```powershell
   python book_work_clock.py
   ```

If PowerShell refuses activation, you can avoid activation entirely:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe book_work_clock.py
```

## How to use

1. Choose **Add book**, browse to its Word document, and give the book a name.
2. Open that document in Microsoft Word (desktop app, **WINWORD.EXE**).
3. Bring Word to the foreground and type or move the mouse; the app starts tracking automatically.
4. If you're interrupted, respond to the prompt or use the Pause/Stop buttons.
5. Select **Pause + don't ask again** (or uncheck automatic tracking) to suppress tracking until you manually press **Start / Resume**.
6. The recent periods and book totals are available in the main window.

## Data and privacy

The app stores data only on the computer in:
`%LOCALAPPDATA%\BookWorkClock\work_clock.sqlite3`

It reads the foreground window's title and process name, and Windows' last keyboard/mouse input timestamp. **It does not record keystrokes, document content, or screenshots.** Nothing is uploaded.

The database persists across restarts. Each running period is committed approximately once per second. Back up this database to preserve your history.

## Optional: build an EXE on Windows

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name BookWorkClock book_work_clock.py
```

The output is in `dist\BookWorkClock.exe`.

## Current limitations / next improvements

- Tested execution on Windows 11 is still needed; development was performed remotely without a Windows GUI.
- Matching uses the Word window title and the registered document filename. Word title variations, similarly named documents, and some unusual Word installations may require improved matching through Word COM automation.
- Tracking is **activity/foreground detection**, not proof of writing or actual document edits. Reading and thinking can also be counted at your choice.
- A paused or stopped session is recorded as one or more time segments. There is not yet a separate pause-event log, manual time correction, or CSV/Excel export.
- When the user chooses *Keep tracking* outside Word, time continues until the registered Word window returns or the timer is manually paused/stopped. Subsequent foreground switches while outside Word are not repeatedly prompted.
- This MVP has a basic interface and needs real-world testing for focus changes, sleep/wake, idle prompt timing, clock accuracy and window title variations.

## Planned improvements

- Better book identification through Microsoft Word automation and multiple documents per book.
- Per-book dashboard with daily/weekly/monthly summaries and export to CSV.
- Explicit pause/resume event log and editing historical entries.
- Configurable idle threshold, desktop notifications, and Windows startup support.
