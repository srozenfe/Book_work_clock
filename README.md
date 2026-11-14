# Book Work Clock

A small, privacy-friendly Windows 11 desktop application that tracks time spent on books in Microsoft Word and supported PDF viewers.

## First version (MVP)

- Add, edit, and archive books, each linked to a local Word or PDF document (.doc, .docx, .docm, .pdf).
- Automatically start/resume tracking when the matching Word or PDF document is in front and the computer is active.
- Automatically switch between registered books.
- Prompt after **15 seconds** without mouse **or keyboard** activity.
- Prompt when switching away from Word; choose **Keep tracking**, **Pause**, **Stop**, or **Pause + don't ask again**.
- Optional automatic-tracking checkbox; manually restart with **Start / Resume**.
- Track sessions in a local SQLite database, with dates, start/end times, and duration.
- Keep historical time when a book is removed (archived), and show recent time entries.
- Run in the Windows notification area when the main window is closed.

## Windows 11 installer (recommended)

No Python installation is required for end users. The Windows build workflow
packages the Python interpreter and all dependencies inside `BookWorkClock.exe`,
then creates `BookWorkClock-Setup.exe` with Inno Setup.

1. Open **GitHub → Actions → Build Windows installer**.
2. Choose the latest successful run and download the **BookWorkClock-Windows** artifact.
3. Extract it and run **BookWorkClock-Setup.exe**.
4. In the installer, keep **Start Book Work Clock in the background when Windows starts**
   selected (enabled by default). This uses the current-user Windows Run entry and
   does **not** need administrator permissions.
5. Launch the app from Start menu or desktop. Thereafter it will start in the
   system tray when you sign in to Windows.

The main window also has a **Start minimized with Windows** checkbox, allowing
you to enable or disable automatic startup whenever you like.

**Note:** This release workflow produces build artifacts, not a public signed
installer. Windows SmartScreen may warn because the executable is unsigned.
Only install binaries from your own trusted GitHub Actions run.

## Manual work time entries

Below the session history, use **+ Add manual time** to select a book,
set start/end date-times, and optionally enter a note. Select a **Manual** entry
in the history and click **Edit manual time** or **Remove manual time**.
Auto-recorded entries are deliberately protected from manual deletion and edits
in this first version. All time, including manual entries, is counted in book totals.
Old databases are migrated automatically while retaining existing records.

## Book payments and effective hourly earnings

Select a book, click **Payment / earnings**, and enter the total payment
received or agreed for that book plus its currency (ILS, USD, EUR, GBP).
The dashboard shows the payment amount, tracked hours, and effective earnings
per hour: **payment / all recorded hours**. Automatic and manual sessions
both count. When the book has zero tracked hours, the hourly rate displays
"N/A". These values remain in your local SQLite database.

## Desktop launcher, single instance, and updates (v0.3)

The installer now **always creates a desktop shortcut**, even if the app is
configured to start automatically with Windows. Startup launches it hidden in
the tray. Double-clicking the desktop shortcut opens the existing running
dashboard (never a second timer) and checks for updates. If the app was not
running, the shortcut launches it and then checks for updates.

Updates are checked against the latest **published GitHub Release** without
authentication. You must make the GitHub repository **public** for this
feature to work. No GitHub token is required, requested, or saved by v0.3.1.

When a newer release contains `BookWorkClock-Setup.exe`, the app asks for
confirmation, downloads the installer, verifies the SHA-256 digest from GitHub,
installs it and relaunches the app. Existing local data remains in
`%LOCALAPPDATA%\\BookWorkClock`. Closing the launcher doesn't create a second
timer: it signals the running background instance.

If you previously used v0.3.0 and saved a GitHub token, v0.3.1 no longer
reads it; you can remove the old `BookWorkClock-GitHub-Readonly` credential
in **Windows Credential Manager** and revoke the token in GitHub.

### Publishing an update (automatic)

When code is committed to `main`, GitHub Actions automatically builds the
Windows executable and installer. If the application version is newer than any
existing release, the workflow publishes a tagged GitHub Release with
`BookWorkClock-Setup.exe`. You no longer need to press **Run workflow**.

For a new version, increment **both** `APP_VERSION` in `updater.py` and
`MyAppVersion` in `installer/BookWorkClock.iss` (for example from
`0.3.2` to `0.3.3`). The release is created automatically after the
Windows build succeeds. Updates are deliberately **not** published when
only the code changes without a version bump, avoiding mismatched versions.

The application checks for published updates when you launch it using the
desktop icon, whether it was closed or already running in the system tray.
When a newer version is found, you can approve the download, installation,
and application restart. Startup at Windows sign-in stays in the background
without interrupting you. The separate **Check GitHub for updates** button
can be used any time.

The application is unsigned and must be tested on Windows; GitHub CI verifies
that the executable and installer build, but does not prove the updater works
correctly in a user's Windows session.

## Developer install on Windows 11

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
2. Open a Word document in Microsoft Word (desktop app, **WINWORD.EXE**), or a registered PDF in Adobe Acrobat, Foxit, SumatraPDF, Microsoft Edge, Chrome, or Firefox.
3. Bring the matching document to the foreground and type, scroll, or move the mouse; the app starts tracking automatically.
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
- Matching uses the foreground window title, process name, and registered document filename (with or without Word extension). Different application title formats may require enhancements. For PDFs opened in a browser, the PDF filename including `.pdf` must appear at the beginning of the tab/window title; an unrelated website is not tracked.
- Tracking is **activity/foreground detection**, not proof of writing or actual document edits. Reading and thinking can also be counted at your choice.
- A paused or stopped session is recorded as one or more time segments. Manual entries are editable, but there is not yet a separate pause-event log or CSV/Excel export.
- When the user chooses *Keep tracking* outside Word, time continues until the registered Word window returns or the timer is manually paused/stopped. Subsequent foreground switches while outside Word are not repeatedly prompted.
- The desktop interface has been refreshed but does **not yet exactly reproduce the earlier design concept** (sidebar, large summary cards, charts). It needs further UI work and real-world testing for focus changes, sleep/wake, idle prompt timing, clock accuracy and window title variations.

## Foreground detection and interruption changes (v0.3.2)

- Recognizes supported desktop PDF readers and PDF files in Edge/Chrome/Firefox.
- Improved Word title matching with or without file extension and with app suffixes.
- Displays a **Detection** diagnostic under the status label. While using a document,
  it shows the matched book; if a title isn't recognized, it shows the foreground
  application's process name and document title, so the matching rule can be fixed.
- Paused tracking resumes on renewed keyboard/mouse input while a registered document
  is foreground, or when switching back to it; automatic tracking must be enabled.
- Interruption choices are now presented using an **always-on-top dialog**, even when
  Book Work Clock is minimized in the tray.
- Choosing **Keep tracking** prevents repeated prompts until normal activity resumes.
- The app still does **not inspect** PDF/Word content. Work time is inferred from the
  foreground window and keyboard/mouse activity; tests on the target Windows 11 device
  are recommended before trusting billed time.

## Planned improvements

- Better book identification through Microsoft Word automation and multiple documents per book.
- Per-book dashboard with daily/weekly/monthly summaries and export to CSV.
- Explicit pause/resume event log and editing **automatically** tracked historical entries.
- Configurable idle threshold, better desktop notifications, and a closer match to the design mockup.
