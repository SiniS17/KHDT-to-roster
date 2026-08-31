# KHDT → Roster

A small Windows desktop tool for VAECO that reads a **KHDT** (training-plan)
workbook and a **Roster** workbook, matches people by Employee ID, and marks
every day of their training on the Roster with the right course code —
so nobody has to do it by hand in Excel.

## What it does

1. Scans the KHDT workbook for its header row (`Mã NV` / `Họ tên` / `Bắt đầu`
   / `Kết thúc`) — no fixed row number required.
2. Scans the Roster workbook the same way (`ID` / `NAME` columns plus a run
   of date columns).
3. For each KHDT row, looks up the Employee ID in the Roster. Name is
   cross-checked as a sanity warning, but ID is authoritative.
4. Marks every day in the person's training range with an abbreviated
   course code (configurable in `config.py`). Same-day courses combine,
   e.g. `H VHNT SÁNG-VHDN CHIỀU`.
5. Saturdays/Sundays are written as a black `N` (day off) instead of a
   course code. Weekday markers are written in red.
6. Produces:
   - `<Roster>_updated.xlsx` — the marked-up roster
   - `<Roster>_updated.log.xlsx` — a full log workbook (per-person events,
     a run summary, and the raw console log)

It also automatically corrects a known KHDT date bug where some cells store
day/month swapped (`YYYY-DD-MM`), using the Roster's own month as a sanity
check.

## Running it

**GUI (default):** just launch `KHDT-to-Roster.exe`. Drag and drop the two
Excel files onto the window (or use Browse…), then click Run. Output goes to
the folder you choose (defaults to `OUTPUT/` next to the exe).

**Zero-click / INPUT folder mode:** create an `INPUT` folder next to the exe
and drop both files in it — one with `KHDT` somewhere in its filename, one
with `Roster` somewhere in its filename (case-insensitive). Run the exe with
no arguments and it processes automatically into `OUTPUT/`.

**Command line:**
```
KHDT-to-Roster.exe path\to\KHDT.xlsx path\to\Roster.xlsx [--force] [-o out.xlsx]
```
`--force` overwrites Roster cells that already have content (by default they
are skipped and logged as conflicts).

## Configuration

Edit `config.py` before building to change:
- Course abbreviations (`COURSE_ABBREVIATIONS`, `COURSE_NAME_ABBREVIATIONS`)
- Marker prefix and log column/event labels
- `APP_VERSION` / `APP_LAST_UPDATE` (bump before every release)
- `GITHUB_REPO` — set to `"owner/repo"` to enable auto-update checks, or
  leave blank to disable them entirely

## Auto-update

If `GITHUB_REPO` is set, the app checks the repo's latest GitHub Release on
startup. If a newer version is tagged (`vX.Y.Z`), it offers to download the
release's `.zip` asset and swap it in automatically (via a small batch
script on Windows), then restarts. Update checks are skipped when running
from source (unfrozen).

## Building

```
pyinstaller --clean --noconfirm khdt_to_roster.spec
```

This produces a **onedir** build at `dist/KHDT-to-Roster/` — keep the whole
folder together when distributing; don't ship just the `.exe`.

To give the build a custom icon, put a `.png`, `.ico`, or `.icns` file in an
`asset/`, `assets/`, or `attached_assets/` folder next to the spec file
before building. PNGs are auto-converted to a multi-resolution `.ico` for
Windows. If you don't see the new icon after a rebuild, it's almost always
Windows' icon cache — run `ie4uinit.exe -show` or delete
`%LocalAppData%\IconCache.db` and restart Explorer.

## Requirements

- `openpyxl` (core processing)
- `PyQt6` (GUI)
- `Pillow` (icon conversion at build time only)

## Project files

| File | Purpose |
|---|---|
| `main.py` | Core matching/marking logic, CLI entry point |
| `gui.py` | PyQt6 drag-and-drop desktop interface |
| `update.py` | GitHub Releases auto-updater |
| `config.py` | User-editable labels, abbreviations, version, repo |
| `khdt_to_roster.spec` | PyInstaller build spec (onedir, custom icon) |
