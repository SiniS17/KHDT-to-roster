# Release Notes

Format follows [Keep a Changelog](https://keepachangelog.com/), and versions
are plain semver (`MAJOR.MINOR.PATCH`) so `update.py` can compare them. Tag
each GitHub release `vX.Y.Z` to match `APP_VERSION` in `config.py`.

## [Unreleased]
- (nothing yet)

## [1.0.3] - 2026-08-31

### Fixed
- `combine_markers` no longer produces a doubled/nested `H` prefix when a
  roster cell already contains a course this run would also add (e.g. from
  re-running the tool on an already-marked roster). Markers are now
  deduplicated at the individual-course level instead of by comparing whole
  marker strings, so combining is idempotent regardless of how many times
  a cell gets processed.
- A malformed cell left over from before this fix (a nested `H` from the
  old duplicate-prefix bug) is automatically normalized back to a single
  `H` prefix the next time it's processed, even if no new course is being
  added that day.
- Removed all remaining "AMOS Validator" branding from `update.py`
  (docstring, updater filenames, `User-Agent` header, background-thread
  name) and a stale example in a `config.py` comment — none of it applied
  to this software.

### Added
- New `already_marked` run stat and log event (`LOG_EVENT_LABELS
  ["already_marked"]`, "Đã có sẵn, bỏ qua"): when a course this run would
  add is already present on a cell, it's recorded in the log workbook,
  console output, and GUI summary instead of being silently skipped or
  silently duplicated.

### Changed
- `khdt_to_roster.spec` now builds into an `EXE/` folder next to the spec
  file instead of the default `dist/`, and deletes the intermediate
  `build/` folder automatically once the build finishes.

## [1.0.0] - 2026-08-29
Initial release.

### Added
- Core matching engine: reads KHDT + Roster workbooks, matches people by
  Employee ID (with Name cross-check), and marks matching day-columns with
  course codes for the full training date range.
- Automatic header-row detection for both KHDT and Roster sheets — no
  dependency on fixed row numbers.
- Same-day course combination (e.g. `H VHNT SÁNG-VHDN CHIỀU`).
- Single-date exception handling when the "Lí do" note names one specific
  date instead of the full Bắt đầu–Kết thúc range.
- Weekend handling: Saturday/Sunday assignments written as a black `N`
  instead of a course marker; conflict detection against existing black/red
  `N` cells (logged as "Overwrite day off" / "Guarantee day off conflict").
- KHDT date bug workaround: auto-corrects day/month-swapped dates using the
  Roster's own month as a plausibility check; anything still implausible is
  logged instead of guessed at.
- Excel log workbook output alongside the updated roster, with three
  sheets: per-person event log, run summary, and full console log.
- PyQt6 desktop GUI with drag-and-drop file zones, Browse buttons, output
  folder picker, force-overwrite checkbox, and a live run log.
- Zero-click `INPUT/` folder mode: drop both workbooks in and run with no
  arguments.
- CLI mode with explicit file paths, `--force`, and `-o/--output`.
- GitHub Releases auto-updater: checks the configured repo on startup,
  downloads newer `.zip` releases, and applies them via a self-restarting
  batch script (Windows) or in-place swap (Linux/macOS).
- PyInstaller onedir build spec with automatic custom-icon support (PNG,
  ICO, or ICNS from `asset/`, `assets/`, or `attached_assets/`, with
  PNG→ICO conversion for Windows builds).
- Configurable course abbreviations, marker prefix, and log labels via
  `config.py`.

### Notes
- First tracked version. Future entries should describe changes relative
  to this baseline.