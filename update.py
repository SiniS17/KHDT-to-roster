"""
AMOS Validator – Auto-Updater
==============================
Checks GitHub Releases for a newer version of the application bundle and
applies it without requiring the user to manually download anything.

How it works (Windows EXE)
--------------------------
1. GET https://api.github.com/repos/<OWNER>/<REPO>/releases/latest
2. Compare the release tag (e.g. "v1.2.0") with APP_VERSION in config.py
3. If newer: download the ZIP bundle next to the current EXE
4. Extract the bundle and write a tiny _amos_updater.bat that waits for the
   current process to exit, copies the files, and restarts
5. Launch the bat detached → exit this process → bat runs → new EXE starts

GitHub setup (one-time)
-----------------------
1. Create a public GitHub repo (or private – add a token to UPDATE_TOKEN secret)
2. Set GITHUB_REPO = "your-username/your-repo" in config.py
3. For every new release:
    - Tag the release  "v1.x.y"  (must match the semver format)
    - Upload a ZIP containing the complete `AMOS Validation/` folder as a
      release asset (for example `AMOS_Validation_v1.1.0.zip`)
   Done. Users get the update automatically on next run.

Source-mode behaviour
---------------------
When running from Python (not frozen EXE), the updater prints a notice but
never tries to replace files – it is safe to use during development.
"""

from __future__ import annotations

import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
import zipfile
from pathlib import Path
from typing import Callable, Optional

# ---------------------------------------------------------------------------
# Lazy import of config so the module is usable even before the package is
# fully initialised (e.g. called from a spec file during build).
# ---------------------------------------------------------------------------
try:
    from doc_validator.config import (
        APP_VERSION,
        BASE_DIR,
        GITHUB_REPO,
        UPDATE_CHECK_ENABLED,
    )
except ImportError:
    APP_VERSION = "0.0.0"
    BASE_DIR = Path(__file__).parent
    GITHUB_REPO = "SiniS17/Amos-filter-software"
    UPDATE_CHECK_ENABLED = bool(GITHUB_REPO)


# ---------------------------------------------------------------------------
# Version helpers
# ---------------------------------------------------------------------------

def _parse_version(tag: str) -> tuple[int, ...]:
    """
    Parse a version string into a comparable tuple of ints.
    Handles: "v1.2.3", "1.2.3", "1.2", "v2.0.0-beta" (pre-release stripped).
    """
    tag = tag.strip().lstrip("vV")
    # Drop pre-release suffix (e.g. "-beta", "-rc1")
    tag = tag.split("-")[0]
    parts = []
    for segment in tag.split("."):
        try:
            parts.append(int(segment))
        except ValueError:
            break
    return tuple(parts) if parts else (0,)


def is_newer(remote_tag: str, local_version: str) -> bool:
    """Return True if remote_tag represents a version newer than local_version."""
    return _parse_version(remote_tag) > _parse_version(local_version)


# ---------------------------------------------------------------------------
# GitHub API
# ---------------------------------------------------------------------------

GITHUB_API = "https://api.github.com/repos/{repo}/releases/latest"
_HEADERS = {
    "User-Agent": "AMOS-Validator-Updater/1.0",
    "Accept": "application/vnd.github+json",
}


def check_for_update(timeout: int = 6, verbose: bool = False) -> Optional[dict]:
    """
    Query GitHub Releases and return update info if a newer version exists.

    Returns
    -------
    dict with keys: version, download_url, asset_name, asset_type,
    release_notes
    None  if already up-to-date, repo not configured, or network error.
    """
    def _log(message: str) -> None:
        if verbose:
            print(f"[Updater] {message}", flush=True)

    _log(f"Current version: v{APP_VERSION}")
    _log(f"Repository: {GITHUB_REPO or '(not configured)'}")

    if not UPDATE_CHECK_ENABLED or not GITHUB_REPO:
        _log("Update checking is disabled.")
        return None

    try:
        _log("Checking GitHub for the latest published release...")
        url = GITHUB_API.format(repo=GITHUB_REPO)
        req = urllib.request.Request(url, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data: dict = json.loads(resp.read())

        tag: str = data.get("tag_name", "")
        if not tag:
            _log("GitHub returned no release tag.")
            return None

        newer = is_newer(tag, APP_VERSION)
        comparison = "update available" if newer else "already current"
        _log(f"Latest release: {tag}")
        _log(f"Version comparison: v{APP_VERSION} → {tag} ({comparison})")
        if not newer:
            return None  # already up to date

        # A PyInstaller onedir build is a directory, not a self-contained EXE.
        # Prefer the complete ZIP bundle. Keep .exe as a fallback for older
        # one-file releases that may still exist on GitHub.
        assets = data.get("assets", [])
        zip_asset = next(
            (
                a for a in assets
                if str(a.get("name", "")).lower().endswith(".zip")
            ),
            None,
        )
        exe_asset = next(
            (
                a for a in assets
                if str(a.get("name", "")).lower().endswith(".exe")
            ),
            None,
        )
        asset = zip_asset or exe_asset
        if not asset:
            available = ", ".join(
                str(a.get("name", "(unnamed)")) for a in assets
            ) or "(none)"
            _log(
                "No compatible update asset found. Supported formats: "
                ".zip or .exe. Available assets: "
                f"{available}"
            )
            return None  # release exists but has no supported application asset

        asset_name = str(asset.get("name", ""))

        return {
            "version": tag,
            "download_url": asset["browser_download_url"],
            "asset_name": asset_name,
            "asset_type": "zip" if asset_name.lower().endswith(".zip") else "exe",
            "file_size": asset.get("size", 0),
            "release_notes": (data.get("body") or "").strip()[:600],
        }

    except Exception as exc:
        # Never crash the app because of an update check
        _log(f"GitHub update check failed: {exc}")
        return None


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

ProgressCallback = Callable[[int, int], None]   # (downloaded_bytes, total_bytes)


def download_update(
    download_url: str,
    dest_path: Path,
    progress_cb: Optional[ProgressCallback] = None,
) -> bool:
    """
    Download the update package to dest_path.

    Parameters
    ----------
    download_url : direct asset download URL from GitHub
    dest_path    : where to save the file (e.g. next to the current EXE)
    progress_cb  : optional callback(downloaded, total) for progress bars

    Returns True on success, False on any error.
    """
    try:
        req = urllib.request.Request(download_url, headers=_HEADERS)
        with urllib.request.urlopen(req) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            chunk = 65_536  # 64 KB
            with open(dest_path, "wb") as fh:
                while True:
                    block = resp.read(chunk)
                    if not block:
                        break
                    fh.write(block)
                    downloaded += len(block)
                    if progress_cb:
                        progress_cb(downloaded, total)
        return True
    except Exception as exc:
        print(f"[Updater] Download failed: {exc}")
        try:
            dest_path.unlink(missing_ok=True)
        except OSError:
            pass
        return False


# ---------------------------------------------------------------------------
# Apply (replace the application bundle)
# ---------------------------------------------------------------------------

def _extract_update_zip(
    archive_path: Path,
    current_exe: Path,
) -> tuple[Path, Path]:
    """Extract a release ZIP safely and return its root and temp folder."""
    extract_dir = archive_path.parent / f"_amos_update_{os.getpid()}"
    shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(archive_path) as archive:
            root = extract_dir.resolve()
            for member in archive.infolist():
                target = (extract_dir / member.filename).resolve()
                if target != root and root not in target.parents:
                    raise ValueError(
                        f"Unsafe path in update archive: {member.filename}"
                    )
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, open(
                    target, "wb"
                ) as destination:
                    shutil.copyfileobj(source, destination)

        # The release may contain either:
        #   AMOS Validation/AMOS Validation.exe
        # or a flat bundle with AMOS Validation.exe at its root.
        candidates = list(extract_dir.rglob(current_exe.name))
        if not candidates:
            candidates = list(extract_dir.rglob("*.exe"))
        if not candidates:
            raise FileNotFoundError(
                f"Update archive does not contain {current_exe.name}"
            )
        return candidates[0].parent, extract_dir
    except Exception:
        shutil.rmtree(extract_dir, ignore_errors=True)
        raise


def apply_update(new_exe_path: Path) -> None:
    """
    Replace the running application with the downloaded update and restart.

    On Windows: writes a detached .bat that waits for this process to exit,
    swaps, cleans up, and restarts.
    On Linux/macOS: direct file replace + os.execv restart.
    Exits the current process after launching the swap mechanism.
    """
    if not getattr(sys, "frozen", False):
        print("[Updater] Not running as EXE – skipping file swap.")
        print(f"          New file is at: {new_exe_path}")
        return

    current_exe = Path(sys.executable).resolve()
    update_path = Path(new_exe_path).resolve()
    is_zip = update_path.suffix.lower() == ".zip"

    # Remove artifacts left by an older updater that stopped before cleanup.
    # Only the updater's own prefixes are targeted; user ZIP files are kept.
    for stale_dir in current_exe.parent.glob("_amos_update_*"):
        if stale_dir.is_dir():
            shutil.rmtree(stale_dir, ignore_errors=True)
    for stale_archive in current_exe.parent.glob("_update_*.zip"):
        if stale_archive.resolve() != update_path:
            stale_archive.unlink(missing_ok=True)

    if is_zip:
        try:
            update_root, extract_dir = _extract_update_zip(update_path, current_exe)
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            print(f"[Updater] Invalid update ZIP: {exc}")
            update_path.unlink(missing_ok=True)
            return
    else:
        # Backwards compatibility for a self-contained one-file EXE release.
        update_root = update_path.parent
        extract_dir = None

    if platform.system() == "Windows":
        bat_path = current_exe.parent / "_amos_updater.bat"
        updater_log = current_exe.parent / "_amos_updater.log"
        if is_zip:
            # robocopy copies the complete onedir bundle and preserves user
            # folders such as INPUT and DATA that are not part of the release.
            copy_command = (
                f'robocopy "{update_root}" "{current_exe.parent}" '
                "/E /R:2 /W:1 /NFL /NDL /NJH /NJS\n"
                "set COPY_EXIT=%ERRORLEVEL%\n"
                "if %COPY_EXIT% GEQ 8 goto update_failed\n"
            )
            cleanup_command = (
                f'rmdir /s /q "{extract_dir}" 2>nul\n'
                f'del /q "{update_path}" 2>nul\n'
            )
        else:
            copy_command = f'move /y "{update_path}" "{current_exe}"\n'
            cleanup_command = f'del /q "{update_path}" 2>nul\n'

        bat_content = (
            "@echo off\n"
            "setlocal EnableExtensions\n"
            f'set "APP_EXE={current_exe}"\n'
            f'set "APP_DIR={current_exe.parent}"\n'
            f'set "APP_PID={os.getpid()}"\n'
            f'set "LOG_FILE={updater_log}"\n'
            'echo [%date% %time%] AMOS Validator updater started.>>"%LOG_FILE%"\n'
            'echo Waiting for the previous process to exit...>>"%LOG_FILE%"\n'
            ":wait_for_exit\n"
            'for /f "tokens=2" %%P in (\'tasklist /FI "PID eq %APP_PID%" /NH\') do (\n'
            '    if "%%P"=="%APP_PID%" (\n'
            "        timeout /t 1 /nobreak > nul\n"
            "        goto wait_for_exit\n"
            "    )\n"
            ")\n"
            ":previous_process_exited\n"
            'echo Previous process exited.>>"%LOG_FILE%"\n'
            + copy_command
            + 'if not exist "%APP_EXE%" goto update_failed\n'
            + 'echo Files copied successfully.>>"%LOG_FILE%"\n'
            + 'start "" /D "%APP_DIR%" "%APP_EXE%" >>"%LOG_FILE%" 2>&1\n'
            + 'if errorlevel 1 goto update_failed\n'
            + 'echo Restart command completed.>>"%LOG_FILE%"\n'
            + cleanup_command
            + 'del /q "%~f0"\n'
            + "exit /b 0\n"
            + ":update_failed\n"
            + cleanup_command
            + 'echo Update failed. The existing installation was not removed.>>"%LOG_FILE%"\n'
            + 'echo Update failed. See "%LOG_FILE%".\n'
            + "exit /b 1\n"
        )
        bat_path.write_text(bat_content, encoding="utf-8")
        # Give the launcher a console, but keep it hidden. Using
        # CREATE_NO_WINDOW here (no console at all) causes the .bat's own
        # child processes (the `for /f`/tasklist loop, and `start`) to pop
        # open their own brand-new *visible* console windows, since they
        # have nothing to inherit. Creating a hidden console instead means
        # every child process shares that same invisible console.
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        subprocess.Popen(
            ["cmd", "/c", str(bat_path)],
            creationflags=(
                subprocess.CREATE_NEW_CONSOLE
                | subprocess.CREATE_NEW_PROCESS_GROUP
            ),
            startupinfo=startupinfo,
            close_fds=True,
        )
        # sys.exit(0) is NOT safe here: if apply_update() is ever called
        # from anything other than the main thread (a Qt slot, a worker
        # thread, etc.), sys.exit() only raises SystemExit in that thread
        # and does not terminate the process. Qt's event loop (and the
        # rest of the app) keeps running under the same PID, and the .bat's
        # "wait for PID to exit" loop then waits forever. os._exit() is a
        # hard, unconditional process kill regardless of which thread calls
        # it - no cleanup needed here since the .bat is already launched.
        os._exit(0)

    else:  # Linux / macOS
        if is_zip:
            shutil.copytree(
                update_root,
                current_exe.parent,
                dirs_exist_ok=True,
            )
            shutil.rmtree(extract_dir, ignore_errors=True)
            update_path.unlink(missing_ok=True)
        else:
            shutil.move(str(update_path), str(current_exe))
        os.chmod(current_exe, 0o755)
        os.execv(str(current_exe), sys.argv)


# ---------------------------------------------------------------------------
# High-level convenience
# ---------------------------------------------------------------------------

def run_update_check_console(interactive: Optional[bool] = None) -> None:
    """
    Full update flow for the console / batch entry point.
    Prints messages to stdout; blocks while downloading if the user agrees.
    Safe to call from a background thread (prints are thread-safe).

    In non-interactive environments such as Replit workflows, the update
    status is printed once but the prompt/download is skipped so processing
    cannot hang waiting for stdin.
    """
    if interactive is None:
        interactive = sys.stdin.isatty()

    print("[Updater] Checking for updates...", flush=True)
    info = check_for_update(verbose=True)
    if info is None:
        print("[Updater] Startup update check complete.\n", flush=True)
        return

    print(
        f"\n{'='*60}\n"
        f"  🚀  Update available: {info['version']}  (current: v{APP_VERSION})\n"
        f"{'='*60}"
    )
    if info["release_notes"]:
        print("  Release notes:")
        for line in info["release_notes"].splitlines():
            print(f"    {line}")
    print()

    if not interactive:
        print(
            "[Updater] Non-interactive startup detected; "
            "skipping download prompt. "
            "Use the desktop GUI to install this update.\n",
            flush=True,
        )
        return

    try:
        answer = input("  Download and install now? [y/N]: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        answer = "n"

    if answer != "y":
        print("[Updater] Skipped. You can update later by restarting.\n")
        return

    dest = Path(sys.executable).parent / f"_update_{info['asset_name']}"

    def _progress(done: int, total: int) -> None:
        if total:
            pct = done * 100 // total
            bar = "█" * (pct // 5) + "░" * (20 - pct // 5)
            print(f"\r  [{bar}] {pct:3d}%", end="", flush=True)

    print(f"  Downloading {info['asset_name']} …")
    ok = download_update(info["download_url"], dest, progress_cb=_progress)
    print()  # newline after progress bar

    if not ok:
        print("[Updater] ❌ Download failed. Continuing with current version.\n")
        return

    print("[Updater] ✅ Download complete. Applying update and restarting…\n")
    apply_update(dest)


def check_in_background(
    callback: Callable[[Optional[dict]], None],
    verbose: bool = False,
) -> threading.Thread:
    """
    Run check_for_update() in a daemon thread and call callback(info) on the
    main thread when done (info is None if no update / error).

    ``verbose=True`` prints the one-time startup check to the console while
    preserving the non-blocking GUI behavior.

    Used by the Qt GUI so the window appears instantly while the check runs.
    """
    def _worker():
        info = check_for_update(verbose=verbose)
        callback(info)

    t = threading.Thread(target=_worker, daemon=True, name="amos-update-check")
    t.start()
    return t