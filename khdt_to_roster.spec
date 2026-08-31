# PyInstaller build specification for the PyQt6 desktop app.
#
# Build from the project root:
#   pyinstaller --clean --noconfirm khdt_to_roster.spec
#
# This intentionally creates an onedir build, not a one-file executable.
# Keep the complete dist/KHDT-to-Roster/ folder when distributing it.
#
# Put a .png, Windows .ico, or macOS .icns in assets/ (or attached_assets/)
# before building if you want a custom application icon. PNG files are
# converted to ICO automatically for Windows builds.

from pathlib import Path
import sys

from PIL import Image
from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata


PROJECT_DIR = Path(SPECPATH).resolve()
ICON_CANDIDATES = [
    path
    for folder in (
        PROJECT_DIR / "assets",
        PROJECT_DIR / "asset",
        PROJECT_DIR / "attached_assets",
    )
    if folder.is_dir()
    for path in sorted(folder.rglob("*"))
    if path.is_file() and path.suffix.lower() in {".png", ".ico", ".icns"}
]
ICON_PATH = None
if ICON_CANDIDATES:
    icon_source = ICON_CANDIDATES[0]
    if icon_source.suffix.lower() == ".png" and sys.platform == "win32":
        converted_icon = PROJECT_DIR / "build" / "khdt_to_roster_icon.ico"
        converted_icon.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(icon_source) as image:
            image.convert("RGBA").save(
                converted_icon,
                format="ICO",
                sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
            )
        ICON_PATH = str(converted_icon)
    elif icon_source.suffix.lower() in {".ico", ".icns"}:
        ICON_PATH = str(icon_source)

pyqt_datas, pyqt_binaries, pyqt_hiddenimports = collect_all("PyQt6")
pyqt_hiddenimports += collect_submodules("PyQt6")

a = Analysis(
    ["gui.py"],
    pathex=[str(PROJECT_DIR)],
    binaries=pyqt_binaries,
    datas=pyqt_datas + copy_metadata("PyQt6"),
    hiddenimports=pyqt_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "tkinterdnd2"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    name="KHDT-to-Roster",
    exclude_binaries=True,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_PATH,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="KHDT-to-Roster",
)