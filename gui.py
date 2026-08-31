"""PyQt6 desktop GUI for marking a Roster workbook from a KHDT workbook."""

from __future__ import annotations

import os
import platform
import subprocess
import sys
import threading
from pathlib import Path

from PyQt6.QtCore import QMimeData, QObject, QRunnable, QThreadPool, QUrl, Qt, pyqtSignal
from PyQt6.QtGui import QDesktopServices, QDragEnterEvent, QDropEvent, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from config import APP_LAST_UPDATE, APP_VERSION
from main import find_input_files, get_base_dir, run, write_log_workbook
import update as updater


EXCEL_FILTER = "Excel workbooks (*.xlsx *.xlsm);;All files (*)"


def _find_app_icon() -> Path | None:
    """Find the source icon in development or the bundled onedir folder."""
    roots = [
        Path(getattr(sys, "_MEIPASS", "")),
        get_base_dir(),
        Path(__file__).resolve().parent,
    ]
    for root in roots:
        if not root:
            continue
        for folder in ("asset", "assets", "attached_assets"):
            for suffix in (".png", ".ico"):
                candidate = root / folder / f"icon{suffix}"
                if candidate.is_file():
                    return candidate
    return None


class DropZone(QFrame):
    """A labeled drop target that also exposes a normal browse button."""

    file_dropped = pyqtSignal(str)

    def __init__(self, title: str, hint: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("dropZone")
        self.setMinimumHeight(130)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.path_label = QLabel(hint)
        self.path_label.setObjectName("filePath")
        self.path_label.setWordWrap(True)
        self.path_label.setMinimumHeight(38)

        title_label = QLabel(title)
        title_label.setObjectName("dropTitle")
        browse_button = QPushButton("Browse…")
        browse_button.clicked.connect(self._browse)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(7)
        layout.addWidget(title_label)
        layout.addWidget(self.path_label)
        layout.addWidget(QLabel("Drop an .xlsx or .xlsm file here"))
        layout.addWidget(browse_button, alignment=Qt.AlignmentFlag.AlignLeft)

    def set_path(self, path: str) -> None:
        self.path_label.setText(path)
        self.path_label.setToolTip(path)
        self.path_label.setObjectName("filePathSelected")
        self.style().unpolish(self.path_label)
        self.style().polish(self.path_label)

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel workbook", "", EXCEL_FILTER)
        if path:
            self.file_dropped.emit(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if _first_local_file(event.mimeData()):
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        path = _first_local_file(event.mimeData())
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        if path and Path(path).suffix.lower() in {".xlsx", ".xlsm"}:
            self.file_dropped.emit(path)
            event.acceptProposedAction()
        else:
            QMessageBox.warning(self, "Excel file required", "Please drop an .xlsx or .xlsm workbook.")
            event.ignore()


def _first_local_file(mime_data: QMimeData) -> str | None:
    if not mime_data.hasUrls():
        return None
    for url in mime_data.urls():
        if url.isLocalFile():
            return url.toLocalFile()
    return None


class WorkerSignals(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(object, object)
    progress = pyqtSignal(int, int)


class ProcessingWorker(QRunnable):
    def __init__(self, khdt: Path, roster: Path, output: Path, force: bool):
        super().__init__()
        self.khdt = khdt
        self.roster = roster
        self.output = output
        self.force = force
        self.signals = WorkerSignals()

    def run(self) -> None:
        log_lines: list[str] = []
        event_log: dict = {}
        try:
            out_path, stats = run(
                self.khdt,
                self.roster,
                self.output,
                force=self.force,
                log=log_lines.append,
                event_log=event_log,
            )
            log_path = Path(out_path).with_suffix(".log.xlsx")
            write_log_workbook(log_path, event_log, log_lines, stats)
            self.signals.finished.emit((Path(out_path), log_path, stats, log_lines))
        except Exception as exc:
            self.signals.failed.emit(exc, log_lines)


class DownloadWorker(QRunnable):
    def __init__(self, info: dict, destination: Path):
        super().__init__()
        self.info = info
        self.destination = destination
        self.signals = WorkerSignals()

    def run(self) -> None:
        ok = updater.download_update(
            self.info["download_url"],
            self.destination,
            progress_cb=lambda done, total: self.signals.progress.emit(done, total),
        )
        self.signals.finished.emit((ok, self.destination))


class UpdateBridge(QObject):
    result = pyqtSignal(object)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KHDT → Roster")
        self.resize(850, 700)
        self.setMinimumSize(720, 580)
        self.thread_pool = QThreadPool.globalInstance()
        self._last_output: Path | None = None
        self.input_dir = get_base_dir() / "INPUT"
        self.input_dir.mkdir(parents=True, exist_ok=True)
        self._update_bridge = UpdateBridge()
        self._update_bridge.result.connect(self._handle_update_result)
        self._build_ui()
        self._load_from_input()
        self._check_for_updates()

    def _build_ui(self) -> None:
        content = QWidget()
        self.setCentralWidget(content)
        outer = QVBoxLayout(content)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("KHDT → Roster")
        title.setObjectName("title")
        title.setFont(QFont("Sans Serif", 21, QFont.Weight.Bold))
        subtitle = QLabel("Mark training days in a roster and save a new Excel workbook.")
        subtitle.setObjectName("subtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        files_group = QGroupBox("1. Add input files")
        files_layout = QVBoxLayout(files_group)
        files_layout.setContentsMargins(14, 18, 14, 14)
        zones_row = QHBoxLayout()
        self.khdt_zone = DropZone("KHDT training plan", "No KHDT file selected")
        self.roster_zone = DropZone("Roster", "No Roster file selected")
        self.khdt_zone.file_dropped.connect(lambda path: self._set_input("khdt", path))
        self.roster_zone.file_dropped.connect(lambda path: self._set_input("roster", path))
        zones_row.addWidget(self.khdt_zone)
        zones_row.addWidget(self.roster_zone)
        files_layout.addLayout(zones_row)

        input_row = QHBoxLayout()
        input_row.addWidget(QLabel(f"Or use files from INPUT/ ({self.input_dir})"))
        input_row.addStretch()
        load_input_button = QPushButton("Load from INPUT")
        load_input_button.clicked.connect(lambda: self._load_from_input(show_message=True))
        input_row.addWidget(load_input_button)
        open_input_button = QPushButton("Open INPUT folder")
        open_input_button.clicked.connect(self._open_input_folder)
        input_row.addWidget(open_input_button)
        files_layout.addLayout(input_row)
        outer.addWidget(files_group)

        options_group = QGroupBox("2. Output options")
        options_layout = QVBoxLayout(options_group)
        options_layout.setContentsMargins(14, 18, 14, 14)
        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("Output folder"))
        self.output_edit = QLineEdit(str(get_base_dir() / "OUTPUT"))
        output_row.addWidget(self.output_edit, stretch=1)
        output_button = QPushButton("Choose…")
        output_button.clicked.connect(self._choose_output_folder)
        output_row.addWidget(output_button)
        options_layout.addLayout(output_row)
        self.force_check = QCheckBox("Force overwrite existing roster cells")
        options_layout.addWidget(self.force_check)
        outer.addWidget(options_group)

        action_row = QHBoxLayout()
        self.run_button = QPushButton("Run and create updated roster")
        self.run_button.setObjectName("runButton")
        self.run_button.clicked.connect(self._start_run)
        action_row.addWidget(self.run_button)
        self.open_button = QPushButton("Open output folder")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self._open_output_folder)
        action_row.addWidget(self.open_button)
        action_row.addStretch()
        update_button = QPushButton("Check for updates")
        update_button.clicked.connect(self._check_for_updates)
        action_row.addWidget(update_button)
        outer.addLayout(action_row)

        self.status_label = QLabel("Add both Excel files to begin.")
        self.status_label.setObjectName("status")
        self.status_label.setWordWrap(True)
        outer.addWidget(self.status_label)

        outer.addWidget(QLabel("Run log"), alignment=Qt.AlignmentFlag.AlignLeft)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Processing messages will appear here.")
        outer.addWidget(self.log_box, stretch=1)

        self.update_label = QLabel(f"Version {APP_VERSION}")
        self.update_label.setObjectName("subtitle")
        outer.addWidget(self.update_label)

    def _set_input(self, kind: str, path: str) -> None:
        if kind == "khdt":
            self.khdt_zone.set_path(path)
        else:
            self.roster_zone.set_path(path)
        if self._khdt_path() and self._roster_path():
            self.status_label.setText("Both files are ready.")
        else:
            self.status_label.setText("Add the other input file to begin.")

    def _load_from_input(self, show_message: bool = False) -> None:
        """Load the uniquely named KHDT and Roster workbooks from INPUT/."""
        khdt, roster, problems = find_input_files(self.input_dir)
        if not problems:
            self._set_input("khdt", str(khdt))
            self._set_input("roster", str(roster))
            self.status_label.setText(f"Loaded both workbooks from {self.input_dir}")
            return
        if show_message:
            QMessageBox.warning(
                self,
                "Could not load INPUT files",
                "Put exactly one workbook with 'KHDT' in its name and one with "
                f"'Roster' in its name into:\n{self.input_dir}\n\n- "
                + "\n- ".join(problems),
            )

    def _open_input_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.input_dir)))

    def _khdt_path(self) -> Path | None:
        return self._path_from_zone(self.khdt_zone)

    def _roster_path(self) -> Path | None:
        return self._path_from_zone(self.roster_zone)

    @staticmethod
    def _path_from_zone(zone: DropZone) -> Path | None:
        text = zone.path_label.text()
        if text.startswith("No ") or not text:
            return None
        return Path(text)

    def _choose_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose output folder", self.output_edit.text())
        if path:
            self.output_edit.setText(path)

    def _start_run(self) -> None:
        khdt = self._khdt_path()
        roster = self._roster_path()
        if khdt is None or roster is None:
            QMessageBox.warning(self, "Missing input", "Add both the KHDT and Roster Excel files first.")
            return
        if not khdt.is_file() or not roster.is_file():
            QMessageBox.critical(self, "File not found", "One of the selected input files no longer exists.")
            return
        if khdt.resolve() == roster.resolve():
            QMessageBox.critical(self, "Invalid input", "KHDT and Roster must be two different files.")
            return

        output_dir = Path(self.output_edit.text()).expanduser()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            QMessageBox.critical(self, "Invalid output folder", str(exc))
            return
        output_path = output_dir / f"{roster.stem}_updated{roster.suffix}"

        self.run_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self.status_label.setText("Processing workbooks…")
        self.log_box.append(f"\nStarting: {khdt.name} + {roster.name}")

        worker = ProcessingWorker(khdt, roster, output_path, self.force_check.isChecked())
        worker.signals.finished.connect(self._run_finished)
        worker.signals.failed.connect(self._run_failed)
        self.thread_pool.start(worker)

    def _run_failed(self, exc: Exception, log_lines: list[str]) -> None:
        self.log_box.append("\n".join(log_lines))
        self.run_button.setEnabled(True)
        self.status_label.setText("Processing failed.")
        QMessageBox.critical(self, "Could not create roster", f"{type(exc).__name__}: {exc}")

    def _run_finished(self, result: tuple) -> None:
        out_path, log_path, stats, log_lines = result
        self.log_box.append("\n".join(log_lines))
        self._last_output = out_path
        self.run_button.setEnabled(True)
        self.open_button.setEnabled(True)
        self.status_label.setText(f"Done — updated roster created at {out_path}")
        summary = (
            "Updated roster created successfully.\n\n"
            f"Cells marked: {stats['marked']}\n"
            f"Weekend N cells: {stats['weekend_n']}\n"
            f"Overlaps combined: {stats['overlaps_combined']}\n"
            f"Conflicts noted: {stats['conflicts']}\n"
            f"IDs not in roster: {stats['no_id_match']}\n"
            f"No date overlap: {stats['no_date_overlap']}\n"
            f"Name/ID mismatches: {stats['name_mismatch']}\n\n"
            f"Roster: {out_path}\n"
            f"Full log: {log_path}"
        )
        QMessageBox.information(self, "Roster updated", summary)

    def _open_output_folder(self) -> None:
        folder = self._last_output.parent if self._last_output else Path(self.output_edit.text())
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _check_for_updates(self) -> None:
        self.update_label.setText("Checking for updates…")
        threading.Thread(
            target=lambda: self._update_bridge.result.emit(
                updater.check_for_update(verbose=False)
            ),
            daemon=True,
        ).start()

    def _handle_update_result(self, info: dict | None) -> None:
        if info is None:
            self.update_label.setText(f"Version {APP_VERSION} • Last update {APP_LAST_UPDATE}")
            return
        version = info["version"]
        if not getattr(sys, "frozen", False):
            self.update_label.setText(f"Update available: {version} • source mode")
            return
        self.update_label.setText(f"Update available: {version}")
        answer = QMessageBox.question(
            self,
            "Update available",
            f"Version {version} is available. Download and install it now?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            self._download_update(info)

    def _download_update(self, info: dict) -> None:
        self.run_button.setEnabled(False)
        self.status_label.setText(f"Downloading {info['asset_name']}…")
        destination = Path(sys.executable).parent / f"_update_{info['asset_name']}"
        worker = DownloadWorker(info, destination)
        worker.signals.progress.connect(self._download_progress)
        worker.signals.finished.connect(self._download_finished)
        self.thread_pool.start(worker)

    def _download_progress(self, done: int, total: int) -> None:
        if total:
            self.status_label.setText(f"Downloading update… {done * 100 // total}%")

    def _download_finished(self, result: tuple[bool, Path]) -> None:
        ok, destination = result
        if not ok:
            self.run_button.setEnabled(True)
            self.status_label.setText("Update download failed.")
            QMessageBox.critical(self, "Update failed", "The update could not be downloaded. Your current version is unchanged.")
            return
        self.status_label.setText("Applying update and restarting…")
        updater.apply_update(destination)

def launch_gui() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("KHDT to Roster")
    icon_path = _find_app_icon()
    if icon_path:
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyle("Fusion")
    app.setStyleSheet(
        """
        QMainWindow, QWidget { background: #f8fafc; color: #1f2937; }
        QGroupBox { border: 1px solid #dbe3ec; border-radius: 8px; margin-top: 10px; padding-top: 8px; font-weight: 600; }
        QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; color: #1f4e78; }
        #title { color: #123b5d; }
        #subtitle, #filePath { color: #64748b; }
        #filePathSelected { color: #1f4e78; font-weight: 600; }
        #dropZone { background: #f1f5f9; border: 1px dashed #a8b8c8; border-radius: 8px; }
        #dropZone[dragActive="true"] { background: #e0f2fe; border: 2px dashed #1683b7; }
        QPushButton { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px 12px; }
        QPushButton:hover { background: #eef6fb; border-color: #6aa6c8; }
        #runButton { background: #1f6f9f; color: white; border: none; font-weight: 700; padding: 10px 18px; }
        #runButton:hover { background: #185a82; }
        #status { color: #1f4e78; padding: 3px 0; }
        QLineEdit, QTextEdit { background: white; border: 1px solid #cbd5e1; border-radius: 5px; padding: 6px; }
        """
    )
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(launch_gui())