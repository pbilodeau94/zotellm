"""
zotellm_gui.py

PyQt6 desktop GUI for zotellm citation formatting.

Usage:
    pip install PyQt6
    python zotellm_gui.py
"""

import argparse
import io
import os
import sys
import threading
from pathlib import Path

# Fix Qt plugin path on Anaconda/macOS before importing QtWidgets
import PyQt6
_qt_plugin_path = os.path.join(os.path.dirname(PyQt6.__file__), "Qt6", "plugins")
if os.path.isdir(_qt_plugin_path):
    os.environ["QT_PLUGIN_PATH"] = _qt_plugin_path

from PyQt6.QtCore import QObject, Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from zotellm import crossref_to_csl, run_zotellm


# ---------------------------------------------------------------------------
# Stdout capture: redirect print() to a Qt signal
# ---------------------------------------------------------------------------

class StdoutSignal(QObject):
    text_written = pyqtSignal(str)


class StdoutRedirector(io.TextIOBase):
    """Replacement for sys.stdout that emits a Qt signal on each write."""

    def __init__(self, signal_obj):
        super().__init__()
        self._signal = signal_obj

    def write(self, text):
        if text:
            self._signal.text_written.emit(text)
        return len(text) if text else 0

    def flush(self):
        pass


# ---------------------------------------------------------------------------
# Worker thread
# ---------------------------------------------------------------------------

class WorkerSignals(QObject):
    log = pyqtSignal(str)
    finished = pyqtSignal(bool, str)  # (success, message)
    resolve_request = pyqtSignal(str, list)  # (citation_text, candidates)


class Worker(threading.Thread):
    """Runs run_zotellm() in a background thread."""

    def __init__(self, args_ns, signals):
        super().__init__(daemon=True)
        self.args_ns = args_ns
        self.signals = signals
        self._resolve_event = threading.Event()
        self._resolve_result = None

    def resolve_callback(self, citation_text, candidates):
        """Called from the worker thread when a match is uncertain."""
        self._resolve_event.clear()
        self._resolve_result = None
        self.signals.resolve_request.emit(citation_text, candidates)
        self._resolve_event.wait()  # blocks worker until GUI responds
        return self._resolve_result

    def set_resolve_result(self, result):
        """Called from the GUI thread to unblock the worker."""
        self._resolve_result = result
        self._resolve_event.set()

    def run(self):
        # Redirect stdout to the log signal
        signal_obj = StdoutSignal()
        signal_obj.text_written.connect(self.signals.log.emit, Qt.ConnectionType.QueuedConnection)
        redirector = StdoutRedirector(signal_obj)
        old_stdout = sys.stdout
        sys.stdout = redirector
        try:
            run_zotellm(self.args_ns, resolve_callback=self.resolve_callback)
            sys.stdout = old_stdout
            self.signals.finished.emit(True, "Formatting complete.")
        except Exception as e:
            sys.stdout = old_stdout
            self.signals.finished.emit(False, str(e))


# ---------------------------------------------------------------------------
# Resolve dialog
# ---------------------------------------------------------------------------

class ResolveDialog(QDialog):
    """Dialog for choosing among uncertain citation matches."""

    def __init__(self, citation_text, candidates, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Resolve Citation Match")
        self.setMinimumWidth(550)
        self.result_value = None  # will be set before accept

        layout = QVBoxLayout(self)

        header = QLabel(f'<b>"{citation_text}"</b> matched to:')
        header.setWordWrap(True)
        layout.addWidget(header)

        self.button_group = QButtonGroup(self)
        self.candidate_items = []  # store crossref items parallel to radio buttons

        for i, (item, score) in enumerate(candidates):
            title = (item.get("title", [""])[0]
                     if isinstance(item.get("title"), list)
                     else item.get("title", "Unknown"))
            authors = item.get("author", [])
            first_author = authors[0].get("family", "") if authors else ""
            year_parts = item.get("issued", {}).get("date-parts", [[]])
            year = str(year_parts[0][0]) if year_parts and year_parts[0] else ""
            doi = item.get("DOI", "")

            label = f"{first_author} ({year}) - {title[:80]}"
            if doi:
                label += f"  [DOI: {doi}]"
            label += f"  (score: {score})"

            rb = QRadioButton(label)
            rb.setWordWrap(True)
            if i == 0:
                rb.setChecked(True)
            self.button_group.addButton(rb, i)
            self.candidate_items.append(item)
            layout.addWidget(rb)

        # Skip option
        self.rb_skip = QRadioButton("Skip this citation")
        self.button_group.addButton(self.rb_skip, len(candidates))
        layout.addWidget(self.rb_skip)

        # Manual DOI/PMID option
        self.rb_manual = QRadioButton("Enter DOI or PMID manually:")
        self.button_group.addButton(self.rb_manual, len(candidates) + 1)
        manual_row = QHBoxLayout()
        manual_row.addWidget(self.rb_manual)
        self.manual_input = QLineEdit()
        self.manual_input.setPlaceholderText("e.g. 10.1000/xyz or 12345678")
        manual_row.addWidget(self.manual_input)
        layout.addLayout(manual_row)

        self.manual_input.textChanged.connect(lambda: self.rb_manual.setChecked(True))

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self._on_accept)
        layout.addWidget(buttons)

    def _on_accept(self):
        checked_id = self.button_group.checkedId()
        if checked_id < len(self.candidate_items):
            self.result_value = self.candidate_items[checked_id]
        elif self.rb_skip.isChecked():
            self.result_value = None
        elif self.rb_manual.isChecked():
            val = self.manual_input.text().strip()
            self.result_value = val if val else None
        self.accept()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class ZotellmWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("zotellm")
        self.setMinimumWidth(600)
        self.worker = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # --- Input / Output ---
        input_row = QHBoxLayout()
        input_row.addWidget(QLabel("Input File:"))
        self.input_edit = QLineEdit()
        input_row.addWidget(self.input_edit)
        input_btn = QPushButton("Browse...")
        input_btn.clicked.connect(self._browse_input)
        input_row.addWidget(input_btn)
        layout.addLayout(input_row)

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("Output File:"))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("(auto: input_zotero.docx)")
        output_row.addWidget(self.output_edit)
        output_btn = QPushButton("Browse...")
        output_btn.clicked.connect(self._browse_output)
        output_row.addWidget(output_btn)
        layout.addLayout(output_row)

        # --- Provider settings ---
        provider_row = QHBoxLayout()
        provider_row.addWidget(QLabel("LLM Provider:"))
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["cli", "openai", "anthropic"])
        self.provider_combo.currentTextChanged.connect(self._on_provider_changed)
        provider_row.addWidget(self.provider_combo)
        provider_row.addStretch()
        layout.addLayout(provider_row)

        cli_row = QHBoxLayout()
        cli_row.addWidget(QLabel("CLI Command:"))
        self.cli_edit = QLineEdit()
        self.cli_edit.setPlaceholderText("(auto-detect: claude, ollama, llm)")
        cli_row.addWidget(self.cli_edit)
        layout.addLayout(cli_row)
        self.cli_label = cli_row.itemAt(0).widget()

        model_row = QHBoxLayout()
        model_row.addWidget(QLabel("Model:"))
        self.model_edit = QLineEdit()
        self.model_edit.setPlaceholderText("(default for provider)")
        model_row.addWidget(self.model_edit)
        layout.addLayout(model_row)
        self.model_label = model_row.itemAt(0).widget()

        key_row = QHBoxLayout()
        key_row.addWidget(QLabel("API Key:"))
        self.key_edit = QLineEdit()
        self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.key_edit.setPlaceholderText("(uses env var if empty)")
        key_row.addWidget(self.key_edit)
        layout.addLayout(key_row)
        self.key_label = key_row.itemAt(0).widget()

        # --- Zotero DB ---
        zotero_row = QHBoxLayout()
        zotero_row.addWidget(QLabel("Zotero DB:"))
        self.zotero_edit = QLineEdit()
        zotero_row.addWidget(self.zotero_edit)
        zotero_btn = QPushButton("Browse...")
        zotero_btn.clicked.connect(self._browse_zotero)
        zotero_row.addWidget(zotero_btn)
        layout.addLayout(zotero_row)

        # Auto-detect Zotero DB
        default_db = Path.home() / "Zotero" / "zotero.sqlite"
        if default_db.exists():
            self.zotero_edit.setText(str(default_db))

        # --- Advanced options (collapsible) ---
        self.advanced_group = QGroupBox("Advanced Options")
        self.advanced_group.setCheckable(True)
        self.advanced_group.setChecked(False)
        adv_layout = QVBoxLayout()

        font_row = QHBoxLayout()
        font_row.addWidget(QLabel("Font:"))
        self.font_edit = QLineEdit("Calibri")
        font_row.addWidget(self.font_edit)
        font_row.addWidget(QLabel("Size:"))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 36)
        self.size_spin.setValue(11)
        font_row.addWidget(self.size_spin)
        adv_layout.addLayout(font_row)

        bib_row = QHBoxLayout()
        bib_row.addWidget(QLabel("Bib Heading:"))
        self.bib_edit = QLineEdit("References")
        bib_row.addWidget(self.bib_edit)
        adv_layout.addLayout(bib_row)

        ref_row = QHBoxLayout()
        ref_row.addWidget(QLabel("Reference Doc:"))
        self.refdoc_edit = QLineEdit()
        self.refdoc_edit.setPlaceholderText("Pandoc reference .docx template")
        ref_row.addWidget(self.refdoc_edit)
        refdoc_btn = QPushButton("Browse...")
        refdoc_btn.clicked.connect(self._browse_refdoc)
        ref_row.addWidget(refdoc_btn)
        adv_layout.addLayout(ref_row)

        checks_row = QHBoxLayout()
        self.no_crossref_cb = QCheckBox("No CrossRef")
        checks_row.addWidget(self.no_crossref_cb)
        self.dry_run_cb = QCheckBox("Dry Run")
        checks_row.addWidget(self.dry_run_cb)
        checks_row.addStretch()
        adv_layout.addLayout(checks_row)

        self.advanced_group.setLayout(adv_layout)
        layout.addWidget(self.advanced_group)

        # --- Format button ---
        self.format_btn = QPushButton("Format Citations")
        self.format_btn.setStyleSheet("font-weight: bold; padding: 8px;")
        self.format_btn.clicked.connect(self._run)
        layout.addWidget(self.format_btn)

        # --- Log ---
        layout.addWidget(QLabel("Log:"))
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.log_edit)

        # Initial provider state
        self._on_provider_changed("cli")

    # --- Provider visibility ---
    def _on_provider_changed(self, provider):
        is_cli = provider == "cli"
        self.cli_edit.setVisible(is_cli)
        self.cli_label.setVisible(is_cli)
        self.model_edit.setVisible(not is_cli)
        self.model_label.setVisible(not is_cli)
        self.key_edit.setVisible(not is_cli)
        self.key_label.setVisible(not is_cli)

    # --- File browsers ---
    def _browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Input File", "",
            "Documents (*.docx *.md *.markdown *.txt);;All Files (*)"
        )
        if path:
            self.input_edit.setText(path)

    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Select Output File", "",
            "Word Documents (*.docx);;All Files (*)"
        )
        if path:
            self.output_edit.setText(path)

    def _browse_zotero(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Zotero Database", "",
            "SQLite (*.sqlite);;All Files (*)"
        )
        if path:
            self.zotero_edit.setText(path)

    def _browse_refdoc(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Reference Doc", "",
            "Word Documents (*.docx);;All Files (*)"
        )
        if path:
            self.refdoc_edit.setText(path)

    # --- Run ---
    def _run(self):
        input_path = self.input_edit.text().strip()
        if not input_path:
            QMessageBox.warning(self, "Missing Input", "Please select an input file.")
            return

        # Build args namespace
        args = argparse.Namespace(
            input=input_path,
            output=self.output_edit.text().strip() or None,
            provider=self.provider_combo.currentText(),
            model=self.model_edit.text().strip() or None,
            api_base=None,
            api_key=self.key_edit.text().strip() or None,
            cli_command=self.cli_edit.text().strip() or None,
            zotero_db=self.zotero_edit.text().strip() or None,
            zotero_api_key=None,
            zotero_library_id=None,
            reference_doc=self.refdoc_edit.text().strip() or None,
            font=self.font_edit.text().strip() or "Calibri",
            size=self.size_spin.value(),
            bib_heading=self.bib_edit.text().strip() or "References",
            no_crossref=self.no_crossref_cb.isChecked(),
            dry_run=self.dry_run_cb.isChecked(),
        )

        self.log_edit.clear()
        self.format_btn.setEnabled(False)
        self.format_btn.setText("Processing...")

        signals = WorkerSignals()
        signals.log.connect(self._append_log)
        signals.finished.connect(self._on_finished)
        signals.resolve_request.connect(self._on_resolve_request)

        self.worker = Worker(args, signals)
        self.worker.start()

    def _append_log(self, text):
        cursor = self.log_edit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        cursor.insertText(text)
        self.log_edit.setTextCursor(cursor)
        self.log_edit.ensureCursorVisible()

    def _on_finished(self, success, message):
        self.format_btn.setEnabled(True)
        self.format_btn.setText("Format Citations")
        if success:
            self._append_log(f"\n{message}\n")
        else:
            self._append_log(f"\nERROR: {message}\n")
            QMessageBox.critical(self, "Error", message)
        self.worker = None

    def _on_resolve_request(self, citation_text, candidates):
        dialog = ResolveDialog(citation_text, candidates, parent=self)
        dialog.exec()
        self.worker.set_resolve_result(dialog.result_value)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("zotellm")
    window = ZotellmWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
