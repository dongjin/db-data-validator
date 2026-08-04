from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QCloseEvent, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .excel_service import ExcelProcessingError, ExcelService, LoadedSheet


PREVIEW_ROW_LIMIT = 1000
MAX_PREVIEW_COLUMN_WIDTH = 520


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.service = ExcelService()
        self.review_progress_dialog: QProgressDialog | None = None
        self.search_match_locations: list[tuple[int, int]] = []
        self.search_match_index = -1
        self.setWindowTitle("DB Data Validator")
        self.resize(1120, 720)
        self.setMinimumSize(860, 560)

        self._create_actions()
        self._create_ui()
        self._set_initial_state()

    def _create_actions(self) -> None:
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu = self.menuBar().addMenu("File")
        file_menu.addAction(exit_action)

    def _create_ui(self) -> None:
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        title = QLabel("DB Data Validator")
        title.setObjectName("titleLabel")
        subtitle = QLabel(
            "Load an Excel or CSV file, choose a primary-key column from sheet 1, review duplicate rows, and save only issue records."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("subtitleLabel")
        root.addWidget(title)
        root.addWidget(subtitle)

        controls = QFrame()
        controls.setObjectName("controlsFrame")
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(14, 12, 14, 12)
        controls_layout.setSpacing(10)

        self.load_button = QPushButton("1. Load Data File")
        self.load_button.clicked.connect(self.load_excel)
        self.process_button = QPushButton("2. Review")
        self.process_button.clicked.connect(self.review_excel)
        self.save_button = QPushButton("3. Save Issue Records")
        self.save_button.clicked.connect(self.save_excel)
        self.primary_key_combo = QComboBox()
        self.primary_key_combo.setMinimumWidth(220)
        self.primary_key_combo.setToolTip("Select the primary-key column from sheet 1")

        controls_layout.addWidget(self.load_button)
        controls_layout.addWidget(QLabel("Primary key column:"))
        controls_layout.addWidget(self.primary_key_combo)
        controls_layout.addWidget(self.process_button)
        controls_layout.addWidget(self.save_button)
        controls_layout.addStretch(1)
        root.addWidget(controls)

        self.file_label = QLabel("No file loaded")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.summary_label = QLabel("")
        summary_row = QHBoxLayout()
        summary_row.addWidget(self.file_label, 1)
        summary_row.addWidget(self.summary_label)
        root.addLayout(summary_row)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        search_row.addWidget(QLabel("Search all columns:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Type text to find in any column")
        self.search_input.setMaxLength(50)
        self.search_input.setFixedWidth(380)
        self.search_input.returnPressed.connect(self._run_search)
        self.find_button = QPushButton("Find")
        self.find_button.clicked.connect(self._run_search)
        self.previous_button = QPushButton("<")
        self.previous_button.setObjectName("navArrowButton")
        self.previous_button.setToolTip("Previous match")
        self.previous_button.setFixedSize(46, 34)
        self.previous_button.clicked.connect(self._go_to_previous_match)
        self.next_button = QPushButton(">")
        self.next_button.setObjectName("navArrowButton")
        self.next_button.setToolTip("Next match")
        self.next_button.setFixedSize(46, 34)
        self.next_button.clicked.connect(self._go_to_next_match)
        self.search_status_label = QLabel("No search")
        search_row.addWidget(self.search_input)
        search_row.addWidget(self.find_button)
        search_row.addWidget(self.previous_button)
        search_row.addWidget(self.next_button)
        search_row.addWidget(self.search_status_label)
        root.addLayout(search_row)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        root.addWidget(self.tabs, 1)

        note = QLabel(
            f"The full first-sheet data is loaded for review. "
            f"For responsiveness, each preview tab displays at most {PREVIEW_ROW_LIMIT:,} rows."
        )
        note.setObjectName("noteLabel")
        note.setWordWrap(True)
        root.addWidget(note)

        self.setCentralWidget(central)
        status = QStatusBar(self)
        self.setStatusBar(status)
        self._apply_styles()

    def _set_initial_state(self) -> None:
        self.process_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.primary_key_combo.setEnabled(False)
        self.previous_button.setEnabled(False)
        self.next_button.setEnabled(False)
        self.statusBar().showMessage("Ready")

    def load_excel(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Data File",
            str(Path.home()),
            "Data files (*.xlsx *.xlsm *.csv)",
        )
        if not file_path:
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            loaded_sheets = self.service.load(file_path)
            self._display_sheets(loaded_sheets)
            self.file_label.setText(str(self.service.input_path))
            self._update_summary(loaded_sheets)
            self._populate_primary_key_columns()
            self.process_button.setEnabled(True)
            self.save_button.setEnabled(False)
            self.statusBar().showMessage(
                f"Loaded {Path(file_path).name}", 5000
            )
        except ExcelProcessingError as exc:
            self._show_error("Could not load data file", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def review_excel(self) -> None:
        selected_column = self.primary_key_combo.currentText().strip()
        if not selected_column:
            self._show_error(
                "Missing primary key column",
                "Choose a primary key column before running review.",
            )
            return

        self._set_review_running_state(True)
        self._start_review_progress_dialog()
        try:
            duplicate_key_count, duplicate_row_count = self.service.review(
                selected_column,
                progress_callback=self._on_review_progress,
            )
            self._display_sheets(self.service.loaded_sheets)
            self._update_summary(self.service.loaded_sheets)
            self.save_button.setEnabled(True)
            QMessageBox.information(
                self,
                "Review complete",
                (
                    f"Duplicate primary-key values found: {duplicate_key_count}\n"
                    f"Rows with issues: {duplicate_row_count}\n\n"
                    "See the 'Duplicate Summary' and 'Issue Records' tabs for details."
                ),
            )
            self.statusBar().showMessage("Review complete", 5000)
        except ExcelProcessingError as exc:
            self._show_error("Review failed", str(exc))
        finally:
            self._finish_review_progress_dialog()
            self._set_review_running_state(False)

    def save_excel(self) -> None:
        try:
            default_path = self.service.default_output_path()
        except ExcelProcessingError as exc:
            self._show_error("Nothing to save", str(exc))
            return

        file_filter = (
            "CSV File (*.csv)"
            if default_path.suffix.lower() == ".csv"
            else (
                "Excel Macro-Enabled Workbook (*.xlsm)"
                if default_path.suffix.lower() == ".xlsm"
                else "Excel Workbook (*.xlsx)"
            )
        )
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Duplicate Issue Workbook",
            str(default_path),
            file_filter,
        )
        if not output_path:
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            saved_path = self.service.save_as(output_path)
            QMessageBox.information(
                self,
                "File saved",
                f"The duplicate issue workbook was saved to:\n\n{saved_path}",
            )
            self.statusBar().showMessage(f"Saved {saved_path.name}", 6000)
        except ExcelProcessingError as exc:
            self._show_error("Could not save file", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def _display_sheets(self, sheets: list[LoadedSheet]) -> None:
        self.tabs.clear()
        for sheet in sheets:
            table = self._create_table(sheet)
            preview_count = min(sheet.row_count, PREVIEW_ROW_LIMIT)
            tab_title = f"{sheet.name} ({sheet.row_count:,} rows)"
            self.tabs.addTab(table, tab_title)
            table.setToolTip(
                f"Showing {preview_count:,} of {sheet.row_count:,} loaded rows"
            )
        self._run_search()

    def _update_summary(self, sheets: list[LoadedSheet]) -> None:
        total_rows = sum(sheet.row_count for sheet in sheets)
        self.summary_label.setText(f"{len(sheets)} sheets · {total_rows:,} loaded rows")

    def _populate_primary_key_columns(self) -> None:
        self.primary_key_combo.clear()
        self.primary_key_combo.addItems(self.service.primary_key_columns)
        self.primary_key_combo.setEnabled(bool(self.service.primary_key_columns))

    def _create_table(self, sheet: LoadedSheet) -> QTableWidget:
        preview_rows = sheet.rows[:PREVIEW_ROW_LIMIT]
        table = QTableWidget(len(preview_rows), sheet.max_columns)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSortingEnabled(False)
        table.setWordWrap(False)

        headers = [get_column_letter(index) for index in range(1, sheet.max_columns + 1)]
        table.setHorizontalHeaderLabels(headers)
        table.setVerticalHeaderLabels([str(index) for index in range(1, len(preview_rows) + 1)])
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)

        for row_index, row in enumerate(preview_rows):
            for column_index in range(sheet.max_columns):
                value: Any = row[column_index] if column_index < len(row) else None
                item = QTableWidgetItem(self._display_value(value))
                if row_index == 0:
                    header_font = item.font()
                    header_font.setBold(True)
                    item.setFont(header_font)
                    item.setBackground(QColor("#f3f6fb"))
                item.setToolTip(item.text())
                table.setItem(row_index, column_index, item)

        table.resizeColumnsToContents()
        for column in range(table.columnCount()):
            if table.columnWidth(column) > MAX_PREVIEW_COLUMN_WIDTH:
                table.setColumnWidth(column, MAX_PREVIEW_COLUMN_WIDTH)
        self._expand_columns_to_available_width(table)
        return table

    @staticmethod
    def _expand_columns_to_available_width(table: QTableWidget) -> None:
        column_count = table.columnCount()
        if column_count == 0:
            return

        used_width = sum(table.columnWidth(column) for column in range(column_count))
        available_width = table.viewport().width()
        if used_width >= available_width:
            return

        extra = available_width - used_width
        add_per_column = extra // column_count
        remainder = extra % column_count
        for column in range(column_count):
            bonus = 1 if column < remainder else 0
            table.setColumnWidth(column, table.columnWidth(column) + add_per_column + bonus)

    @staticmethod
    def _display_value(value: Any) -> str:
        if value is None:
            return ""
        return str(value)

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)
        self.statusBar().showMessage(message, 7000)

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 (Qt naming)
        self.service.close()
        event.accept()

    def _set_review_running_state(self, is_running: bool) -> None:
        self.load_button.setEnabled(not is_running)
        self.process_button.setEnabled(not is_running and bool(self.service.primary_key_columns))
        self.primary_key_combo.setEnabled(not is_running and bool(self.service.primary_key_columns))
        self.save_button.setEnabled(not is_running and self.service.reviewed)
        self.find_button.setEnabled(not is_running)
        self.search_input.setEnabled(not is_running)
        self.previous_button.setEnabled(not is_running and len(self.search_match_locations) > 0)
        self.next_button.setEnabled(not is_running and len(self.search_match_locations) > 0)

    def _start_review_progress_dialog(self) -> None:
        dialog = QProgressDialog("Reviewing records...", "", 0, 100, self)
        dialog.setWindowTitle("Review in progress")
        dialog.setWindowModality(Qt.WindowModal)
        dialog.setMinimumDuration(0)
        dialog.setCancelButton(None)
        dialog.setAutoClose(False)
        dialog.setAutoReset(False)
        dialog.setValue(0)
        dialog.show()
        self.review_progress_dialog = dialog
        self.statusBar().showMessage("Review started...", 2000)

    def _on_review_progress(self, value: int) -> None:
        if self.review_progress_dialog is None:
            return
        self.review_progress_dialog.setValue(max(0, min(100, value)))
        QApplication.processEvents()

    def _finish_review_progress_dialog(self) -> None:
        if self.review_progress_dialog is not None:
            self.review_progress_dialog.setValue(100)
            self.review_progress_dialog.close()
            self.review_progress_dialog.deleteLater()
            self.review_progress_dialog = None

    def _on_tab_changed(self, _index: int) -> None:
        self._run_search()

    def _run_search(self) -> None:
        query = self.search_input.text().strip().casefold()
        table = self._current_table()
        self.search_match_locations = []
        self.search_match_index = -1

        if table is None:
            self._update_search_controls()
            return
        if not query:
            table.clearSelection()
            self.search_status_label.setText("No search")
            self._update_search_controls()
            return

        for row_index in range(1, table.rowCount()):
            first_match_column: int | None = None
            for column_index in range(table.columnCount()):
                item = table.item(row_index, column_index)
                value = item.text().casefold() if item is not None else ""
                if query in value:
                    first_match_column = column_index
                    break
            if first_match_column is not None:
                self.search_match_locations.append((row_index, first_match_column))

        if self.search_match_locations:
            self.search_match_index = 0
            self._focus_match(self.search_match_locations[0])
            self.search_status_label.setText(
                f"{len(self.search_match_locations)} matches · 1/{len(self.search_match_locations)}"
            )
        else:
            table.clearSelection()
            self.search_status_label.setText("No matches")
        self._update_search_controls()

    def _go_to_previous_match(self) -> None:
        if not self.search_match_locations:
            return
        self.search_match_index = (self.search_match_index - 1) % len(self.search_match_locations)
        self._focus_current_match()

    def _go_to_next_match(self) -> None:
        if not self.search_match_locations:
            return
        self.search_match_index = (self.search_match_index + 1) % len(self.search_match_locations)
        self._focus_current_match()

    def _focus_current_match(self) -> None:
        if not self.search_match_locations:
            return
        self._focus_match(self.search_match_locations[self.search_match_index])
        self.search_status_label.setText(
            f"{len(self.search_match_locations)} matches · {self.search_match_index + 1}/{len(self.search_match_locations)}"
        )
        self._update_search_controls()

    def _focus_match(self, match_location: tuple[int, int]) -> None:
        table = self._current_table()
        if table is None or table.columnCount() == 0:
            return
        row_index, column_index = match_location
        item = table.item(row_index, column_index)
        if item is None:
            return
        table.setCurrentCell(row_index, column_index)
        table.selectRow(row_index)
        table.scrollToItem(item, QAbstractItemView.ScrollHint.PositionAtCenter)

    def _update_search_controls(self) -> None:
        has_matches = len(self.search_match_locations) > 0
        review_running = self.review_progress_dialog is not None
        self.previous_button.setEnabled(has_matches and not review_running)
        self.next_button.setEnabled(has_matches and not review_running)

    def _current_table(self) -> QTableWidget | None:
        widget = self.tabs.currentWidget()
        if isinstance(widget, QTableWidget):
            return widget
        return None

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #f6f7f9; }
            QLabel#titleLabel { font-size: 24px; font-weight: 700; color: #172033; }
            QLabel#subtitleLabel { font-size: 13px; color: #546174; }
            QLabel#noteLabel { color: #667085; font-size: 12px; }
            QFrame#controlsFrame {
                background: white;
                border: 1px solid #dfe3e8;
                border-radius: 8px;
            }
            QPushButton {
                min-height: 34px;
                padding: 0 14px;
                border: 1px solid #bcc4ce;
                border-radius: 6px;
                background: white;
            }
            QPushButton:hover:enabled { background: #eef4ff; border-color: #7aa2e3; }
            QPushButton:disabled { color: #9aa3af; background: #f0f1f3; }
            QPushButton#navArrowButton { font-size: 19px; font-weight: 700; padding: 0; }
            QTabWidget::pane { border: 1px solid #dfe3e8; background: white; }
            QTableWidget { background: white; gridline-color: #e6e9ee; }
            """
        )


def run() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("DB Data Validator")
    window = MainWindow()
    window.show()
    return app.exec()
