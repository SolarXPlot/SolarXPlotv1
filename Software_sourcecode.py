import os
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
import sys
import pandas as pd
import numpy as np
from PyQt5 import QtWidgets, QtGui, QtCore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, NullLocator
import matplotlib.font_manager as fm
from peakutils import baseline
import re
from xrd_plotter import XRDPlotter
from eqe_plotter import EQEPlotter
from widgets import CollapsibleSection
import seaborn as sns

class FileItemWidget(QtWidgets.QWidget):
    def __init__(self, filename, device_name=None, parent=None):
        super().__init__(parent)
        layout = QtWidgets.QHBoxLayout()
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(4)
        
        # Editable filename
        self.label = QtWidgets.QLineEdit(filename)
        self.label.setStyleSheet("QLineEdit { border: none; background: transparent; }")
        self.label.editingFinished.connect(self.on_edit_finished)
        
        # Editable device name
        self.deviceEdit = QtWidgets.QLineEdit(device_name if device_name is not None else "Device 1")
        self.deviceEdit.setFixedWidth(90)
        self.deviceEdit.setPlaceholderText("Device name")
        self.deviceEdit.setStyleSheet("QLineEdit { border: 1px solid #ccc; border-radius: 4px; background: #f9f9f9; padding: 2px 4px; }")
        
        self.closeButton = QtWidgets.QPushButton("×")
        self.closeButton.setFixedSize(20, 20)
        self.closeButton.setStyleSheet("""
            QPushButton {
                background-color: #ff6b6b;
                border: none;
                color: white;
                font-weight: bold;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #ff4757;
            }
        """)
        
        layout.addWidget(self.label)
        layout.addWidget(self.deviceEdit)
        layout.addWidget(self.closeButton)
        layout.addStretch()
        self.setLayout(layout)
        
        # Store the original filename
        self.original_filename = filename
        
    def on_edit_finished(self):
        # Ensure the label is not empty
        if not self.label.text():
            self.label.setText(self.original_filename)
        # Optionally: ensure device name is not empty
        if not self.deviceEdit.text():
            self.deviceEdit.setText("Device 1")
        
    def get_label_text(self):
        return self.label.text()
    
    def get_device_name(self):
        return self.deviceEdit.text()

class PlotEditDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Plot Settings")
        self.setModal(True)
        
        # Get current values from parent
        self.xmax = parent.xmaxSpin.value()
        self.ymax = parent.ymaxSpin.value()
        self.xstep = parent.xStepSpin.value()
        self.ystep = parent.yStepSpin.value()
        self.graph_type = parent.scaleCombo.currentText()
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Create form layout for settings
        form = QtWidgets.QFormLayout()
        
        # X Max spinner
        self.xmaxSpin = QtWidgets.QDoubleSpinBox()
        self.xmaxSpin.setRange(0, 1e4)
        self.xmaxSpin.setDecimals(3)
        self.xmaxSpin.setValue(self.xmax)
        form.addRow("X Max (V):", self.xmaxSpin)
        
        # Y Max spinner
        self.ymaxSpin = QtWidgets.QDoubleSpinBox()
        self.ymaxSpin.setRange(0, 1e6)
        self.ymaxSpin.setValue(self.ymax)
        form.addRow("Y Max:", self.ymaxSpin)
        
        # X Tick Step spinner
        self.xStepSpin = QtWidgets.QDoubleSpinBox()
        self.xStepSpin.setRange(0.001, 1e3)
        self.xStepSpin.setDecimals(3)
        self.xStepSpin.setSingleStep(0.1)
        self.xStepSpin.setValue(self.xstep)
        form.addRow("X Tick Step (V):", self.xStepSpin)
        
        # Y Tick Step spinner
        self.yStepSpin = QtWidgets.QDoubleSpinBox()
        self.yStepSpin.setRange(0.1, 1e6)
        self.yStepSpin.setDecimals(3)
        self.yStepSpin.setSingleStep(1)
        self.yStepSpin.setValue(self.ystep)
        form.addRow("Y Tick Step:", self.yStepSpin)
        
        # Graph Type combo
        self.scaleCombo = QtWidgets.QComboBox()
        self.scaleCombo.addItems(["Linear", "Log X", "Log Y", "Log XY"])
        self.scaleCombo.setCurrentText(self.graph_type)
        form.addRow("Graph Type:", self.scaleCombo)
        
        layout.addLayout(form)
        
        # Add buttons
        buttons = QtWidgets.QHBoxLayout()
        self.okButton = QtWidgets.QPushButton("OK")
        self.cancelButton = QtWidgets.QPushButton("Cancel")
        buttons.addWidget(self.okButton)
        buttons.addWidget(self.cancelButton)
        layout.addLayout(buttons)
        
        # Connect buttons
        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        
        # Apply same styling as parent
        self.setStyleSheet(
            "QLabel, QSpinBox, QDoubleSpinBox, QComboBox, QLineEdit, QListWidget, QCheckBox { font-size: 13pt; }"
            "QPushButton { padding: 8px; font-size: 12pt; }"
        )

class BackgroundSettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Background Settings")
        self.setModal(True)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Create form layout for settings
        form = QtWidgets.QFormLayout()
        
        # Background type selector
        self.bgTypeCombo = QtWidgets.QComboBox()
        self.bgTypeCombo.addItems(["Solid", "Gradient"])
        form.addRow("Background Type:", self.bgTypeCombo)
        
        # Solid color selector
        color_layout = QtWidgets.QHBoxLayout()
        self.bgColorBtn = QtWidgets.QPushButton()
        self.bgColorBtn.setFixedSize(30, 30)
        self.bgColor = parent.bgColor if parent else QtGui.QColor(255, 255, 255)
        self._update_bg_color_button()
        color_layout.addWidget(self.bgColorBtn)
        color_layout.addStretch()
        form.addRow("Background Color:", color_layout)
        
        # Gradient settings
        self.gradientGroup = QtWidgets.QGroupBox("Gradient Settings")
        gradient_layout = QtWidgets.QVBoxLayout()
        
        # Gradient presets
        self.gradientPresetsCombo = QtWidgets.QComboBox()
        self.gradientPresetsCombo.addItems([
            "Custom",
            "Grade Gray (#bdc3c7 → #2c3e50)",
            "Ocean (#2980b9 → #2c3e50)",
            "Sunset (#e74c3c → #2c3e50)",
            "Forest (#27ae60 → #2c3e50)",
            "Purple (#9b59b6 → #2c3e50)"
        ])
        gradient_layout.addWidget(QtWidgets.QLabel("Preset Gradients:"))
        gradient_layout.addWidget(self.gradientPresetsCombo)
        
        # Custom gradient colors
        custom_colors = QtWidgets.QHBoxLayout()
        self.gradientStart = QtWidgets.QPushButton()
        self.gradientStart.setFixedSize(30, 30)
        self.gradientEnd = QtWidgets.QPushButton()
        self.gradientEnd.setFixedSize(30, 30)
        
        self.startColor = parent.bgColor if parent else QtGui.QColor("#bdc3c7")
        self.endColor = parent.bgColor2 if parent else QtGui.QColor("#2c3e50")
        
        self._update_gradient_buttons()
        
        custom_colors.addWidget(QtWidgets.QLabel("Start:"))
        custom_colors.addWidget(self.gradientStart)
        custom_colors.addWidget(QtWidgets.QLabel("End:"))
        custom_colors.addWidget(self.gradientEnd)
        custom_colors.addStretch()
        
        gradient_layout.addLayout(custom_colors)
        self.gradientGroup.setLayout(gradient_layout)
        form.addRow(self.gradientGroup)
        
        layout.addLayout(form)
        
        # Add buttons
        buttons = QtWidgets.QHBoxLayout()
        self.okButton = QtWidgets.QPushButton("OK")
        self.cancelButton = QtWidgets.QPushButton("Cancel")
        buttons.addWidget(self.okButton)
        buttons.addWidget(self.cancelButton)
        layout.addLayout(buttons)
        
        # Connect signals
        self.bgTypeCombo.currentTextChanged.connect(self.on_bg_type_changed)
        self.bgColorBtn.clicked.connect(self.pick_bg_color)
        self.gradientStart.clicked.connect(self.pick_gradient_start)
        self.gradientEnd.clicked.connect(self.pick_gradient_end)
        self.gradientPresetsCombo.currentTextChanged.connect(self.on_preset_changed)
        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        
        # Initial state
        self.on_bg_type_changed(self.bgTypeCombo.currentText())
        
    def _update_bg_color_button(self):
        pix = QtGui.QPixmap(30, 30)
        pix.fill(self.bgColor)
        self.bgColorBtn.setIcon(QtGui.QIcon(pix))
        
    def _update_gradient_buttons(self):
        pix_start = QtGui.QPixmap(30, 30)
        pix_start.fill(self.startColor)
        self.gradientStart.setIcon(QtGui.QIcon(pix_start))
        
        pix_end = QtGui.QPixmap(30, 30)
        pix_end.fill(self.endColor)
        self.gradientEnd.setIcon(QtGui.QIcon(pix_end))
        
    def on_bg_type_changed(self, bg_type):
        self.gradientGroup.setVisible(bg_type == "Gradient")
        self.adjustSize()
        
    def pick_bg_color(self):
        color = QtWidgets.QColorDialog.getColor(self.bgColor, self)
        if color.isValid():
            self.bgColor = color
            self._update_bg_color_button()
            
    def pick_gradient_start(self):
        color = QtWidgets.QColorDialog.getColor(self.startColor, self)
        if color.isValid():
            self.startColor = color
            self._update_gradient_buttons()
            self.gradientPresetsCombo.setCurrentText("Custom")
            
    def pick_gradient_end(self):
        color = QtWidgets.QColorDialog.getColor(self.endColor, self)
        if color.isValid():
            self.endColor = color
            self._update_gradient_buttons()
            self.gradientPresetsCombo.setCurrentText("Custom")
            
    def on_preset_changed(self, preset):
        if preset == "Custom":
            return
            
        # Extract colors from preset name
        if "Grade Gray" in preset:
            self.startColor = QtGui.QColor("#bdc3c7")
            self.endColor = QtGui.QColor("#2c3e50")
        elif "Ocean" in preset:
            self.startColor = QtGui.QColor("#2980b9")
            self.endColor = QtGui.QColor("#2c3e50")
        elif "Sunset" in preset:
            self.startColor = QtGui.QColor("#e74c3c")
            self.endColor = QtGui.QColor("#2c3e50")
        elif "Forest" in preset:
            self.startColor = QtGui.QColor("#27ae60")
            self.endColor = QtGui.QColor("#2c3e50")
        elif "Purple" in preset:
            self.startColor = QtGui.QColor("#9b59b6")
            self.endColor = QtGui.QColor("#2c3e50")
            
        self._update_gradient_buttons()

class DataSelectionDialog(QtWidgets.QDialog):
    def __init__(self, data, filename, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Data Columns")
        self.setMinimumSize(800, 600)
        self.data = data
        self.filename = filename
        self.selections = []  # Store multiple selections
        self.current_selection_number = 1
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Remove window frame and add custom title bar
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Dialog)

        # Custom title bar
        self.title_bar = QtWidgets.QFrame()
        self.title_bar.setStyleSheet("background: #f5f5f5; border-top-left-radius: 8px; border-top-right-radius: 8px;")
        self.title_bar.setFixedHeight(36)
        title_layout = QtWidgets.QHBoxLayout(self.title_bar)
        title_layout.setContentsMargins(8, 0, 8, 0)
        title_layout.setSpacing(0)
        self.title_label = QtWidgets.QLabel("Select Data Columns")
        self.title_label.setStyleSheet("font-size: 13pt; font-weight: bold; color: #333;")
        title_layout.addWidget(self.title_label)
        title_layout.addStretch()
        # Edit Mode tool button (icon only)
        self.edit_mode_btn = QtWidgets.QToolButton()
        self.edit_mode_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogDetailedView))
        self.edit_mode_btn.setToolTip("Edit Mode (Green: ON)")
        self.edit_mode_btn.setCheckable(True)
        self.edit_mode_btn.setFixedSize(28, 28)
        self.edit_mode_btn.setStyleSheet("")
        self.edit_mode_btn.clicked.connect(self.toggle_edit_mode)
        title_layout.addWidget(self.edit_mode_btn)
        # Minimize button
        self.min_btn = QtWidgets.QToolButton()
        self.min_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarMinButton))
        self.min_btn.setToolTip("Minimize")
        self.min_btn.setFixedSize(28, 28)
        self.min_btn.clicked.connect(self.showMinimized)
        title_layout.addWidget(self.min_btn)
        # Maximize/restore button
        self.max_btn = QtWidgets.QToolButton()
        self.max_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarMaxButton))
        self.max_btn.setToolTip("Full Screen")
        self.max_btn.setFixedSize(28, 28)
        self.max_btn.clicked.connect(self.toggle_fullscreen)
        title_layout.addWidget(self.max_btn)
        # Close button
        self.close_btn = QtWidgets.QToolButton()
        self.close_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarCloseButton))
        self.close_btn.setToolTip("Close")
        self.close_btn.setFixedSize(28, 28)
        self.close_btn.clicked.connect(self.reject)
        title_layout.addWidget(self.close_btn)
        # Add drag support
        self.title_bar.mousePressEvent = self._title_mouse_press
        self.title_bar.mouseMoveEvent = self._title_mouse_move
        self._drag_pos = None
        # Insert title bar at the top
        layout.insertWidget(0, self.title_bar)
        
        # Add toolbar with full screen and other options
        toolbar_layout = QtWidgets.QHBoxLayout()
        
        # Add selection counter label
        self.selection_counter = QtWidgets.QLabel(f"Plot Selection #{self.current_selection_number}")
        self.selection_counter.setStyleSheet("font-size: 14pt; font-weight: bold; color: #2c3e50;")
        toolbar_layout.addWidget(self.selection_counter)
        
        layout.addLayout(toolbar_layout)
        
        # Add table widget
        self.table = QtWidgets.QTableWidget()
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)  # Start with no edit
        self.table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.table)
        
        # Selection info
        selection_layout = QtWidgets.QHBoxLayout()
        self.voltage_label = QtWidgets.QLabel("Voltage Column: None")
        self.current_label = QtWidgets.QLabel("Current Column: None")
        selection_layout.addWidget(self.voltage_label)
        selection_layout.addWidget(self.current_label)
        layout.addLayout(selection_layout)
        
        # Add row range selection widgets
        row_layout = QtWidgets.QHBoxLayout()
        
        # Start row selection
        start_row_label = QtWidgets.QLabel("Start Row:")
        self.start_row_spin = QtWidgets.QSpinBox()
        self.start_row_spin.setRange(1, len(data))
        self.start_row_spin.setValue(1)
        self.start_row_spin.setToolTip("Select the first row of data to include")
        row_layout.addWidget(start_row_label)
        row_layout.addWidget(self.start_row_spin)
        
        # End row selection
        end_row_label = QtWidgets.QLabel("End Row:")
        self.end_row_spin = QtWidgets.QSpinBox()
        self.end_row_spin.setRange(1, len(data))
        self.end_row_spin.setValue(len(data))
        self.end_row_spin.setToolTip("Select the last row of data to include")
        row_layout.addWidget(end_row_label)
        row_layout.addWidget(self.end_row_spin)
        
        # Add "Use Selected Range" button
        self.use_selection_btn = QtWidgets.QPushButton("Use Selected Range")
        self.use_selection_btn.setToolTip("Set row range based on current table selection")
        self.use_selection_btn.clicked.connect(self.use_selected_range)
        row_layout.addWidget(self.use_selection_btn)
        
        row_layout.addStretch()
        layout.addLayout(row_layout)
        
        # Add custom name field
        name_layout = QtWidgets.QHBoxLayout()
        name_layout.addWidget(QtWidgets.QLabel("Plot Name:"))
        self.name_edit = QtWidgets.QLineEdit()
        self.name_edit.setPlaceholderText("Enter a name for this selection")
        self.name_edit.setText(f"{os.path.splitext(self.filename)[0]}_{self.current_selection_number}")
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)
        
        # Instructions label
        instructions = QtWidgets.QLabel(
            "Instructions:\n"
            "1. Click column header to select Voltage column (blue)\n"
            "2. Ctrl+Click column header to select Current column (green)\n"
            "3. Select rows in the table OR use the spinboxes to set row range\n"
            "4. Click 'Use Selected Range' to set row range from table selection\n"
            "5. Enter a name for this selection\n"
            "6. Click 'Next Plot' to add another selection or 'OK' when done\n"
            "7. Use 'Full Screen' for better data viewing\n"
            "8. Enable 'Edit Mode' to edit cell values directly\n"
            "9. Right-click cells for context menu options"
        )
        instructions.setStyleSheet("QLabel { background-color: #f0f0f0; padding: 10px; }")
        layout.addWidget(instructions)
        
        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        self.next_button = QtWidgets.QPushButton("Next Plot")
        self.next_button.setEnabled(False)
        self.next_button.clicked.connect(self.save_and_continue)
        self.ok_button = QtWidgets.QPushButton("OK")
        self.ok_button.setEnabled(False)
        cancel_button = QtWidgets.QPushButton("Cancel")
        button_layout.addWidget(self.next_button)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        # Add status bar
        self.status_bar = QtWidgets.QLabel("Ready - Press F2 to edit cells, Escape to exit full screen")
        self.status_bar.setStyleSheet("QLabel { background-color: #f0f0f0; padding: 5px; border-top: 1px solid #ccc; }")
        layout.addWidget(self.status_bar)
        
        # Connect buttons
        self.ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        # Connect spinbox value changes
        self.start_row_spin.valueChanged.connect(self.validate_row_range)
        self.end_row_spin.valueChanged.connect(self.validate_row_range)
        
        # Setup table
        self.setup_table(data)
        
        # Store selected columns
        self.voltage_col = None
        self.current_col = None
        
        # Connect selection changed signal
        self.table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)
        
        # Connect cell editing signals
        self.table.itemChanged.connect(self.on_cell_changed)
        
        # Track edit mode state
        self.edit_mode = False

    def toggle_fullscreen(self):
        """Toggle full screen mode for the dialog and update icon"""
        if self.isFullScreen():
            self.showNormal()
            self.max_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarMaxButton))
            self.max_btn.setToolTip("Full Screen")
            self.status_bar.setText("Exited full screen mode")
        else:
            self.showFullScreen()
            self.max_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarNormalButton))
            self.max_btn.setToolTip("Restore Down")
            self.status_bar.setText("Entered full screen mode - Press Escape to exit")

    def toggle_edit_mode(self):
        """Toggle edit mode for the table"""
        self.edit_mode = not self.edit_mode
        if self.edit_mode:
            self.table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | 
                                     QtWidgets.QAbstractItemView.EditKeyPressed)
            self.edit_mode_btn.setChecked(True)
            self.edit_mode_btn.setStyleSheet("QToolButton { background-color: #4CAF50; color: white; border-radius: 6px; }")
            self.status_bar.setText("Edit mode enabled - Double-click or press F2 to edit cells")
        else:
            self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            self.edit_mode_btn.setChecked(False)
            self.edit_mode_btn.setStyleSheet("")
            self.status_bar.setText("Edit mode disabled")

    def show_context_menu(self, position):
        """Show context menu for right-click on table"""
        context_menu = QtWidgets.QMenu(self)
        
        # Get the item at the clicked position
        item = self.table.itemAt(position)
        if item is not None:
            # Edit cell action
            edit_action = context_menu.addAction("Edit Cell")
            edit_action.triggered.connect(lambda: self.edit_cell_at_position(position))
            
            # Delete cell action
            delete_action = context_menu.addAction("Delete Cell")
            delete_action.triggered.connect(lambda: self.delete_cell_at_position(position))
            
            context_menu.addSeparator()
        
        # Add general actions
        copy_action = context_menu.addAction("Copy Selection")
        copy_action.triggered.connect(self.copy_selection)
        
        paste_action = context_menu.addAction("Paste")
        paste_action.triggered.connect(self.paste_selection)
        
        context_menu.addSeparator()
        
        # Add column selection actions
        select_voltage_action = context_menu.addAction("Select as Voltage Column")
        select_voltage_action.triggered.connect(lambda: self.select_column_as_voltage(position))
        
        select_current_action = context_menu.addAction("Select as Current Column")
        select_current_action.triggered.connect(lambda: self.select_column_as_current(position))
        
        context_menu.exec_(self.table.mapToGlobal(position))

    def edit_cell_at_position(self, position):
        """Edit the cell at the given position"""
        item = self.table.itemAt(position)
        if item is not None:
            self.table.editItem(item)
            self.status_bar.setText(f"Editing cell at row {item.row()+1}, column {item.column()+1}")

    def delete_cell_at_position(self, position):
        """Delete the cell at the given position"""
        item = self.table.itemAt(position)
        if item is not None:
            item.setText("")
            # Update the underlying data
            row = item.row()
            col = item.column()
            if isinstance(self.data, pd.DataFrame):
                # For DataFrame, we need to handle this carefully
                try:
                    # Try to convert to numeric and set to NaN
                    self.data.iloc[row, col] = np.nan
                except:
                    self.data.iloc[row, col] = np.nan
            else:
                # For numpy array
                self.data[row, col] = np.nan
            self.status_bar.setText(f"Deleted cell at row {row+1}, column {col+1}")

    def copy_selection(self):
        """Copy selected cells to clipboard"""
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            return
            
        clipboard_text = ""
        for range_obj in selected_ranges:
            for row in range(range_obj.topRow(), range_obj.bottomRow() + 1):
                row_data = []
                for col in range(range_obj.leftColumn(), range_obj.rightColumn() + 1):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                clipboard_text += "\t".join(row_data) + "\n"
        
        clipboard = QtWidgets.QApplication.clipboard()
        clipboard.setText(clipboard_text)
        self.status_bar.setText(f"Copied {len(selected_ranges)} selection(s) to clipboard")

    def paste_selection(self):
        """Paste clipboard content to selected cells"""
        clipboard = QtWidgets.QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
            
        # Get current selection
        current_item = self.table.currentItem()
        if current_item is None:
            return
            
        start_row = current_item.row()
        start_col = current_item.column()
        
        # Parse clipboard text
        lines = text.strip().split('\n')
        for i, line in enumerate(lines):
            if start_row + i >= self.table.rowCount():
                break
            cells = line.split('\t')
            for j, cell_text in enumerate(cells):
                if start_col + j >= self.table.columnCount():
                    break
                item = self.table.item(start_row + i, start_col + j)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    self.table.setItem(start_row + i, start_col + j, item)
                item.setText(cell_text)
                
                # Update underlying data
                if isinstance(self.data, pd.DataFrame):
                    try:
                        self.data.iloc[start_row + i, start_col + j] = cell_text
                    except:
                        pass
                else:
                    try:
                        self.data[start_row + i, start_col + j] = float(cell_text) if cell_text else np.nan
                    except:
                        self.data[start_row + i, start_col + j] = cell_text
        self.status_bar.setText(f"Pasted data to {len(lines)} rows")

    def select_column_as_voltage(self, position):
        """Select the column at position as voltage column"""
        item = self.table.itemAt(position)
        if item is not None:
            col = item.column()
            self.on_header_clicked(col)

    def select_column_as_current(self, position):
        """Select the column at position as current column (green highlight)"""
        item = self.table.itemAt(position)
        if item is not None:
            col = item.column()
            # Deselect previous current_col highlight
            def set_column_color(col_index, color=None):
                if col_index is None: return
                if color is None: color = QtGui.QColor("white")
                for i in range(self.table.rowCount()):
                    self.table.item(i, col_index).setBackground(color)
                self.table.horizontalHeaderItem(col_index).setBackground(color)
            if self.current_col != col:
                set_column_color(self.current_col)
            if self.voltage_col == col:
                self.voltage_col = None
                self.voltage_label.setText("Voltage Column: None")
            self.current_col = col
            self.current_label.setText(f"Current Column: {self.table.horizontalHeaderItem(col).text()}")
            # Redraw highlights
            set_column_color(self.voltage_col, QtGui.QColor("lightblue"))
            set_column_color(self.current_col, QtGui.QColor("lightgreen"))
            has_selection = self.voltage_col is not None and self.current_col is not None
            self.next_button.setEnabled(has_selection)
            self.ok_button.setEnabled(has_selection)

    def on_cell_changed(self, item):
        """Handle cell content changes"""
        if not self.edit_mode:
            return
            
        row = item.row()
        col = item.column()
        new_value = item.text()
        
        # Update the underlying data
        if isinstance(self.data, pd.DataFrame):
            try:
                # Try to convert to numeric if possible
                if new_value.strip() == "":
                    self.data.iloc[row, col] = np.nan
                else:
                    try:
                        numeric_value = float(new_value)
                        self.data.iloc[row, col] = numeric_value
                    except ValueError:
                        self.data.iloc[row, col] = new_value
            except Exception as e:
                print(f"Error updating DataFrame: {e}")
        else:
            try:
                if new_value.strip() == "":
                    self.data[row, col] = np.nan
                else:
                    self.data[row, col] = float(new_value)
            except ValueError:
                # If conversion fails, store as string (for numpy array this might cause issues)
                print(f"Could not convert '{new_value}' to numeric value")
                
        # Refresh the table display
        self.refresh_table_display()

    def refresh_table_display(self):
        """Refresh the table display to reflect data changes"""
        # This method can be called to update the visual representation
        # when underlying data is modified
        pass

    def setup_keyboard_shortcuts(self):
        """Set up keyboard shortcuts for the table"""
        # Create shortcuts
        edit_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("F2"), self.table)
        edit_shortcut.activated.connect(self.edit_current_cell)
        
        delete_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self.table)
        delete_shortcut.activated.connect(self.delete_current_cell)
        
        copy_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+C"), self.table)
        copy_shortcut.activated.connect(self.copy_selection)
        
        paste_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+V"), self.table)
        paste_shortcut.activated.connect(self.paste_selection)
        
        # Enter key to edit cell
        enter_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Return"), self.table)
        enter_shortcut.activated.connect(self.edit_current_cell)
        
        # Escape key to exit full screen
        escape_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Escape"), self)
        escape_shortcut.activated.connect(self.handle_escape_key)

    def edit_current_cell(self):
        """Edit the currently selected cell"""
        current_item = self.table.currentItem()
        if current_item is not None and self.edit_mode:
            self.table.editItem(current_item)

    def delete_current_cell(self):
        """Delete the currently selected cell"""
        current_item = self.table.currentItem()
        if current_item is not None:
            self.delete_cell_at_position(self.table.visualItemRect(current_item).center())

    def handle_escape_key(self):
        """Handle escape key press"""
        if self.isFullScreen():
            self.toggle_fullscreen()
        elif self.edit_mode:
            self.toggle_edit_mode()

    def save_and_continue(self):
        """Save current selection and prepare for next one"""
        current_selection = self.get_current_selection()
        if current_selection:
            self.selections.append(current_selection)
            self.current_selection_number += 1
            
            # Reset selection state
            self.voltage_col = None
            self.current_col = None
            self.voltage_label.setText("Voltage Column: None")
            self.current_label.setText("Current Column: None")
            self.start_row_spin.setValue(1)
            self.end_row_spin.setValue(len(self.data))
            self.name_edit.setText(f"{os.path.splitext(self.filename)[0]}_{self.current_selection_number}")
            
            # Reset column headers
            for i in range(self.table.columnCount()):
                self.table.horizontalHeaderItem(i).setBackground(QtGui.QColor("white"))
            
            # Update selection counter
            self.selection_counter.setText(f"Plot Selection #{self.current_selection_number}")
            
            # Disable next/ok buttons until new selection is made
            self.next_button.setEnabled(False)
            self.ok_button.setEnabled(False)
    
    def get_current_selection(self):
        """Get the current selection details"""
        if self.voltage_col is not None and self.current_col is not None:
            return {
                'name': self.name_edit.text(),
                'voltage_col': self.table.horizontalHeaderItem(self.voltage_col).text(),
                'current_col': self.table.horizontalHeaderItem(self.current_col).text(),
                'start_row': self.start_row_spin.value(),
                'end_row': self.end_row_spin.value()
            }
        return None
    
    def get_all_selections(self):
        """Get all selections including the current one"""
        current = self.get_current_selection()
        if current:
            return self.selections + [current]
        return self.selections
    
    def validate_row_range(self):
        """Ensure start row is not greater than end row"""
        if self.start_row_spin.value() > self.end_row_spin.value():
            if self.sender() == self.start_row_spin:
                self.start_row_spin.setValue(self.end_row_spin.value())
            else:
                self.end_row_spin.setValue(self.start_row_spin.value())
    
    def use_selected_range(self):
        """Update row range based on current table selection"""
        selected_ranges = self.table.selectedRanges()
        if selected_ranges:
            min_row = min(r.topRow() for r in selected_ranges) + 1  # Convert to 1-based
            max_row = max(r.bottomRow() for r in selected_ranges) + 1  # Convert to 1-based
            self.start_row_spin.setValue(min_row)
            self.end_row_spin.setValue(max_row)
    
    def setup_table(self, data):
        # Set up the table with the data
        if isinstance(data, pd.DataFrame):
            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(data.columns))
            self.table.setHorizontalHeaderLabels(data.columns)
            
            # Fill the table
            for i in range(len(data)):
                for j in range(len(data.columns)):
                    value = data.iloc[i, j]
                    # Handle NaN values
                    if pd.isna(value):
                        item = QtWidgets.QTableWidgetItem("")
                    else:
                        item = QtWidgets.QTableWidgetItem(str(value))
                    self.table.setItem(i, j, item)
            
            # Add row numbers
            self.table.setVerticalHeaderLabels([str(i+1) for i in range(len(data))])
        else:
            # Handle numpy array
            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(data[0]) if len(data) > 0 else 0)
            self.table.setHorizontalHeaderLabels([f"Col {i+1}" for i in range(len(data[0]) if len(data) > 0 else 0)])
            
            # Fill the table
            for i in range(len(data)):
                for j in range(len(data[i])):
                    value = data[i, j]
                    # Handle NaN values
                    if np.isnan(value):
                        item = QtWidgets.QTableWidgetItem("")
                    else:
                        item = QtWidgets.QTableWidgetItem(str(value))
                    self.table.setItem(i, j, item)
            
            # Add row numbers
            self.table.setVerticalHeaderLabels([str(i+1) for i in range(len(data))])
        
        # Resize columns to content
        self.table.resizeColumnsToContents()
        
        # Set up keyboard shortcuts
        self.table.setFocusPolicy(QtCore.Qt.StrongFocus)
        
        # Add keyboard shortcuts
        self.setup_keyboard_shortcuts()
        
    def on_header_clicked(self, column):
        modifiers = QtWidgets.QApplication.keyboardModifiers()

        def set_column_color(col_index, color=None):
            if col_index is None: return
            if color is None: color = QtGui.QColor("white")
            for i in range(self.table.rowCount()):
                self.table.item(i, col_index).setBackground(color)
            self.table.horizontalHeaderItem(col_index).setBackground(color)

        if modifiers == QtCore.Qt.ControlModifier:
            # Current (green) selection
            if self.current_col != column:
                set_column_color(self.current_col)
            if self.voltage_col == column:
                self.voltage_col = None
                self.voltage_label.setText("Voltage Column: None")
            self.current_col = column
            self.current_label.setText(f"Current Column: {self.table.horizontalHeaderItem(column).text()}")
        else:
            # Voltage (blue) selection
            if self.voltage_col != column:
                set_column_color(self.voltage_col)
            if self.current_col == column:
                self.current_col = None
                self.current_label.setText("Current Column: None")
            self.voltage_col = column
            self.voltage_label.setText(f"Voltage Column: {self.table.horizontalHeaderItem(column).text()}")

        # Redraw highlights
        set_column_color(self.voltage_col, QtGui.QColor("lightblue"))
        set_column_color(self.current_col, QtGui.QColor("lightgreen"))

        has_selection = self.voltage_col is not None and self.current_col is not None
        self.next_button.setEnabled(has_selection)
        self.ok_button.setEnabled(has_selection)

    def _title_mouse_press(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
    def _title_mouse_move(self, event):
        if self._drag_pos and event.buttons() & QtCore.Qt.LeftButton:
            self.move(event.globalPos() - self._drag_pos)
            event.accept()

class JVCalculator:
    """Class to handle all J-V curve calculations separate from UI logic"""
    
    @staticmethod
    def calculate_parameters(V, J):
        """Calculate Voc, Jsc, FF, and Efficiency from V-J data
        
        Args:
            V (np.array): Voltage data
            J (np.array): Current density data
            
        Returns:
            tuple: (Voc, Jsc, FF, Eff)
        """
        try:
            # Round input data to 4 decimals
            V = np.round(V, 4)
            J = np.round(J, 4)
            
            # Find all zero crossings for Voc
            zero_cross = np.where(np.diff(np.sign(J)))[0]
            
            if len(zero_cross) > 0:
                # Calculate Voc at each zero crossing
                voc_values = []
                for idx in zero_cross:
                    v1, v2 = V[idx], V[idx + 1]
                    j1, j2 = J[idx], J[idx + 1]
                    # Linear interpolation to find exact zero crossing
                    voc = v1 - j1 * (v2 - v1) / (j2 - j1)
                    voc_values.append(round(voc, 4))
                
                # Take the maximum absolute Voc value
                Voc = voc_values[np.argmax(np.abs(voc_values))]
            else:
                Voc = 0
            
            # Calculate Jsc by interpolating at V=0
            if V[0] <= 0 <= V[-1]:  # Only if data crosses V=0
                Jsc = -np.round(np.interp(0, V, J), 4)  # Negative sign added to correct the sign
            else:
                Jsc = 0
            
            # Calculate power at each point
            P = np.round(V * J, 4)
            # Find maximum power point (could be positive or negative)
            Pmax = P[np.argmax(np.abs(P))]
            
            # Calculate fill factor considering signs
            if Voc != 0 and Jsc != 0:
                FF = round(abs(Pmax / (Voc * Jsc)) * 100, 4)  # Multiply by 100 to convert to percentage
            else:
                FF = 0
            
            # Calculate efficiency (use absolute value of Pmax)
            Eff = round(abs(Pmax / 100) * 100, 4)
            
            # Return values with 4 decimal precision
            return round(Voc, 4), round(Jsc, 4), round(FF, 4), round(Eff, 4)
            
        except Exception as e:
            print(f"Error in calculations: {str(e)}")
            return 0, 0, 0, 0
    
    @staticmethod
    def calculate_statistics(data_list):
        """Calculate statistics for all parameters across datasets
        
        Args:
            data_list (list): List of tuples containing (path, data)
            
        Returns:
            dict: Dictionary containing statistics for each parameter
        """
        voc_values = []
        jsc_values = []
        ff_values = []
        eff_values = []
        
        for _, data in data_list:
            V, J = np.round(data[:, 0], 4), np.round(data[:, 1], 4)
            voc, jsc, ff, eff = JVCalculator.calculate_parameters(V, J)
            if voc != 0:  # Only include non-zero values
                voc_values.append(voc)
            if jsc != 0:
                jsc_values.append(jsc)
            if ff != 0:
                ff_values.append(ff)
            if eff != 0:
                eff_values.append(eff)
        
        stats = {
            'Voc': {
                'max': round(max(voc_values), 4) if voc_values else 0,
                'min': round(min(voc_values), 4) if voc_values else 0,
                'avg': round(sum(voc_values) / len(voc_values), 4) if voc_values else 0
            },
            'Jsc': {
                'max': round(max(jsc_values), 4) if jsc_values else 0,
                'min': round(min(jsc_values), 4) if jsc_values else 0,
                'avg': round(sum(jsc_values) / len(jsc_values), 4) if jsc_values else 0
            },
            'FF': {
                'max': round(max(ff_values), 4) if ff_values else 0,
                'min': round(min(ff_values), 4) if ff_values else 0,
                'avg': round(sum(ff_values) / len(ff_values), 4) if ff_values else 0
            },
            'Eff': {
                'max': round(max(eff_values), 4) if eff_values else 0,
                'min': round(min(eff_values), 4) if eff_values else 0,
                'avg': round(sum(eff_values) / len(eff_values), 4) if eff_values else 0
            }
        }
        
        return stats
    
    @staticmethod
    def process_data(V, I, area, convert_to_ma=False):
        """Process voltage and current data
        
        Args:
            V (np.array): Voltage data
            I (np.array): Current data
            area (float): Device area in cm²
            convert_to_ma (bool): Whether to convert current from A to mA
            
        Returns:
            tuple: (processed_V, processed_J)
        """
        # Convert current to mA if needed
        if convert_to_ma:
            I = I * 1000
            
        # Convert current to current density
        J = I / area
        
        # Filter out NaN and Inf values
        valid_mask = np.isfinite(V) & np.isfinite(J)
        V = V[valid_mask]
        J = J[valid_mask]
        
        return V, J

class JVPlotter(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Solar Cell J–V Plotter")
        self.setMinimumSize(1200, 900)
        self.data_list = []
        self.file_widgets = []
        # --- Simple, minimal style ---
        self.setStyleSheet('''
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f6f7fa, stop:1 #dadbde);
                border: 1px solid #bfc1c2;
                border-radius: 4px;
                padding: 4px 10px;
                font-size: 9.5pt;
                min-width: 60px;
                min-height: 22px;
                font-weight: normal;
                color: #222;
                box-shadow: 1px 1px 2px #eee;
            }
            QPushButton:hover {
                background: #e6e6e6;
            }
            QPushButton:pressed {
                background: #dadbde;
            }
        ''')
        # --- end stylesheet ---
        # Initialize background colors
        self.bgColor = QtGui.QColor(255, 255, 255)  # Default white
        self.bgColor2 = QtGui.QColor("#2c3e50")  # Default gradient end
        self.use_gradient = False
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QVBoxLayout(central)
        # Create tab widget
        self.tab_widget = QtWidgets.QTabWidget()
        main_layout.addWidget(self.tab_widget)
        # Create JV Plot tab
        self.jv_tab = QtWidgets.QWidget()
        self.setup_jv_tab()
        self.tab_widget.addTab(self.jv_tab, "J-V Plot")
        # Create XRD Plot tab
        self.xrd_tab = XRDPlotter()
        self.tab_widget.addTab(self.xrd_tab, "XRD Plot")
        # Create EQE Plot tab
        self.eqe_tab = EQEPlotter(self)
        self.tab_widget.addTab(self.eqe_tab, "EQE Plot")
        self.status = self.statusBar()
        self.status.showMessage("Ready")

    def setup_jv_tab(self):
        layout = QtWidgets.QVBoxLayout(self.jv_tab)
        # Add tab-style parameter buttons at the top
        param_button_layout = QtWidgets.QHBoxLayout()
        param_button_layout.setSpacing(0)
        # Create a button group for exclusive selection
        self.param_button_group = QtWidgets.QButtonGroup(self)
        self.param_button_group.setExclusive(True)
        # Create tab-style buttons
        self.param_buttons = {}
        for param in ['JV Plot', 'Voc Stats', 'Jsc Stats', 'FF Stats', 'PCE Stats']:
            btn = QtWidgets.QPushButton(param)
            btn.setCheckable(True)
            btn.setProperty("tabButton", True)
            btn.setStyleSheet("")  # Use global QSS
            param_button_layout.addWidget(btn)
            self.param_button_group.addButton(btn)
            self.param_buttons[param] = btn
        param_button_layout.addStretch()
        layout.addLayout(param_button_layout)
        # Create content widget with border to match tab style
        content_widget = QtWidgets.QWidget()
        content_widget.setStyleSheet("")  # Remove global QSS for data preview/input area
        content_layout = QtWidgets.QHBoxLayout(content_widget)
        layout.addWidget(content_widget)
        # Create a container widget for the controls section
        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        # File management section at the top
        file_section = QtWidgets.QGroupBox("File Management")
        file_section.setStyleSheet("""
            QGroupBox { 
                font-size: 11pt; 
                font-weight: bold; 
                padding-top: 10px; 
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                          stop:0 #f6f7fa, stop:1 #dadbde);
                border: 1px solid #c2c2c2;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 11pt;
                font-weight: bold;
                color: #2c3e50;
                min-width: 100px;
                margin: 5px;
            }
        """)
        file_layout = QtWidgets.QVBoxLayout(file_section)
        
        buttons_layout = QtWidgets.QHBoxLayout()
        buttons_layout.setSpacing(10)
        buttons_layout.setContentsMargins(5, 5, 5, 5)
        
        self.addFileBtn = QtWidgets.QPushButton("Add File…")
        self.addFileBtn.setIcon(QtGui.QIcon.fromTheme("document-open"))
        self.clearFilesBtn = QtWidgets.QPushButton("Clear Files")
        self.clearFilesBtn.setIcon(QtGui.QIcon.fromTheme("edit-clear"))
        buttons_layout.addWidget(self.addFileBtn)
        buttons_layout.addWidget(self.clearFilesBtn)
        file_layout.addLayout(buttons_layout)
        
        self.fileList = QtWidgets.QListWidget()
        self.fileList.setMaximumHeight(150)
        file_layout.addWidget(self.fileList)
        
        container_layout.addWidget(file_section)
        
        # Results section
        results_section = QtWidgets.QGroupBox("Calculation Results")
        results_section.setStyleSheet("QGroupBox { font-size: 11pt; font-weight: bold; padding-top: 10px; }")
        results_layout = QtWidgets.QVBoxLayout(results_section)
        
        self.resultTable = QtWidgets.QTableWidget(0, 5)
        self.resultTable.setHorizontalHeaderLabels(["Label", "Voc (V)", "Jsc (mA/cm²)", "FF", "Eff (%)"])
        self.resultTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.resultTable.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.resultTable.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.resultTable.setAlternatingRowColors(True)
        results_layout.addWidget(self.resultTable)
        
        container_layout.addWidget(results_section)

        # Create horizontal layout for Tools, Style, and Plot sections
        controls_layout = QtWidgets.QHBoxLayout()
        
        # Tools Section
        tools_section = CollapsibleSection("Tools")
        tools_form = QtWidgets.QFormLayout()
        tools_form.setVerticalSpacing(10)
        
        self.areaEdit = QtWidgets.QLineEdit("0.1")
        tools_form.addRow("Device Area (cm²):", self.areaEdit)
        
        self.unitCombo = QtWidgets.QComboBox()
        self.unitCombo.addItems(["mA", "A"])
        tools_form.addRow("Current Unit:", self.unitCombo)
        
        self.xLabelEdit = QtWidgets.QLineEdit("Voltage (V)")
        tools_form.addRow("X-axis Label:", self.xLabelEdit)
        
        self.yLabelEdit = QtWidgets.QLineEdit(r"Current Density (mA/cm$^2$)")
        tools_form.addRow("Y-axis Label:", self.yLabelEdit)
        
        self.titleEdit = QtWidgets.QLineEdit("J–V Comparison")
        tools_form.addRow("Plot Title:", self.titleEdit)
        
        tools_section.addLayout(tools_form)
        controls_layout.addWidget(tools_section)
        
        # Style Section
        style_section = CollapsibleSection("Style")
        style_form = QtWidgets.QFormLayout()
        style_form.setVerticalSpacing(10)
        
        self.styleCombo = QtWidgets.QComboBox()
        self.styleCombo.addItems(["Line", "Scatter", "Line + Marker", "Bubble"])
        style_form.addRow("Plot Style:", self.styleCombo)
        
        self.lineStyleCombo = QtWidgets.QComboBox()
        self.lineStyleCombo.addItems(["Solid", "Dashed", "Dotted", "Dash-dot", "Step"])
        style_form.addRow("Line Style:", self.lineStyleCombo)
        
        self.boldChk = QtWidgets.QCheckBox()
        style_form.addRow("Bold Labels:", self.boldChk)
        
        self.majorChk = QtWidgets.QCheckBox()
        self.majorChk.setChecked(True)
        style_form.addRow("Show Major Ticks:", self.majorChk)
        
        self.minorChk = QtWidgets.QCheckBox()
        style_form.addRow("Show Minor Ticks:", self.minorChk)
        
        self.boldScaleChk = QtWidgets.QCheckBox()
        style_form.addRow("Bold Scale:", self.boldScaleChk)
        
        self.bgSettingsBtn = QtWidgets.QPushButton("Background Settings...")
        style_form.addRow("Background:", self.bgSettingsBtn)
        
        color_layout = QtWidgets.QHBoxLayout()
        self.colorBtn = QtWidgets.QPushButton()
        self.colorBtn.setFixedSize(30, 30)
        self.color = QtGui.QColor(31, 119, 180)
        self._update_color_button()
        color_layout.addWidget(self.colorBtn)
        color_layout.addStretch()
        style_form.addRow("Line Color:", color_layout)
        
        self.fontCombo = QtWidgets.QComboBox()
        self.fontCombo.addItems([
            "Default", "Arial", "Helvetica", "Times New Roman",
            "Computer Modern", "Calibri", "Georgia"
        ])
        style_form.addRow("Font:", self.fontCombo)
        
        self.fontSizeCombo = QtWidgets.QComboBox()
        self.fontSizeCombo.addItems(['8', '9', '10', '11', '12', '14', '16', '18', '20'])
        self.fontSizeCombo.setCurrentText('12')
        style_form.addRow("Size:", self.fontSizeCombo)
        
        style_section.addLayout(style_form)
        controls_layout.addWidget(style_section)
        
        # Plot Settings Section
        plot_section = CollapsibleSection("Plot")
        plot_form = QtWidgets.QFormLayout()
        plot_form.setVerticalSpacing(10)
        
        self.dpiSpin = QtWidgets.QSpinBox()
        self.dpiSpin.setRange(50, 2400)
        self.dpiSpin.setValue(1200)
        plot_form.addRow("DPI:", self.dpiSpin)
        
        self.xmaxSpin = QtWidgets.QDoubleSpinBox()
        self.xmaxSpin.setRange(0, 1e4)
        self.xmaxSpin.setDecimals(3)
        plot_form.addRow("X Max (V):", self.xmaxSpin)
        
        self.ymaxSpin = QtWidgets.QDoubleSpinBox()
        self.ymaxSpin.setRange(0, 1e6)
        plot_form.addRow("Y Max:", self.ymaxSpin)
        
        self.xStepSpin = QtWidgets.QDoubleSpinBox()
        self.xStepSpin.setRange(0.001, 1e3)
        self.xStepSpin.setDecimals(3)
        self.xStepSpin.setSingleStep(0.1)
        self.xStepSpin.setValue(0.2)
        plot_form.addRow("X Tick Step:", self.xStepSpin)
        
        self.yStepSpin = QtWidgets.QDoubleSpinBox()
        self.yStepSpin.setRange(0.1, 1e6)
        self.yStepSpin.setDecimals(3)
        self.yStepSpin.setSingleStep(1)
        self.yStepSpin.setValue(5)
        plot_form.addRow("Y Tick Step:", self.yStepSpin)
        
        self.scaleCombo = QtWidgets.QComboBox()
        self.scaleCombo.addItems(["Linear", "Log X", "Log Y", "Log XY"])
        plot_form.addRow("Graph Type:", self.scaleCombo)
        
        self.bubbleSpin = QtWidgets.QSpinBox()
        self.bubbleSpin.setRange(1, 1000)
        self.bubbleSpin.setValue(50)
        plot_form.addRow("Bubble Size:", self.bubbleSpin)
        
        # Add legend settings
        legend_group = QtWidgets.QGroupBox("Legend Settings")
        legend_layout = QtWidgets.QVBoxLayout(legend_group)
        
        # Legend background transparency slider
        self.legendBgAlphaSlider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.legendBgAlphaSlider.setRange(0, 100)
        self.legendBgAlphaSlider.setValue(100)
        legend_bg_alpha_layout = QtWidgets.QHBoxLayout()
        legend_bg_alpha_layout.addWidget(QtWidgets.QLabel("Background Opacity:"))
        legend_bg_alpha_layout.addWidget(self.legendBgAlphaSlider)
        self.legendBgAlphaValue = QtWidgets.QLabel("100%")
        legend_bg_alpha_layout.addWidget(self.legendBgAlphaValue)
        legend_layout.addLayout(legend_bg_alpha_layout)
        
        # Add legend position controls
        position_layout = QtWidgets.QHBoxLayout()
        self.legendPosCombo = QtWidgets.QComboBox()
        self.legendPosCombo.addItems(["best", "upper right", "upper left", "lower left", 
                                     "lower right", "right", "center left", "center right", 
                                     "lower center", "upper center", "center"])
        position_layout.addWidget(QtWidgets.QLabel("Position:"))
        position_layout.addWidget(self.legendPosCombo)
        legend_layout.addLayout(position_layout)
        
        plot_form.addRow(legend_group)
        
        # Add annotation button
        self.annotateBtn = QtWidgets.QPushButton("Add Annotation")
        plot_form.addRow("Annotation:", self.annotateBtn)
        
        plot_section.addLayout(plot_form)
        controls_layout.addWidget(plot_section)
        
        # Add the horizontal controls layout to the container
        container_layout.addLayout(controls_layout)
        
        # Add Parameter Analysis section
        param_section = CollapsibleSection("Parameter Analysis")
        param_form = QtWidgets.QFormLayout()
        param_form.setVerticalSpacing(10)
        # Add parameter analysis controls
        param_stats_group = QtWidgets.QGroupBox("Statistics Display")
        param_stats_layout = QtWidgets.QVBoxLayout(param_stats_group)
        self.statsLabel = QtWidgets.QLabel()
        self.statsLabel.setStyleSheet("")
        self.statsLabel.setWordWrap(True)
        param_stats_layout.addWidget(self.statsLabel)
        # Add legend visibility checkbox
        self.showParamLegendChk = QtWidgets.QCheckBox("Show Parameter Legend")
        self.showParamLegendChk.setChecked(True)
        param_stats_layout.addWidget(self.showParamLegendChk)
        param_form.addRow(param_stats_group)
        param_section.addLayout(param_form)
        container_layout.addWidget(param_section)
        
        # Add action buttons at the bottom
        button_layout = QtWidgets.QHBoxLayout()
        self.plotBtn = QtWidgets.QPushButton("Generate Plot")
        self.saveBtn = QtWidgets.QPushButton("Save Plot…")
        self.exportTableBtn = QtWidgets.QPushButton("Export Table…")
        
        for btn in [self.plotBtn, self.saveBtn, self.exportTableBtn]:
            btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                              stop:0 #f6f7fa, stop:1 #dadbde);
                    border: 1px solid #c2c2c2;
                    border-radius: 6px;
                    padding: 8px 16px;
                    font-size: 11pt;
                    font-weight: bold;
                    color: #2c3e50;
                    min-width: 120px;
                    margin: 5px;
                }
            """)
            button_layout.addWidget(btn)
        
        container_layout.addLayout(button_layout)
        container_layout.addStretch()
        
        # Create splitter and add container and plot
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(container)
        
        # Plot area
        plot_frame = QtWidgets.QFrame()
        plot_frame.setStyleSheet("background:#fff; border:1px solid #ccc; padding:4px;")
        pl = QtWidgets.QVBoxLayout(plot_frame)
        plt.rcParams.update({
            'font.family': 'serif',
            'font.size': 12,
            'axes.linewidth': 1,
            'xtick.direction': 'in',
            'ytick.direction': 'in'
        })
        self.fig, self.ax = plt.subplots(figsize=(6, 6), dpi=200)
        self.canvas = FigureCanvasQTAgg(self.fig)
        self.canvas.setMinimumSize(600, 600)  # Fix minimum plot area size
        pl.addWidget(self.canvas)
        splitter.addWidget(plot_frame)
        
        content_layout.addWidget(splitter)
        
        # Connect all signals
        self.addFileBtn.clicked.connect(self.add_files)
        self.clearFilesBtn.clicked.connect(self.clear_files)
        self.colorBtn.clicked.connect(self.pick_color)
        self.plotBtn.clicked.connect(self.do_compare_plot)
        self.saveBtn.clicked.connect(self.save_plot)
        self.annotateBtn.clicked.connect(self.add_annotation)
        self.exportTableBtn.clicked.connect(self.export_table)
        self.bgSettingsBtn.clicked.connect(self.show_background_settings)
        self.unitCombo.currentTextChanged.connect(self.do_compare_plot)
        self.fontCombo.currentTextChanged.connect(self.update_font_style)
        self.fontSizeCombo.currentTextChanged.connect(self.update_font_style)
        self.legendBgAlphaSlider.valueChanged.connect(self.update_legend_transparency)
        self.legendPosCombo.currentTextChanged.connect(self.update_legend_position)
        self.showParamLegendChk.stateChanged.connect(self.on_param_legend_changed)
        
        # Connect the button group signal
        self.param_button_group.buttonClicked.connect(self.on_param_button_clicked)
        
        # Set JV Plot as default selected
        self.param_buttons['JV Plot'].setChecked(True)

    def on_param_button_clicked(self, button):
        """Handle parameter button clicks"""
        param = button.text()
        if param == 'JV Plot':
            self.do_compare_plot()
        elif param == 'Voc Stats':
            self.plot_parameter('Voc')
        elif param == 'Jsc Stats':
            self.plot_parameter('Jsc')
        elif param == 'FF Stats':
            self.plot_parameter('FF')
        elif param == 'PCE Stats':
            self.plot_parameter('Eff')

    def show_background_settings(self):
        dialog = BackgroundSettingsDialog(self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.use_gradient = dialog.bgTypeCombo.currentText() == "Gradient"
            if self.use_gradient:
                self.bgColor = dialog.startColor
                self.bgColor2 = dialog.endColor
            else:
                self.bgColor = dialog.bgColor
            self.do_compare_plot()

    def _update_plot_background(self):
        if self.use_gradient:
            # Create a smooth gradient background for the entire figure
            points = 100
            gradient = np.linspace(0, 1, points).reshape(-1, 1)
            
            # Get RGB values for start and end colors
            start_color = np.array([self.bgColor.red()/255.0, self.bgColor.green()/255.0, self.bgColor.blue()/255.0])
            end_color = np.array([self.bgColor2.red()/255.0, self.bgColor2.green()/255.0, self.bgColor2.blue()/255.0])
            
            # Create smooth color transition
            colors = start_color + gradient * (end_color - start_color)
            
            # Create the gradient array
            gradient_array = np.zeros((points, 1, 3))
            for i in range(points):
                gradient_array[i, 0] = colors[i]
            
            # Create a new axes for the background that covers the entire figure
            if not hasattr(self, 'bg_axes'):
                self.bg_axes = self.fig.add_axes([0, 0, 1, 1])
                self.bg_axes.set_zorder(-1)  # Put it behind everything
            else:
                self.bg_axes.clear()
            
            # Plot the gradient in the background axes
            self.bg_axes.imshow(gradient_array, aspect='auto', extent=[0, 1, 0, 1],
                              interpolation='gaussian')
            self.bg_axes.axis('off')  # Hide axes
            
            # Make the main plot area transparent to show gradient
            self.ax.set_facecolor('none')
            self.ax.patch.set_alpha(0.0)  # Make plot background transparent
            
            # Set figure background to transparent to show gradient
            self.fig.patch.set_facecolor('none')
            
        else:
            # For solid color, just set both backgrounds
            if hasattr(self, 'bg_axes'):
                self.bg_axes.remove()
                delattr(self, 'bg_axes')
            self.fig.patch.set_facecolor(self.bgColor.name())
            self.ax.set_facecolor(self.bgColor.name())
            
        # Ensure proper plot positioning
        self.ax.set_position([0.15, 0.1, 0.8, 0.8])
        
        # Make sure the main plot frame is visible
        for spine in self.ax.spines.values():
            spine.set_visible(True)
            spine.set_linewidth(1.0)

    def add_files(self):
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Select Data Files", "", "All Supported (*.txt *.dat *.csv *.xls *.xlsx);;Text/Data Files (*.txt *.dat);;CSV Files (*.csv);;Excel Files (*.xls *.xlsx)"
        )
        if paths:
            added = 0
            for path in paths:
                try:
                    ext = os.path.splitext(path)[1].lower()
                    # Determine default device name
                    default_device = f"Device {len(self.data_list) + 1}"
                    if ext in ['.csv']:
                        # Try different delimiters
                        try:
                            df = pd.read_csv(path, delimiter=',')
                        except:
                            try:
                                df = pd.read_csv(path, delimiter=';')
                            except:
                                df = pd.read_csv(path, delimiter='\t')
                    elif ext in ['.xls', '.xlsx']:
                        df = pd.read_excel(path)
                    else:
                        with open(path) as f:
                            first_line = f.readline()
                            has_header = not all(c.isdigit() or c.isspace() or c.isspace() or c == '.' or c == '-' for c in first_line)

                        if has_header:
                            df = pd.read_csv(path, delim_whitespace=True)
                        else:
                            data = np.loadtxt(path)
                            self.data_list.append((path, data, default_device))
                            # Create custom widget item with close button and device name
                            item = QtWidgets.QListWidgetItem()
                            file_widget = FileItemWidget(os.path.basename(path), default_device)
                            item.setSizeHint(file_widget.sizeHint())
                            self.fileList.addItem(item)
                            self.fileList.setItemWidget(item, file_widget)
                            # Connect close button with explicit reference to self
                            def create_remove_callback(p, i):
                                def callback():
                                    self.remove_file(p, i)
                                return callback
                            file_widget.closeButton.clicked.connect(create_remove_callback(path, item))
                            self.file_widgets.append(file_widget)
                            added += 1
                            continue
                    # Show data selection dialog
                    dialog = DataSelectionDialog(df, path, self)
                    if dialog.exec_() == QtWidgets.QDialog.Accepted:
                        selections = dialog.get_all_selections()
                        for selection in selections:
                            try:
                                v_col = selection['voltage_col']
                                i_col = selection['current_col']
                                start_row = selection['start_row'] - 1  # Convert to 0-based index
                                end_row = selection['end_row']  # Convert to 0-based index
                                plot_name = selection['name']
                                # Skip rows before the data start row and after end row
                                V = df[v_col].iloc[start_row:end_row].astype(float).to_numpy()
                                I = df[i_col].iloc[start_row:end_row].astype(float).to_numpy()
                                if self.unitCombo.currentText() == "A":
                                    I *= 1000  # Convert A to mA
                                data = np.column_stack((V, I))
                                # Use default device name for each new dataset
                                self.data_list.append((path, data, default_device))
                                # Create custom widget item with close button and device name
                                item = QtWidgets.QListWidgetItem()
                                file_widget = FileItemWidget(plot_name, default_device)  # Use the custom name and device
                                item.setSizeHint(file_widget.sizeHint())
                                self.fileList.addItem(item)
                                self.fileList.setItemWidget(item, file_widget)
                                def create_remove_callback(p, i):
                                    def callback():
                                        self.remove_file(p, i)
                                    return callback
                                file_widget.closeButton.clicked.connect(create_remove_callback(path, item))
                                self.file_widgets.append(file_widget)
                                added += 1
                            except Exception as e:
                                QtWidgets.QMessageBox.warning(self, "Error", f"Failed to process selection:\n{str(e)}")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "Error", f"Failed {path}:\n{str(e)}")
            self.status.showMessage(f"Added {added} files, total {len(self.data_list)}")

    def remove_file(self, path, item):
        # Find and remove the data from data_list (now includes device_name)
        for i, (p, _, _) in enumerate(self.data_list):
            if p == path:
                self.data_list.pop(i)
                break
        # Remove the item from the list widget
        row = self.fileList.row(item)
        self.fileList.takeItem(row)
        # Remove from file_widgets list
        for widget in self.file_widgets:
            if widget.label.text() == os.path.basename(path):
                self.file_widgets.remove(widget)
                break
        # Update the plot
        self.do_compare_plot()
        self.status.showMessage(f"Removed file {os.path.basename(path)}")

    def clear_files(self):
        self.data_list.clear(); self.fileList.clear(); self.ax.clear(); self.canvas.draw(); self.status.showMessage("Cleared files and plot area")

    def pick_color(self):
        c = QtWidgets.QColorDialog.getColor(self.color, self)
        if c.isValid():
            self.color = c
            self._update_color_button()
            # Update all lines in the plot to use the new color
            if hasattr(self, 'ax'):
                for line in self.ax.lines:
                    line.set_color(self.color.name())
                for collection in self.ax.collections:  # For scatter plots
                    collection.set_color(self.color.name())
                self.canvas.draw()
            # Update the plot with new color
            self.do_compare_plot()

    def _update_color_button(self):
        pix = QtGui.QPixmap(30,30); pix.fill(self.color); self.colorBtn.setIcon(QtGui.QIcon(pix))

    def add_annotation(self):
        x, ok1 = QtWidgets.QInputDialog.getDouble(self, "Annotation X", "X Coordinate:", 0, -1e6, 1e6, 3)
        if not ok1: return
        y, ok2 = QtWidgets.QInputDialog.getDouble(self, "Annotation Y", "Y Coordinate:", 0, -1e6, 1e6, 3)
        if not ok2: return
        text, ok3 = QtWidgets.QInputDialog.getText(self, "Annotation Text", "Label:")
        if not ok3: return
        
        # Add dotted reference lines
        xlims = self.ax.get_xlim()
        ylims = self.ax.get_ylim()
        
        # Add vertical dotted line
        self.ax.axvline(x=x, color='gray', linestyle=':', alpha=0.5, zorder=1)
        # Add horizontal dotted line
        self.ax.axhline(y=y, color='gray', linestyle=':', alpha=0.5, zorder=1)
        
        # Add the annotation with arrow
        self.ax.annotate(text, 
                        xy=(x, y),  # Point to annotate
                        xytext=(x + (xlims[1]-xlims[0])*0.05, y + (ylims[1]-ylims[0])*0.05),  # Text offset
                        arrowprops=dict(
                            arrowstyle="->",
                            connectionstyle="arc3,rad=.2",
                            color='black'
                        ),
                        fontsize=10,
                        color='black',
                        zorder=2
                        )
        
        self.canvas.draw()

    def update_font_style(self):
        font = self.fontCombo.currentText()
        font_size = float(self.fontSizeCombo.currentText())
        
        if font == "Default":
            plt.rcParams.update({
                'font.family': 'serif',
                'font.size': font_size
            })
        else:
            try:
                # Check if font is available
                if font == "Computer Modern":
                    plt.rcParams.update({
                        'font.family': 'serif',
                        'font.serif': ['Computer Modern Roman'],
                        'text.usetex': True,
                        'font.size': font_size
                    })
                else:
                    plt.rcParams.update({
                        'font.family': font,
                        'text.usetex': False,
                        'font.size': font_size
                    })
                
                # Update the plot if it exists
                if hasattr(self, 'ax'):
                    # Store current limits
                    xlim = self.ax.get_xlim()
                    ylim = self.ax.get_ylim()
                    
                    # Update font for all text elements
                    for text in self.ax.texts:
                        if font == "Computer Modern":
                            text.set_fontname("Computer Modern Roman")
                        else:
                            text.set_fontname(font)
                        text.set_fontsize(font_size)
                    
                    # Update axis labels and title with new size
                    self.ax.set_xlabel(self.ax.get_xlabel(), fontfamily=font, fontsize=font_size * 1.2)
                    self.ax.set_ylabel(self.ax.get_ylabel(), fontfamily=font, fontsize=font_size * 1.2)
                    self.ax.set_title(self.ax.get_title(), fontfamily=font, fontsize=font_size * 1.5)
                    
                    # Update tick labels
                    for label in self.ax.get_xticklabels() + self.ax.get_yticklabels():
                        if font == "Computer Modern":
                            label.set_fontname("Computer Modern Roman")
                        else:
                            label.set_fontname(font)
                        label.set_fontsize(font_size)
                    
                    # Update legend
                    if self.ax.get_legend():
                        for text in self.ax.get_legend().get_texts():
                            if font == "Computer Modern":
                                text.set_fontname("Computer Modern Roman")
                            else:
                                text.set_fontname(font)
                            text.set_fontsize(font_size * 0.85)  # Make legend slightly smaller
                    
                    # Restore limits
                    self.ax.set_xlim(xlim)
                    self.ax.set_ylim(ylim)
                    
                    # Adjust layout to prevent text overlap
                    self.fig.tight_layout()
                    
                    # Redraw
                    self.canvas.draw()
                    
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Font Error", 
                    f"Could not set font to {font}. Error: {str(e)}\nReverting to default font.")
                self.fontCombo.setCurrentText("Default")
                plt.rcParams.update({'font.family': 'serif', 'font.size': font_size})

    def do_compare_plot(self):
        if not self.data_list: return
        try: area=float(self.areaEdit.text()); assert area>0
        except: QtWidgets.QMessageBox.warning(self,"Invalid Area","Enter positive area"); return
        self.update_font_style()
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self._update_plot_background()
        self.ax.grid(False)
        scale=self.scaleCombo.currentText()
        self.ax.set_xscale('log' if 'Log X' in scale else 'linear')
        self.ax.set_yscale('log' if 'Log Y' in scale else 'linear')
        cmap=plt.cm.tab10
        style_map={'Solid':'-','Dashed':'--','Dotted':':','Dash-dot':'-.','Step':'steps-post'}
        ls = style_map.get(self.lineStyleCombo.currentText(),' -')
        result_text = []
        x_min, x_max = 1e10, -1e10
        y_min, y_max = 1e10, -1e10
        # --- Group datasets by device name ---
        device_to_indices = {}
        for i, (path, data, _) in enumerate(self.data_list):
            widget = self.fileList.itemWidget(self.fileList.item(i))
            device = widget.get_device_name() if widget else "Device 1"
            if device not in device_to_indices:
                device_to_indices[device] = []
            device_to_indices[device].append(i)
        # First pass: collect valid min/max values
        for indices in device_to_indices.values():
            for i in indices:
                _, data, _ = self.data_list[i]
                V, I = data[:, 0], data[:, 1]
                convert_to_ma = self.unitCombo.currentText() == "A"
                V, J = JVCalculator.process_data(V, I, area, convert_to_ma)
                if len(V) > 0:
                    x_min = min(x_min, np.min(V))
                    x_max = max(x_max, np.max(V))
                if len(J) > 0:
                    y_min = min(y_min, np.min(J))
                    y_max = max(y_max, np.max(J))
        if x_min > x_max or not np.isfinite(x_min) or not np.isfinite(x_max):
            x_min, x_max = -1, 1
        if y_min > y_max or not np.isfinite(y_min) or not np.isfinite(y_max):
            y_min, y_max = -1, 1
        # --- Plot by device group ---
        for device_idx, (device, indices) in enumerate(device_to_indices.items()):
            color = cmap(device_idx % 10)
            for i in indices:
                _, data, _ = self.data_list[i]
                V, I = data[:, 0], data[:, 1]
                convert_to_ma = self.unitCombo.currentText() == "A"
                V, J = JVCalculator.process_data(V, I, area, convert_to_ma)
                if len(V) == 0:
                    continue
                plot_type = self.styleCombo.currentText()
                # Only add label for the first dataset in the group (for legend)
                label = device if i == indices[0] else None
                if plot_type == 'Scatter':
                    self.ax.scatter(V, J, color=color, s=35, label=label)
                elif plot_type == 'Bubble':
                    self.ax.scatter(V, J, color=color, s=self.bubbleSpin.value(), alpha=0.75, label=label)
                else:
                    marker = 'o' if 'Marker' in plot_type else ''
                    self.ax.plot(V, J, linestyle=ls, marker=marker, color=color, linewidth=1.4, markersize=7, label=label)
                Voc, Jsc, FF, Eff = JVCalculator.calculate_parameters(V, J)
                result_text.append(f"{device}: Voc={Voc:.2f} V, Jsc={Jsc:.2f} mA/cm², FF={FF:.2f}, Eff={Eff:.2f}%")
        user_x_max = self.xmaxSpin.value()
        user_y_max = self.ymaxSpin.value()
        if user_x_max > 0:
            self.ax.set_xlim(x_min, user_x_max)
        else:
            x_padding = (x_max - x_min) * 0.05
            self.ax.set_xlim(x_min - x_padding, x_max + x_padding)
        if user_y_max > 0:
            self.ax.set_ylim(y_min, user_y_max)
        else:
            y_padding = (y_max - y_min) * 0.05
            self.ax.set_ylim(y_min - y_padding, y_max + y_padding)
        self.fig.subplots_adjust(left=0.15)
        self.ax.yaxis.set_label_coords(-0.08, 0.5)
        for spine in self.ax.spines.values():
            spine.set_linewidth(1.0)
        self.ax.tick_params(axis='both', which='major', width=1.0, length=5, pad=4)
        self.ax.tick_params(axis='both', which='minor', width=0.5, length=3)
        if self.majorChk.isChecked():
            self.ax.xaxis.set_major_locator(MultipleLocator(self.xStepSpin.value()))
            self.ax.yaxis.set_major_locator(MultipleLocator(self.yStepSpin.value()))
        else:
            self.ax.xaxis.set_major_locator(NullLocator())
            self.ax.yaxis.set_major_locator(NullLocator())
        if self.minorChk.isChecked():
            self.ax.minorticks_on()
        else:
            self.ax.minorticks_off()
        self.ax.axhline(y=0, color='none')
        self.ax.axvline(x=0, color='none')
        weight = 'bold' if self.boldScaleChk.isChecked() else 'normal'
        for lbl in self.ax.get_xticklabels() + self.ax.get_yticklabels():
            lbl.set_fontweight(weight)
        lwt = 'bold' if self.boldChk.isChecked() else 'normal'
        self.ax.set_xlabel(self.xLabelEdit.text(), fontsize=14, fontweight=lwt)
        self.ax.set_ylabel(self.yLabelEdit.text(), fontsize=14, fontweight=lwt)
        self.ax.set_title(self.titleEdit.text(), fontsize=18, fontweight=lwt)
        legend_position = self.legendPosCombo.currentText() if hasattr(self, 'legendPosCombo') else 'best'
        legend = self.ax.legend(fontsize=10, loc=legend_position)
        if hasattr(self, 'legendBgAlphaSlider'):
            alpha = self.legendBgAlphaSlider.value() / 100
            legend.get_frame().set_alpha(alpha)
        for sp in ['bottom', 'left', 'top', 'right']:
            self.ax.spines[sp].set_visible(True)
            self.ax.spines[sp].set_linewidth(1.0)
        self.canvas.draw()
        self.status.showMessage("Plot updated")
        # Update result table
        self.resultTable.setRowCount(0)
        for res in result_text:
            label, voc, jsc, ff, eff = res.split(': ')[0], *map(str.strip, res.split(': ')[1].replace('V', '').replace('mA/cm²', '').replace('%', '').split(','))
            row = self.resultTable.rowCount()
            self.resultTable.insertRow(row)
            self.resultTable.setItem(row, 0, QtWidgets.QTableWidgetItem(label))
            self.resultTable.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{float(voc.split('=')[1]):.4f}"))
            self.resultTable.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{float(jsc.split('=')[1]):.4f}"))
            self.resultTable.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{float(ff.split('=')[1]):.4f}"))
            self.resultTable.setItem(row, 4, QtWidgets.QTableWidgetItem(f"{float(eff.split('=')[1]):.4f}"))

    def on_click(self, event):
        """Handle click events on the plot"""
        if event.inaxes:
            QtWidgets.QMessageBox.information(self, "Clicked Point", f"X = {event.xdata:.3f}, Y = {event.ydata:.3f}")

    def export_table(self):
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export Table", "", "Excel Files (*.xls *.xlsx);;Text Files (*.txt);;CSV Files (*.csv)")
        if fname:
            try:
                sep = ',' if fname.endswith('.csv') else '\t'
                with open(fname, 'w') as f:
                    headers = [self.resultTable.horizontalHeaderItem(i).text() for i in range(self.resultTable.columnCount())]
                    f.write(sep.join(headers) + '\n')
                    for row in range(self.resultTable.rowCount()):
                        values = [self.resultTable.item(row, col).text() if self.resultTable.item(row, col) else '' for col in range(self.resultTable.columnCount())]
                        f.write(sep.join(values) + '\n')
                self.status.showMessage(f"Table exported to {fname}")
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Export Error", f"Could not export table:\n{str(e)}")
                self.status.showMessage("Export failed")

    def save_plot(self):
        fname,_=QtWidgets.QFileDialog.getSaveFileName(self,"Save Plot","","PNG Files (*.png);;TIFF Files (*.tiff *.tif);;JPEG Files (*.jpeg *.jpg);;PDF Files (*.pdf)")
        if fname:
            try:
                self.fig.savefig(fname, dpi=self.dpiSpin.value(), bbox_inches='tight')
                self.status.showMessage(f"Saved to {fname}")
            except Exception as e:
                QtWidgets.QMessageBox.warning(self,"Error",f"Save failed:\n{e}")
                self.status.showMessage("Save failed")

    def show_plot_settings(self):
        dialog = PlotEditDialog(self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            # Update stored values
            self.xmaxSpin.setValue(dialog.xmaxSpin.value())
            self.ymaxSpin.setValue(dialog.ymaxSpin.value())
            self.xStepSpin.setValue(dialog.xStepSpin.value())
            self.yStepSpin.setValue(dialog.yStepSpin.value())
            self.scaleCombo.setCurrentText(dialog.scaleCombo.currentText())
            # Update plot
            self.do_compare_plot()

    def update_legend_transparency(self):
        """Update the legend transparency based on slider values"""
        # Update label
        self.legendBgAlphaValue.setText(f"{self.legendBgAlphaSlider.value()}%")
        
        if hasattr(self, 'ax') and self.ax.get_legend() is not None:
            # Get alpha value (convert percentage to 0-1 range)
            bg_alpha = self.legendBgAlphaSlider.value() / 100
            
            # Update legend properties
            legend = self.ax.get_legend()
            legend.set_frame_on(True)  # Ensure frame is visible
            legend.get_frame().set_alpha(bg_alpha)  # Set background transparency
            
            # Update the canvas
            self.canvas.draw()

    def update_legend_position(self, position):
        """Update the legend position"""
        try:
            # Simply redraw the plot with the new legend position
            if hasattr(self, 'ax'):
                self.do_compare_plot()
        except Exception as e:
            print(f"Error updating legend position: {str(e)}")
            # If there's an error, try to recover by setting position to 'best'
            self.legendPosCombo.setCurrentText('best')
            self.do_compare_plot()

        # Ensure the canvas is updated
        if hasattr(self, 'canvas'):
            self.canvas.draw()

    def plot_parameter(self, param_type):
        if not self.data_list:
            QtWidgets.QMessageBox.warning(self, "No Data", "Please load data files first.")
            return
        try:
            area = float(self.areaEdit.text())
            if area <= 0:
                QtWidgets.QMessageBox.warning(self, "Invalid Area", "Enter positive area")
                return
            # --- Group parameter values by device ---
            device_to_values = {}
            device_to_labels = {}
            for i, (path, data, _) in enumerate(self.data_list):
                V, I = data[:, 0], data[:, 1]
                convert_to_ma = self.unitCombo.currentText() == "A"
                V, J = JVCalculator.process_data(V, I, area, convert_to_ma)
                voc, jsc, ff, eff = JVCalculator.calculate_parameters(V, J)
                widget = self.fileList.itemWidget(self.fileList.item(i))
                device = widget.get_device_name() if widget else "Device 1"
                if device not in device_to_values:
                    device_to_values[device] = []
                    device_to_labels[device] = []
                if param_type == 'Voc':
                    device_to_values[device].append(voc)
                elif param_type == 'Jsc':
                    device_to_values[device].append(jsc)
                elif param_type == 'FF':
                    device_to_values[device].append(ff)
                else:  # Eff/PCE
                    device_to_values[device].append(eff)
                # Store dataset label for possible overlay
                label = widget.get_label_text() if widget else os.path.basename(path)
                device_to_labels[device].append(label)
            # --- Prepare data for seaborn ---
            devices = list(device_to_values.keys())
            values = []
            device_names = []
            for dev in devices:
                values.extend(device_to_values[dev])
                device_names.extend([dev] * len(device_to_values[dev]))
            df = pd.DataFrame({
                'Device': device_names,
                'Value': values
            })
            self.fig.clear()
            self.ax = self.fig.add_subplot(111)
            self._update_plot_background()
            cmap = plt.cm.tab10
            palette = {dev: cmap(i % 10) for i, dev in enumerate(devices)}
            # --- Draw violin plot ---
            sns.violinplot(
                x='Device', y='Value', data=df, ax=self.ax,
                inner=None, linewidth=1.5, palette=palette, cut=0, scale='width', width=0.7
            )
            # --- Overlay boxplot ---
            sns.boxplot(
                x='Device', y='Value', data=df, ax=self.ax,
                showcaps=True, boxprops={'facecolor':'none', 'edgecolor':'k', 'linewidth':1.5},
                whiskerprops={'color':'k', 'linewidth':1.2},
                capprops={'color':'k', 'linewidth':1.2},
                medianprops={'color':'k', 'linewidth':2},
                flierprops={'marker':'o', 'markersize':6, 'markerfacecolor':'gray', 'alpha':0.7},
                showfliers=True, width=0.25
            )
            # --- Overlay individual data points ---
            sns.stripplot(
                x='Device', y='Value', data=df, ax=self.ax,
                jitter=0.18, size=7, edgecolor='k', linewidth=0.5, palette=palette, zorder=3
            )
            # Hide x-axis tick labels
            self.ax.set_xticklabels(['' for _ in devices])
            # Add device names as colored text near each group
            y_max = df['Value'].max()
            y_min = df['Value'].min()
            y_range = y_max - y_min
            for i, dev in enumerate(devices):
                color = palette[dev]
                # Place label to the left or right of the group, at the median value
                group_vals = df[df['Device'] == dev]['Value']
                median = np.median(group_vals)
                # Offset for label placement
                x = i
                y = median + 0.15 * y_range if i % 2 == 0 else median - 0.15 * y_range
                self.ax.text(
                    x, y, dev,
                    color=color,
                    fontsize=15,
                    fontweight='bold',
                    ha='center',
                    va='center'
                )
            # Axis labels and title
            units = {
                'Voc': 'V',
                'Jsc': 'mA/cm²',
                'FF': '%',
                'Eff': '%'
            }
            self.ax.set_ylabel(f"{param_type} ({units[param_type]})", fontsize=14)
            self.ax.set_xlabel("")
            self.ax.set_title(f"Statistical Distribution of {param_type}", fontsize=16)
            self.ax.tick_params(axis='x', length=0)
            self.fig.subplots_adjust(left=0.18, bottom=0.18, right=0.97, top=0.92)
            self.ax.grid(axis='y', linestyle='--', alpha=0.4)
            # Show stats in statsLabel
            all_values = df['Value'].values
            avg = np.mean(all_values) if len(all_values) else 0
            std = np.std(all_values) if len(all_values) else 0
            max_val = np.max(all_values) if len(all_values) else 0
            min_val = np.min(all_values) if len(all_values) else 0
            self.statsLabel.setText(f"{param_type} Statistics:\nMaximum: {max_val:.4f} {units[param_type]}\nMinimum: {min_val:.4f} {units[param_type]}\nAverage: {avg:.4f} {units[param_type]}\nStd Dev: {std:.4f} {units[param_type]}")
            self.canvas.draw()
            self.status.showMessage(f"Generated {param_type} violin+box plot")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Error generating plot:\n{str(e)}")
            self.status.showMessage("Plot generation failed")

    def on_param_legend_changed(self):
        # Redraw the current parameter plot with/without legend
        for param, btn in self.param_buttons.items():
            if btn.isChecked() and param != 'JV Plot':
                self.plot_parameter(param.split()[0])
                break

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    QtWidgets.QApplication.setStyle('Fusion')
    win = JVPlotter()
    win.show()
    sys.exit(app.exec_())
