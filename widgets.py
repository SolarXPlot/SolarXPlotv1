from PyQt5 import QtWidgets, QtGui, QtCore

class CollapsibleSection(QtWidgets.QWidget):
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setStyleSheet("")  # Use global QSS
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        # Header button with icon
        self.toggle_button = QtWidgets.QPushButton(title)
        self.toggle_button.setProperty("sectionHeader", True)
        self.toggle_button.setIcon(QtGui.QIcon.fromTheme("go-down"))
        self.toggle_button.setIconSize(QtCore.QSize(16, 16))
        layout.addWidget(self.toggle_button)
        # Content
        self.content = QtWidgets.QFrame()
        self.content_layout = QtWidgets.QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(10, 5, 5, 5)
        layout.addWidget(self.content)
        # Connect toggle button
        self.toggle_button.clicked.connect(self.toggle_section)
        self.is_expanded = True
        
    def toggle_section(self):
        self.is_expanded = not self.is_expanded
        self.content.setVisible(self.is_expanded)
        self.toggle_button.setIcon(QtGui.QIcon.fromTheme("go-down" if self.is_expanded else "go-next"))
        
    def addWidget(self, widget):
        self.content_layout.addWidget(widget)
        
    def addLayout(self, layout):
        self.content_layout.addLayout(layout) 

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
        self.bgColor = parent.bgColor if parent and hasattr(parent, 'bgColor') else QtGui.QColor(255, 255, 255)
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
        
        self.startColor = parent.bgColor if parent and hasattr(parent, 'bgColor') else QtGui.QColor("#bdc3c7")
        self.endColor = parent.bgColor2 if parent and hasattr(parent, 'bgColor2') else QtGui.QColor("#2c3e50")
        
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