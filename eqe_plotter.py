import os
import pandas as pd
import numpy as np
from PyQt5 import QtWidgets, QtGui, QtCore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, NullLocator

from widgets import CollapsibleSection, BackgroundSettingsDialog

class EQEPlotter(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.eqe_data_list = []
        # --- Add background color attributes ---
        self.bgColor = QtGui.QColor(255, 255, 255)  # Default white
        self.bgColor2 = QtGui.QColor("#2c3e50")    # Default gradient end
        self.use_gradient = False
        main_layout = QtWidgets.QVBoxLayout(self)
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_layout.addWidget(splitter)

        # --- Reorganized Controls Panel ---
        controls_container = QtWidgets.QWidget()
        cl = QtWidgets.QVBoxLayout(controls_container)
        cl.setContentsMargins(5, 5, 5, 5)

        # File Management Section
        file_section = QtWidgets.QGroupBox("File Management")
        file_layout = QtWidgets.QVBoxLayout(file_section)
        
        file_buttons = QtWidgets.QHBoxLayout()
        self.eqeAddBtn = QtWidgets.QPushButton("Add EQE File…")
        self.eqeClearBtn = QtWidgets.QPushButton("Clear Files")
        file_buttons.addWidget(self.eqeAddBtn)
        file_buttons.addWidget(self.eqeClearBtn)
        file_layout.addLayout(file_buttons)

        self.eqeFileList = QtWidgets.QListWidget()
        self.eqeFileList.setToolTip("Double-click to edit legend")
        self.eqeFileList.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked)
        self.eqeFileList.setMaximumHeight(150)
        file_layout.addWidget(self.eqeFileList)
        cl.addWidget(file_section)

        # --- Horizontal layout for Tools, Plot, Style ---
        top_sections_layout = QtWidgets.QHBoxLayout()
        # Tools Section (was Customization)
        tools_section = CollapsibleSection("Tools")
        tools_form = QtWidgets.QFormLayout()
        self.titleEdit = QtWidgets.QLineEdit("EQE Comparison")
        tools_form.addRow("Plot Title:", self.titleEdit)
        self.xLabelEdit = QtWidgets.QLineEdit("Wavelength (nm)")
        tools_form.addRow("X-axis Label:", self.xLabelEdit)
        self.yLabelEditLeft = QtWidgets.QLineEdit("EQE (%)")
        tools_form.addRow("Left Y-axis Label:", self.yLabelEditLeft)
        self.yLabelEditRight = QtWidgets.QLineEdit("Integrated current (mA/cm²)")
        tools_form.addRow("Right Y-axis Label:", self.yLabelEditRight)
        self.lineStyleCombo = QtWidgets.QComboBox(); self.lineStyleCombo.addItems(["Solid","Dashed","Dotted","Dash-dot"])
        tools_form.addRow("Line Style:", self.lineStyleCombo)
        self.markerCombo = QtWidgets.QComboBox(); self.markerCombo.addItems(["None","Circle","Square","Triangle","Diamond","Cross"])
        tools_form.addRow("Marker Style:", self.markerCombo)
        tools_section.addLayout(tools_form)
        top_sections_layout.addWidget(tools_section)
        # Plot Section (was Axes & Ticks)
        plot_section = CollapsibleSection("Plot")
        plot_form = QtWidgets.QFormLayout()
        self.xminSpin = QtWidgets.QDoubleSpinBox(); self.xminSpin.setRange(-1e4, 1e4); self.xminSpin.setDecimals(3)
        plot_form.addRow("X Min:", self.xminSpin)
        self.xmaxSpin = QtWidgets.QDoubleSpinBox(); self.xmaxSpin.setRange(-1e4, 1e4); self.xmaxSpin.setDecimals(3)
        plot_form.addRow("X Max:", self.xmaxSpin)
        self.yminSpin = QtWidgets.QDoubleSpinBox(); self.yminSpin.setRange(-1e4, 1e4); self.yminSpin.setDecimals(3)
        plot_form.addRow("Left Y Min:", self.yminSpin)
        self.ymaxSpinLeft = QtWidgets.QDoubleSpinBox(); self.ymaxSpinLeft.setRange(-1e4, 1e4); self.ymaxSpinLeft.setDecimals(3)
        plot_form.addRow("Left Y Max:", self.ymaxSpinLeft)
        self.yminRightSpin = QtWidgets.QDoubleSpinBox(); self.yminRightSpin.setRange(-1e4, 1e4); self.yminRightSpin.setDecimals(3)
        plot_form.addRow("Right Y Min:", self.yminRightSpin)
        self.ymaxSpinRight = QtWidgets.QDoubleSpinBox(); self.ymaxSpinRight.setRange(-1e4, 1e4); self.ymaxSpinRight.setDecimals(3)
        plot_form.addRow("Right Y Max:", self.ymaxSpinRight)
        self.xStepSpin = QtWidgets.QDoubleSpinBox(); self.xStepSpin.setRange(0.001,1e3); self.xStepSpin.setDecimals(3)
        plot_form.addRow("X Tick Step:", self.xStepSpin)
        self.yStepSpinLeft = QtWidgets.QDoubleSpinBox(); self.yStepSpinLeft.setRange(0.1,1e3); self.yStepSpinLeft.setDecimals(3)
        self.yStepSpinLeft.setValue(10)
        plot_form.addRow("Left Y Tick Step:", self.yStepSpinLeft)
        self.yStepSpinRight = QtWidgets.QDoubleSpinBox(); self.yStepSpinRight.setRange(0.1,1e3); self.yStepSpinRight.setDecimals(3)
        self.yStepSpinRight.setValue(5)
        plot_form.addRow("Right Y Tick Step:", self.yStepSpinRight)
        self.scaleCombo = QtWidgets.QComboBox(); self.scaleCombo.addItems(["Linear","Log X","Log Y","Log XY"])
        plot_form.addRow("Graph Type:", self.scaleCombo)
        self.majorChk = QtWidgets.QCheckBox("Show Major Ticks"); self.majorChk.setChecked(True)
        plot_form.addRow(self.majorChk)
        self.minorChk = QtWidgets.QCheckBox("Show Minor Ticks"); plot_form.addRow(self.minorChk)
        self.boldChk = QtWidgets.QCheckBox("Bold Labels"); plot_form.addRow(self.boldChk)
        plot_section.addLayout(plot_form)
        top_sections_layout.addWidget(plot_section)
        # Style Section (was Style & Export)
        style_section = CollapsibleSection("Style")
        style_form = QtWidgets.QFormLayout()
        cdl = QtWidgets.QHBoxLayout()
        self.colorBtn = QtWidgets.QPushButton(); self.colorBtn.setFixedSize(30,30)
        self.color = QtGui.QColor(31,119,180); self._update_color_button()
        cdl.addWidget(QtWidgets.QLabel("Base Color:")); cdl.addWidget(self.colorBtn)
        cdl.addStretch()
        style_form.addRow(cdl)
        # --- Font controls ---
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
        # --- End font controls ---
        # --- Add Background Settings Button ---
        self.bgSettingsBtn = QtWidgets.QPushButton("Background Settings...")
        style_form.addRow("Background:", self.bgSettingsBtn)
        self.dpiSpin = QtWidgets.QSpinBox(); self.dpiSpin.setRange(50,2400); self.dpiSpin.setValue(300)
        style_form.addRow("Save DPI:", self.dpiSpin)
        self.annotateBtn = QtWidgets.QPushButton("Add Annotation")
        style_form.addRow(self.annotateBtn)
        # --- Legend Settings for EQE ---
        legend_group = QtWidgets.QGroupBox("Legend Settings")
        legend_layout = QtWidgets.QVBoxLayout(legend_group)
        # Transparency slider
        self.eqeLegendBgAlphaSlider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.eqeLegendBgAlphaSlider.setRange(0, 100)
        self.eqeLegendBgAlphaSlider.setValue(100)
        legend_bg_alpha_layout = QtWidgets.QHBoxLayout()
        legend_bg_alpha_layout.addWidget(QtWidgets.QLabel("Background Opacity:"))
        legend_bg_alpha_layout.addWidget(self.eqeLegendBgAlphaSlider)
        self.eqeLegendBgAlphaValue = QtWidgets.QLabel("100%")
        legend_bg_alpha_layout.addWidget(self.eqeLegendBgAlphaValue)
        legend_layout.addLayout(legend_bg_alpha_layout)
        # Position combo
        position_layout = QtWidgets.QHBoxLayout()
        self.eqeLegendPosCombo = QtWidgets.QComboBox()
        self.eqeLegendPosCombo.addItems([
            "best", "upper right", "upper left", "lower left", 
            "lower right", "right", "center left", "center right", 
            "lower center", "upper center", "center"
        ])
        position_layout.addWidget(QtWidgets.QLabel("Position:"))
        position_layout.addWidget(self.eqeLegendPosCombo)
        legend_layout.addLayout(position_layout)
        style_form.addRow(legend_group)
        style_section.addLayout(style_form)
        top_sections_layout.addWidget(style_section)
        cl.addLayout(top_sections_layout)
        
        cl.addStretch() # Pushes buttons to the bottom
        
        # Action Buttons
        self.plotBtn = QtWidgets.QPushButton("Generate EQE Plot")
        self.saveBtn = QtWidgets.QPushButton("Save Plot…")
        cl.addWidget(self.plotBtn)
        cl.addWidget(self.saveBtn)
        
        splitter.addWidget(controls_container)

        # Plot canvas
        plot_frame = QtWidgets.QFrame()
        plot_frame.setStyleSheet("background:#fff; border:1px solid #ccc; padding:4px;")
        pl = QtWidgets.QVBoxLayout(plot_frame)
        plt.rcParams.update({'font.family':'serif','font.size':12,'axes.linewidth':1,'xtick.direction':'in','ytick.direction':'in'})
        self.fig, self.ax1 = plt.subplots(figsize=(6,6), dpi=200)
        self.ax2 = self.ax1.twinx()
        self.canvas = FigureCanvasQTAgg(self.fig)
        pl.addWidget(self.canvas)
        splitter.addWidget(plot_frame)
        
        # Connect signals
        self.eqeAddBtn.clicked.connect(self.eqe_add_files)
        self.eqeClearBtn.clicked.connect(self.eqe_clear_files)
        self.colorBtn.clicked.connect(self.pick_color)
        self.annotateBtn.clicked.connect(self.add_annotation)
        self.plotBtn.clicked.connect(self.eqe_do_plot)
        self.saveBtn.clicked.connect(self.save_plot)
        self.eqeLegendBgAlphaSlider.valueChanged.connect(self.update_eqe_legend_transparency)
        self.eqeLegendPosCombo.currentTextChanged.connect(self.update_eqe_legend_position)
        # --- Connect background settings button ---
        self.bgSettingsBtn.clicked.connect(self.show_background_settings)
        # --- Connect font controls ---
        self.fontCombo.currentTextChanged.connect(self.update_font_style)
        self.fontSizeCombo.currentTextChanged.connect(self.update_font_style)
        
    def _update_color_button(self):
        pix = QtGui.QPixmap(30,30); pix.fill(self.color); self.colorBtn.setIcon(QtGui.QIcon(pix))
        
    def pick_color(self):
        c = QtWidgets.QColorDialog.getColor(self.color, self)
        if c.isValid():
            self.color = c
            self._update_color_button()
            
    def eqe_add_files(self):
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Select EQE Data Files", "", "Data Files (*.txt *.dat)")
        if not paths:
            return
        added = 0
        for path in paths:
            try:
                header_idx = None
                with open(path, 'r') as f:
                    for i, line in enumerate(f):
                        if line.strip().startswith('Lambda'):
                            header_idx = i
                            break
                if header_idx is None:
                    raise RuntimeError("No 'Lambda' header found")
                df = pd.read_csv(
                    path,
                    sep='\s+',
                    skiprows=header_idx,
                    header=0,
                    usecols=[0,1,2,3],
                    names=['Wavelength','EQE','SR','Jsc']
                ).astype(float)
                df['Integrated'] = df['Jsc'].cumsum()
                self.eqe_data_list.append((os.path.basename(path), df))
                itm = QtWidgets.QListWidgetItem(os.path.basename(path))
                itm.setFlags(itm.flags() | QtCore.Qt.ItemIsEditable)
                self.eqeFileList.addItem(itm)
                added += 1
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Error", f"Failed {path}:\n{e}")
        if self.parent() and hasattr(self.parent(), 'status'):
            self.parent().status.showMessage(f"Added {added} files")
    def eqe_clear_files(self):
        self.eqe_data_list.clear()
        self.eqeFileList.clear()
        self.ax1.clear()
        self.ax2.clear()
        self.canvas.draw()
        if self.parent() and hasattr(self.parent(), 'status'):
            self.parent().status.showMessage("Cleared files and plot area")
    def show_background_settings(self):
        dialog = BackgroundSettingsDialog(self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.use_gradient = dialog.bgTypeCombo.currentText() == "Gradient"
            if self.use_gradient:
                self.bgColor = dialog.startColor
                self.bgColor2 = dialog.endColor
            else:
                self.bgColor = dialog.bgColor
            self.eqe_do_plot()
    def _update_plot_background(self):
        if self.use_gradient:
            # Create a smooth gradient background for the entire figure
            points = 100
            gradient = np.linspace(0, 1, points).reshape(-1, 1)
            start_color = np.array([self.bgColor.red()/255.0, self.bgColor.green()/255.0, self.bgColor.blue()/255.0])
            end_color = np.array([self.bgColor2.red()/255.0, self.bgColor2.green()/255.0, self.bgColor2.blue()/255.0])
            colors = start_color + gradient * (end_color - start_color)
            gradient_array = np.zeros((points, 1, 3))
            for i in range(points):
                gradient_array[i, 0] = colors[i]
            # Create a new axes for the background that covers the entire figure
            if not hasattr(self, 'bg_axes'):
                self.bg_axes = self.fig.add_axes([0, 0, 1, 1])
                self.bg_axes.set_zorder(-1)
            else:
                self.bg_axes.clear()
            self.bg_axes.imshow(gradient_array, aspect='auto', extent=[0, 1, 0, 1], interpolation='gaussian')
            self.bg_axes.axis('off')
            self.ax1.set_facecolor('none')
            self.ax1.patch.set_alpha(0.0)
            self.fig.patch.set_facecolor('none')
        else:
            if hasattr(self, 'bg_axes'):
                self.bg_axes.remove()
                delattr(self, 'bg_axes')
            self.fig.patch.set_facecolor(self.bgColor.name())
            self.ax1.set_facecolor(self.bgColor.name())
        # Ensure proper plot positioning
        for spine in self.ax1.spines.values():
            spine.set_visible(True)
            spine.set_linewidth(1.0)
    def eqe_do_plot(self):
        if not self.eqe_data_list:
            return
        self.ax1.clear()
        self.ax2.clear()
        # --- Update background before plotting ---
        self._update_plot_background()
        # --- Update font style before plotting ---
        self.update_font_style()
        xmin, xmax = self.xminSpin.value(), self.xmaxSpin.value()
        yminL, ymaxL = self.yminSpin.value(), self.ymaxSpinLeft.value()
        yminR, ymaxR = self.yminRightSpin.value(), self.ymaxSpinRight.value()
        if xmin != xmax:
            self.ax1.set_xlim(xmin, xmax)
        if yminL != ymaxL:
            self.ax1.set_ylim(yminL, ymaxL)
        if yminR != ymaxR:
            self.ax2.set_ylim(yminR, ymaxR)
        def safe_locator(axis, step):
            rng = axis.get_view_interval()
            if step > 0 and (rng[1]-rng[0])/step < 1000:
                return MultipleLocator(step)
            return NullLocator()
        scale = self.scaleCombo.currentText()
        self.ax1.set_xscale('log' if 'Log X' in scale else 'linear')
        self.ax1.set_yscale('log' if 'Log Y' in scale else 'linear')
        self.ax2.set_yscale('log' if 'Log Y' in scale else 'linear')
        cmap = plt.cm.tab10
        ls_map = {'Solid':'-','Dashed':'--','Dotted':':','Dash-dot':'-.'}
        mk_map = {'None':'','Circle':'o','Square':'s','Triangle':'^','Diamond':'D','Cross':'x'}
        for i, (label, df) in enumerate(self.eqe_data_list):
            color = cmap(i % 10)
            ls = ls_map[self.lineStyleCombo.currentText()]
            mk = mk_map[self.markerCombo.currentText()]
            self.ax1.plot(df['Wavelength'].to_numpy(), df['EQE'].to_numpy(), linestyle=ls, marker=mk, color=self.color.name(), label=f"{label} EQE")
            self.ax2.plot(df['Wavelength'].to_numpy(), df['Integrated'].to_numpy(), linestyle=ls, marker=mk, color=color, alpha=0.7, label=f"{label} Int.")
        weight = 'bold' if self.boldChk.isChecked() else 'normal'
        self.ax1.set_xlabel(self.xLabelEdit.text(), fontsize=14, fontweight=weight)
        self.ax1.set_ylabel(self.yLabelEditLeft.text(), fontsize=14, fontweight=weight)
        self.ax2.set_ylabel(self.yLabelEditRight.text(), fontsize=14, fontweight=weight)
        self.ax2.yaxis.set_label_position('right')
        self.ax2.yaxis.set_ticks_position('right')
        self.ax1.set_title(self.titleEdit.text(), fontsize=18, fontweight=weight)
        if self.majorChk.isChecked():
            self.ax1.xaxis.set_major_locator(safe_locator(self.ax1.xaxis, self.xStepSpin.value()))
            self.ax1.yaxis.set_major_locator(safe_locator(self.ax1.yaxis, self.yStepSpinLeft.value()))
            self.ax2.yaxis.set_major_locator(safe_locator(self.ax2.yaxis, self.yStepSpinRight.value()))
        else:
            self.ax1.xaxis.set_major_locator(NullLocator())
            self.ax1.yaxis.set_major_locator(NullLocator())
            self.ax2.yaxis.set_major_locator(NullLocator())
        if self.minorChk.isChecked():
            self.ax1.minorticks_on()
            self.ax2.minorticks_on()
        else:
            self.ax1.minorticks_off()
            self.ax2.minorticks_off()
        lines1, labels1 = self.ax1.get_legend_handles_labels()
        lines2, labels2 = self.ax2.get_legend_handles_labels()
        legend = self.ax1.legend(lines1 + lines2, labels1 + labels2, loc=self.eqeLegendPosCombo.currentText(), fontsize=10)
        # Apply legend transparency
        alpha = self.eqeLegendBgAlphaSlider.value() / 100
        legend.get_frame().set_alpha(alpha)
        self.eqeLegendBgAlphaValue.setText(f"{self.eqeLegendBgAlphaSlider.value()}%")
        self.fig.tight_layout()
        self.canvas.draw()
        if self.parent() and hasattr(self.parent(), 'status'):
            self.parent().status.showMessage("EQE plot updated")
    def add_annotation(self):
        x, ok = QtWidgets.QInputDialog.getDouble(self, "Annotation X", "X:", 0, -1e6, 1e6, 3)
        if not ok:
            return
        y, ok = QtWidgets.QInputDialog.getDouble(self, "Annotation Y", "Y:", 0, -1e6, 1e6, 3)
        if not ok:
            return
        txt, ok = QtWidgets.QInputDialog.getText(self, "Annotation Text", "Label:")
        if ok:
            self.ax1.annotate(txt, xy=(x, y), xytext=(x+0.1, y+0.1), arrowprops=dict(arrowstyle="->"), fontsize=10)
            self.canvas.draw()
    def save_plot(self):
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Plot", "", "PNG Files (*.png);;PDF Files (*.pdf)")
        if not fname:
            return
        try:
            self.fig.savefig(fname, dpi=self.dpiSpin.value(), bbox_inches='tight')
            if self.parent() and hasattr(self.parent(), 'status'):
                self.parent().status.showMessage(f"Saved to {fname}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Save failed:\n{e}")
            if self.parent() and hasattr(self.parent(), 'status'):
                self.parent().status.showMessage("Save failed")
    def update_eqe_legend_transparency(self):
        self.eqeLegendBgAlphaValue.setText(f"{self.eqeLegendBgAlphaSlider.value()}%")
        # Redraw plot to update legend transparency
        self.eqe_do_plot()
    def update_eqe_legend_position(self, position):
        # Redraw plot to update legend position
        self.eqe_do_plot() 
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
                if hasattr(self, 'ax1'):
                    # Store current limits
                    xlim1 = self.ax1.get_xlim()
                    ylim1 = self.ax1.get_ylim()
                    xlim2 = self.ax2.get_xlim()
                    ylim2 = self.ax2.get_ylim()
                    # Update font for all text elements on ax1 and ax2
                    for ax in [self.ax1, self.ax2]:
                        for text in ax.texts:
                            if font == "Computer Modern":
                                text.set_fontname("Computer Modern Roman")
                            else:
                                text.set_fontname(font)
                            text.set_fontsize(font_size)
                        # Update axis labels and title with new size
                        ax.set_xlabel(ax.get_xlabel(), fontfamily=font, fontsize=font_size * 1.2)
                        ax.set_ylabel(ax.get_ylabel(), fontfamily=font, fontsize=font_size * 1.2)
                        ax.set_title(ax.get_title(), fontfamily=font, fontsize=font_size * 1.5)
                        # Update tick labels
                        for label in ax.get_xticklabels() + ax.get_yticklabels():
                            if font == "Computer Modern":
                                label.set_fontname("Computer Modern Roman")
                            else:
                                label.set_fontname(font)
                            label.set_fontsize(font_size)
                        # Update legend
                        legend = ax.get_legend()
                        if legend:
                            for text in legend.get_texts():
                                if font == "Computer Modern":
                                    text.set_fontname("Computer Modern Roman")
                                else:
                                    text.set_fontname(font)
                                text.set_fontsize(font_size * 0.85)
                    # Restore limits
                    self.ax1.set_xlim(xlim1)
                    self.ax1.set_ylim(ylim1)
                    self.ax2.set_xlim(xlim2)
                    self.ax2.set_ylim(ylim2)
                    self.fig.tight_layout()
                    self.canvas.draw()
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Font Error", f"Could not set font to {font}. Error: {str(e)}\nReverting to default font.")
                self.fontCombo.setCurrentText("Default")
                plt.rcParams.update({'font.family': 'serif', 'font.size': font_size}) 