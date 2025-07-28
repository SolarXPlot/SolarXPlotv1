import os
import pandas as pd
import numpy as np
from PyQt5 import QtWidgets, QtGui, QtCore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import matplotlib.pyplot as plt
from peakutils import baseline
import re

class XRDPlotter(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        
        # Create controls
        controls_layout = QtWidgets.QHBoxLayout()
        
        # File selection button
        self.load_button = QtWidgets.QPushButton("Load Data File")
        self.load_button.clicked.connect(self.load_data)
        controls_layout.addWidget(self.load_button)
        
        # Degree slider
        degree_layout = QtWidgets.QHBoxLayout()
        degree_layout.addWidget(QtWidgets.QLabel("Baseline Degree:"))
        self.degree_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.degree_slider.setMinimum(1)
        self.degree_slider.setMaximum(20)
        self.degree_slider.setValue(6)
        self.degree_slider.valueChanged.connect(self.update_plot)
        degree_layout.addWidget(self.degree_slider)
        self.degree_label = QtWidgets.QLabel("6")
        degree_layout.addWidget(self.degree_label)
        controls_layout.addLayout(degree_layout)
        
        # Baseline toggle
        self.baseline_checkbox = QtWidgets.QCheckBox("Show Baseline")
        self.baseline_checkbox.setChecked(True)
        self.baseline_checkbox.stateChanged.connect(self.update_plot)
        controls_layout.addWidget(self.baseline_checkbox)
        
        # Style options
        style_layout = QtWidgets.QHBoxLayout()
        
        # Bold labels checkbox
        self.bold_labels_checkbox = QtWidgets.QCheckBox("Bold Labels")
        self.bold_labels_checkbox.stateChanged.connect(self.update_plot)
        style_layout.addWidget(self.bold_labels_checkbox)
        
        # Major ticks checkbox
        self.major_ticks_checkbox = QtWidgets.QCheckBox("Major Ticks")
        self.major_ticks_checkbox.setChecked(True)
        self.major_ticks_checkbox.stateChanged.connect(self.update_plot)
        style_layout.addWidget(self.major_ticks_checkbox)
        
        # Minor ticks checkbox
        self.minor_ticks_checkbox = QtWidgets.QCheckBox("Minor Ticks")
        self.minor_ticks_checkbox.stateChanged.connect(self.update_plot)
        style_layout.addWidget(self.minor_ticks_checkbox)
        
        # Bold scale checkbox
        self.bold_scale_checkbox = QtWidgets.QCheckBox("Bold Scale")
        self.bold_scale_checkbox.stateChanged.connect(self.update_plot)
        style_layout.addWidget(self.bold_scale_checkbox)
        
        controls_layout.addLayout(style_layout)
        
        # Save button
        self.save_button = QtWidgets.QPushButton("Save Corrected Data")
        self.save_button.clicked.connect(self.save_data)
        controls_layout.addWidget(self.save_button)
        
        layout.addLayout(controls_layout)
        
        # Create matplotlib figure
        self.figure = plt.Figure(figsize=(12, 8))
        self.canvas = FigureCanvasQTAgg(self.figure)
        layout.addWidget(self.canvas)
        
        # Initialize variables
        self.df = None
        self.deg = 6
        self.show_baseline = True
        
    def find_data_start(self, path):
        """Return the zero-based line number where lines begin with two floats."""
        pat = re.compile(r'^\s*[-+]?\d+(\.\d+)?\s+[-+]?\d+(\.\d+)?')
        with open(path) as f:
            for i, line in enumerate(f):
                if pat.match(line):
                    return i
        raise ValueError("No numeric data found.")
    
    def load_data(self):
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Data File", "", "Text Files (*.txt)")
        if file_name:
            try:
                start = self.find_data_start(file_name)
                self.df = pd.read_csv(
                    file_name,
                    sep=r'\s+',
                    skiprows=start,
                    header=None,
                    names=['two_theta','intensity']
                )
                self.df.set_index('two_theta', inplace=True)
                self.update_plot()
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Error", f"Error loading file: {str(e)}")
    
    def update_plot(self):
        if self.df is None:
            return
            
        self.deg = self.degree_slider.value()
        self.degree_label.setText(str(self.deg))
        self.show_baseline = self.baseline_checkbox.isChecked()
        
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        bl = baseline(self.df.intensity, self.deg)
        
        if self.show_baseline:
            # Plot raw data, baseline, and corrected
            ax.plot(self.df.index.to_numpy(), self.df.intensity.to_numpy(), alpha=0.2, label='Raw Data')
            ax.plot(self.df.index.to_numpy(), bl, alpha=0.2, label='Baseline')
            ax.plot(self.df.index.to_numpy(), (self.df.intensity - bl).to_numpy(), color='green', label='Corrected')
        else:
            # Only plot corrected data
            ax.plot(self.df.index.to_numpy(), (self.df.intensity - bl).to_numpy(), color='green', label='Corrected')
        
        # Apply style settings
        fontweight = 'bold' if self.bold_labels_checkbox.isChecked() else 'normal'
        ax.set_xlabel(r'2$\theta$ ($^{\circ}$)', fontweight=fontweight)
        ax.set_ylabel('Intensity (a.u.)', fontweight=fontweight)
        
        # Configure ticks
        if self.major_ticks_checkbox.isChecked():
            ax.tick_params(axis='both', which='major', width=1.0, length=5)
        else:
            ax.tick_params(axis='both', which='major', width=0, length=0)
            
        if self.minor_ticks_checkbox.isChecked():
            ax.minorticks_on()
            ax.tick_params(axis='both', which='minor', width=0.5, length=3)
        else:
            ax.minorticks_off()
            
        # Apply bold scale if enabled
        if self.bold_scale_checkbox.isChecked():
            for spine in ax.spines.values():
                spine.set_linewidth(2.0)
        else:
            for spine in ax.spines.values():
                spine.set_linewidth(1.0)
        
        ax.legend()
        self.canvas.draw()
    
    def save_data(self):
        if self.df is None:
            return
            
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Corrected Data", "", "CSV Files (*.csv)")
        if file_name:
            try:
                bl = baseline(self.df.intensity, self.deg)
                corrected_df = self.df.copy()
                corrected_df['intensity'] = corrected_df['intensity'] - bl
                corrected_df.to_csv(file_name)
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Error", f"Error saving file: {str(e)}") 