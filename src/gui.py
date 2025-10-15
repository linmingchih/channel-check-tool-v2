import os
import json
import re
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QLineEdit,
    QTabWidget,
    QListWidget,
    QListWidgetItem,
    QGroupBox,
    QCheckBox,
    QStatusBar,
    QFileDialog,
    QScrollArea,
    QAbstractItemView,
    QGridLayout,
    QTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QSpinBox,
    QRadioButton,
)
from PySide6.QtCore import Qt, QProcess
from PySide6.QtGui import QColor


class NetListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Space:
            self.toggle_selected_items_check_state()
        else:
            super().keyPressEvent(event)

    def toggle_selected_items_check_state(self):
        selected_items = self.selectedItems()
        if not selected_items:
            return

        target_state = (
            Qt.Checked
            if selected_items[0].checkState() == Qt.Unchecked
            else Qt.Unchecked
        )

        for item in selected_items:
            item.setCheckState(target_state)


class AEDBCCTCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AEDB CCT Calculator")
        self.setGeometry(100, 100, 1200, 800)
        self.pcb_data = None
        self.all_components = []
        self.config_file = os.path.join(os.path.dirname(__file__), "..", "data", "config.json")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        self.tabs = QTabWidget()
        self.import_tab = QWidget()
        self.port_setup_tab = QWidget()
        self.simulation_tab = QWidget()
        self.cct_tab = QWidget()
        self.tabs.addTab(self.import_tab, "Import")
        self.tabs.addTab(self.port_setup_tab, "Port Setup")
        self.tabs.addTab(self.simulation_tab, "Simulation")
        self.tabs.addTab(self.cct_tab, "CCT")
        main_layout.addWidget(self.tabs)

        log_group = QGroupBox("Information")
        log_group.setObjectName("logGroup")
        log_layout = QVBoxLayout(log_group)
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        self.log_window.setObjectName("logWindow")
        log_layout.addWidget(self.log_window)
        main_layout.addWidget(log_group)

        self.setup_import_tab()
        self.setup_port_setup_tab()
        self.setup_simulation_tab()
        self.setup_cct_tab()
        self.apply_styles()

    def setup_import_tab(self):
        import_layout = QVBoxLayout(self.import_tab)
        import_group = QGroupBox("Layout Import")
        import_group_layout = QGridLayout(import_group)

        # Row 1: Layout type selection
        self.brd_radio = QRadioButton(".brd")
        self.aedb_radio = QRadioButton(".aedb")
        self.brd_radio.setChecked(True)
        import_group_layout.addWidget(self.brd_radio, 0, 0)
        import_group_layout.addWidget(self.aedb_radio, 0, 1)

        # Row 2: Path selection
        self.open_layout_button = QPushButton("Open...")
        self.layout_path_label = QLabel("No design loaded")
        import_group_layout.addWidget(QLabel("Design:"), 1, 0)
        import_group_layout.addWidget(self.layout_path_label, 1, 1)
        import_group_layout.addWidget(self.open_layout_button, 1, 2)

        # Row 3: Stackup selection
        self.stackup_path_input = QLineEdit()
        self.browse_stackup_button = QPushButton("Browse...")
        import_group_layout.addWidget(QLabel("Stackup (.xml):"), 2, 0)
        import_group_layout.addWidget(self.stackup_path_input, 2, 1)
        import_group_layout.addWidget(self.browse_stackup_button, 2, 2)
        
        # EDB version input
        import_group_layout.addWidget(QLabel("EDB version:"), 3, 0)
        self.edb_version_input = QLineEdit()
        self.edb_version_input.setFixedWidth(60)
        import_group_layout.addWidget(self.edb_version_input, 3, 1)

        import_layout.addWidget(import_group)

        self.apply_import_button = QPushButton("Apply")
        import_layout.addWidget(self.apply_import_button, alignment=Qt.AlignRight)

        import_layout.addStretch()

    def apply_styles(self):
        self.setStyleSheet("""
            QPushButton {
                padding: 5px 10px;
                border: 1px solid #ccc;
                border-radius: 3px;
                background-color: #f0f0f0;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
            QGroupBox#logGroup {
                padding: 12px 2px 2px 2px;
                margin: 10px 0 0 0;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QGroupBox#logGroup::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                left: 4px;
            }
            QTextEdit#logWindow {
                border: none;
                padding: 0;
                margin: 0;
            }
        """)
        primary_style = "background-color: #007bff; color: white; border: none;"
        self.calculate_button.setStyleSheet(primary_style)
        self.calculate_button_original_style = primary_style
        self.apply_button.setStyleSheet(primary_style)
        self.apply_simulation_button.setStyleSheet(primary_style)
        self.apply_simulation_button_original_style = primary_style
        self.apply_import_button.setStyleSheet(primary_style)
        self.apply_import_button_original_style = primary_style

        secondary_style = "background-color: #6c757d; color: white; border: none;"
        self.prerun_button.setStyleSheet(secondary_style)
        self.prerun_button_original_style = secondary_style

    def setup_port_setup_tab(self):
        port_setup_layout = QVBoxLayout(self.port_setup_tab)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Component filter (regex):"))
        self.component_filter_input = QLineEdit("^[UJ]")
        filter_layout.addWidget(self.component_filter_input)
        port_setup_layout.addLayout(filter_layout)

        components_layout = QHBoxLayout()
        self.controller_components_list = QListWidget()
        self.controller_components_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.dram_components_list = QListWidget()
        self.dram_components_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        controller_group = QGroupBox("Controller Components")
        controller_layout = QVBoxLayout(controller_group)
        controller_layout.addWidget(self.controller_components_list)
        dram_group = QGroupBox("DRAM Components")
        dram_layout = QVBoxLayout(dram_group)
        dram_layout.addWidget(self.dram_components_list)
        components_layout.addWidget(controller_group)
        components_layout.addWidget(dram_group)
        port_setup_layout.addLayout(components_layout)

        ref_net_layout = QHBoxLayout()
        ref_net_layout.addWidget(QLabel("Reference net:"))
        self.ref_net_combo = QComboBox()
        self.ref_net_combo.setMinimumWidth(150)
        self.ref_net_combo.addItems(["GND"])
        ref_net_layout.addWidget(self.ref_net_combo)
        ref_net_layout.addStretch()
        self.checked_nets_label = QLabel("Checked nets: 0 | Ports: 0")
        ref_net_layout.addWidget(self.checked_nets_label)
        port_setup_layout.addLayout(ref_net_layout)

        nets_layout = QHBoxLayout()
        single_ended_group = QGroupBox("Single-Ended Nets")
        self.single_ended_list = NetListWidget()
        single_ended_layout = QVBoxLayout(single_ended_group)
        single_ended_layout.addWidget(self.single_ended_list)
        differential_pairs_group = QGroupBox("Differential Pairs")
        self.differential_pairs_list = NetListWidget()
        differential_pairs_layout = QVBoxLayout(differential_pairs_group)
        differential_pairs_layout.addWidget(self.differential_pairs_list)
        nets_layout.addWidget(single_ended_group)
        nets_layout.addWidget(differential_pairs_group)
        port_setup_layout.addLayout(nets_layout)

        self.apply_button = QPushButton("Apply")
        self.apply_button.setEnabled(False)
        port_setup_layout.addWidget(self.apply_button, alignment=Qt.AlignRight)

    def setup_simulation_tab(self):
        simulation_layout = QVBoxLayout(self.simulation_tab)

        cutout_group = QGroupBox("Cutout")
        cutout_layout = QGridLayout(cutout_group)
        self.enable_cutout_checkbox = QCheckBox("Enable cutout")
        self.expansion_size_input = QLineEdit("0.005000")
        self.signal_nets_label = QLabel("(not set)")
        self.signal_nets_label.setWordWrap(True)
        self.reference_net_label = QLabel("(not set)")
        cutout_layout.addWidget(self.enable_cutout_checkbox, 0, 0)
        cutout_layout.addWidget(QLabel("Expansion size (m)"), 1, 0)
        cutout_layout.addWidget(self.expansion_size_input, 1, 1)
        cutout_layout.addWidget(QLabel("Signal nets"), 2, 0)
        cutout_layout.addWidget(self.signal_nets_label, 2, 1)
        cutout_layout.addWidget(QLabel("Reference net"), 3, 0)
        cutout_layout.addWidget(self.reference_net_label, 3, 1)
        simulation_layout.addWidget(cutout_group)

        solver_group = QGroupBox("Solver")
        solver_layout = QHBoxLayout(solver_group)
        self.siwave_radio = QRadioButton("SIwave")
        self.hfss_radio = QRadioButton("HFSS")
        self.siwave_radio.setChecked(True)
        solver_layout.addWidget(self.siwave_radio)
        solver_layout.addWidget(self.hfss_radio)
        simulation_layout.addWidget(solver_group)

        sweeps_group = QGroupBox("Frequency Sweeps")
        sweeps_layout = QVBoxLayout(sweeps_group)
        self.sweeps_table = QTableWidget()
        self.sweeps_table.setColumnCount(4)
        self.sweeps_table.setHorizontalHeaderLabels(["Sweep Type", "Start", "Stop", "Step/Count"])
        self.sweeps_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.add_sweep(["linear count", "0", "1kHz", "3"])
        self.add_sweep(["log scale", "1kHz", "0.1GHz", "10"])
        self.add_sweep(["linear scale", "0.1GHz", "10GHz", "0.1GHz"])
        sweeps_layout.addWidget(self.sweeps_table)
        
        sweep_buttons_layout = QHBoxLayout()
        add_sweep_button = QPushButton("Add Sweep")
        add_sweep_button.clicked.connect(self.add_sweep)
        remove_sweep_button = QPushButton("Remove Selected")
        remove_sweep_button.clicked.connect(self.remove_selected_sweep)
        sweep_buttons_layout.addWidget(add_sweep_button)
        sweep_buttons_layout.addWidget(remove_sweep_button)
        sweep_buttons_layout.addStretch()
        sweeps_layout.addLayout(sweep_buttons_layout)
        simulation_layout.addWidget(sweeps_group)

        self.apply_simulation_button = QPushButton("Apply Simulation")
        simulation_layout.addWidget(self.apply_simulation_button, alignment=Qt.AlignRight)
        simulation_layout.addStretch()

    def add_sweep(self, sweep_data=None):
        if sweep_data is None:
            sweep_data = ["linear count", "", "", ""]
        row_position = self.sweeps_table.rowCount()
        self.sweeps_table.insertRow(row_position)

        sweep_type_combo = QComboBox()
        sweep_type_combo.addItems(["linear count", "log scale", "linear scale"])
        sweep_type_combo.setCurrentText(sweep_data[0])

        self.sweeps_table.setCellWidget(row_position, 0, sweep_type_combo)
        self.sweeps_table.setItem(row_position, 1, QTableWidgetItem(str(sweep_data[1])))
        self.sweeps_table.setItem(row_position, 2, QTableWidgetItem(str(sweep_data[2])))
        self.sweeps_table.setItem(row_position, 3, QTableWidgetItem(str(sweep_data[3])))

    def remove_selected_sweep(self):
        # Get all selected ranges
        selected_ranges = self.sweeps_table.selectedRanges()
        if not selected_ranges:
            return

        # Collect all unique rows to be removed from all selected ranges
        rows_to_remove = set()
        for s_range in selected_ranges:
            for row in range(s_range.topRow(), s_range.bottomRow() + 1):
                rows_to_remove.add(row)

        # Sort rows in descending order to avoid index shifting issues
        for row in sorted(list(rows_to_remove), reverse=True):
            self.sweeps_table.removeRow(row)

    def setup_cct_tab(self):
        cct_layout = QVBoxLayout(self.cct_tab)
        file_input_layout = QGridLayout()
        self.touchstone_path_input = QLineEdit()
        self.port_metadata_path_input = QLineEdit()
        self.browse_touchstone_button = QPushButton("Browse")
        self.browse_metadata_button = QPushButton("Browse")
        file_input_layout.addWidget(QLabel("Touchstone (.sNp):"), 0, 0)
        file_input_layout.addWidget(self.touchstone_path_input, 0, 1)
        file_input_layout.addWidget(self.browse_touchstone_button, 0, 2)
        file_input_layout.addWidget(QLabel("Port metadata (.json):"), 1, 0)
        file_input_layout.addWidget(self.port_metadata_path_input, 1, 1)
        file_input_layout.addWidget(self.browse_metadata_button, 1, 2)
        cct_layout.addLayout(file_input_layout)

        config_panels_layout = QHBoxLayout()

        def add_unit_widget(layout, row, label, value, unit):
            layout.addWidget(QLabel(label), row, 0)
            widget_layout = QHBoxLayout()
            line_edit = QLineEdit(value)
            widget_layout.addWidget(line_edit)
            widget_layout.addWidget(QLabel(unit))
            layout.addLayout(widget_layout, row, 1)
            return line_edit

        self.cct_defaults = {
            "tx_vhigh": "0.800",
            "tx_rise_time": "30.000",
            "tx_unit_interval": "133.000",
            "tx_resistance": "40.000",
            "tx_capacitance": "1.000",
            "rx_resistance": "30.000",
            "rx_capacitance": "1.800",
            "transient_step": "100.000",
            "transient_stop": "3.000",
            "aedt_version": "2025.2",
            "threshold": "-40.0",
        }

        tx_group = QGroupBox("TX Settings")
        tx_layout = QGridLayout(tx_group)
        self.tx_vhigh = add_unit_widget(tx_layout, 0, "TX Vhigh", self.cct_defaults["tx_vhigh"], "V")
        self.tx_rise_time = add_unit_widget(tx_layout, 1, "TX Rise Time", self.cct_defaults["tx_rise_time"], "ps")
        self.tx_unit_interval = add_unit_widget(tx_layout, 2, "Unit Interval", self.cct_defaults["tx_unit_interval"], "ps")
        self.tx_resistance = add_unit_widget(tx_layout, 3, "TX Resistance", self.cct_defaults["tx_resistance"], "ohm")
        self.tx_capacitance = add_unit_widget(tx_layout, 4, "TX Capacitance", self.cct_defaults["tx_capacitance"], "pF")
        tx_layout.setRowStretch(5, 1)
        config_panels_layout.addWidget(tx_group)

        rx_group = QGroupBox("RX Settings")
        rx_layout = QGridLayout(rx_group)
        self.rx_resistance = add_unit_widget(rx_layout, 0, "RX Resistance", self.cct_defaults["rx_resistance"], "ohm")
        self.rx_capacitance = add_unit_widget(rx_layout, 1, "RX Capacitance", self.cct_defaults["rx_capacitance"], "pF")
        rx_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(rx_group)

        transient_group = QGroupBox("Transient Settings")
        transient_layout = QGridLayout(transient_group)
        self.transient_step = add_unit_widget(transient_layout, 0, "Transient Step", self.cct_defaults["transient_step"], "ps")
        self.transient_stop = add_unit_widget(transient_layout, 1, "Transient Stop", self.cct_defaults["transient_stop"], "ns")
        transient_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(transient_group)

        options_group = QGroupBox("Options")
        options_layout = QGridLayout(options_group)
        self.aedt_version = QLineEdit(self.cct_defaults["aedt_version"])
        options_layout.addWidget(QLabel("AEDT Version"), 0, 0)
        options_layout.addWidget(self.aedt_version, 0, 1)
        self.threshold = add_unit_widget(options_layout, 1, "Threshold", self.cct_defaults["threshold"], "dB")
        options_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(options_group)

        config_buttons_layout = QVBoxLayout()
        self.save_config_button = QPushButton("Save Config")
        self.load_config_button = QPushButton("Load Config")
        self.reset_defaults_button = QPushButton("Reset Defaults")
        config_buttons_layout.addStretch()
        config_buttons_layout.addWidget(self.save_config_button)
        config_buttons_layout.addWidget(self.load_config_button)
        config_buttons_layout.addWidget(self.reset_defaults_button)
        config_panels_layout.addLayout(config_buttons_layout)
        cct_layout.addLayout(config_panels_layout)

        port_group = QGroupBox("Port Information")
        port_layout = QVBoxLayout(port_group)
        self.port_table = QTableWidget()
        self.port_table.setColumnCount(5)
        self.port_table.setHorizontalHeaderLabels(["", "TX Port", "RX Port", "Type", "Pair"])
        header = self.port_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.port_table.verticalHeader().setVisible(False)
        self.port_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.port_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        port_layout.addWidget(self.port_table)
        cct_layout.addWidget(port_group)

        action_buttons_layout = QHBoxLayout()
        action_buttons_layout.addStretch()
        self.prerun_button = QPushButton("Pre-run")
        self.calculate_button = QPushButton("Calculate")
        action_buttons_layout.addWidget(self.prerun_button)
        action_buttons_layout.addWidget(self.calculate_button)
        cct_layout.addLayout(action_buttons_layout)
