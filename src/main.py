import sys
import os
import json
import re
import subprocess
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
        self.load_config()

    def setup_import_tab(self):
        import_layout = QVBoxLayout(self.import_tab)
        import_group = QGroupBox("Layout Import")
        import_group_layout = QGridLayout(import_group)

        # Row 1: Layout type selection
        self.brd_radio = QRadioButton(".brd")
        self.aedb_radio = QRadioButton(".aedb")
        self.brd_radio.setChecked(True)
        self.brd_radio.toggled.connect(self.on_layout_type_changed)
        self.aedb_radio.toggled.connect(self.on_layout_type_changed)
        import_group_layout.addWidget(self.brd_radio, 0, 0)
        import_group_layout.addWidget(self.aedb_radio, 0, 1)

        # Row 2: Path selection
        self.open_layout_button = QPushButton("Open...")
        self.open_layout_button.clicked.connect(self.open_layout)
        self.layout_path_label = QLabel("No design loaded")
        import_group_layout.addWidget(QLabel("Design:"), 1, 0)
        import_group_layout.addWidget(self.layout_path_label, 1, 1)
        import_group_layout.addWidget(self.open_layout_button, 1, 2)

        # Row 3: Stackup selection
        self.stackup_path_input = QLineEdit()
        self.browse_stackup_button = QPushButton("Browse...")
        self.browse_stackup_button.clicked.connect(self.browse_stackup)
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
        self.apply_import_button.clicked.connect(lambda: self.run_get_edb(self.layout_path_label.text()))
        import_layout.addWidget(self.apply_import_button, alignment=Qt.AlignRight)

        import_layout.addStretch()

    def on_layout_type_changed(self):
        if self.sender().isChecked():
            self.layout_path_label.setText("No design loaded")
            self.stackup_path_input.clear()

    def closeEvent(self, event):
        self.save_config()
        super().closeEvent(event)

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    config = json.load(f)
                    self.edb_version_input.setText(config.get("edb_version", "2024.1"))
            else:
                self.edb_version_input.setText("2024.1")
        except (IOError, json.JSONDecodeError) as e:
            self.log(f"Could not load config file: {e}", "orange")
            self.edb_version_input.setText("2024.1")

    def save_config(self):
        try:
            config = {"edb_version": self.edb_version_input.text()}
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, "w") as f:
                json.dump(config, f, indent=2)
        except IOError as e:
            self.log(f"Could not save config file: {e}", "red")

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
        self.apply_import_button.setStyleSheet(primary_style)

        secondary_style = "background-color: #6c757d; color: white; border: none;"
        self.prerun_button.setStyleSheet(secondary_style)

    def setup_port_setup_tab(self):
        port_setup_layout = QVBoxLayout(self.port_setup_tab)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Component filter (regex):"))
        self.component_filter_input = QLineEdit("^[UJ]")
        self.component_filter_input.textChanged.connect(self.filter_components)
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

        self.controller_components_list.itemSelectionChanged.connect(self.update_nets)
        self.dram_components_list.itemSelectionChanged.connect(self.update_nets)

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

        self.single_ended_list.itemChanged.connect(self.update_checked_count)
        self.differential_pairs_list.itemChanged.connect(self.update_checked_count)

        self.apply_button = QPushButton("Apply")
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self.apply_settings)
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
        self.apply_simulation_button.clicked.connect(self.apply_simulation_settings)
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
        self.touchstone_path_input.textChanged.connect(self.check_paths_and_load_ports)
        self.port_metadata_path_input = QLineEdit()
        self.port_metadata_path_input.textChanged.connect(self.check_paths_and_load_ports)
        self.browse_touchstone_button = QPushButton("Browse")
        self.browse_touchstone_button.clicked.connect(self.browse_touchstone)
        self.browse_metadata_button = QPushButton("Browse")
        self.browse_metadata_button.clicked.connect(self.browse_port_metadata)
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
        self.save_config_button.clicked.connect(self.save_cct_config)
        self.load_config_button = QPushButton("Load Config")
        self.load_config_button.clicked.connect(self.load_cct_config)
        self.reset_defaults_button = QPushButton("Reset Defaults")
        self.reset_defaults_button.clicked.connect(self.reset_cct_defaults)
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
        self.prerun_button.clicked.connect(self.run_prerun)
        self.calculate_button.clicked.connect(self.run_calculate)
        action_buttons_layout.addWidget(self.prerun_button)
        action_buttons_layout.addWidget(self.calculate_button)
        cct_layout.addLayout(action_buttons_layout)

    def browse_touchstone(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Touchstone File", "", "Touchstone files (*.s*p)")
        if file_path:
            self.touchstone_path_input.setText(file_path)

    def browse_port_metadata(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Port Metadata File", "", "JSON files (*.json)")
        if file_path:
            self.port_metadata_path_input.setText(file_path)

    def check_paths_and_load_ports(self):
        touchstone_path = self.touchstone_path_input.text()
        metadata_path = self.port_metadata_path_input.text()
        if os.path.exists(touchstone_path) and os.path.exists(metadata_path):
            self.load_port_data(metadata_path)
        else:
            self.port_table.setRowCount(0)

    def load_port_data(self, metadata_path):
        try:
            with open(metadata_path, "r") as f:
                data = json.load(f)
            ports = data.get("ports", [])
            tx_ports, rx_ports = {}, {}
            for port in ports:
                role, net_type = port.get("component_role"), port.get("net_type")
                key = port.get("pair") if net_type == "differential" else port.get("net")
                if role == "controller":
                    if key not in tx_ports: tx_ports[key] = []
                    tx_ports[key].append(port)
                elif role == "dram":
                    if key not in rx_ports: rx_ports[key] = []
                    rx_ports[key].append(port)
            
            matched_ports = []
            for key, tx_port_list in tx_ports.items():
                if key in rx_ports:
                    rx_port_list = rx_ports[key]
                    if tx_port_list and rx_port_list:
                        matched_ports.append({
                            "tx": tx_port_list, "rx": rx_port_list,
                            "type": tx_port_list[0].get("net_type"),
                            "pair_name": tx_port_list[0].get("pair")
                        })
            
            self.port_table.setRowCount(len(matched_ports))
            for i, item in enumerate(matched_ports):
                tx_names = " / ".join(p['name'] for p in item['tx'])
                rx_names = " / ".join(p['name'] for p in item['rx'])
                port_type = item['type'].capitalize()
                pair_name = item['pair_name'] or f"M_DQ<{i}>"
                self.port_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                self.port_table.setItem(i, 1, QTableWidgetItem(tx_names))
                self.port_table.setItem(i, 2, QTableWidgetItem(rx_names))
                self.port_table.setItem(i, 3, QTableWidgetItem(port_type))
                self.port_table.setItem(i, 4, QTableWidgetItem(pair_name))
            self.log(f"Loaded {len(ports)} ports from {os.path.basename(metadata_path)}")
        except Exception as e:
            self.log(f"Error loading port data: {e}", color="red")

    def save_cct_config(self):
        config_data = {
            "tx_vhigh": self.tx_vhigh.text(),
            "tx_rise_time": self.tx_rise_time.text(),
            "tx_unit_interval": self.tx_unit_interval.text(),
            "tx_resistance": self.tx_resistance.text(),
            "tx_capacitance": self.tx_capacitance.text(),
            "rx_resistance": self.rx_resistance.text(),
            "rx_capacitance": self.rx_capacitance.text(),
            "transient_step": self.transient_step.text(),
            "transient_stop": self.transient_stop.text(),
            "aedt_version": self.aedt_version.text(),
            "threshold": self.threshold.text(),
        }
        file_path, _ = QFileDialog.getSaveFileName(self, "Save CCT Config", "", "JSON files (*.json)")
        if file_path:
            try:
                with open(file_path, "w") as f:
                    json.dump(config_data, f, indent=2)
                self.log(f"CCT configuration saved to {file_path}")
            except Exception as e:
                self.log(f"Error saving CCT config: {e}", color="red")

    def load_cct_config(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Load CCT Config", "", "JSON files (*.json)")
        if file_path:
            try:
                with open(file_path, "r") as f:
                    config_data = json.load(f)
                
                self.tx_vhigh.setText(config_data.get("tx_vhigh", ""))
                self.tx_rise_time.setText(config_data.get("tx_rise_time", ""))
                self.tx_unit_interval.setText(config_data.get("tx_unit_interval", ""))
                self.tx_resistance.setText(config_data.get("tx_resistance", ""))
                self.tx_capacitance.setText(config_data.get("tx_capacitance", ""))
                self.rx_resistance.setText(config_data.get("rx_resistance", ""))
                self.rx_capacitance.setText(config_data.get("rx_capacitance", ""))
                self.transient_step.setText(config_data.get("transient_step", ""))
                self.transient_stop.setText(config_data.get("transient_stop", ""))
                self.aedt_version.setText(config_data.get("aedt_version", ""))
                self.threshold.setText(config_data.get("threshold", ""))
                
                self.log(f"CCT configuration loaded from {file_path}")
            except Exception as e:
                self.log(f"Error loading CCT config: {e}", color="red")

    def reset_cct_defaults(self):
        self.tx_vhigh.setText(self.cct_defaults["tx_vhigh"])
        self.tx_rise_time.setText(self.cct_defaults["tx_rise_time"])
        self.tx_unit_interval.setText(self.cct_defaults["tx_unit_interval"])
        self.tx_resistance.setText(self.cct_defaults["tx_resistance"])
        self.tx_capacitance.setText(self.cct_defaults["tx_capacitance"])
        self.rx_resistance.setText(self.cct_defaults["rx_resistance"])
        self.rx_capacitance.setText(self.cct_defaults["rx_capacitance"])
        self.transient_step.setText(self.cct_defaults["transient_step"])
        self.transient_stop.setText(self.cct_defaults["transient_stop"])
        self.aedt_version.setText(self.cct_defaults["aedt_version"])
        self.threshold.setText(self.cct_defaults["threshold"])
        self.log("CCT settings reset to defaults.")

    def get_cct_settings(self):
        return {
            "tx": {
                "vhigh": self.tx_vhigh.text() + "V", "t_rise": self.tx_rise_time.text() + "ps",
                "ui": self.tx_unit_interval.text() + "ps", "res_tx": self.tx_resistance.text() + "ohm",
                "cap_tx": self.tx_capacitance.text() + "pF",
            },
            "rx": {
                "res_rx": self.rx_resistance.text() + "ohm", "cap_rx": self.rx_capacitance.text() + "pF",
            },
            "run": {
                "tstep": self.transient_step.text() + "ps", "tstop": self.transient_stop.text() + "ns",
            },
            "options": {
                "circuit_version": self.aedt_version.text(), "threshold_db": self.threshold.text(),
            },
        }

    def run_cct_process(self, mode):
        touchstone_path = self.touchstone_path_input.text()
        metadata_path = self.port_metadata_path_input.text()
        if not (os.path.exists(touchstone_path) and os.path.exists(metadata_path)):
            self.log("Please provide valid paths for Touchstone and Port Metadata files.", color="red")
            return

        self.log(f"Starting CCT {mode}...")
        self.prerun_button.setEnabled(False)
        self.calculate_button.setEnabled(False)
        if mode == 'run':
            self.calculate_button.setText("Running")
            self.calculate_button.setStyleSheet("background-color: yellow; color: black;")

        script_path = os.path.join(os.path.dirname(__file__), "cct_runner.py")
        python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
        workdir = os.path.join(os.path.dirname(metadata_path), "cct_work")
        settings_str = json.dumps(self.get_cct_settings())
        command = [
            python_executable, script_path, "--touchstone-path", touchstone_path,
            "--metadata-path", metadata_path, "--workdir", workdir,
            "--settings", settings_str, "--mode", mode,
        ]
        if mode == 'run':
            output_path = os.path.join(os.path.dirname(metadata_path), "cct_results.csv")
            command.extend(["--output-path", output_path])

        self.process = QProcess()
        self.process.readyReadStandardOutput.connect(self.handle_stdout)
        self.process.readyReadStandardError.connect(self.handle_stderr)
        self.process.finished.connect(self.cct_finished)
        self.process.start(command[0], command[1:])

    def handle_stdout(self):
        data = self.process.readAllStandardOutput().data().decode().strip()
        for line in data.splitlines(): self.log(line)

    def handle_stderr(self):
        data = self.process.readAllStandardError().data().decode().strip()
        for line in data.splitlines(): self.log(line, color="red")

    def cct_finished(self):
        self.log("CCT process finished.")
        self.prerun_button.setEnabled(True)
        self.calculate_button.setEnabled(True)
        self.calculate_button.setText("Calculate")
        self.calculate_button.setStyleSheet(self.calculate_button_original_style)

    def run_prerun(self): self.run_cct_process("prerun")
    def run_calculate(self): self.run_cct_process("run")

    def apply_simulation_settings(self):
        aedb_path = self.layout_path_label.text()
        if not os.path.isdir(aedb_path):
            self.log("Please open an .aedb project first.", "red")
            return

        output_dir = os.path.dirname(aedb_path)
        file_path = os.path.join(output_dir, "simulation.json")

        sweeps = []
        for row in range(self.sweeps_table.rowCount()):
            sweep_type = self.sweeps_table.cellWidget(row, 0).currentText()
            start = self.sweeps_table.item(row, 1).text()
            stop = self.sweeps_table.item(row, 2).text()
            step = self.sweeps_table.item(row, 3).text()
            sweeps.append([sweep_type, start, stop, step])

        settings = {
            "aedb_path": aedb_path,
            "edb_version": self.edb_version_input.text(),
            "cutout": {
                "enabled": self.enable_cutout_checkbox.isChecked(),
                "expansion_size": self.expansion_size_input.text(),
                "signal_nets": self.signal_nets_label.text().split(", "),
                "reference_net": self.reference_net_label.text(),
            },
            "solver": "SIwave" if self.siwave_radio.isChecked() else "HFSS",
            "frequency_sweeps": sweeps,
        }

        try:
            with open(file_path, "w") as f:
                json.dump(settings, f, indent=2)
            self.log(f"Simulation settings saved to {file_path}")

            self.log("Applying simulation settings to EDB...")
            script_path = os.path.join(os.path.dirname(__file__), "set_sim.py")
            python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
            command = [python_executable, script_path, file_path]
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            stdout, stderr = process.communicate()

            if process.returncode == 0:
                self.log("Successfully applied simulation settings.")
                if stdout: self.log(stdout)
            else:
                self.log(f"Error running set_sim.py: {stderr.strip()}", "red")
        except Exception as e:
            self.log(f"Error applying simulation settings: {e}", color="red")

    def log(self, message, color=None):
        if color: self.log_window.setTextColor(QColor(color))
        self.log_window.append(message)
        self.log_window.setTextColor(QColor("black"))
        self.log_window.verticalScrollBar().setValue(self.log_window.verticalScrollBar().maximum())

    def filter_components(self):
        pattern = self.component_filter_input.text()
        try: regex = re.compile(pattern)
        except re.error: return
        self.controller_components_list.clear()
        self.dram_components_list.clear()
        for comp_name, pin_count in self.all_components:
            if regex.search(comp_name):
                item_text = f"{comp_name} ({pin_count})"
                self.controller_components_list.addItem(item_text)
                self.dram_components_list.addItem(item_text)

    def update_checked_count(self):
        checked_single = sum(1 for i in range(self.single_ended_list.count()) if self.single_ended_list.item(i).checkState() == Qt.Checked)
        checked_diff = sum(1 for i in range(self.differential_pairs_list.count()) if self.differential_pairs_list.item(i).checkState() == Qt.Checked)
        checked_nets = checked_single + (checked_diff * 2)
        ports = (checked_single * 2) + (checked_diff * 4)
        self.checked_nets_label.setText(f"Checked nets: {checked_nets} | Ports: {ports}")
        self.apply_button.setEnabled(checked_nets > 0)

    def apply_settings(self):
        if not self.pcb_data:
            self.log("No PCB data loaded.", "red")
            return
        aedb_path = self.layout_path_label.text()
        if not os.path.isdir(aedb_path):
            self.log("Invalid AEDB path.", "red")
            return
        output_path = os.path.join(os.path.dirname(aedb_path), "ports.json")
        data = {
            "aedb_path": aedb_path, "reference_net": self.ref_net_combo.currentText(),
            "controller_components": [item.text().split(" ")[0] for item in self.controller_components_list.selectedItems()],
            "dram_components": [item.text().split(" ")[0] for item in self.dram_components_list.selectedItems()],
            "ports": [],
        }
        sequence = 1
        diff_pairs_info = self.pcb_data.get("diff", {})
        net_to_diff_pair = {p_net: (pair_name, "positive") for pair_name, (p_net, n_net) in diff_pairs_info.items()}
        net_to_diff_pair.update({n_net: (pair_name, "negative") for pair_name, (p_net, n_net) in diff_pairs_info.items()})

        signal_nets = []
        for i in range(self.single_ended_list.count()):
            item = self.single_ended_list.item(i)
            if item.checkState() == Qt.Checked:
                net_name = item.text()
                signal_nets.append(net_name)
                for comp in data["controller_components"]:
                    if any(pin[1] == net_name for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{net_name}", "component": comp, "component_role": "controller", "net": net_name, "net_type": "single", "pair": None, "polarity": None, "reference_net": data["reference_net"]}); sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == net_name for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{net_name}", "component": comp, "component_role": "dram", "net": net_name, "net_type": "single", "pair": None, "polarity": None, "reference_net": data["reference_net"]}); sequence += 1
        
        for i in range(self.differential_pairs_list.count()):
            item = self.differential_pairs_list.item(i)
            if item.checkState() == Qt.Checked:
                pair_name = item.text()
                p_net, n_net = diff_pairs_info[pair_name]
                signal_nets.extend([p_net, n_net])
                for comp in data["controller_components"]:
                    if any(pin[1] == p_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{p_net}", "component": comp, "component_role": "controller", "net": p_net, "net_type": "differential", "pair": pair_name, "polarity": "positive", "reference_net": data["reference_net"]}); sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == p_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{p_net}", "component": comp, "component_role": "dram", "net": p_net, "net_type": "differential", "pair": pair_name, "polarity": "positive", "reference_net": data["reference_net"]}); sequence += 1
                for comp in data["controller_components"]:
                    if any(pin[1] == n_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{n_net}", "component": comp, "component_role": "controller", "net": n_net, "net_type": "differential", "pair": pair_name, "polarity": "negative", "reference_net": data["reference_net"]}); sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == n_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({"sequence": sequence, "name": f"{sequence}_{comp}_{n_net}", "component": comp, "component_role": "dram", "net": n_net, "net_type": "differential", "pair": pair_name, "polarity": "negative", "reference_net": data["reference_net"]}); sequence += 1
        
        self.signal_nets_label.setText(", ".join(sorted(signal_nets)))
        self.reference_net_label.setText(data["reference_net"])

        try:
            with open(output_path, "w") as f: json.dump(data, f, indent=2)
            self.log(f"Successfully saved to {output_path}. Now applying to EDB...")
            script_path = os.path.join(os.path.dirname(__file__), "set_edb.py")
            python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
            edb_version = self.edb_version_input.text()
            command = [python_executable, script_path, output_path, edb_version]
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            stdout, stderr = process.communicate()
            if process.returncode == 0:
                new_aedb_path = aedb_path.replace('.aedb', '_applied.aedb')
                self.log(f"Successfully created {new_aedb_path}")
                if stdout: self.log(stdout)
            else:
                self.log(f"Error running set_edb.py: {stderr.strip()}", "red")
        except Exception as e:
            self.log(f"Error during apply: {e}", "red")

    def open_layout(self):
        path = ""
        if self.brd_radio.isChecked():
            path, _ = QFileDialog.getOpenFileName(self, "Select .brd file", "", "BRD files (*.brd)")
        elif self.aedb_radio.isChecked():
            path = QFileDialog.getExistingDirectory(self, "Select .aedb directory", ".", QFileDialog.ShowDirsOnly)

        if path:
            self.layout_path_label.setText(path)

    def browse_stackup(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Stackup File", "", "XML files (*.xml)")
        if file_path:
            self.stackup_path_input.setText(file_path)

    def run_get_edb(self, layout_path):
        self.log(f"Opening layout: {layout_path}")
        script_path = os.path.join(os.path.dirname(__file__), "get_edb.py")
        python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
        edb_version = self.edb_version_input.text()
        
        command = [python_executable, script_path, layout_path, edb_version]
        
        stackup_path = self.stackup_path_input.text()
        if stackup_path and os.path.exists(stackup_path):
            command.append(stackup_path)

        self.log(f"Running command: {' '.join(command)}")
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            stdout, stderr = process.communicate()
            
            if layout_path.endswith('.aedb'):
                json_output_path = layout_path.replace('.aedb', '.json')
            else:
                json_output_path = os.path.splitext(layout_path)[0] + '.json'

            if process.returncode == 0:
                self.log(f"Successfully generated {os.path.basename(json_output_path)}")
                if stdout: self.log(stdout)
                self.load_pcb_data(json_output_path)
            else:
                self.log(f"Error running get_edb.py: {stderr.strip()}", "red")
        except Exception as e:
            self.log(f"Failed to execute get_edb.py: {e}", "red")

    def load_pcb_data(self, json_path):
        try:
            with open(json_path, "r") as f: self.pcb_data = json.load(f)
            self.controller_components_list.clear()
            self.dram_components_list.clear()
            if "component" in self.pcb_data:
                self.all_components = [(name, len(pins)) for name, pins in self.pcb_data["component"].items()]
                self.all_components.sort(key=lambda x: x[1], reverse=True)
                self.filter_components()
        except FileNotFoundError: self.log("pcb.json not found.", "red"); self.pcb_data = None
        except json.JSONDecodeError: self.log("Error decoding pcb.json.", "red"); self.pcb_data = None
        except Exception as e: self.log(f"Error loading data: {e}", "red"); self.pcb_data = None

    def update_nets(self):
        if not self.pcb_data or "component" not in self.pcb_data: return
        selected_controllers = [item.text().split(" ")[0] for item in self.controller_components_list.selectedItems()]
        selected_drams = [item.text().split(" ")[0] for item in self.dram_components_list.selectedItems()]
        self.single_ended_list.clear()
        self.differential_pairs_list.clear()
        self.ref_net_combo.clear()
        if not selected_controllers or not selected_drams:
            self.ref_net_combo.addItem("GND")
            self.update_checked_count()
            return
        controller_nets = set(pin[1] for comp in selected_controllers for pin in self.pcb_data["component"].get(comp, []))
        dram_nets = set(pin[1] for comp in selected_drams for pin in self.pcb_data["component"].get(comp, []))
        common_nets = controller_nets.intersection(dram_nets)
        selected_components = selected_controllers + selected_drams
        net_pin_counts = {net: sum(1 for comp_name in selected_components for pin in self.pcb_data["component"].get(comp_name, []) if pin[1] == net) for net in common_nets}
        sorted_nets = sorted(net_pin_counts.items(), key=lambda item: item[1], reverse=True)
        for net_name, count in sorted_nets: self.ref_net_combo.addItem(net_name)
        if sorted_nets: self.ref_net_combo.setCurrentIndex(0)
        diff_pairs_info = self.pcb_data.get("diff", {})
        diff_pair_nets = {net for pos_net, neg_net in diff_pairs_info.values() for net in (pos_net, neg_net)}
        single_nets = sorted([net for net in common_nets if net not in diff_pair_nets and net.upper() != "GND"])
        for net in single_nets:
            item = QListWidgetItem(net)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.single_ended_list.addItem(item)
        for pair_name, (pos_net, neg_net) in sorted(diff_pairs_info.items()):
            if pos_net in common_nets and neg_net in common_nets:
                item = QListWidgetItem(pair_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.differential_pairs_list.addItem(item)
        self.update_checked_count()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AEDBCCTCalculator()
    window.show()
    sys.exit(app.exec())
