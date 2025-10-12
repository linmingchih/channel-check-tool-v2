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

        # Determine the target state based on the first selected item
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
        self.setGeometry(100, 100, 1200, 800)  # Adjusted window size
        self.pcb_data = None
        self.all_components = []

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # Top panel
        top_panel_layout = QHBoxLayout()
        self.edb_version_combo = QComboBox()
        self.edb_version_combo.addItems(["2024.1"])
        self.open_aedb_button = QPushButton("Open .aedb?")
        self.open_aedb_button.clicked.connect(self.open_aedb)
        self.aedb_path_label = QLabel("No design loaded")
        top_panel_layout.addWidget(QLabel("EDB version:"))
        top_panel_layout.addWidget(self.edb_version_combo)
        top_panel_layout.addWidget(self.open_aedb_button)
        top_panel_layout.addWidget(self.aedb_path_label)
        top_panel_layout.addStretch()
        main_layout.addLayout(top_panel_layout)

        # Tabs
        self.tabs = QTabWidget()
        self.port_setup_tab = QWidget()
        self.simulation_tab = QWidget()
        self.cct_tab = QWidget()
        self.tabs.addTab(self.port_setup_tab, "Port Setup")
        self.tabs.addTab(self.simulation_tab, "Simulation")
        self.tabs.addTab(self.cct_tab, "CCT")
        main_layout.addWidget(self.tabs)

        self.setup_port_setup_tab()
        self.setup_cct_tab()
        self.apply_styles()

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(
            "Controllers: 0 | DRAMs: 0 | Shared nets: 0 | Shared differential pairs: 0"
        )

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
        """)
        # Primary action buttons
        primary_style = "background-color: #007bff; color: white; border: none;"
        self.calculate_button.setStyleSheet(primary_style)
        self.apply_button.setStyleSheet(primary_style)

        # Secondary action buttons
        secondary_style = "background-color: #6c757d; color: white; border: none;"
        self.prerun_button.setStyleSheet(secondary_style)
        
    def setup_port_setup_tab(self):
        # Port Setup Tab Layout
        port_setup_layout = QVBoxLayout(self.port_setup_tab)

        # Component filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Component filter (regex):"))
        self.component_filter_input = QLineEdit("^[UJ]")
        self.component_filter_input.textChanged.connect(self.filter_components)
        filter_layout.addWidget(self.component_filter_input)
        port_setup_layout.addLayout(filter_layout)

        # Components panel
        components_layout = QHBoxLayout()
        self.controller_components_list = QListWidget()
        self.controller_components_list.setSelectionMode(
            QAbstractItemView.ExtendedSelection
        )
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

        # Connect selection changed signals
        self.controller_components_list.itemSelectionChanged.connect(self.update_nets)
        self.dram_components_list.itemSelectionChanged.connect(self.update_nets)

        # Reference net
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

        # Nets panel
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

        # Connect item changed signals
        self.single_ended_list.itemChanged.connect(self.update_checked_count)
        self.differential_pairs_list.itemChanged.connect(self.update_checked_count)

        # Apply button
        self.apply_button = QPushButton("Apply")
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self.apply_settings)
        port_setup_layout.addWidget(self.apply_button, alignment=Qt.AlignRight)

    def setup_cct_tab(self):
        cct_layout = QVBoxLayout(self.cct_tab)

        # File Inputs
        file_input_layout = QGridLayout()
        self.touchstone_path_input = QLineEdit()
        self.touchstone_path_input.textChanged.connect(self.check_paths_and_load_ports)
        self.port_metadata_path_input = QLineEdit()
        self.port_metadata_path_input.textChanged.connect(
            self.check_paths_and_load_ports
        )
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

        # Configuration Panels
        config_panels_layout = QHBoxLayout()

        def add_unit_widget(layout, row, label, value, unit):
            layout.addWidget(QLabel(label), row, 0)
            widget_layout = QHBoxLayout()
            line_edit = QLineEdit(value)
            widget_layout.addWidget(line_edit)
            widget_layout.addWidget(QLabel(unit))
            layout.addLayout(widget_layout, row, 1)
            return line_edit

        # TX Settings
        tx_group = QGroupBox("TX Settings")
        tx_layout = QGridLayout(tx_group)
        self.tx_vhigh = add_unit_widget(tx_layout, 0, "TX Vhigh", "0.800", "V")
        self.tx_rise_time = add_unit_widget(tx_layout, 1, "TX Rise Time", "30.000", "ps")
        self.tx_unit_interval = add_unit_widget(tx_layout, 2, "Unit Interval", "133.000", "ps")
        self.tx_resistance = add_unit_widget(tx_layout, 3, "TX Resistance", "40.000", "ohm")
        self.tx_capacitance = add_unit_widget(tx_layout, 4, "TX Capacitance", "1.000", "pF")
        tx_layout.setRowStretch(5, 1)
        config_panels_layout.addWidget(tx_group)

        # RX Settings
        rx_group = QGroupBox("RX Settings")
        rx_layout = QGridLayout(rx_group)
        self.rx_resistance = add_unit_widget(rx_layout, 0, "RX Resistance", "30.000", "ohm")
        self.rx_capacitance = add_unit_widget(rx_layout, 1, "RX Capacitance", "1.800", "pF")
        rx_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(rx_group)

        # Transient Settings
        transient_group = QGroupBox("Transient Settings")
        transient_layout = QGridLayout(transient_group)
        self.transient_step = add_unit_widget(transient_layout, 0, "Transient Step", "100.000", "ps")
        self.transient_stop = add_unit_widget(transient_layout, 1, "Transient Stop", "3.000", "ns")
        transient_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(transient_group)

        # Options
        options_group = QGroupBox("Options")
        options_layout = QGridLayout(options_group)
        self.aedt_version = QLineEdit("2025.2")
        options_layout.addWidget(QLabel("AEDT Version"), 0, 0)
        options_layout.addWidget(self.aedt_version, 0, 1)
        self.threshold = add_unit_widget(options_layout, 1, "Threshold", "-40.0", "dB")
        options_layout.setRowStretch(2, 1)
        config_panels_layout.addWidget(options_group)

        # Config buttons
        config_buttons_layout = QVBoxLayout()
        self.save_config_button = QPushButton("Save Config")
        self.load_config_button = QPushButton("Load Config")
        self.reset_defaults_button = QPushButton("Reset Defaults")
        config_buttons_layout.addWidget(self.save_config_button)
        config_buttons_layout.addWidget(self.load_config_button)
        config_buttons_layout.addWidget(self.reset_defaults_button)
        config_buttons_layout.addStretch()
        config_panels_layout.addLayout(config_buttons_layout)

        cct_layout.addLayout(config_panels_layout)

        # Port Table
        self.port_table = QTableWidget()
        self.port_table.setColumnCount(5)
        self.port_table.setHorizontalHeaderLabels(
            ["", "TX Port", "RX Port", "Type", "Pair"]
        )
        self.port_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )
        self.port_table.verticalHeader().setVisible(False)
        self.port_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.port_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        cct_layout.addWidget(self.port_table)

        # Action Buttons
        action_buttons_layout = QHBoxLayout()
        action_buttons_layout.addStretch()
        self.prerun_button = QPushButton("Pre-run")
        self.calculate_button = QPushButton("Calculate")
        self.prerun_button.clicked.connect(self.run_prerun)
        self.calculate_button.clicked.connect(self.run_calculate)
        action_buttons_layout.addWidget(self.prerun_button)
        action_buttons_layout.addWidget(self.calculate_button)
        cct_layout.addLayout(action_buttons_layout)

        # Log Window
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        cct_layout.addWidget(self.log_window)

    def browse_touchstone(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Touchstone File", "", "Touchstone files (*.s*p)"
        )
        if file_path:
            self.touchstone_path_input.setText(file_path)

    def browse_port_metadata(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Port Metadata File", "", "JSON files (*.json)"
        )
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
            
            # Group ports by pair for differential or by net for single
            tx_ports = {}
            rx_ports = {}

            for port in ports:
                role = port.get("component_role")
                net_type = port.get("net_type")
                key = port.get("pair") if net_type == "differential" else port.get("net")
                
                if role == "controller":
                    if key not in tx_ports:
                        tx_ports[key] = []
                    tx_ports[key].append(port)
                elif role == "dram":
                    if key not in rx_ports:
                        rx_ports[key] = []
                    rx_ports[key].append(port)

            # Create a list of matched pairs/nets
            matched_ports = []
            for key, tx_port_list in tx_ports.items():
                if key in rx_ports:
                    rx_port_list = rx_ports[key]
                    # Simple 1-to-1 matching for now
                    if tx_port_list and rx_port_list:
                        matched_ports.append({
                            "tx": tx_port_list,
                            "rx": rx_port_list,
                            "type": tx_port_list[0].get("net_type"),
                            "pair_name": tx_port_list[0].get("pair")
                        })
            
            self.port_table.setRowCount(len(matched_ports))
            for i, item in enumerate(matched_ports):
                tx_names = " / ".join(p['name'] for p in item['tx'])
                rx_names = " / ".join(p['name'] for p in item['rx'])
                port_type = item['type'].capitalize()
                pair_name = item['pair_name'] or f"M_DQ<{i}>" # Placeholder

                self.port_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                self.port_table.setItem(i, 1, QTableWidgetItem(tx_names))
                self.port_table.setItem(i, 2, QTableWidgetItem(rx_names))
                self.port_table.setItem(i, 3, QTableWidgetItem(port_type))
                self.port_table.setItem(i, 4, QTableWidgetItem(pair_name))

            self.log(f"Loaded {len(ports)} ports from {os.path.basename(metadata_path)}")

        except Exception as e:
            self.log(f"Error loading port data: {e}", color="red")

    def get_cct_settings(self):
        return {
            "tx": {
                "vhigh": self.tx_vhigh.text() + "V",
                "t_rise": self.tx_rise_time.text() + "ps",
                "ui": self.tx_unit_interval.text() + "ps",
                "res_tx": self.tx_resistance.text() + "ohm",
                "cap_tx": self.tx_capacitance.text() + "pF",
            },
            "rx": {
                "res_rx": self.rx_resistance.text() + "ohm",
                "cap_rx": self.rx_capacitance.text() + "pF",
            },
            "run": {
                "tstep": self.transient_step.text() + "ps",
                "tstop": self.transient_stop.text() + "ns",
            },
            "options": {
                "circuit_version": self.aedt_version.text(),
                "threshold_db": self.threshold.text(),
            },
        }

    def run_cct_process(self, mode):
        touchstone_path = self.touchstone_path_input.text()
        metadata_path = self.port_metadata_path_input.text()

        if not os.path.exists(touchstone_path) or not os.path.exists(metadata_path):
            self.log("Please provide valid paths for Touchstone and Port Metadata files.", color="red")
            return

        self.log(f"Starting CCT {mode}...")
        self.prerun_button.setEnabled(False)
        self.calculate_button.setEnabled(False)

        script_path = os.path.join(os.path.dirname(__file__), "cct_runner.py")
        python_executable = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe"
        )
        workdir = os.path.join(os.path.dirname(metadata_path), "cct_work")
        
        settings_str = json.dumps(self.get_cct_settings())

        command = [
            python_executable,
            script_path,
            "--touchstone-path", touchstone_path,
            "--metadata-path", metadata_path,
            "--workdir", workdir,
            "--settings", settings_str,
            "--mode", mode,
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
        for line in data.splitlines():
            self.log(line)

    def handle_stderr(self):
        data = self.process.readAllStandardError().data().decode().strip()
        for line in data.splitlines():
            self.log(line, color="red")

    def cct_finished(self):
        self.log("CCT process finished.")
        self.prerun_button.setEnabled(True)
        self.calculate_button.setEnabled(True)

    def run_prerun(self):
        self.run_cct_process("prerun")

    def run_calculate(self):
        self.run_cct_process("run")

    def log(self, message, color=None):
        if color:
            self.log_window.setTextColor(QColor(color))
        self.log_window.append(message)
        self.log_window.setTextColor(QColor("black")) # Reset to default
        self.log_window.verticalScrollBar().setValue(self.log_window.verticalScrollBar().maximum())

    def filter_components(self):
        pattern = self.component_filter_input.text()
        try:
            regex = re.compile(pattern)
        except re.error:
            # Invalid regex, do nothing
            return

        self.controller_components_list.clear()
        self.dram_components_list.clear()

        for comp_name, pin_count in self.all_components:
            if regex.search(comp_name):
                item_text = f"{comp_name} ({pin_count})"
                self.controller_components_list.addItem(item_text)
                self.dram_components_list.addItem(item_text)

    def update_checked_count(self):
        checked_single = sum(
            1
            for i in range(self.single_ended_list.count())
            if self.single_ended_list.item(i).checkState() == Qt.Checked
        )
        checked_diff = sum(
            1
            for i in range(self.differential_pairs_list.count())
            if self.differential_pairs_list.item(i).checkState() == Qt.Checked
        )

        checked_nets = checked_single + (checked_diff * 2)
        ports = (checked_single * 2) + (checked_diff * 4)

        self.checked_nets_label.setText(
            f"Checked nets: {checked_nets} | Ports: {ports}"
        )
        self.apply_button.setEnabled(checked_nets > 0)

    def apply_settings(self):
        if not self.pcb_data:
            self.status_bar.showMessage("No PCB data loaded.")
            return

        aedb_path = self.aedb_path_label.text()
        if not os.path.isdir(aedb_path):
            self.status_bar.showMessage("Invalid AEDB path.")
            return

        output_path = os.path.join(os.path.dirname(aedb_path), "ports.json")

        data = {
            "aedb_path": aedb_path,
            "reference_net": self.ref_net_combo.currentText(),
            "controller_components": [
                item.text().split(" ")[0]
                for item in self.controller_components_list.selectedItems()
            ],
            "dram_components": [
                item.text().split(" ")[0]
                for item in self.dram_components_list.selectedItems()
            ],
            "ports": [],
        }

        sequence = 1
        diff_pairs_info = self.pcb_data.get("diff", {})
        
        # Invert the diff_pairs_info for easier lookup
        net_to_diff_pair = {}
        for pair_name, (p_net, n_net) in diff_pairs_info.items():
            net_to_diff_pair[p_net] = (pair_name, "positive")
            net_to_diff_pair[n_net] = (pair_name, "negative")

        # Process single-ended nets
        for i in range(self.single_ended_list.count()):
            item = self.single_ended_list.item(i)
            if item.checkState() == Qt.Checked:
                net_name = item.text()
                for comp in data["controller_components"]:
                    if any(pin[1] == net_name for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence,
                            "name": f"{sequence}_{comp}_{net_name}",
                            "component": comp,
                            "component_role": "controller",
                            "net": net_name, "net_type": "single", "pair": None, "polarity": None,
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == net_name for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence,
                            "name": f"{sequence}_{comp}_{net_name}",
                            "component": comp,
                            "component_role": "dram",
                            "net": net_name, "net_type": "single", "pair": None, "polarity": None,
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1
        
        # Process differential pairs
        for i in range(self.differential_pairs_list.count()):
            item = self.differential_pairs_list.item(i)
            if item.checkState() == Qt.Checked:
                pair_name = item.text()
                p_net, n_net = diff_pairs_info[pair_name]
                
                # Positive net
                for comp in data["controller_components"]:
                    if any(pin[1] == p_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence, "name": f"{sequence}_{comp}_{p_net}",
                            "component": comp, "component_role": "controller", "net": p_net,
                            "net_type": "differential", "pair": pair_name, "polarity": "positive",
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == p_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence, "name": f"{sequence}_{comp}_{p_net}",
                            "component": comp, "component_role": "dram", "net": p_net,
                            "net_type": "differential", "pair": pair_name, "polarity": "positive",
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1
                
                # Negative net
                for comp in data["controller_components"]:
                    if any(pin[1] == n_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence, "name": f"{sequence}_{comp}_{n_net}",
                            "component": comp, "component_role": "controller", "net": n_net,
                            "net_type": "differential", "pair": pair_name, "polarity": "negative",
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1
                for comp in data["dram_components"]:
                    if any(pin[1] == n_net for pin in self.pcb_data["component"].get(comp, [])):
                        data["ports"].append({
                            "sequence": sequence, "name": f"{sequence}_{comp}_{n_net}",
                            "component": comp, "component_role": "dram", "net": n_net,
                            "net_type": "differential", "pair": pair_name, "polarity": "negative",
                            "reference_net": data["reference_net"]
                        })
                        sequence += 1

        try:
            with open(output_path, "w") as f:
                json.dump(data, f, indent=2)
            self.status_bar.showMessage(f"Successfully saved to {output_path}. Now applying to EDB...")

            # Run set_edb.py
            script_path = os.path.join(os.path.dirname(__file__), "set_edb.py")
            python_executable = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                ".venv",
                "Scripts",
                "python.exe",
            )
            command = [python_executable, script_path, output_path]
            
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            stdout, stderr = process.communicate()

            if process.returncode == 0:
                new_aedb_path = aedb_path.replace('.aedb', '_applied.aedb')
                self.status_bar.showMessage(f"Successfully created {new_aedb_path}")
            else:
                self.status_bar.showMessage(f"Error running set_edb.py: {stderr.strip()}")

        except Exception as e:
            self.status_bar.showMessage(f"Error during apply: {e}")

    def open_aedb(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Select .aedb directory", ".", QFileDialog.ShowDirsOnly
        )
        if dir_path and dir_path.endswith(".aedb"):
            self.aedb_path_label.setText(dir_path)
            self.run_get_edb(dir_path)

    def run_get_edb(self, aedb_path):
        script_path = os.path.join(os.path.dirname(__file__), "get_edb.py")
        json_output_path = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "data", "pcb.json"
        )
        python_executable = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            ".venv",
            "Scripts",
            "python.exe",
        )

        command = [python_executable, script_path, aedb_path, json_output_path]

        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            stdout, stderr = process.communicate()
            if process.returncode == 0:
                self.status_bar.showMessage("Successfully generated pcb.json")
                self.load_pcb_data(json_output_path)
            else:
                self.status_bar.showMessage(
                    f"Error running get_edb.py: {stderr.strip()}"
                )
        except Exception as e:
            self.status_bar.showMessage(f"Failed to execute get_edb.py: {e}")

    def load_pcb_data(self, json_path):
        try:
            with open(json_path, "r") as f:
                self.pcb_data = json.load(f)

            self.controller_components_list.clear()
            self.dram_components_list.clear()

            if "component" in self.pcb_data:
                self.all_components = [
                    (name, len(pins))
                    for name, pins in self.pcb_data["component"].items()
                ]
                self.all_components.sort(key=lambda x: x[1], reverse=True)
                self.filter_components()

        except FileNotFoundError:
            self.status_bar.showMessage("pcb.json not found.")
            self.pcb_data = None
        except json.JSONDecodeError:
            self.status_bar.showMessage("Error decoding pcb.json.")
            self.pcb_data = None
        except Exception as e:
            self.status_bar.showMessage(f"Error loading data: {e}")
            self.pcb_data = None

    def update_nets(self):
        if not self.pcb_data or "component" not in self.pcb_data:
            return

        selected_controllers = [
            item.text().split(" ")[0]
            for item in self.controller_components_list.selectedItems()
        ]
        selected_drams = [
            item.text().split(" ")[0]
            for item in self.dram_components_list.selectedItems()
        ]

        self.single_ended_list.clear()
        self.differential_pairs_list.clear()
        self.ref_net_combo.clear()

        if not selected_controllers or not selected_drams:
            self.ref_net_combo.addItem("GND")
            self.update_checked_count()
            return

        # Get all nets for selected controllers and drams
        controller_nets = set()
        for comp in selected_controllers:
            controller_nets.update(
                pin[1] for pin in self.pcb_data["component"].get(comp, [])
            )

        dram_nets = set()
        for comp in selected_drams:
            dram_nets.update(pin[1] for pin in self.pcb_data["component"].get(comp, []))

        # Find the intersection of nets
        common_nets = controller_nets.intersection(dram_nets)
        selected_components = selected_controllers + selected_drams

        # Calculate total pin counts for each common net across selected components
        net_pin_counts = {}
        for net in common_nets:
            count = 0
            for comp_name in selected_components:
                count += sum(
                    1
                    for pin in self.pcb_data["component"].get(comp_name, [])
                    if pin[1] == net
                )
            net_pin_counts[net] = count

        sorted_nets = sorted(
            net_pin_counts.items(), key=lambda item: item[1], reverse=True
        )

        # Populate ref_net_combo
        for net_name, count in sorted_nets:
            self.ref_net_combo.addItem(net_name)

        if sorted_nets:
            self.ref_net_combo.setCurrentIndex(0)

        diff_pairs_info = self.pcb_data.get("diff", {})
        diff_pair_nets = set()
        for pos_net, neg_net in diff_pairs_info.values():
            diff_pair_nets.add(pos_net)
            diff_pair_nets.add(neg_net)

        # Populate UI lists
        single_nets = sorted(
            [
                net
                for net in common_nets
                if net not in diff_pair_nets and net.upper() != "GND"
            ]
        )
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
