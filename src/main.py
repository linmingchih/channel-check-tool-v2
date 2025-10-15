import sys
import os
import json
import re
import subprocess
from PySide6.QtWidgets import QApplication, QFileDialog, QListWidgetItem, QTableWidgetItem
from PySide6.QtCore import QProcess, Qt
from PySide6.QtGui import QColor
from gui import AEDBCCTCalculator

class MainController(AEDBCCTCalculator):
    def __init__(self):
        super().__init__()
        self.load_config()
        self.connect_signals()

    def connect_signals(self):
        self.brd_radio.toggled.connect(self.on_layout_type_changed)
        self.aedb_radio.toggled.connect(self.on_layout_type_changed)
        self.open_layout_button.clicked.connect(self.open_layout)
        self.browse_stackup_button.clicked.connect(self.browse_stackup)
        self.apply_import_button.clicked.connect(lambda: self.run_get_edb(self.layout_path_label.text()))
        self.component_filter_input.textChanged.connect(self.filter_components)
        self.controller_components_list.itemSelectionChanged.connect(self.update_nets)
        self.dram_components_list.itemSelectionChanged.connect(self.update_nets)
        self.single_ended_list.itemChanged.connect(self.update_checked_count)
        self.differential_pairs_list.itemChanged.connect(self.update_checked_count)
        self.apply_button.clicked.connect(self.apply_settings)
        self.apply_simulation_button.clicked.connect(self.apply_simulation_settings)
        self.browse_touchstone_button.clicked.connect(self.browse_touchstone)
        self.browse_metadata_button.clicked.connect(self.browse_port_metadata)
        self.touchstone_path_input.textChanged.connect(self.check_paths_and_load_ports)
        self.port_metadata_path_input.textChanged.connect(self.check_paths_and_load_ports)
        self.save_config_button.clicked.connect(self.save_cct_config)
        self.load_config_button.clicked.connect(self.load_cct_config)
        self.reset_defaults_button.clicked.connect(self.reset_cct_defaults)
        self.prerun_button.clicked.connect(self.run_prerun)
        self.calculate_button.clicked.connect(self.run_calculate)

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
        if not layout_path or layout_path == "No design loaded":
            self.log("Please select a design first.", "red")
            return

        self.log(f"Opening layout: {layout_path}")
        self.apply_import_button.setEnabled(False)
        self.apply_import_button.setText("Running...")
        self.apply_import_button.setStyleSheet("background-color: yellow; color: black;")
        self.current_layout_path = layout_path # Store for finished handler

        script_path = os.path.join(os.path.dirname(__file__), "get_edb.py")
        python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
        edb_version = self.edb_version_input.text()
        
        stackup_path = self.stackup_path_input.text()
        if not (stackup_path and os.path.exists(stackup_path)):
            stackup_path = ""

        command = [python_executable, script_path, layout_path, edb_version, stackup_path]
        self.log(f"Running command: {' '.join(command)}")

        self.get_edb_process = QProcess()
        self.get_edb_process.readyReadStandardOutput.connect(self.handle_get_edb_stdout)
        self.get_edb_process.readyReadStandardError.connect(self.handle_get_edb_stderr)
        self.get_edb_process.finished.connect(self.get_edb_finished)
        self.get_edb_process.start(command[0], command[1:])

    def handle_get_edb_stdout(self):
        data = self.get_edb_process.readAllStandardOutput().data().decode().strip()
        for line in data.splitlines(): self.log(line)

    def handle_get_edb_stderr(self):
        data = self.get_edb_process.readAllStandardError().data().decode().strip()
        for line in data.splitlines(): self.log(line, color="red")

    def get_edb_finished(self):
        self.log("Get EDB process finished.")
        self.apply_import_button.setEnabled(True)
        self.apply_import_button.setText("Apply")
        self.apply_import_button.setStyleSheet(self.apply_import_button_original_style)
        exit_code = self.get_edb_process.exitCode()

        if exit_code == 0:
            layout_path = self.current_layout_path
            if layout_path.endswith('.aedb'):
                json_output_path = layout_path.replace('.aedb', '.json')
            else:
                json_output_path = os.path.splitext(layout_path)[0] + '.json'

            self.log(f"Successfully generated {os.path.basename(json_output_path)}")
            
            if layout_path.lower().endswith('.brd'):
                new_aedb_path = os.path.splitext(layout_path)[0] + '.aedb'
                self.layout_path_label.setText(new_aedb_path)
                self.log(f"Design path has been updated to: {new_aedb_path}")

            self.load_pcb_data(json_output_path)
        else:
            self.log(f"Get EDB process failed with exit code {exit_code}.", "red")

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
        
        # Block signals to prevent excessive updates
        self.single_ended_list.blockSignals(True)
        self.differential_pairs_list.blockSignals(True)
        
        self.single_ended_list.clear()
        self.differential_pairs_list.clear()
        self.ref_net_combo.clear()

        try:
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
        finally:
            # Unblock signals
            self.single_ended_list.blockSignals(False)
            self.differential_pairs_list.blockSignals(False)

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
                "reference_net": [self.reference_net_label.text()],
            },
            "solver": "SIwave" if self.siwave_radio.isChecked() else "HFSS",
            "frequency_sweeps": sweeps,
        }

        try:
            with open(file_path, "w") as f:
                json.dump(settings, f, indent=2)
            self.log(f"Simulation settings saved to {file_path}")

            self.log("Applying simulation settings to EDB...")
            self.apply_simulation_button.setEnabled(False)
            self.apply_simulation_button.setText("Running...")
            self.apply_simulation_button.setStyleSheet("background-color: yellow; color: black;")

            script_path = os.path.join(os.path.dirname(__file__), "set_sim.py")
            python_executable = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".venv", "Scripts", "python.exe")
            command = [python_executable, script_path, file_path]
            
            self.set_sim_process = QProcess()
            self.set_sim_process.readyReadStandardOutput.connect(self.handle_set_sim_stdout)
            self.set_sim_process.readyReadStandardError.connect(self.handle_set_sim_stderr)
            self.set_sim_process.finished.connect(self.set_sim_finished)
            self.set_sim_process.start(command[0], command[1:])

        except Exception as e:
            self.log(f"Error applying simulation settings: {e}", color="red")
            self.apply_simulation_button.setEnabled(True)
            self.apply_simulation_button.setText("Apply Simulation")
            self.apply_simulation_button.setStyleSheet(self.apply_simulation_button_original_style)

    def handle_set_sim_stdout(self):
        data = self.set_sim_process.readAllStandardOutput().data().decode().strip()
        for line in data.splitlines(): self.log(line)

    def handle_set_sim_stderr(self):
        data = self.set_sim_process.readAllStandardError().data().decode().strip()
        for line in data.splitlines(): self.log(line, color="red")

    def set_sim_finished(self):
        self.log("Set simulation process finished.")
        self.apply_simulation_button.setEnabled(True)
        self.apply_simulation_button.setText("Apply Simulation")
        self.apply_simulation_button.setStyleSheet(self.apply_simulation_button_original_style)
        if self.set_sim_process.exitCode() == 0:
            self.log("Successfully applied simulation settings.")
        else:
            self.log(f"Set simulation process failed with exit code {self.set_sim_process.exitCode()}.", "red")

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
            self.calculate_button.setText("Running...")
            self.calculate_button.setStyleSheet("background-color: yellow; color: black;")
        elif mode == 'prerun':
            self.prerun_button.setText("Running...")
            self.prerun_button.setStyleSheet("background-color: yellow; color: black;")

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
        self.prerun_button.setText("Pre-run")
        self.prerun_button.setStyleSheet(self.prerun_button_original_style)

    def run_prerun(self): self.run_cct_process("prerun")
    def run_calculate(self): self.run_cct_process("run")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainController()
    window.show()
    sys.exit(app.exec())