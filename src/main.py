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
)
from PySide6.QtCore import Qt


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
        self.setGeometry(100, 100, 800, 600)
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
        self.aedb_path_label = QLabel("D:\\...")
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

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(
            "Controllers: 0 | DRAMs: 0 | Shared nets: 0 | Shared differential pairs: 0"
        )

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
