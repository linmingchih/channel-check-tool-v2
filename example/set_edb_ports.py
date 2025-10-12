"""Configure circuit ports in an AEDB design using PyEDB."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

try:
    from pyedb import Edb  # type: ignore
except ImportError as exc:  # pragma: no cover - runtime dependency
    raise SystemExit(f"PyEDB is required to use this tool: {exc}")

try:  # pragma: no cover - optional dependency for consistent port names
    from cct import prefix_port_name  # type: ignore
except ImportError:  # pragma: no cover - allow running without CCT installed
    def prefix_port_name(name: str, sequence: int) -> str:
        base = str(name or "").strip()
        match = re.match(r"^\d+_(.*)$", base)
        if match:
            base = match.group(1)
        return f"{sequence}_{base}" if base else str(sequence)


def _iter_nets(edb: "Edb") -> Dict[str, object]:
    container = getattr(edb.nets, "nets", {})
    if isinstance(container, dict):
        return container
    items = getattr(container, "items", None)
    if callable(items):
        candidates = items()
        if isinstance(candidates, dict):
            return candidates
        return {name: obj for name, obj in candidates}
    return {}


def _pin_group_from_result(result: object) -> Optional[object]:
    candidates: List[object]
    if isinstance(result, (list, tuple)):
        candidates = list(result)
    else:
        candidates = [result]
    for entry in reversed(candidates):
        if hasattr(entry, "create_port_terminal"):
            return entry
    return None


def _sanitized_group_name(component_name: str, net_name: str, suffix: Optional[str] = None) -> str:
    base = f"{component_name}_{net_name}"
    safe = re.sub(r"[^A-Za-z0-9_]+", "_", base)
    safe = re.sub(r"_+", "_", safe).strip("_")
    if suffix:
        safe = f"{safe}_{suffix}" if safe else suffix
    if not safe:
        fallback = re.sub(r"[^A-Za-z0-9_]+", "_", component_name).strip("_") or "comp"
        suffix_part = suffix or "pg"
        safe = f"{fallback}_{suffix_part}"
    return safe


def _ensure_pin_group(edb: "Edb", component_name: str, net_name: str, group_name: str) -> Optional[object]:
    result = edb.core_siwave.create_pin_group_on_net(component_name, net_name, group_name)
    return _pin_group_from_result(result)


def _create_reference_terminal(edb: "Edb", component_name: str, reference_net: str) -> Optional[object]:
    group_name = _sanitized_group_name(component_name, reference_net, suffix="ref")
    pin_group = _ensure_pin_group(edb, component_name, reference_net, group_name)
    if pin_group is None:
        return None
    terminal = pin_group.create_port_terminal(50)
    if hasattr(terminal, "SetName"):
        terminal.SetName(f"ref;{component_name};{reference_net}")
    return terminal


def _create_signal_terminal(edb: "Edb", component_name: str, net_name: str) -> Optional[object]:
    group_name = _sanitized_group_name(component_name, net_name)
    pin_group = _ensure_pin_group(edb, component_name, net_name, group_name)
    if pin_group is None:
        return None
    return pin_group.create_port_terminal(50)


def _create_ports(
    edb: "Edb",
    *,
    net_names: Iterable[str],
    reference_net: str,
    components: Iterable[str],
    component_roles: Dict[str, str],
    net_metadata: Dict[str, Dict[str, Optional[str]]],
) -> List[Dict[str, Optional[str]]]:
    nets_lookup = _iter_nets(edb)
    component_sequence = list(components)
    component_set = set(component_sequence)
    component_order = {name: idx for idx, name in enumerate(component_sequence)}
    reference_terminals: Dict[str, object] = {}
    metadata: List[Dict[str, Optional[str]]] = []

    for net_name in net_names:
        net_obj = nets_lookup.get(net_name)
        if net_obj is None:
            continue
        components_map = getattr(net_obj, "components", {})
        if isinstance(components_map, dict):
            component_iterable = list(components_map.keys())
        else:
            component_iterable = list(components_map)
        component_iterable.sort(key=lambda name: component_order.get(name, len(component_order)))
        for component_name in component_iterable:
            if component_name not in component_set:
                continue
            if component_name not in reference_terminals:
                reference_terminal = _create_reference_terminal(edb, component_name, reference_net)
                if reference_terminal is None:
                    continue
                reference_terminals[component_name] = reference_terminal
            signal_terminal = _create_signal_terminal(edb, component_name, net_name)
            if signal_terminal is None:
                continue
            base_port_name = f"{component_name}_{net_name}"
            sequence = len(metadata) + 1
            port_name = prefix_port_name(base_port_name, sequence)
            if hasattr(signal_terminal, "SetName"):
                signal_terminal.SetName(port_name)
            if hasattr(signal_terminal, "SetReferenceTerminal"):
                signal_terminal.SetReferenceTerminal(reference_terminals[component_name])
            net_info = net_metadata.get(net_name, {"type": "single"})
            metadata.append(
                {
                    "sequence": sequence,
                    "name": port_name,
                    "component": component_name,
                    "component_role": component_roles.get(component_name, "unknown"),
                    "net": net_name,
                    "net_type": net_info.get("type", "single"),
                    "pair": net_info.get("pair"),
                    "polarity": net_info.get("polarity"),
                }
            )
    return metadata


def _default_output_path(source: Path) -> Path:
    stem = source.stem
    stem = re.sub(r"(_applied)+$", "", stem)
    if not stem:
        stem = source.stem
    return source.with_name(f"{stem}_applied.aedb")


def _parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create circuit ports in an AEDB design")
    parser.add_argument("aedb_path", type=Path, help="Source AEDB directory")
    parser.add_argument("--config", type=Path, required=True, help="JSON file describing the ports to create")
    parser.add_argument("--output", type=Path, help="Destination AEDB directory")
    parser.add_argument("--metadata-output", type=Path, help="Optional JSON path to store metadata about the ports")
    parser.add_argument("--version", dest="version", default=None, help="Optional AEDT version for PyEDB")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = _parse_args(argv)
    aedb_path: Path = args.aedb_path
    config_path: Path = args.config
    output_path: Optional[Path] = args.output
    metadata_path: Optional[Path] = args.metadata_output
    version = (args.version or "").strip()
    if not version or version.lower() == "none":
        version = None

    if not aedb_path.exists():
        raise SystemExit(f"AEDB path does not exist: {aedb_path}")
    if not config_path.exists():
        raise SystemExit(f"Configuration JSON not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)

    signal_nets = config.get("signal_nets", [])
    reference_net = config.get("reference_net")
    components = config.get("components", [])
    component_roles = config.get("component_roles", {})
    net_metadata = config.get("net_metadata", {})

    if not isinstance(signal_nets, list) or not all(isinstance(net, str) for net in signal_nets):
        raise SystemExit("Configuration 'signal_nets' must be a list of strings")
    if not isinstance(reference_net, str) or not reference_net:
        raise SystemExit("Configuration must include a non-empty 'reference_net'")
    if not isinstance(components, list) or not all(isinstance(name, str) for name in components):
        raise SystemExit("Configuration 'components' must be a list of strings")
    if not components:
        raise SystemExit("Configuration must specify at least one component")

    if output_path is None:
        output_path = Path(config.get("output_path") or _default_output_path(aedb_path))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if metadata_path is not None:
        metadata_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"MESSAGE:Opening {aedb_path}")
    edb = Edb(str(aedb_path), edbversion=version)
    try:
        metadata = _create_ports(
            edb,
            net_names=signal_nets,
            reference_net=reference_net,
            components=components,
            component_roles=component_roles,
            net_metadata=net_metadata,
        )
        print(f"MESSAGE:Created {len(metadata)} ports")
        edb.save_edb_as(str(output_path))
        if metadata_path is not None:
            with metadata_path.open("w", encoding="utf-8") as handle:
                json.dump(metadata, handle, indent=2)
        print(f"FINISHED:{output_path}")
        if metadata_path is not None:
            print(f"METADATA:{metadata_path}")
    finally:
        try:
            edb.close()
        except Exception:
            pass
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())
