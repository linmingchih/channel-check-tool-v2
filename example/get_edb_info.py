"""Extract information from an AEDB design using PyEDB."""
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


def _iter_component_items(edb: "Edb") -> Iterator[Tuple[str, object]]:
    container = getattr(edb.components, "components", {})
    if isinstance(container, dict):
        yield from container.items()
        return

    items = getattr(container, "items", None)
    if callable(items):
        candidates = items()
        if isinstance(candidates, dict):
            yield from candidates.items()
        else:
            for item in candidates:
                if isinstance(item, (list, tuple)) and len(item) == 2:
                    yield item[0], item[1]
    elif isinstance(container, Iterable):
        for component in container:
            name = getattr(component, "name", None)
            if isinstance(name, str):
                yield name, component


def _extract_net_names(component: object) -> Set[str]:
    names: Set[str] = set()
    nets = getattr(component, "nets", [])
    if isinstance(nets, dict):
        iterable = nets.values()
    else:
        iterable = nets
    for net in iterable:
        if isinstance(net, str):
            names.add(net)
        else:
            name = getattr(net, "name", None)
            if name:
                names.add(name)
    return names


def _pin_count(component: object) -> int:
    pins = getattr(component, "pins", None)
    if pins is None:
        return 0
    try:
        keys = getattr(pins, "keys", None)
        if callable(keys):
            return len(list(keys()))
    except Exception:  # pragma: no cover - PyEDB objects are dynamic
        pass
    try:
        return len(pins)  # type: ignore[arg-type]
    except TypeError:
        values = getattr(pins, "values", None)
        if callable(values):
            return sum(1 for _ in values())
    return 0


def _iter_nets(edb: "Edb") -> Iterator[str]:
    container = getattr(edb.nets, "nets", {})
    if isinstance(container, dict):
        for name in container:
            if isinstance(name, str):
                yield name
        return

    items = getattr(container, "items", None)
    if callable(items):
        for name, _ in items() or []:
            if isinstance(name, str):
                yield name
        return

    if isinstance(container, Iterable):
        for net in container:
            if isinstance(net, str):
                yield net
            else:
                name = getattr(net, "name", None)
                if isinstance(name, str):
                    yield name


def _differential_pairs(edb: "Edb") -> List[Dict[str, str]]:
    container = getattr(edb, "differential_pairs", None)
    if container is None:
        return []

    try:
        items_attr = container.items
    except AttributeError:
        return []

    if callable(items_attr):
        candidates = items_attr()
        iterator: Iterable[Tuple[str, object]]
        if isinstance(candidates, dict):
            iterator = candidates.items()
        else:
            iterator = candidates
    elif isinstance(items_attr, dict):
        iterator = items_attr.items()
    else:
        iterator = []

    pairs: List[Dict[str, str]] = []
    for name, diff in iterator:
        pos = getattr(diff, "positive_net", None)
        neg = getattr(diff, "negative_net", None)
        pos_name = getattr(pos, "name", None)
        neg_name = getattr(neg, "name", None)
        if isinstance(name, str) and isinstance(pos_name, str) and isinstance(neg_name, str):
            pairs.append({"name": name, "positive": pos_name, "negative": neg_name})
    pairs.sort(key=lambda item: item["name"].lower())
    return pairs


def _match_pattern(name: str, pattern: Optional[re.Pattern[str]]) -> bool:
    if pattern is None:
        return True
    return bool(pattern.search(name))


def _collect_components(
    edb: "Edb", pattern: Optional[re.Pattern[str]]
) -> Tuple[List[Dict[str, object]], Dict[str, List[str]]]:
    components: List[Dict[str, object]] = []
    component_nets: Dict[str, List[str]] = {}
    for name, component in _iter_component_items(edb):
        if not isinstance(name, str) or not name:
            continue
        if not _match_pattern(name, pattern):
            continue
        nets = sorted(_extract_net_names(component))
        component_nets[name] = nets
        components.append(
            {
                "name": name,
                "pin_count": _pin_count(component),
                "nets": nets,
            }
        )
    components.sort(key=lambda item: (-int(item["pin_count"]), item["name"].lower()))
    ordered_nets: Dict[str, List[str]] = {item["name"]: component_nets[item["name"]] for item in components}
    return components, ordered_nets


def _parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Collect AEDB component and net information")
    parser.add_argument("aedb_path", type=Path, help="Path to the .aedb directory to inspect")
    parser.add_argument("--output", type=Path, required=True, help="Destination JSON file for the extracted data")
    parser.add_argument("--version", dest="version", help="Optional AEDT version to use with PyEDB", default=None)
    parser.add_argument(
        "--component-filter",
        dest="component_filter",
        default=None,
        help="Regex pattern (case-insensitive) used to filter component names",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = _parse_args(argv)
    aedb_path: Path = args.aedb_path
    output_path: Path = args.output
    component_filter: Optional[str] = args.component_filter
    version = (args.version or "").strip()
    if not version or version.lower() == "none":
        version = None

    if not aedb_path.exists():
        raise SystemExit(f"AEDB path does not exist: {aedb_path}")

    pattern: Optional[re.Pattern[str]] = None
    if component_filter:
        try:
            pattern = re.compile(component_filter, re.IGNORECASE)
        except re.error as exc:
            raise SystemExit(f"Invalid component filter '{component_filter}': {exc}")

    print(f"MESSAGE:Loading {aedb_path}")
    edb = Edb(str(aedb_path), edbversion=version)
    try:
        components, component_nets = _collect_components(edb, pattern)
        all_nets = sorted(set(net for nets in component_nets.values() for net in nets))
        diff_pairs = _differential_pairs(edb)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "aedb_path": str(aedb_path),
            "components": components,
            "component_nets": component_nets,
            "nets": all_nets,
            "differential_pairs": diff_pairs,
        }
        with output_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
        print(f"FINISHED:{output_path}")
    finally:
        try:
            edb.close()
        except Exception:
            pass
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())
