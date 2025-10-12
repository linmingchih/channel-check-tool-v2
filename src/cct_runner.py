import argparse
import json
import logging
import sys
import traceback
from pathlib import Path

# Ensure the 'src' directory is in the Python path
# to allow importing the 'cct' module.
ROOT_DIR = Path(__file__).resolve().parent
if ROOT_DIR.name == 'src':
    SRC_DIR = ROOT_DIR
else:
    SRC_DIR = ROOT_DIR / 'src'

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from cct import CCT, DEFAULT_CIRCUIT_VERSION


def main():
    """
    Executes a CCT (Channel Check Tool) run as a standalone process.

    This script is designed to be called from another process (like a GUI)
    to isolate the AEDT session and prevent conflicts with other libraries
    like pyedb.

    It communicates progress and results via stdout and stderr.
    - stdout is used for structured messages:
      - "MESSAGE: <text>" for status updates.
      - "PROGRESS: <step>" for progress bar updates.
      - "FINISHED: <payload>" on successful completion.
    - stderr is used for error reporting.
    """
    parser = argparse.ArgumentParser(description="CCT Runner")
    parser.add_argument("--touchstone-path", required=True, type=Path)
    parser.add_argument("--metadata-path", required=True, type=Path)
    parser.add_argument("--output-path", type=Path, default=None)
    parser.add_argument("--workdir", required=True, type=Path)
    parser.add_argument("--settings", required=True, type=str, help="JSON string of CCT settings")
    parser.add_argument("--mode", required=True, choices=['run', 'prerun'])
    args = parser.parse_args()

    # Ensure the working directory exists before setting up logging
    args.workdir.mkdir(parents=True, exist_ok=True)

    log_file = args.workdir / 'cct_runner.log'
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        filename=str(log_file),
                        filemode='w')

    logging.info("CCT Runner started.")
    logging.info(f"Args: {args}")

    try:
        settings = json.loads(args.settings)

        print("MESSAGE: Preparing CCT inputs...")
        print("PROGRESS: 0")
        logging.info("Preparing CCT inputs.")

        options = settings.get('options') or settings.get('prune', {})
        threshold_raw = options.get('threshold_db') if isinstance(options, dict) else None
        try:
            threshold_value = float(threshold_raw) if threshold_raw is not None else None
        except (TypeError, ValueError):
            threshold_value = None

        circuit_version = None
        if isinstance(options, dict):
            version_candidate = options.get('circuit_version')
            if version_candidate is not None:
                circuit_version = str(version_candidate).strip() or None

        logging.info("Initializing CCT object.")
        cct = CCT(
            str(args.touchstone_path),
            str(args.metadata_path),
            workdir=args.workdir,
            threshold_db=threshold_value,
            circuit_version=circuit_version,
        )
        logging.info("CCT object initialized.")

        print("MESSAGE: Configuring transmit settings...")
        print("PROGRESS: 1")
        logging.info("Configuring transmit settings.")
        tx = settings.get('tx', {})
        cct.set_txs(
            vhigh=tx.get('vhigh', ''),
            t_rise=tx.get('t_rise', ''),
            ui=tx.get('ui', ''),
            res_tx=tx.get('res_tx', ''),
            cap_tx=tx.get('cap_tx', ''),
        )
        logging.info("Transmit settings configured.")

        print("MESSAGE: Configuring receive settings...")
        print("PROGRESS: 2")
        logging.info("Configuring receive settings.")
        rx = settings.get('rx', {})
        cct.set_rxs(
            res_rx=rx.get('res_rx', ''),
            cap_rx=rx.get('cap_rx', ''),
        )
        logging.info("Receive settings configured.")

        if args.mode == 'prerun':
            print("MESSAGE: Running pre-run threshold analysis...")
            print("PROGRESS: 3")
            logging.info("Running pre-run threshold analysis.")
            summaries = cct.pre_run()
            summary_text = _summarize_prerun(summaries, threshold_value)
            print("PROGRESS: 4")
            print(f"FINISHED: {summary_text}")
            logging.info("Pre-run finished.")
            return

        print("MESSAGE: Running transient simulation...")
        print("PROGRESS: 3")
        logging.info("Running transient simulation.")
        run_params = settings.get('run', {})
        cct.run(
            tstep=run_params.get('tstep', ''),
            tstop=run_params.get('tstop', ''),
        )
        logging.info("Transient simulation finished.")

        print("MESSAGE: Generating CCT report...")
        print("PROGRESS: 4")
        logging.info("Generating CCT report.")
        if args.output_path is None:
            raise RuntimeError('Output path not provided for CCT run')
        cct.calculate(output_path=str(args.output_path))
        logging.info(f"CCT report saved to {args.output_path}")

        print(f"MESSAGE: CCT results saved to {args.output_path}")
        print(f"FINISHED: {args.output_path}")

    except Exception:
        logging.error("An exception occurred in CCT Runner.", exc_info=True)
        # Use stderr for exceptions to separate them from normal output.
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)

    logging.info("CCT Runner finished successfully.")


def _summarize_prerun(summaries, threshold_value):
    """Helper to generate the same summary text as the original worker."""
    if not summaries:
        if threshold_value is None:
            return 'Pre-run complete. No transmitters evaluated.'
        return f'Pre-run complete at threshold {threshold_value:.1f} dB. No transmitters evaluated.'

    lines = []
    if threshold_value is None:
        lines.append('Pre-run complete. Using full network (no threshold applied).')
    else:
        lines.append(f'Pre-run complete at threshold {threshold_value:.1f} dB.')

    port_ratios = []
    rx_ratios = []
    for stats in summaries:
        total_ports = int(stats.get('total_port_count', 0) or 0)
        kept_ports = int(stats.get('kept_port_count', 0) or 0)
        port_ratio = (kept_ports / total_ports) if total_ports else 0.0
        port_ratios.append(port_ratio)

        total_rx = int(stats.get('total_rx_port_count', 0) or 0)
        kept_rx = int(stats.get('kept_rx_port_count', 0) or 0)
        rx_ratio = (kept_rx / total_rx) if total_rx else None
        if rx_ratio is not None:
            rx_ratios.append(rx_ratio)

        label = str(stats.get('tx_label', 'tx'))
        line = f"{label}: ports {kept_ports}/{total_ports}"
        if total_ports:
            line += f" ({port_ratio:.1%})"
        if total_rx:
            line += f", rx {kept_rx}/{total_rx} ({rx_ratio:.1%})"
        lines.append(line)

    if port_ratios:
        avg_port = sum(port_ratios) / len(port_ratios)
        lines.insert(1, f"Average kept ports: {avg_port:.1%}")
    if rx_ratios:
        avg_rx = sum(rx_ratios) / len(rx_ratios)
        insert_at = 2 if port_ratios else 1
        lines.insert(insert_at, f"Average kept RX ports: {avg_rx:.1%}")

    return "\n".join(lines)


if __name__ == "__main__":
    main()
