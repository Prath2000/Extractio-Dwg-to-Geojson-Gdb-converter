"""
extractio — launcher
====================
Convenience wrapper around extractio.py.

Usage:
    python launcher.py                           # run all unlocked layers
    python launcher.py --layers "Block Boundaries"
    python launcher.py --layers "Solar Tracker" "HT Trench"
    python launcher.py --list                    # show all layers
    python launcher.py --config path/to/config.yaml
"""

import sys
import os
import argparse

# ── CONFIG PATH ─────────────────────────────────────────────
# Point this to your project YAML config file.
DEFAULT_CONFIG = "./config.yaml"

# ── EXECUTOR PATH ────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXECUTOR   = os.path.join(SCRIPT_DIR, "extractio.py")


def main():
    parser = argparse.ArgumentParser(
        prog="extractio",
        description="extractio — DWG to GeoJSON extraction engine",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""
Examples:
  python launcher.py                           Interactive layer selector
  python launcher.py --layers "Solar Tracker"  Run one layer
  python launcher.py --layers "HT" "cable"     Run multiple (fuzzy names ok)
  python launcher.py --run all                 Run all unlocked layers
  python launcher.py --list                    Show all layer names + aliases
  python launcher.py --config custom.yaml      Use a different config file
        """
    )
    parser.add_argument(
        "--config", "-c",
        default=DEFAULT_CONFIG,
        help=f"Path to YAML config (default: {DEFAULT_CONFIG})"
    )
    parser.add_argument(
        "--layers", "-l",
        nargs="*",
        metavar="LAYER",
        help="Layer name(s) to run — fuzzy match, aliases accepted"
    )
    parser.add_argument(
        "--run",
        nargs="?",
        const="all",
        metavar="all",
        help="Run unlocked layers. Use --run all"
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List all defined layers and exit"
    )

    args = parser.parse_args()

    if not os.path.exists(args.config):
        print(f"\n  Config not found: {args.config}")
        print(f"  Edit DEFAULT_CONFIG in launcher.py or use: --config path/to/config.yaml\n")
        sys.exit(1)

    if not os.path.exists(EXECUTOR):
        print(f"\n  Executor not found: {EXECUTOR}")
        print(f"  Make sure extractio.py is in the same folder as launcher.py\n")
        sys.exit(1)

    sys.argv = [EXECUTOR, args.config]

    if args.list:
        sys.argv.append("--list")
    elif args.run:
        sys.argv += ["--run", args.run]
    elif args.layers is not None:
        sys.argv += ["--layers"] + args.layers

    import importlib.util
    spec   = importlib.util.spec_from_file_location("extractio", EXECUTOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.main()


if __name__ == "__main__":
    main()
