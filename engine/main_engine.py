
#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from input_validation import InputValidationError

try:
    from .exports import print_console_summary
    from .orchestrator import run_pipeline
except ImportError:
    from exports import print_console_summary
    from orchestrator import run_pipeline

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MarginRail engine v2")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", default=Path("./output_controllo_margini_v3"), type=Path)
    parser.add_argument("--config", default=None, type=Path)
    parser.add_argument("--review-file", default=None, type=Path)
    return parser.parse_args()

def run_engine(args: argparse.Namespace) -> None:
    if not args.input.exists():
        raise InputValidationError(f"File input non trovato: {args.input}")
    result = run_pipeline(input_path=args.input, output_dir=args.output_dir, config_path=args.config, review_file=args.review_file)
    print_console_summary(result.kpis, result.summary_tables, result.review_pack_df)
    print("\n=== File generati ===")
    for name, path in result.outputs.items():
        print(f"- {name}: {path.resolve()}")

def main() -> int:
    args = parse_args()
    try:
        run_engine(args)
        return 0
    except InputValidationError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except Exception as exc:
        print("Errore durante l’analisi MarginRail. Verifica input, configurazione e struttura progetto.", file=sys.stderr)
        print(f"Dettaglio sintetico: {exc}", file=sys.stderr)
        return 1
if __name__ == "__main__":
    raise SystemExit(main())
