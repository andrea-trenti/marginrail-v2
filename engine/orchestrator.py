
from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional
import pandas as pd

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from input_validation import raise_if_invalid_file

try:
    from .config import RuleConfig, load_json_config
    from .exports import build_kpis, build_rule_analysis, build_summary_tables, export_outputs
    from .review_memory import build_review_pack, load_review_file, merge_reviews
    from .rules import evaluate_cases
    from .validation import build_base_dataframe, load_workbook_data
except ImportError:
    from config import RuleConfig, load_json_config
    from exports import build_kpis, build_rule_analysis, build_summary_tables, export_outputs
    from review_memory import build_review_pack, load_review_file, merge_reviews
    from rules import evaluate_cases
    from validation import build_base_dataframe, load_workbook_data

@dataclass
class EngineRunResult:
    config: RuleConfig
    data: Dict[str, pd.DataFrame]
    base_df: pd.DataFrame
    cases_df: pd.DataFrame
    review_df: pd.DataFrame
    review_pack_df: pd.DataFrame
    rule_analysis_df: pd.DataFrame
    kpis: Dict[str, object]
    summary_tables: Dict[str, pd.DataFrame]
    outputs: Dict[str, Path]

def run_pipeline(input_path: Path, output_dir: Path, config_path: Optional[Path] = None, review_file: Optional[Path] = None) -> EngineRunResult:
    raise_if_invalid_file(input_path)
    config = load_json_config(config_path)
    data = load_workbook_data(input_path)
    base_df = build_base_dataframe(data)
    cases_df = evaluate_cases(base_df, config)
    review_df = load_review_file(review_file)
    cases_df = merge_reviews(cases_df, review_df)
    review_pack_df = build_review_pack(cases_df, config)
    rule_analysis_df = build_rule_analysis(cases_df)
    kpis = build_kpis(cases_df, base_df)
    summary_tables = build_summary_tables(cases_df)
    outputs = export_outputs(output_dir=output_dir, cases_df=cases_df, review_pack_df=review_pack_df, rule_analysis_df=rule_analysis_df, kpis=kpis, summary_tables=summary_tables, config=config)
    return EngineRunResult(config=config, data=data, base_df=base_df, cases_df=cases_df, review_df=review_df, review_pack_df=review_pack_df, rule_analysis_df=rule_analysis_df, kpis=kpis, summary_tables=summary_tables, outputs=outputs)
