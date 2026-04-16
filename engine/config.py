from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional


@dataclass
class RuleConfig:
    extra_discount_tolerance: float = 0.02
    low_margin_floor: float = 0.12
    low_margin_gap_vs_target: float = 0.05
    cost_increase_trigger: float = 0.05
    severe_floor_gap_pct: float = 0.05
    high_risk_value_eur: float = 1000.0
    medium_risk_value_eur: float = 250.0
    discount_role_buffer_pct: float = 0.015
    discount_role_min_gap_pct: float = 0.015
    low_margin_min_risk_eur: float = 150.0
    cost_pass_min_risk_eur: float = 150.0
    return_note_min_abs_value_eur: float = 250.0
    credit_risk_min_discount_pct: float = 0.00
    high_risk_only_credit_rule: bool = True
    skip_discount_if_active_promo: bool = True
    skip_discount_if_active_contract: bool = True
    review_sample_discount: int = 20
    review_sample_low_margin: int = 20
    review_sample_cost_pass: int = 20
    review_sample_credit: int = 10
    review_sample_returns: int = 10


def load_json_config(config_path: Optional[Path]) -> RuleConfig:
    config = RuleConfig()
    if not config_path or not config_path.exists():
        return config

    payload = json.loads(config_path.read_text(encoding="utf-8"))
    allowed = set(asdict(config).keys())
    updates = {k: v for k, v in payload.items() if k in allowed}
    return RuleConfig(**{**asdict(config), **updates})
