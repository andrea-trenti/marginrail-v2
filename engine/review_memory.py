from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd

try:
    from .config import RuleConfig
    from .constants import REVIEW_COLUMNS
except ImportError:  # pragma: no cover
    from config import RuleConfig
    from constants import REVIEW_COLUMNS


def load_review_file(review_path: Optional[Path]) -> pd.DataFrame:
    if not review_path or not review_path.exists():
        return pd.DataFrame()
    if review_path.suffix.lower() == ".csv":
        review_df = pd.read_csv(review_path)
    else:
        review_df = pd.read_excel(review_path)
    return review_df


def merge_reviews(cases_df: pd.DataFrame, review_df: pd.DataFrame) -> pd.DataFrame:
    if review_df.empty:
        return cases_df
    key = None
    if "NaturalCaseKey" in review_df.columns:
        key = "NaturalCaseKey"
    elif "CaseID" in review_df.columns:
        key = "CaseID"
    else:
        return cases_df

    keep_cols = [c for c in review_df.columns if c in REVIEW_COLUMNS or c == key]
    review_df = review_df[keep_cols].copy()
    merged = cases_df.drop(columns=[c for c in REVIEW_COLUMNS if c in cases_df.columns], errors="ignore").merge(review_df, on=key, how="left")
    for col in REVIEW_COLUMNS:
        if col not in merged.columns:
            merged[col] = pd.NA
    return merged


def build_review_pack(cases_df: pd.DataFrame, config: RuleConfig) -> pd.DataFrame:
    frames = []
    critical = cases_df[cases_df["Severity"] == "critical"].copy()
    frames.append(critical)

    def top_by_rule(rule_code: str, n: int) -> pd.DataFrame:
        subset = cases_df[(cases_df["RuleCode"] == rule_code) & (cases_df["Severity"] == "high")].copy()
        if subset.empty:
            return subset
        return subset.sort_values(["MarginRiskEUR", "DataDocumento"], ascending=[False, True]).head(n)

    frames.append(top_by_rule("DISCOUNT_OVER_ROLE_THRESHOLD", config.review_sample_discount))
    frames.append(top_by_rule("LOW_MARGIN_VS_TARGET", config.review_sample_low_margin))
    frames.append(top_by_rule("COST_PASS_THROUGH_RISK", config.review_sample_cost_pass))
    frames.append(cases_df[cases_df["RuleCode"] == "CREDIT_RISK_ON_CONCESSION"].sort_values(["MarginRiskEUR"], ascending=False).head(config.review_sample_credit))
    frames.append(cases_df[cases_df["RuleCode"] == "RETURN_OR_CREDIT_NOTE"].sort_values(["MarginRiskEUR"], ascending=False).head(config.review_sample_returns))

    review_pack = pd.concat(frames, ignore_index=True)
    review_pack = review_pack.drop_duplicates(subset=["NaturalCaseKey"]).copy()
    review_pack["ReviewPriority"] = review_pack["Severity"].map({"critical": 1, "high": 2, "medium": 3, "low": 4})
    review_pack = review_pack.sort_values(["ReviewPriority", "MarginRiskEUR"], ascending=[True, False]).reset_index(drop=True)

    cols = [
        "CaseID", "NaturalCaseKey", "RuleCode", "RuleName", "CaseBucket", "Severity", "Status", "Owner", "MarginRiskEUR",
        "HeuristicReviewHint", "Reason", "SuggestedAction", "ClienteID", "Cliente", "Venditore", "ProdottoID", "Prodotto",
        "Categoria", "NumeroOrdine", "RigaOrdine", "DataDocumento", "ActivePromoFlag", "ActiveAccordoFlag",
    ] + REVIEW_COLUMNS + ["ReviewPriority"]
    keep = [c for c in cols if c in review_pack.columns]
    return review_pack[keep]
