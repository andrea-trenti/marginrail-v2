from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Dict

import pandas as pd

try:
    from .config import RuleConfig
    from .rules import choose_bucket
except ImportError:  # pragma: no cover
    from config import RuleConfig
    from rules import choose_bucket


def build_kpis(cases_df: pd.DataFrame, base_df: pd.DataFrame) -> Dict[str, object]:
    if cases_df.empty:
        return {
            "totale_righe_vendita": int(len(base_df)),
            "totale_casi": 0,
            "totale_rischio_eur": 0.0,
            "casi_critical": 0,
            "casi_high": 0,
            "clienti_coinvolti": 0,
            "venditori_coinvolti": 0,
        }

    kpis = {
        "totale_righe_vendita": int(len(base_df)),
        "totale_casi": int(len(cases_df)),
        "totale_rischio_eur": round(float(cases_df["MarginRiskEUR"].sum()), 2),
        "casi_critical": int((cases_df["Severity"] == "critical").sum()),
        "casi_high": int((cases_df["Severity"] == "high").sum()),
        "casi_medium": int((cases_df["Severity"] == "medium").sum()),
        "casi_low": int((cases_df["Severity"] == "low").sum()),
        "clienti_coinvolti": int(cases_df["ClienteID"].nunique()),
        "venditori_coinvolti": int(cases_df["Venditore"].nunique()),
        "ordini_coinvolti": int(cases_df["NumeroOrdine"].nunique()),
        "regole_attive": int(cases_df["RuleCode"].nunique()),
    }
    if "ReviewOutcome" in cases_df.columns:
        reviewed = cases_df["ReviewOutcome"].notna().sum()
        kpis["casi_reviewati"] = int(reviewed)
        if reviewed:
            kpis["true_issue_reviewed"] = int((cases_df["ReviewOutcome"] == "true_issue").sum())
            kpis["false_positive_reviewed"] = int((cases_df["ReviewOutcome"] == "false_positive").sum())
            kpis["approved_exception_reviewed"] = int((cases_df["ReviewOutcome"] == "approved_exception").sum())
    return kpis


def build_summary_tables(cases_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if cases_df.empty:
        empty = pd.DataFrame()
        return {
            "per_regola": empty,
            "per_venditore": empty,
            "per_cliente": empty,
            "top_prodotti": empty,
            "timeline_mensile": empty,
            "review_quality": empty,
        }

    per_regola = cases_df.groupby(["RuleCode", "RuleName", "Severity", "Status"], dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum")).reset_index().sort_values(["RischioEUR", "Casi"], ascending=[False, False])
    per_venditore = cases_df.groupby(["Venditore"], dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum"), Clienti=("ClienteID", "nunique")).reset_index().sort_values(["RischioEUR", "Casi"], ascending=[False, False])
    per_cliente = cases_df.groupby(["ClienteID", "Cliente"], dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum"), Regole=("RuleCode", "nunique")).reset_index().sort_values(["RischioEUR", "Casi"], ascending=[False, False])
    top_prodotti = cases_df.groupby(["ProdottoID", "Prodotto", "Categoria"], dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum")).reset_index().sort_values(["RischioEUR", "Casi"], ascending=[False, False])

    timeline = cases_df.copy()
    timeline["Mese"] = pd.to_datetime(timeline["DataDocumento"], errors="coerce").dt.to_period("M").astype(str)
    timeline_mensile = timeline.groupby("Mese", dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum")).reset_index().sort_values("Mese")

    if "ReviewOutcome" in cases_df.columns and cases_df["ReviewOutcome"].notna().any():
        review_quality = cases_df[cases_df["ReviewOutcome"].notna()].groupby(["RuleCode", "ReviewOutcome"], dropna=False).agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum")).reset_index().sort_values(["RuleCode", "Casi"], ascending=[True, False])
    else:
        review_quality = pd.DataFrame()

    return {
        "per_regola": per_regola,
        "per_venditore": per_venditore,
        "per_cliente": per_cliente,
        "top_prodotti": top_prodotti,
        "timeline_mensile": timeline_mensile,
        "review_quality": review_quality,
    }


def build_rule_analysis(cases_df: pd.DataFrame) -> pd.DataFrame:
    grouped = cases_df.groupby("RuleCode", dropna=False).agg(RuleName=("RuleName", "first"), Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum"), Critical=("Severity", lambda s: int((s == "critical").sum())), High=("Severity", lambda s: int((s == "high").sum())), Medium=("Severity", lambda s: int((s == "medium").sum())), Low=("Severity", lambda s: int((s == "low").sum())), ApprovedException=("Status", lambda s: int((s == "approved_exception").sum())), OpenReview=("Status", lambda s: int((s == "open_review").sum()))).reset_index()
    grouped["ApprovedShare"] = (grouped["ApprovedException"] / grouped["Casi"]).round(3)
    grouped["AvgRiskEUR"] = (grouped["RischioEUR"] / grouped["Casi"]).round(2)
    grouped["CaseBucket"] = grouped["RuleCode"].map(choose_bucket)

    recommendations = []
    hypotheses = []
    for _, row in grouped.iterrows():
        rule = row["RuleCode"]
        if rule == "PRICE_BELOW_FLOOR":
            recommendations.append("TENERE COME REGOLA CORE E METTERE IN DEMO")
            hypotheses.append("Segnale forte: pochi casi, severità alta, buona leggibilità.")
        elif rule == "DISCOUNT_OVER_ROLE_THRESHOLD":
            recommendations.append("RESTRINGERE E AGGIUNGERE BUFFER/FILTRI")
            hypotheses.append("Approved share elevata: probabile rumore da soglie ampie o deroghe già note.")
        elif rule == "LOW_MARGIN_VS_TARGET":
            recommendations.append("TENERE, MA SEGMENTARE TARGET")
            hypotheses.append("Rischio totale molto alto: utile, ma target probabilmente troppo uniforme.")
        elif rule == "COST_PASS_THROUGH_RISK":
            recommendations.append("TENERE, MA INTRODURRE FILTRI OPERATIVI")
            hypotheses.append("Serve distinguere ritardo fisiologico vs vero mancato pass-through.")
        elif rule == "CREDIT_RISK_ON_CONCESSION":
            recommendations.append("SEPARARE IN BUCKET CREDITO/GOVERNANCE")
            hypotheses.append("Più regola di rischio commerciale che di leakage pricing puro.")
        elif rule == "RETURN_OR_CREDIT_NOTE":
            recommendations.append("SEPARARE IN POST-MORTEM / AFTER-SALES")
            hypotheses.append("Molto utile, ma non va mischiata al backlog operativo pricing.")
        else:
            recommendations.append("REVIEW MANUALE")
            hypotheses.append("Serve campione review per capire precisione effettiva.")
    grouped["Recommendation"] = recommendations
    grouped["Hypothesis"] = hypotheses
    grouped = grouped.sort_values(["RischioEUR", "Casi"], ascending=[False, False])
    return grouped


def export_outputs(output_dir: Path, cases_df: pd.DataFrame, review_pack_df: pd.DataFrame, rule_analysis_df: pd.DataFrame, kpis: Dict[str, object], summary_tables: Dict[str, pd.DataFrame], config: RuleConfig) -> Dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    cases_csv = output_dir / "casi_controllo_margini_v3.csv"
    cases_df.to_csv(cases_csv, index=False)

    review_csv = output_dir / "review_pack_step1.csv"
    review_pack_df.to_csv(review_csv, index=False)

    kpi_json = output_dir / "kpi_controllo_margini_v3.json"
    with kpi_json.open("w", encoding="utf-8") as f:
        json.dump(kpis, f, ensure_ascii=False, indent=2, default=str)

    config_json = output_dir / "config_effettiva_usata.json"
    with config_json.open("w", encoding="utf-8") as f:
        json.dump(asdict(config), f, ensure_ascii=False, indent=2, default=str)

    report_xlsx = output_dir / "report_controllo_margini_v3.xlsx"
    with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
        cases_df.to_excel(writer, sheet_name="casi", index=False)
        pd.DataFrame([kpis]).to_excel(writer, sheet_name="kpi", index=False)
        review_pack_df.to_excel(writer, sheet_name="review_pack", index=False)
        rule_analysis_df.to_excel(writer, sheet_name="analisi_regole", index=False)
        for name, table in summary_tables.items():
            if not table.empty:
                table.to_excel(writer, sheet_name=name[:31], index=False)

    return {
        "cases_csv": cases_csv,
        "review_csv": review_csv,
        "kpi_json": kpi_json,
        "config_json": config_json,
        "report_xlsx": report_xlsx,
    }


def print_console_summary(kpis: Dict[str, object], summary_tables: Dict[str, pd.DataFrame], review_pack_df: pd.DataFrame) -> None:
    print("\n=== KPI principali ===")
    for key, value in kpis.items():
        print(f"- {key}: {value}")

    print(f"\n- review_pack_size: {len(review_pack_df)}")

    if not summary_tables["per_regola"].empty:
        print("\n=== Top regole per rischio ===")
        print(summary_tables["per_regola"].head(10).to_string(index=False))
