from __future__ import annotations

from typing import List

import pandas as pd

try:
    from .config import RuleConfig
    from .constants import REVIEW_COLUMNS
    from .utils import build_case_id, build_natural_case_key, clean_text, nz
except ImportError:  # pragma: no cover
    from config import RuleConfig
    from constants import REVIEW_COLUMNS
    from utils import build_case_id, build_natural_case_key, clean_text, nz


def choose_owner(row: pd.Series, rule_code: str) -> str:
    if rule_code in {"PRICE_BELOW_FLOOR", "LOW_MARGIN_VS_TARGET", "COST_PASS_THROUGH_RISK"}:
        return "Pricing Manager"
    if rule_code in {"DISCOUNT_WITHOUT_COVERAGE", "DISCOUNT_OVER_ROLE_THRESHOLD"}:
        return "Sales Manager"
    if rule_code == "RETURN_OR_CREDIT_NOTE":
        return "Customer Service"
    if rule_code == "CREDIT_RISK_ON_CONCESSION":
        return "Credit Manager"
    return "Commercial Excellence"


def choose_bucket(rule_code: str) -> str:
    if rule_code in {"PRICE_BELOW_FLOOR", "LOW_MARGIN_VS_TARGET", "DISCOUNT_OVER_ROLE_THRESHOLD", "DISCOUNT_WITHOUT_COVERAGE"}:
        return "core_pricing"
    if rule_code == "COST_PASS_THROUGH_RISK":
        return "cost_pass_through"
    if rule_code == "CREDIT_RISK_ON_CONCESSION":
        return "credit_governance"
    if rule_code == "RETURN_OR_CREDIT_NOTE":
        return "post_sales"
    return "other"


def severity_from_risk(margin_risk_eur: float, config: RuleConfig, critical: bool = False, high: bool = False) -> str:
    if critical:
        return "critical"
    if high:
        return "high"
    if margin_risk_eur >= config.high_risk_value_eur:
        return "high"
    if margin_risk_eur >= config.medium_risk_value_eur:
        return "medium"
    return "low"


def evaluate_cases(df: pd.DataFrame, config: RuleConfig) -> pd.DataFrame:
    cases: List[dict] = []

    for _, row in df.iterrows():
        tipo_documento = clean_text(row.get("TipoDocumento")) or "Fattura"
        motivo_deroga = clean_text(row.get("MotivoDeroga"))
        stato_deroga = clean_text(row.get("StatoDeroga"))
        promo_id = clean_text(row.get("PromoID"))
        accordo_id = clean_text(row.get("AccordoID"))

        qty = nz(row.get("qta_rilevante"))
        prezzo_netto = nz(row.get("PrezzoNettoUnit"))
        floor_price = nz(row.get("prezzo_atteso_minimo"))
        sconto_totale = nz(row.get("ScontoTotalePct"))
        soglia_ruolo = row.get("SogliaScontoRuoloPct")
        floor_gap_eur = nz(row.get("floor_gap_eur"))
        margine_pct = nz(row.get("MarginePct"))
        target_margin = nz(row.get("margine_target_eff"), 0.22)
        gm_gap_eur = nz(row.get("gm_gap_eur"))
        var_costo = nz(row.get("VariazioneCostoPct"))
        extra_discount = nz(row.get("ScontoExtraPct"))
        rischio_credito = clean_text(row.get("RischioCredito_cliente"))
        approved_exception = stato_deroga.lower().startswith("approv")
        active_promo = bool(row.get("ActivePromoFlag", False))
        active_contract = bool(row.get("ActiveAccordoFlag", False))

        if floor_price and prezzo_netto < floor_price:
            critical = floor_price > 0 and ((floor_price - prezzo_netto) / floor_price) >= config.severe_floor_gap_pct
            cases.append(
                _make_case(
                    row=row,
                    rule_code="PRICE_BELOW_FLOOR",
                    rule_name="Prezzo netto sotto floor",
                    reason=f"Prezzo netto {prezzo_netto:.2f} < floor {floor_price:.2f} su qty {qty:.0f}",
                    owner=choose_owner(row, "PRICE_BELOW_FLOOR"),
                    margin_risk_eur=floor_gap_eur,
                    severity=severity_from_risk(floor_gap_eur, config, critical=critical, high=not critical),
                    suggested_action="Bloccare o riesaminare la riga; verificare se la deroga è realmente autorizzata.",
                    status="approved_exception" if approved_exception else "open_review",
                )
            )

        discount_exception_allowed = (config.skip_discount_if_active_promo and active_promo) or (config.skip_discount_if_active_contract and active_contract)
        if not pd.isna(soglia_ruolo) and sconto_totale > (float(soglia_ruolo) + config.discount_role_buffer_pct):
            breach_gap = sconto_totale - float(soglia_ruolo)
            risk_eur = max(nz(row.get("discount_gap_eur")), nz(row.get("RicavoRiga")) * breach_gap)
            if breach_gap >= config.discount_role_min_gap_pct and not discount_exception_allowed:
                cases.append(
                    _make_case(
                        row=row,
                        rule_code="DISCOUNT_OVER_ROLE_THRESHOLD",
                        rule_name="Sconto oltre soglia ruolo",
                        reason=f"Sconto totale {sconto_totale:.1%} > soglia ruolo {float(soglia_ruolo):.1%}",
                        owner=choose_owner(row, "DISCOUNT_OVER_ROLE_THRESHOLD"),
                        margin_risk_eur=round(risk_eur, 2),
                        severity=severity_from_risk(risk_eur, config, high=breach_gap >= 0.03),
                        suggested_action="Verificare escalation e approvatore corretto; se il caso è legittimo, aggiornare soglia o policy.",
                        status="approved_exception" if approved_exception else "open_review",
                    )
                )

        uncovered_extra = (
            extra_discount > config.extra_discount_tolerance
            and not promo_id
            and not accordo_id
            and not motivo_deroga
        )
        if uncovered_extra:
            risk_eur = max(nz(row.get("discount_gap_eur")), extra_discount * nz(row.get("RicavoRiga")))
            cases.append(
                _make_case(
                    row=row,
                    rule_code="DISCOUNT_WITHOUT_COVERAGE",
                    rule_name="Sconto extra senza copertura",
                    reason=(
                        f"Sconto extra {extra_discount:.1%} > tolleranza {config.extra_discount_tolerance:.1%} "
                        "senza promo/accordo/motivo deroga."
                    ),
                    owner=choose_owner(row, "DISCOUNT_WITHOUT_COVERAGE"),
                    margin_risk_eur=round(risk_eur, 2),
                    severity=severity_from_risk(risk_eur, config, high=extra_discount >= 0.05),
                    suggested_action="Richiedere motivazione commerciale o riallineare il prezzo alla policy.",
                    status="open_review",
                )
            )

        low_margin = margine_pct < max(config.low_margin_floor, target_margin - config.low_margin_gap_vs_target)
        if tipo_documento.lower() == "fattura" and low_margin and gm_gap_eur >= config.low_margin_min_risk_eur:
            margin_gap = target_margin - margine_pct
            cases.append(
                _make_case(
                    row=row,
                    rule_code="LOW_MARGIN_VS_TARGET",
                    rule_name="Margine sotto target",
                    reason=f"Margine {margine_pct:.1%} < target effettivo {target_margin:.1%}",
                    owner=choose_owner(row, "LOW_MARGIN_VS_TARGET"),
                    margin_risk_eur=round(gm_gap_eur, 2),
                    severity=severity_from_risk(gm_gap_eur, config, high=margin_gap >= 0.08),
                    suggested_action="Verificare combinazione prezzo/costo/logistica; valutare repricing o blocco eccezioni simili.",
                    status="approved_exception" if approved_exception else "open_review",
                )
            )

        cost_risk = (
            tipo_documento.lower() == "fattura"
            and var_costo >= config.cost_increase_trigger
            and margine_pct < target_margin
            and not active_promo
            and not active_contract
        )
        if cost_risk:
            pass_through_risk_eur = round(max(gm_gap_eur, nz(row.get("RicavoRiga")) * max(target_margin - margine_pct, 0)), 2)
            if pass_through_risk_eur >= config.cost_pass_min_risk_eur:
                cases.append(
                    _make_case(
                        row=row,
                        rule_code="COST_PASS_THROUGH_RISK",
                        rule_name="Aumento costo non trasferito",
                        reason=f"Variazione costo {var_costo:.1%} con margine {margine_pct:.1%} sotto target {target_margin:.1%}",
                        owner=choose_owner(row, "COST_PASS_THROUGH_RISK"),
                        margin_risk_eur=pass_through_risk_eur,
                        severity=severity_from_risk(pass_through_risk_eur, config, high=var_costo >= 0.08),
                        suggested_action="Rivedere listino / floor / accordi per recuperare il delta costo.",
                        status="approved_exception" if approved_exception else "open_review",
                    )
                )

        if tipo_documento.lower() == "notacredito":
            margin_loss = abs(nz(row.get("MargineContributivo")))
            if margin_loss >= config.return_note_min_abs_value_eur:
                reason_bits = ["Nota credito emessa"]
                if motivo_deroga:
                    reason_bits.append(f"Motivo deroga: {motivo_deroga}")
                cases.append(
                    _make_case(
                        row=row,
                        rule_code="RETURN_OR_CREDIT_NOTE",
                        rule_name="Nota credito / reso",
                        reason="; ".join(reason_bits),
                        owner=choose_owner(row, "RETURN_OR_CREDIT_NOTE"),
                        margin_risk_eur=round(margin_loss, 2),
                        severity=severity_from_risk(margin_loss, config, high=margin_loss >= config.medium_risk_value_eur),
                        suggested_action="Analizzare root cause del reso e collegare il caso all’ordine originale.",
                        status="open_review",
                    )
                )

        credit_condition = (extra_discount > config.credit_risk_min_discount_pct) or low_margin
        if (not config.high_risk_only_credit_rule or rischio_credito.lower() == "alto") and credit_condition:
            credit_risk = max(nz(row.get("RicavoRiga")) * 0.02, nz(row.get("discount_gap_eur")))
            cases.append(
                _make_case(
                    row=row,
                    rule_code="CREDIT_RISK_ON_CONCESSION",
                    rule_name="Concessione commerciale su cliente rischio alto",
                    reason="Cliente a rischio credito alto con sconto/margine fragile.",
                    owner=choose_owner(row, "CREDIT_RISK_ON_CONCESSION"),
                    margin_risk_eur=round(credit_risk, 2),
                    severity=severity_from_risk(credit_risk, config, high=True),
                    suggested_action="Confermare coerenza tra concessione, rischio credito e fido disponibile.",
                    status="open_review",
                )
            )

    cases_df = pd.DataFrame(cases)
    if cases_df.empty:
        columns = [
            "CaseID", "NaturalCaseKey", "RuleCode", "RuleName", "Severity", "Status", "Owner",
            "MarginRiskEUR", "Reason", "SuggestedAction",
        ] + REVIEW_COLUMNS
        return pd.DataFrame(columns=columns)

    cases_df = cases_df.sort_values(by=["SeverityRank", "MarginRiskEUR", "DataDocumento"], ascending=[False, False, True]).reset_index(drop=True)
    cases_df["CaseID"] = [build_case_id(i + 1) for i in range(len(cases_df))]
    for col in REVIEW_COLUMNS:
        if col not in cases_df.columns:
            cases_df[col] = pd.NA
    return cases_df


def _make_case(
    row: pd.Series,
    rule_code: str,
    rule_name: str,
    reason: str,
    owner: str,
    margin_risk_eur: float,
    severity: str,
    suggested_action: str,
    status: str,
) -> dict:
    severity_rank = {"critical": 4, "high": 3, "medium": 2, "low": 1}.get(severity, 0)

    out = {
        "CaseID": None,
        "NaturalCaseKey": build_natural_case_key(row, rule_code),
        "RuleCode": rule_code,
        "RuleName": rule_name,
        "Severity": severity,
        "SeverityRank": severity_rank,
        "Status": status,
        "Owner": owner,
        "CaseBucket": choose_bucket(rule_code),
        "HeuristicReviewHint": (
            "likely_true_issue" if rule_code == "PRICE_BELOW_FLOOR"
            else "likely_legitimate_exception" if rule_code == "DISCOUNT_OVER_ROLE_THRESHOLD" and (bool(row.get("ActivePromoFlag", False)) or bool(row.get("ActiveAccordoFlag", False)) or clean_text(row.get("StatoDeroga")).lower().startswith("approv"))
            else "needs_human_review" if rule_code in {"LOW_MARGIN_VS_TARGET", "COST_PASS_THROUGH_RISK"}
            else "separate_bucket_review" if rule_code in {"CREDIT_RISK_ON_CONCESSION", "RETURN_OR_CREDIT_NOTE"}
            else "needs_human_review"
        ),
        "MarginRiskEUR": round(margin_risk_eur, 2),
        "Reason": reason,
        "SuggestedAction": suggested_action,
        "DataDocumento": row.get("DataDocumento"),
        "NumeroOrdine": row.get("NumeroOrdine"),
        "RigaOrdine": row.get("RigaOrdine"),
        "TipoDocumento": row.get("TipoDocumento"),
        "NumeroFattura": row.get("NumeroFattura"),
        "NumeroNotaCredito": row.get("NumeroNotaCredito"),
        "ClienteID": row.get("ClienteID"),
        "Cliente": row.get("Cliente"),
        "GruppoCliente": row.get("GruppoCliente"),
        "Regione": row.get("Regione"),
        "Provincia": row.get("Provincia"),
        "Venditore": row.get("Venditore"),
        "CanaleVendita": row.get("CanaleVendita"),
        "ProdottoID": row.get("ProdottoID"),
        "Prodotto": row.get("Prodotto"),
        "Categoria": row.get("Categoria"),
        "Qta": row.get("qta_rilevante"),
        "PrezzoListinoUnit": row.get("PrezzoListinoUnit"),
        "PrezzoNettoUnit": row.get("PrezzoNettoUnit"),
        "FloorPriceUnit": row.get("prezzo_atteso_minimo"),
        "ScontoTotalePct": row.get("ScontoTotalePct"),
        "SogliaScontoRuoloPct": row.get("SogliaScontoRuoloPct"),
        "MarginePct": row.get("MarginePct"),
        "MargineTargetPct": row.get("margine_target_eff"),
        "VariazioneCostoPct": row.get("VariazioneCostoPct"),
        "PromoID": row.get("PromoID"),
        "AccordoID": row.get("AccordoID"),
        "MotivoDeroga": row.get("MotivoDeroga"),
        "StatoDeroga": row.get("StatoDeroga"),
        "ApprovatoDa": row.get("ApprovatoDa"),
        "ActivePromoFlag": row.get("ActivePromoFlag"),
        "ActiveAccordoFlag": row.get("ActiveAccordoFlag"),
    }
    for col in REVIEW_COLUMNS:
        out[col] = pd.NA
    return out
