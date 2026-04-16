
from __future__ import annotations

import io
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
ENGINE_SCRIPT = BASE_DIR / "controllo_margini_mvp_v3.py"
CONFIG_PATH = BASE_DIR / "rules_config_v3.json"
TEMPLATE_PATH = BASE_DIR / "MarginRail_Input_Template_v1.xlsx"

st.set_page_config(page_title="MarginRail v1", page_icon="📈", layout="wide")


def eur(x: float) -> str:
    try:
        return f"€{float(x):,.0f}"
    except Exception:
        return "€0"


def build_zip_bytes(folder: Path) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in folder.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, arcname=file_path.relative_to(folder))
    buffer.seek(0)
    return buffer.getvalue()


def run_marginrail(uploaded_bytes: bytes, original_name: str) -> dict:
    with tempfile.TemporaryDirectory(prefix="marginrail_") as tmp:
        tmp_dir = Path(tmp)
        input_path = tmp_dir / original_name
        output_dir = tmp_dir / "output"
        input_path.write_bytes(uploaded_bytes)

        cmd = [
            sys.executable,
            str(ENGINE_SCRIPT),
            "--input",
            str(input_path),
            "--output-dir",
            str(output_dir),
            "--config",
            str(CONFIG_PATH),
        ]

        proc = subprocess.run(cmd, capture_output=True, text=True)

        if proc.returncode != 0:
            raise RuntimeError(proc.stderr or proc.stdout or "Errore durante l'analisi.")

        kpi_path = output_dir / "kpi_controllo_margini_v3.json"
        cases_path = output_dir / "casi_controllo_margini_v3.csv"
        review_path = output_dir / "review_pack_step1.csv"
        report_path = output_dir / "report_controllo_margini_v3.xlsx"
        config_used_path = output_dir / "config_effettiva_usata.json"

        if not kpi_path.exists() or not cases_path.exists():
            raise RuntimeError("Output non generati correttamente.")

        with kpi_path.open("r", encoding="utf-8") as f:
            kpis = json.load(f)

        cases_df = pd.read_csv(cases_path)
        review_df = pd.read_csv(review_path) if review_path.exists() else pd.DataFrame()

        top_rules = (
            cases_df.groupby("RuleCode", dropna=False)
            .agg(Casi=("CaseID", "count"), RischioEUR=("MarginRiskEUR", "sum"))
            .reset_index()
            .sort_values(["RischioEUR", "Casi"], ascending=[False, False])
        )

        top_cases = cases_df.sort_values("MarginRiskEUR", ascending=False).head(50)

        zip_bytes = build_zip_bytes(output_dir)

        files_bytes = {}
        for p in [cases_path, review_path, report_path, kpi_path, config_used_path]:
            if p.exists():
                files_bytes[p.name] = p.read_bytes()

        return {
            "kpis": kpis,
            "cases_df": cases_df,
            "review_df": review_df,
            "top_rules_df": top_rules,
            "top_cases_df": top_cases,
            "zip_bytes": zip_bytes,
            "files_bytes": files_bytes,
            "stdout": proc.stdout,
        }


def render_results(results: dict) -> None:
    kpis = results["kpis"]
    cases_df = results["cases_df"]
    top_rules_df = results["top_rules_df"]
    top_cases_df = results["top_cases_df"]

    st.success("Analisi completata.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Casi totali", f"{int(kpis.get('totale_casi', 0)):,}")
    c2.metric("Rischio totale", eur(kpis.get("totale_rischio_eur", 0)))
    c3.metric("Critical", f"{int(kpis.get('casi_critical', 0)):,}")
    c4.metric("High", f"{int(kpis.get('casi_high', 0)):,}")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Righe vendita", f"{int(kpis.get('totale_righe_vendita', 0)):,}")
    c6.metric("Clienti coinvolti", f"{int(kpis.get('clienti_coinvolti', 0)):,}")
    c7.metric("Ordini coinvolti", f"{int(kpis.get('ordini_coinvolti', 0)):,}")
    c8.metric("Regole attive", f"{int(kpis.get('regole_attive', 0)):,}")

    st.subheader("Top regole per rischio")
    st.dataframe(top_rules_df.head(10), use_container_width=True, hide_index=True)

    st.subheader("Top casi")
    severity_values = sorted([x for x in cases_df["Severity"].dropna().unique().tolist()])
    rule_values = sorted([x for x in cases_df["RuleCode"].dropna().unique().tolist()])

    f1, f2 = st.columns(2)
    selected_severity = f1.multiselect("Severity", severity_values, default=severity_values)
    selected_rules = f2.multiselect("RuleCode", rule_values, default=rule_values[: min(6, len(rule_values))])

    filtered = cases_df.copy()
    if selected_severity:
        filtered = filtered[filtered["Severity"].isin(selected_severity)]
    if selected_rules:
        filtered = filtered[filtered["RuleCode"].isin(selected_rules)]

    filtered = filtered.sort_values("MarginRiskEUR", ascending=False)
    show_cols = [
        c for c in [
            "CaseID", "RuleCode", "Severity", "Owner", "MarginRiskEUR",
            "Cliente", "Prodotto", "NumeroOrdine", "Reason", "SuggestedAction"
        ] if c in filtered.columns
    ]
    st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True, height=450)

    st.subheader("Download output")
    col_a, col_b, col_c = st.columns(3)
    col_a.download_button(
        "Scarica pacchetto completo (.zip)",
        data=results["zip_bytes"],
        file_name="marginrail_output_pack.zip",
        mime="application/zip",
        use_container_width=True,
    )
    files_bytes = results["files_bytes"]
    if "report_controllo_margini_v3.xlsx" in files_bytes:
        col_b.download_button(
            "Scarica report Excel",
            data=files_bytes["report_controllo_margini_v3.xlsx"],
            file_name="MarginRail_Executive_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    if "casi_controllo_margini_v3.csv" in files_bytes:
        col_c.download_button(
            "Scarica casi CSV",
            data=files_bytes["casi_controllo_margini_v3.csv"],
            file_name="MarginRail_Cases.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with st.expander("Log tecnico run"):
        st.code(results["stdout"] or "Nessun log disponibile.")


st.title("MarginRail v1")
st.caption("Upload Excel standardizzato → analisi automatica → dashboard + report scaricabili")

left, right = st.columns([2, 1])

with left:
    st.subheader("1) Carica il file Excel")
    uploaded_file = st.file_uploader("Excel cliente", type=["xlsx"])
    run_clicked = st.button("Esegui analisi", type="primary", use_container_width=True, disabled=uploaded_file is None)

with right:
    st.subheader("Template input")
    if TEMPLATE_PATH.exists():
        st.download_button(
            "Scarica template Excel",
            data=TEMPLATE_PATH.read_bytes(),
            file_name=TEMPLATE_PATH.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.info(
        "Il file deve usare il template ufficiale con i fogli richiesti. "
        "La V1 è pensata per input standardizzati."
    )

if run_clicked and uploaded_file is not None:
    with st.spinner("Sto eseguendo l'analisi..."):
        try:
            st.session_state["mr_results"] = run_marginrail(uploaded_file.getvalue(), uploaded_file.name)
        except Exception as e:
            st.error(str(e))

if "mr_results" in st.session_state:
    render_results(st.session_state["mr_results"])
