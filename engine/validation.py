
from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict

import pandas as pd

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    from .constants import DEFAULT_SHEETS, NUMERIC_COLUMNS
    from .utils import safe_read_excel, to_datetime, to_numeric
except ImportError:  # pragma: no cover
    from constants import DEFAULT_SHEETS, NUMERIC_COLUMNS
    from utils import safe_read_excel, to_datetime, to_numeric

from input_validation import InputValidationError, WORKBOOK_SCHEMA, ensure_optional_columns


def load_workbook_data(file_path: Path) -> Dict[str, pd.DataFrame]:
    return {key: safe_read_excel(file_path, sheet_name) for key, sheet_name in DEFAULT_SHEETS.items()}


def build_base_dataframe(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    vendite = data["vendite"].copy()
    if vendite.empty:
        raise InputValidationError("Foglio Vendite_2025 presente ma senza righe dati. MarginRail ha bisogno di almeno una riga vendita per eseguire l’analisi.")
    vendite = ensure_optional_columns(vendite, WORKBOOK_SCHEMA["Vendite_2025"].optional_columns)
    vendite = to_numeric(vendite, NUMERIC_COLUMNS)
    vendite = to_datetime(vendite, ["DataDocumento", "DataOrdine"])
    vendite["RigaOrdine"] = pd.to_numeric(vendite["RigaOrdine"], errors="coerce").fillna(0).astype(int)
    workflow = data["workflow"].copy()
    if not workflow.empty:
        workflow = ensure_optional_columns(workflow, WORKBOOK_SCHEMA["Workflow_Deroghe_2025"].optional_columns)
        workflow["RigaOrdine"] = pd.to_numeric(workflow["RigaOrdine"], errors="coerce").fillna(0).astype(int)
        workflow = to_numeric(workflow, ["SogliaScontoRuoloPct"])
        workflow = to_datetime(workflow, ["DataApprovazione"])
        workflow = workflow.drop_duplicates(subset=["NumeroOrdine", "RigaOrdine"])
        vendite = vendite.merge(workflow[["NumeroOrdine", "RigaOrdine", "SogliaScontoRuoloPct", "StatoDeroga", "ApprovatoDa", "DataApprovazione"]], on=["NumeroOrdine", "RigaOrdine"], how="left")
    else:
        vendite["SogliaScontoRuoloPct"] = pd.NA; vendite["StatoDeroga"] = pd.NA; vendite["ApprovatoDa"] = pd.NA; vendite["DataApprovazione"] = pd.NaT
    clienti = data["clienti"].copy()
    if not clienti.empty:
        clienti = ensure_optional_columns(clienti, WORKBOOK_SCHEMA["Clienti"].optional_columns)
        clienti = clienti.rename(columns={"ScontoBase": "ScontoBase_cliente", "RischioCredito": "RischioCredito_cliente", "AreaCommerciale": "AreaCommerciale_cliente"})
        clienti = to_numeric(clienti, ["ScontoBase_cliente"])
        clienti = clienti.drop_duplicates(subset=["ClienteID"])
        vendite = vendite.merge(clienti[["ClienteID", "ScontoBase_cliente", "RischioCredito_cliente", "AreaCommerciale_cliente"]], on="ClienteID", how="left")
    else:
        vendite["ScontoBase_cliente"] = pd.NA; vendite["RischioCredito_cliente"] = pd.NA; vendite["AreaCommerciale_cliente"] = pd.NA
    prodotti = data["prodotti"].copy()
    if not prodotti.empty:
        prodotti = ensure_optional_columns(prodotti, WORKBOOK_SCHEMA["Prodotti"].optional_columns)
        prodotti = prodotti.rename(columns={"MargineTarget": "MargineTarget_prodotto", "ClasseBrand": "ClasseBrand_prodotto", "StatoProdotto": "StatoProdotto_prodotto"})
        prodotti = to_numeric(prodotti, ["MargineTarget_prodotto"])
        prodotti = prodotti.drop_duplicates(subset=["ProdottoID"])
        vendite = vendite.merge(prodotti[["ProdottoID", "MargineTarget_prodotto", "ClasseBrand_prodotto", "StatoProdotto_prodotto"]], on="ProdottoID", how="left")
    else:
        vendite["MargineTarget_prodotto"] = pd.NA; vendite["ClasseBrand_prodotto"] = pd.NA; vendite["StatoProdotto_prodotto"] = pd.NA
    accordi = data["accordi"].copy()
    if not accordi.empty and "AccordoID" in vendite.columns:
        accordi = ensure_optional_columns(accordi, WORKBOOK_SCHEMA["Accordi_Commerciali"].optional_columns)
        accordi = accordi.rename(columns={"PrezzoContrattualeUnit": "PrezzoContrattualeUnit_acc", "FloorPriceUnit": "FloorPriceUnit_acc", "ScontoContrattualePct": "ScontoContrattualePct_acc", "StatoAccordo": "StatoAccordo_acc", "ValidoDa": "ValidoDa_acc", "ValidoA": "ValidoA_acc"})
        accordi = to_numeric(accordi, ["PrezzoContrattualeUnit_acc", "FloorPriceUnit_acc", "ScontoContrattualePct_acc"])
        accordi = to_datetime(accordi, ["ValidoDa_acc", "ValidoA_acc"])
        accordi = accordi.drop_duplicates(subset=["AccordoID"])
        vendite = vendite.merge(accordi[["AccordoID", "PrezzoContrattualeUnit_acc", "FloorPriceUnit_acc", "ScontoContrattualePct_acc", "StatoAccordo_acc", "ValidoDa_acc", "ValidoA_acc"]], on="AccordoID", how="left")
    else:
        vendite["PrezzoContrattualeUnit_acc"] = pd.NA; vendite["FloorPriceUnit_acc"] = pd.NA; vendite["ScontoContrattualePct_acc"] = pd.NA; vendite["StatoAccordo_acc"] = pd.NA; vendite["ValidoDa_acc"] = pd.NaT; vendite["ValidoA_acc"] = pd.NaT
    promo = data["promo"].copy()
    if not promo.empty and "PromoID" in vendite.columns:
        promo = ensure_optional_columns(promo, WORKBOOK_SCHEMA["Promo_2025"].optional_columns)
        promo = promo.rename(columns={"ScontoExtraPct": "ScontoExtraPct_promo", "MotivoPromo": "MotivoPromo_promo", "DataInizio": "DataInizio_promo", "DataFine": "DataFine_promo", "CanaleValido": "CanaleValido_promo"})
        promo = to_numeric(promo, ["ScontoExtraPct_promo"])
        promo = to_datetime(promo, ["DataInizio_promo", "DataFine_promo"])
        promo = promo.drop_duplicates(subset=["PromoID"])
        vendite = vendite.merge(promo[["PromoID", "ScontoExtraPct_promo", "MotivoPromo_promo", "DataInizio_promo", "DataFine_promo", "CanaleValido_promo"]], on="PromoID", how="left")
    else:
        vendite["ScontoExtraPct_promo"] = pd.NA; vendite["MotivoPromo_promo"] = pd.NA; vendite["DataInizio_promo"] = pd.NaT; vendite["DataFine_promo"] = pd.NaT; vendite["CanaleValido_promo"] = pd.NA
    vendite = to_numeric(vendite, NUMERIC_COLUMNS)
    vendite = to_datetime(vendite, ["DataDocumento", "DataOrdine", "DataApprovazione", "ValidoDa_acc", "ValidoA_acc", "DataInizio_promo", "DataFine_promo"])
    vendite["qta_rilevante"] = vendite["QtaDocumento"].fillna(vendite["QtaOrdinata"]).fillna(0)
    vendite["prezzo_atteso_minimo"] = vendite["FloorPriceUnit_acc"].fillna(vendite["FloorPriceUnit"])
    vendite["margine_target_eff"] = vendite["MargineTarget_prodotto"].fillna(0.22)
    data_doc = pd.to_datetime(vendite["DataDocumento"], errors="coerce")
    vendite["ActiveAccordoFlag"] = (vendite.get("AccordoID").notna() & vendite.get("StatoAccordo_acc", pd.Series(index=vendite.index, dtype=object)).fillna("").astype(str).str.lower().eq("attivo") & (vendite.get("ValidoDa_acc").isna() | (data_doc >= pd.to_datetime(vendite.get("ValidoDa_acc"), errors="coerce"))) & (vendite.get("ValidoA_acc").isna() | (data_doc <= pd.to_datetime(vendite.get("ValidoA_acc"), errors="coerce"))))
    canale_promo = vendite.get("CanaleValido_promo", pd.Series(index=vendite.index, dtype=object)).fillna("").astype(str).str.lower()
    canale_row = vendite.get("CanaleVendita", pd.Series(index=vendite.index, dtype=object)).fillna("").astype(str).str.lower()
    vendite["ActivePromoFlag"] = (vendite.get("PromoID").notna() & (vendite.get("DataInizio_promo").isna() | (data_doc >= pd.to_datetime(vendite.get("DataInizio_promo"), errors="coerce"))) & (vendite.get("DataFine_promo").isna() | (data_doc <= pd.to_datetime(vendite.get("DataFine_promo"), errors="coerce"))) & ((canale_promo == "") | (canale_promo == canale_row)))
    vendite["discount_gap_eur"] = (vendite["ScontoExtraPct"].fillna(0) * vendite["RicavoRiga"].fillna(0)).round(2)
    vendite["floor_gap_eur"] = ((vendite["prezzo_atteso_minimo"].fillna(0) - vendite["PrezzoNettoUnit"].fillna(0)).clip(lower=0) * vendite["qta_rilevante"].fillna(0)).round(2)
    vendite["gm_gap_eur"] = ((vendite["margine_target_eff"].fillna(0.22) - vendite["MarginePct"].fillna(0)).clip(lower=0) * vendite["RicavoRiga"].fillna(0)).round(2)
    for col in ["Cliente", "Venditore", "Prodotto", "Categoria"]:
        if col not in vendite.columns:
            vendite[col] = pd.NA
    return vendite
