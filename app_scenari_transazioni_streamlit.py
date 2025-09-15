# app_scenari_transazioni_streamlit.py
# -------------------------------------------------------------
# Web app Streamlit per analizzare due file Excel con intestazione
# alla TERZA riga e generare fogli di risultato per due scenari.
# -------------------------------------------------------------
# Requisiti: streamlit, pandas, openpyxl, xlsxwriter
# Avvio:  streamlit run app_scenari_transazioni_streamlit.py
# -------------------------------------------------------------

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analisi Inbound/Outbound – Scenari", layout="wide")

# ---------------------- Utility ----------------------
HEADER_ROW_INDEX = 2  # 0-based -> terza riga

@st.cache_data(show_spinner=False)
def read_excel_third_header(file_bytes: bytes, filename: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    df = pd.read_excel(bio, header=HEADER_ROW_INDEX, engine="openpyxl")
    df = df.dropna(how="all")
    df.columns = (
        df.columns
        .map(lambda c: str(c).strip())
        .map(lambda c: re.sub(r"\s+", " ", c))
    )
    df.insert(0, "Source File", filename)
    df.insert(1, "Source Row", range(HEADER_ROW_INDEX + 2, HEADER_ROW_INDEX + 2 + len(df)))
    return df

def infer_direction_from_name(filename: str) -> Optional[str]:
    name = filename.lower()
    if "inbound" in name:
        return "Inbound"
    if "outbound" in name:
        return "Outbound"
    return None

# ---------------------- Scenari e funzioni helper ----------------------
DATE_CANDIDATES = ["Input Date (dd Mon yyyy)", "Input Date", "Date", "InputDate"]
COUNTRY_CANDIDATES = ["Counterparty Country Name","Creditor FI Country Name","Debtor FI Country Name","Country"]
AMOUNT_CANDIDATES_EUR_PREF = ["Amount received (converted in EUR)","Amount sent (converted in EUR)"]
AMOUNT_CANDIDATES_FALLBACK = ["Amount received", "Amount sent", "Amount"]
CURRENCY_CANDIDATES = ["Currency Code", "Currency", "Ccy"]
RISK_RATING_CANDIDATES = ["Counterparty Country Risk Rating","Creditor FI Country Risk Rating","Debtor FI Country Risk Rating","Country Risk Rating"]

def pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def normalize_and_prepare(df: pd.DataFrame):
    df = df.copy()
    date_col = pick_col(df, DATE_CANDIDATES)
    country_col = pick_col(df, COUNTRY_CANDIDATES)
    currency_col = pick_col(df, CURRENCY_CANDIDATES)
    amount_col = pick_col(df, AMOUNT_CANDIDATES_EUR_PREF) or pick_col(df, AMOUNT_CANDIDATES_FALLBACK)

    if date_col is not None:
        def _parse(x):
            if pd.isna(x):
                return pd.NaT
            for fmt in ("%d %b %Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    return datetime.strptime(str(x), fmt)
                except Exception:
                    continue
            return pd.to_datetime(x, errors="coerce")
        dt = df[date_col].apply(_parse)
        df["Month"] = dt.dt.to_period("M").astype(str)
    else:
        df["Month"] = None

    if amount_col is not None:
        df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce")
    return df, date_col, country_col, amount_col, currency_col

# ---- Helper: derive_file_highrisk_map ----
def derive_file_highrisk_map(inbound_df, outbound_df) -> dict:
    """
    Restituisce una mappa Country -> True se nei dati (inbound/outbound)
    il Paese risulta 'High' in almeno una riga, usando le colonne di rating se presenti.
    """
    import pandas as pd

    pairs = [
        ("Creditor FI Country Name",  "Creditor FI Country Risk Rating"),
        ("Debtor FI Country Name",    "Debtor FI Country Risk Rating"),
        ("Counterparty Country Name", "Counterparty Country Risk Rating"),
        ("Country",                   "Country Risk Rating"),
    ]

    res = {}
    for df in (inbound_df, outbound_df):
        if df is None or getattr(df, "empty", True):
            continue
        for country_col, risk_col in pairs:
            if country_col in df.columns and risk_col in df.columns:
                tmp = df[[country_col, risk_col]].dropna()
                if tmp.empty:
                    continue
                mask = (
                    tmp[risk_col]
                    .astype(str).str.strip().str.lower()
                    .str.contains("high")
                )
                if mask.any():
                    for c in (
                        tmp.loc[mask, country_col]
                        .astype(str).str.strip()
                    ):
                        if c:
                            res[c] = True
    return res

    # Normalizza colonne
    cols = {c.lower().strip(): c for c in df.columns}
    if "country" not in cols:
        return {}
    country_col = cols["country"]
    highrisk_col = None
    if "highrisk" in cols:
        highrisk_col = cols["highrisk"]
    elif "risk" in cols:
        highrisk_col = cols["risk"]
        # mappa testuale High/Low
        df[highrisk_col] = df[highrisk_col].astype(str).str.strip().str.lower().map({"high": True, "low": False, "no": False, "norisk": False})
    if highrisk_col is None:
        return {}
    mapping = {}
    for _, r in df.iterrows():
        ctry = str(r[country_col]).strip()
        try:
            val = bool(r[highrisk_col])
        except Exception:
            val = False
        if ctry:
            mapping[ctry] = val
    return mapping
DATE_CANDIDATES = ["Input Date (dd Mon yyyy)", "Input Date", "Date", "InputDate"]
COUNTRY_CANDIDATES = ["Counterparty Country Name","Creditor FI Country Name","Debtor FI Country Name","Country"]
AMOUNT_CANDIDATES_EUR_PREF = ["Amount received (converted in EUR)","Amount sent (converted in EUR)"]
AMOUNT_CANDIDATES_FALLBACK = ["Amount received", "Amount sent", "Amount"]
CURRENCY_CANDIDATES = ["Currency Code", "Currency", "Ccy"]
RISK_RATING_CANDIDATES = ["Counterparty Country Risk Rating","Creditor FI Country Risk Rating","Debtor FI Country Risk Rating","Country Risk Rating"]

def pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def normalize_and_prepare(df: pd.DataFrame):
    df = df.copy()
    date_col = pick_col(df, DATE_CANDIDATES)
    country_col = pick_col(df, COUNTRY_CANDIDATES)
    currency_col = pick_col(df, CURRENCY_CANDIDATES)
    amount_col = pick_col(df, AMOUNT_CANDIDATES_EUR_PREF) or pick_col(df, AMOUNT_CANDIDATES_FALLBACK)

    if date_col is not None:
        def _parse(x):
            if pd.isna(x):
                return pd.NaT
            for fmt in ("%d %b %Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    return datetime.strptime(str(x), fmt)
                except Exception:
                    continue
            return pd.to_datetime(x, errors="coerce")
        dt = df[date_col].apply(_parse)
        df["Month"] = dt.dt.to_period("M").astype(str)
    else:
        df["Month"] = None

    if amount_col is not None:
        df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce")
    return df, date_col, country_col, amount_col, currency_col

def collect_highrisk_from_files(inbound_df: Optional[pd.DataFrame], outbound_df: Optional[pd.DataFrame]) -> set:
    """Ricava l'elenco di Paesi marcati High nei file, guardando le colonne di rischio se presenti.
    Ritorna un set di nomi Paese."""
    countries: set = set()
    def _extract(df: Optional[pd.DataFrame]):
        if df is None or df.empty:
            return
        for rcol in RISK_RATING_CANDIDATES:
            if rcol in df.columns:
                mask = df[rcol].astype(str).str.lower().str.contains("high", na=False)
                if mask.any():
                    for ccol in ["Creditor FI Country Name","Debtor FI Country Name","Counterparty Country Name","Country"]:
                        if ccol in df.columns:
                            countries.update(df.loc[mask, ccol].dropna().astype(str).unique().tolist())
    _extract(inbound_df)
    _extract(outbound_df)
    return countries

# ---------- Scenario A (sintesi) ----------
def _scenario_A_for_direction(df: pd.DataFrame, direction_label: str, pct_threshold: float, eur_threshold: float, top_n: int) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df, date_col, country_col, amount_col, _ = normalize_and_prepare(df)
    if country_col is None or amount_col is None or "Month" not in df.columns:
        return pd.DataFrame()

    valid_months = df.loc[df["Month"].notna(), "Month"]
    if valid_months.empty:
        return pd.DataFrame()

    current_month = sorted(valid_months.unique())[-1]
    months_sorted = sorted(valid_months.unique())
    prev_month = months_sorted[-2] if len(months_sorted) > 1 else None

    # Aggregazione importi e conteggi per mese/paese
    grp_amt = (
        df.groupby(["Month", country_col], dropna=False)[amount_col]
          .sum(min_count=1).reset_index()
          .rename(columns={country_col: "Country", amount_col: "Amount (EUR)"})
    )
    grp_cnt = (
        df.groupby(["Month", country_col], dropna=False)[amount_col]
          .count().reset_index()
          .rename(columns={country_col: "Country", amount_col: "Rows"})
    )
    grp = grp_amt.merge(grp_cnt, on=["Month","Country"], how="left")

    # Top-N paesi sul mese corrente per importo
    cur = grp[grp["Month"] == current_month]
    top_countries = (
        cur.sort_values("Amount (EUR)", ascending=False)
           .head(top_n)["Country"].tolist()
    )

    cur_sel = grp[(grp["Month"] == current_month) & (grp["Country"].isin(top_countries))]
    prev_sel = grp[(grp["Month"] == prev_month) & (grp["Country"].isin(top_countries))] if prev_month else grp.iloc[0:0]

    res = cur_sel.merge(
        prev_sel[["Country","Amount (EUR)","Rows"]],
        on="Country",
        how="left",
        suffixes=("", " (prev)")
    )
    res.rename(columns={"Amount (EUR) (prev)": "Prev Amount (EUR)", "Rows (prev)": "Rows Prev"}, inplace=True)
    res["Prev Amount (EUR)"] = res["Prev Amount (EUR)"].fillna(0.0)
    res["Rows Prev"] = res["Rows Prev"].fillna(0).astype(int)

    # Delta
    res["Δ Amount"] = res["Amount (EUR)"] - res["Prev Amount (EUR)"]
    res["Δ %"] = res.apply(lambda r: (r["Δ Amount"] / r["Prev Amount (EUR)"] * 100.0) if r["Prev Amount (EUR)"] > 0 else (100.0 if r["Amount (EUR)"] > 0 else 0.0), axis=1)
    res.rename(columns={"Rows": "Rows Curr"}, inplace=True)

    # Metadati
    res.insert(0, "Direction", direction_label)
    res.insert(2, "Prev Month", prev_month)
    res["Flag >= Thresholds"] = (res["Δ %"] >= pct_threshold) & (res["Amount (EUR)"] >= eur_threshold)

    # Ordina e ritorna
    res = res.sort_values(["Flag >= Thresholds","Amount (EUR)"], ascending=[False, False]).reset_index(drop=True)
    return res

def scenario_a(inbound_df, outbound_df, pct_threshold=20.0, eur_threshold=5_000_000.0, top_n=10):
    res_in = _scenario_A_for_direction(inbound_df, "Inbound", pct_threshold, eur_threshold, top_n)
    res_out = _scenario_A_for_direction(outbound_df, "Outbound", pct_threshold, eur_threshold, top_n)
    return pd.concat([res_in, res_out], ignore_index=True, sort=False)

# ---------- Scenario A (details) ----------
def scenario_a_details(inbound_df, outbound_df, pct_threshold=20.0, eur_threshold=5_000_000.0, top_n=10):
    def build_details(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, date_col, country_col, amount_col, _ = normalize_and_prepare(df)
        if country_col is None or amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        months = df2.loc[df2["Month"].notna(), "Month"]
        if months.empty:
            return pd.DataFrame()
        current_month = sorted(months.unique())[-1]
        grp = df2.groupby(["Month", country_col], dropna=False)[amount_col].sum(min_count=1).reset_index().rename(columns={country_col: "Country", amount_col: "Amount (EUR)"})
        top_countries = grp[grp["Month"] == current_month].sort_values("Amount (EUR)", ascending=False).head(top_n)["Country"].tolist()
        prev_month = sorted(months.unique())[-2] if len(months.unique()) > 1 else None
        mask = df2[country_col].isin(top_countries) & df2["Month"].isin([m for m in [current_month, prev_month] if m])
        out = df2.loc[mask].copy()
        out.insert(0, "Direction", direction)
        out.sort_values(["Direction", country_col, "Month"], inplace=True)
        return out
    din = build_details(inbound_df, "Inbound")
    dout = build_details(outbound_df, "Outbound")
    return pd.concat([din, dout], ignore_index=True, sort=False)

# ---------- Scenario B (sintesi) ----------
def scenario_b(inbound_df, outbound_df):
    def build(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df, date_col, country_col, amount_col, currency_col = normalize_and_prepare(df)
        if country_col is None or amount_col is None or "Month" not in df.columns:
            return pd.DataFrame()
        keys = ["Month", country_col]
        if currency_col: keys.append(currency_col)
        agg = (
            df.groupby(keys, dropna=False)[amount_col]
              .agg([("Amount (EUR)", "sum"), ("Rows", "count")])
              .reset_index()
              .rename(columns={country_col:"Country"})
        )
        agg.insert(0, "Direction", direction)
        return agg
    part_in = build(inbound_df, "Inbound")
    part_out = build(outbound_df, "Outbound")
    return pd.concat([part_in, part_out], ignore_index=True, sort=False)

# ---------- Scenario B (details) ----------
def scenario_b_details(inbound_df, outbound_df):
    def det(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, date_col, country_col, amount_col, currency_col = normalize_and_prepare(df)
        if country_col is None or amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        out = df2.copy()
        out.insert(0, "Direction", direction)
        return out
    din = det(inbound_df, "Inbound")
    dout = det(outbound_df, "Outbound")
    return pd.concat([din, dout], ignore_index=True, sort=False)

# ---------- Top Risk Countries ----------
def top_risk_countries_sheet(inbound_df, outbound_df):
    def extract(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, date_col, country_col, amount_col, _ = normalize_and_prepare(df)
        risk_col = pick_col(df2, RISK_RATING_CANDIDATES)
        if country_col is None or amount_col is None or risk_col is None:
            return pd.DataFrame()
        tmp = df2[["Month", country_col, amount_col, risk_col]].copy()
        tmp.rename(columns={country_col: "Country", amount_col: "Amount (EUR)", risk_col: "Risk Rating"}, inplace=True)
        tmp.insert(0, "Direction", direction)
        return tmp
    df_in = extract(inbound_df, "Inbound")
    df_out = extract(outbound_df, "Outbound")
    all_df = pd.concat([df_in, df_out], ignore_index=True, sort=False)
    if all_df.empty:
        return pd.DataFrame({"Info":["Nessuna colonna di rischio paese trovata."]})
    hr = all_df[all_df["Risk Rating"].astype(str).str.lower().str.contains("high")].copy()
    if hr.empty:
        return pd.DataFrame({"Info":["Nessun paese marcato High risk."]})
    latest_month = sorted(hr["Month"].dropna().unique())[-1]
    cur = hr[hr["Month"] == latest_month]
    res = cur.groupby(["Direction","Country"], dropna=False)["Amount (EUR)"].sum(min_count=1).reset_index().sort_values(["Direction","Amount (EUR)"], ascending=[True, False])
    res.insert(2, "Month", latest_month)
    return res

# ---------- By Country (A e B) + Role-specific ----------

def role_country_inbound(inbound_df, risk_override: Optional[dict] = None):
    """Aggrega INBOUND per 'Creditor FI Country Name' con Amount, Rows, HighRisk per mese.
    Indica esplicitamente lo **scope** e il **mese**.
    """
    if inbound_df is None or inbound_df.empty:
        return pd.DataFrame()
    df2, _, _, amount_col, _ = normalize_and_prepare(inbound_df)
    if amount_col is None or "Month" not in df2.columns:
        return pd.DataFrame()
    if "Creditor FI Country Name" not in df2.columns:
        return pd.DataFrame({"Info":["Colonna 'Creditor FI Country Name' non trovata"]})
    risk_col = "Creditor FI Country Risk Rating" if "Creditor FI Country Risk Rating" in df2.columns else None
    temp = df2.copy()
    if risk_override:
        temp["HighRisk"] = temp["Creditor FI Country Name"].astype(str).map(lambda x: bool(risk_override.get(x, False)))
    else:
        temp["HighRisk"] = temp[risk_col].astype(str).str.lower().str.contains("high") if risk_col else False
    agg = temp.groupby(["Month","Creditor FI Country Name"], dropna=False).agg(**{
        "Amount (EUR)": (amount_col, "sum"),
        "Rows": (amount_col, "count"),
        "HighRisk": ("HighRisk", "any"),
    }).reset_index().rename(columns={"Creditor FI Country Name":"Country"})
    agg.insert(0, "Direction", "Inbound")
    agg["Aggregation"] = "Role-based by month (Creditor FI Country)"
    agg["Risk Label"] = agg["HighRisk"].map(lambda x: "High Risk" if bool(x) else "No/Low Risk")
    agg.sort_values(["Month","Amount (EUR)"], ascending=[True, False], inplace=True)
    cols = ["Direction","Month","Country","Amount (EUR)","Rows","HighRisk","Risk Label","Aggregation"]
    agg = agg.reindex(columns=cols)
    return agg


def role_country_outbound(outbound_df, risk_override: Optional[dict] = None):
    """Aggrega OUTBOUND per 'Debtor FI Country Name' con Amount, Rows, HighRisk per mese.
    Indica esplicitamente lo **scope** e il **mese**.
    """
    if outbound_df is None or outbound_df.empty:
        return pd.DataFrame()
    df2, _, _, amount_col, _ = normalize_and_prepare(outbound_df)
    if amount_col is None or "Month" not in df2.columns:
        return pd.DataFrame()
    if "Debtor FI Country Name" not in df2.columns:
        return pd.DataFrame({"Info":["Colonna 'Debtor FI Country Name' non trovata"]})
    risk_col = "Debtor FI Country Risk Rating" if "Debtor FI Country Risk Rating" in df2.columns else None
    temp = df2.copy()
    if risk_override:
        temp["HighRisk"] = temp["Debtor FI Country Name"].astype(str).map(lambda x: bool(risk_override.get(x, False)))
    else:
        temp["HighRisk"] = temp[risk_col].astype(str).str.lower().str.contains("high") if risk_col else False
    agg = temp.groupby(["Month","Debtor FI Country Name"], dropna=False).agg(**{
        "Amount (EUR)": (amount_col, "sum"),
        "Rows": (amount_col, "count"),
        "HighRisk": ("HighRisk", "any"),
    }).reset_index().rename(columns={"Debtor FI Country Name":"Country"})
    agg.insert(0, "Direction", "Outbound")
    agg["Aggregation"] = "Role-based by month (Debtor FI Country)"
    agg["Risk Label"] = agg["HighRisk"].map(lambda x: "High Risk" if bool(x) else "No/Low Risk")
    agg.sort_values(["Month","Amount (EUR)"], ascending=[True, False], inplace=True)
    cols = ["Direction","Month","Country","Amount (EUR)","Rows","HighRisk","Risk Label","Aggregation"]
    agg = agg.reindex(columns=cols)
    return agg

# ---------- By Country (A e B) ----------
def by_country_scenario_a(inbound_df, outbound_df, top_n=10, risk_override: Optional[dict] = None):
    """Aggregato per Paese (preferendo colonne role-based):
    - INBOUND → 'Creditor FI Country Name' se presente
    - OUTBOUND → 'Debtor FI Country Name' se presente
    Include `Rows` e flag `HighRisk`. Indica esplicitamente il **mese di riferimento** e lo **scope** dell'aggregazione.
    """
    def build(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, _, _, amount_col, _ = normalize_and_prepare(df)
        if amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        if direction == "Inbound":
            country_col = "Creditor FI Country Name" if "Creditor FI Country Name" in df2.columns else pick_col(df2, COUNTRY_CANDIDATES)
            risk_col = "Creditor FI Country Risk Rating" if "Creditor FI Country Risk Rating" in df2.columns else pick_col(df2, RISK_RATING_CANDIDATES)
        else:
            country_col = "Debtor FI Country Name" if "Debtor FI Country Name" in df2.columns else pick_col(df2, COUNTRY_CANDIDATES)
            risk_col = "Debtor FI Country Risk Rating" if "Debtor FI Country Risk Rating" in df2.columns else pick_col(df2, RISK_RATING_CANDIDATES)
        if country_col is None:
            return pd.DataFrame()
        months = df2.loc[df2["Month"].notna(), "Month"]
        if months.empty:
            return pd.DataFrame()
        current_month = sorted(months.unique())[-1]
        cur = df2[df2["Month"] == current_month].copy()
        # High risk flag per riga: override UI > rating col > False
        if risk_override and country_col in cur.columns:
            cur["HighRisk"] = cur[country_col].astype(str).map(lambda x: bool(risk_override.get(x, False)))
        elif risk_col:
            cur["HighRisk"] = cur[risk_col].astype(str).str.lower().str.contains("high")
        else:
            cur["HighRisk"] = False
        agg = cur.groupby(country_col, dropna=False).agg(**{
            "Amount (EUR)": (amount_col, "sum"),
            "Rows": (amount_col, "count"),
            "HighRisk": ("HighRisk", "any"),
        }).reset_index().rename(columns={country_col: "Country"})
        agg.insert(0, "Direction", direction)
        agg.insert(1, "Ref Month", current_month)
        agg["Aggregation"] = "Top N current month (per direction)"
        agg["Risk Label"] = agg["HighRisk"].map(lambda x: "High Risk" if bool(x) else "No/Low Risk")
        agg.sort_values("Amount (EUR)", ascending=False, inplace=True)
        # Limita ai Top N e riordina colonne per chiarezza
        agg = agg.head(top_n)
        cols = [c for c in ["Direction","Ref Month","Country","Amount (EUR)","Rows","HighRisk","Risk Label","Aggregation"] if c in agg.columns]
        agg = agg.reindex(columns=cols)
        return agg
    a = build(inbound_df, "Inbound")
    b = build(outbound_df, "Outbound")
    return pd.concat([a, b], ignore_index=True, sort=False)

def by_country_scenario_b(inbound_df, outbound_df, risk_override: Optional[dict] = None):
    """Aggregato per Paese per ogni mese, preferendo colonne role-based; include `Rows` e `HighRisk`.
    Indica esplicitamente il **mese** e lo **scope** dell'aggregazione.
    """
    def build(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, _, _, amount_col, _ = normalize_and_prepare(df)
        if amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        if direction == "Inbound":
            country_col = "Creditor FI Country Name" if "Creditor FI Country Name" in df2.columns else pick_col(df2, COUNTRY_CANDIDATES)
            risk_col = "Creditor FI Country Risk Rating" if "Creditor FI Country Risk Rating" in df2.columns else pick_col(df2, RISK_RATING_CANDIDATES)
        else:
            country_col = "Debtor FI Country Name" if "Debtor FI Country Name" in df2.columns else pick_col(df2, COUNTRY_CANDIDATES)
            risk_col = "Debtor FI Country Risk Rating" if "Debtor FI Country Risk Rating" in df2.columns else pick_col(df2, RISK_RATING_CANDIDATES)
        if country_col is None:
            return pd.DataFrame()
        temp = df2.copy()
        if risk_override and country_col in temp.columns:
            temp["HighRisk"] = temp[country_col].astype(str).map(lambda x: bool(risk_override.get(x, False)))
        elif risk_col:
            temp["HighRisk"] = temp[risk_col].astype(str).str.lower().str.contains("high")
        else:
            temp["HighRisk"] = False
        agg = temp.groupby(["Month", country_col], dropna=False).agg(**{
            "Amount (EUR)": (amount_col, "sum"),
            "Rows": (amount_col, "count"),
            "HighRisk": ("HighRisk", "any"),
        }).reset_index().rename(columns={country_col: "Country"})
        agg.insert(0, "Direction", direction)
        agg["Aggregation"] = "Per-month (independent)"
        agg["Risk Label"] = agg["HighRisk"].map(lambda x: "High Risk" if bool(x) else "No/Low Risk")
        agg.sort_values(["Month", "Amount (EUR)"], ascending=[True, False], inplace=True)
        cols = [c for c in ["Direction","Month","Country","Amount (EUR)","Rows","HighRisk","Risk Label","Aggregation"] if c in agg.columns]
        agg = agg.reindex(columns=cols)
        return agg
    a = build(inbound_df, "Inbound")
    b = build(outbound_df, "Outbound")
    return pd.concat([a, b], ignore_index=True, sort=False)

# ---------- MoM sheets ----------

def build_mom_sheets(inbound_df, outbound_df):
    def prep(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, _, country_col, amount_col, _ = normalize_and_prepare(df)
        if country_col is None or amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        agg = df2.groupby(["Month", country_col], dropna=False).agg(**{
            "Amount": (amount_col, "sum"),
            "Rows": (amount_col, "count"),
        }).reset_index().rename(columns={country_col: "Country"})
        agg.insert(0, "Direction", direction)
        return agg
    ina = prep(inbound_df, "Inbound")
    oua = prep(outbound_df, "Outbound")
    base = pd.concat([ina, oua], ignore_index=True, sort=False)
    if base.empty or base["Month"].dropna().empty:
        return {}

    months_sorted = sorted(base["Month"].dropna().unique())
    sheets = {}
    for i in range(1, len(months_sorted)):
        prev_m = months_sorted[i-1]
        curr_m = months_sorted[i]
        left = base[base["Month"] == prev_m][["Direction","Country","Amount","Rows"]].rename(columns={"Amount":"Amount_prev","Rows":"Rows_prev"})
        right = base[base["Month"] == curr_m][["Direction","Country","Amount","Rows"]].rename(columns={"Amount":"Amount_curr","Rows":"Rows_curr"})
        merged = pd.merge(left, right, on=["Direction","Country"], how="outer")
        merged["Amount_prev"] = merged["Amount_prev"].fillna(0.0)
        merged["Amount_curr"] = merged["Amount_curr"].fillna(0.0)
        merged["Rows_prev"] = merged["Rows_prev"].fillna(0).astype(int)
        merged["Rows_curr"] = merged["Rows_curr"].fillna(0).astype(int)
        merged.insert(2, "Prev Month", prev_m)
        merged.insert(3, "Month", curr_m)
        merged["Δ Amount"] = merged["Amount_curr"] - merged["Amount_prev"]
        merged["Δ %"] = merged.apply(lambda r: (r["Δ Amount"]/r["Amount_prev"]*100.0) if r["Amount_prev"]>0 else (100.0 if r["Amount_curr"]>0 else 0.0), axis=1)
        merged["Δ Rows"] = merged["Rows_curr"] - merged["Rows_prev"]
        merged = merged[["Direction","Country","Prev Month","Month","Amount_prev","Amount_curr","Δ Amount","Δ %","Rows_prev","Rows_curr","Δ Rows"]]
        merged.sort_values(["Direction","Amount_curr"], ascending=[True, False], inplace=True)
        sheet_name = f"MoM_{curr_m}_vs_{prev_m}"
        sheets[sheet_name] = merged
    return sheets

# ---------- MoM detail sheets (righe originali) ----------

def build_mom_detail_sheets(inbound_df, outbound_df):
    """Crea, per ogni coppia mese t/t-1, un foglio con le **righe originali** usate nei due mesi,
    annotate con Direction e Month, in modo da verificare facilmente gli aggregati."""
    def prep_rows(df: pd.DataFrame, direction: str):
        if df is None or df.empty:
            return pd.DataFrame()
        df2, _, country_col, amount_col, _ = normalize_and_prepare(df)
        if country_col is None or amount_col is None or "Month" not in df2.columns:
            return pd.DataFrame()
        out = df2.copy()
        out.insert(0, "Direction", direction)
        return out

    rin = prep_rows(inbound_df, "Inbound")
    rout = prep_rows(outbound_df, "Outbound")
    rows = pd.concat([rin, rout], ignore_index=True, sort=False)
    if rows.empty or rows["Month"].dropna().empty:
        return {}

    months_sorted = sorted(rows["Month"].dropna().unique())
    details = {}
    for i in range(1, len(months_sorted)):
        prev_m = months_sorted[i-1]
        curr_m = months_sorted[i]
        subset = rows[rows["Month"].isin([prev_m, curr_m])].copy()
        # Ordina per Direction, Country se disponibile, quindi Month
        sort_keys = ["Direction"]
        if "Counterparty Country Name" in subset.columns:
            sort_keys.append("Counterparty Country Name")
        elif "Country" in subset.columns:
            sort_keys.append("Country")
        sort_keys.append("Month")
        subset.sort_values(sort_keys, inplace=True)
        sheet_name = f"MoM_Details_{curr_m}_vs_{prev_m}"
        details[sheet_name] = subset
    return details

# ---------- Role-based PIVOT with MoM Δ and % ----------

def _role_country_pivot(df: pd.DataFrame, direction: str, country_col: str, risk_col: Optional[str] = None, delta_threshold_eur: float = 5_000_000.0, risk_override: Optional[dict] = None) -> pd.DataFrame:
    """Tabella pivot (Paesi in riga, Mesi in colonna) con:
    - Amount per mese + Δ e Δ% mese-su-mese
    - Rows (conteggio righe) per mese + Δ e Δ% mese-su-mese
    - Flag HighRisk a livello Paese (true se almeno una riga del Paese ha rating High)
    - Flag soglia Δ>threshold (da UI) per Amount mese-su-mese
    Usa la policy di UI (override) se fornita; altrimenti usa i rating presenti nei file.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    df2, _, _, amount_col, _ = normalize_and_prepare(df)
    if amount_col is None or "Month" not in df2.columns or country_col not in df2.columns:
        return pd.DataFrame()

    # Calcolo mappa HighRisk per Country
    if risk_override:
        high_map = pd.Series({k: bool(v) for k, v in risk_override.items()})
    elif risk_col and risk_col in df2.columns:
        tmp_r = df2[[country_col, risk_col]].copy()
        tmp_r["HighRisk"] = tmp_r[risk_col].astype(str).str.lower().str.contains("high")
        high_map = tmp_r.groupby(country_col, dropna=False)["HighRisk"].any()
    else:
        high_map = pd.Series(dtype=bool)

    # Aggrega importi e righe per Month/Country
    agg = (
        df2.groupby(["Month", country_col], dropna=False)[amount_col]
           .agg(Amount="sum", Rows="count")
           .reset_index()
           .rename(columns={country_col: "Country"})
    )
    months = sorted(agg["Month"].dropna().unique())
    if not months:
        return pd.DataFrame()

    # Pivot per Amount e per Rows
    p_amt = agg.pivot_table(index="Country", columns="Month", values="Amount", aggfunc="sum").fillna(0.0)
    p_rows = agg.pivot_table(index="Country", columns="Month", values="Rows", aggfunc="sum").fillna(0.0)
    p_amt = p_amt.reindex(columns=months)
    p_rows = p_rows.reindex(columns=months)

    # Output base
    out = pd.DataFrame(index=p_amt.index)
    out.insert(0, "Direction", direction)
    out.insert(1, "Country", out.index)
    if not high_map.empty:
        out.insert(2, "HighRisk", out["Country"].map(high_map).fillna(False).astype(bool))
    else:
        out.insert(2, "HighRisk", False)
    # Etichetta rischio
    out.insert(3, "Risk Label", out["HighRisk"].map(lambda x: "High Risk" if bool(x) else "No/Low Risk"))

    # Etichetta soglia Δ>threshold per titolo colonne
    thr_label = f"{int(delta_threshold_eur):,}".replace(",", ".")

    # Aggiungi colonne per ogni mese
    prev = None
    for m in months:
        # Amount
        out[f"{m} Amount"] = p_amt[m]
        if prev is not None:
            d = p_amt[m] - p_amt[prev]
            denom = p_amt[prev].replace({0: pd.NA})
            pct = (d / denom) * 100.0
            pct = pct.fillna((p_amt[m] > 0).astype(float) * 100.0)
            out[f"{m} Δ Amount vs {prev}"] = d
            out[f"{m} Δ% Amount vs {prev}"] = pct
            out[f"{m} Δ>{thr_label} vs {prev}"] = (d > float(delta_threshold_eur))
        # Rows
        out[f"{m} Rows"] = p_rows[m].astype(int)
        if prev is not None:
            dr = (p_rows[m] - p_rows[prev]).astype(int)
            denom_r = p_rows[prev].replace({0: pd.NA})
            pr = (dr / denom_r) * 100.0
            pr = pr.fillna((p_rows[m] > 0).astype(float) * 100.0)
            out[f"{m} Δ Rows vs {prev}"] = dr
            out[f"{m} Δ% Rows vs {prev}"] = pr
        prev = m

    # Ordina per Amount dell'ultimo mese
    out = out.sort_values(f"{months[-1]} Amount", ascending=False).reset_index(drop=True)
    return out


def role_country_inbound_pivot(inbound_df: pd.DataFrame, delta_threshold_eur: float, risk_override: Optional[dict] = None) -> pd.DataFrame:
    if inbound_df is None or inbound_df.empty:
        return pd.DataFrame()
    col = "Creditor FI Country Name" if "Creditor FI Country Name" in inbound_df.columns else None
    if not col:
        return pd.DataFrame({"Info":["Colonna 'Creditor FI Country Name' non trovata"]})
    risk = "Creditor FI Country Risk Rating" if "Creditor FI Country Risk Rating" in inbound_df.columns else None
    return _role_country_pivot(inbound_df, "Inbound", col, risk, delta_threshold_eur, risk_override)


def role_country_outbound_pivot(outbound_df: pd.DataFrame, delta_threshold_eur: float, risk_override: Optional[dict] = None) -> pd.DataFrame:
    if outbound_df is None or outbound_df.empty:
        return pd.DataFrame()
    col = "Debtor FI Country Name" if "Debtor FI Country Name" in outbound_df.columns else None
    if not col:
        return pd.DataFrame({"Info":["Colonna 'Debtor FI Country Name' non trovata"]})
    risk = "Debtor FI Country Risk Rating" if "Debtor FI Country Risk Rating" in outbound_df.columns else None
    return _role_country_pivot(outbound_df, "Outbound", col, risk, delta_threshold_eur, risk_override)

# ---------- Utility: conditional formatting for pivot sheets ----------

def apply_pivot_conditional_formatting(writer, sheet_name: str, df: pd.DataFrame):
    """Apply conditional formatting to a pivot worksheet: green for positive deltas, red for negative,
    highlight boolean threshold hits and HighRisk."""
    if df is None or df.empty:
        return
    if sheet_name not in writer.sheets:
        return
    ws = writer.sheets[sheet_name]
    wb = writer.book

    nrows = len(df)
    ncols = len(df.columns)

    fmt_pos = wb.add_format({"font_color": "#1a7f37"})  # green
    fmt_neg = wb.add_format({"font_color": "#cf222e"})  # red
    fmt_bool = wb.add_format({"bg_color": "#FFF2CC"})   # light yellow
    fmt_hr   = wb.add_format({"bg_color": "#F4CCCC"})   # light red
    fmt_pct  = wb.add_format({"num_format": '0.0"%"'})

    # find columns by patterns
    for j, col in enumerate(df.columns):
        # Percent columns: set number format and color pos/neg
        if "Δ%" in col:
            ws.set_column(j, j, None, fmt_pct)
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": ">", "value": 0, "format": fmt_pos})
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": "<", "value": 0, "format": fmt_neg})
        # Amount / Rows deltas: color pos/neg
        if "Δ Amount" in col or "Δ Rows" in col:
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": ">", "value": 0, "format": fmt_pos})
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": "<", "value": 0, "format": fmt_neg})
        # Threshold boolean hits
        if "Δ>" in col and "vs" in col:
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": "==", "value": True, "format": fmt_bool})
        # HighRisk column
        if col == "HighRisk":
            ws.conditional_format(1, j, nrows, j, {"type": "cell", "criteria": "==", "value": True, "format": fmt_hr})

    # Freeze header row
    ws.freeze_panes(1, 0)

# ---------------------- UI ----------------------
st.title("🔎 Analisi transazioni – Scenari Inbound/Outbound")
st.caption("Intestazione alla terza riga. Carica i due file (uno inbound, uno outbound).")

col1, col2 = st.columns(2)
with col1:
    inbound_file = st.file_uploader("Carica file INBOUND (nome contenente 'inbound')", type=["xlsx", "xls"], key="inb")
with col2:
    outbound_file = st.file_uploader("Carica file OUTBOUND (nome contenente 'outbound')", type=["xlsx", "xls"], key="outb")

with st.expander("⚙️ Opzioni di import"):
    override_detect = st.checkbox("Permetti override manuale della direzione se i nomi non contengono inbound/outbound", value=True)

logs = []
def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logs.append(f"[{ts}] {msg}")

inbound_df = None
outbound_df = None
if inbound_file is not None:
    inbound_df = read_excel_third_header(inbound_file.getvalue(), inbound_file.name)
    log(f"Letto INBOUND: {inbound_file.name} con {len(inbound_df)} righe.")
if outbound_file is not None:
    outbound_df = read_excel_third_header(outbound_file.getvalue(), outbound_file.name)
    log(f"Letto OUTBOUND: {outbound_file.name} con {len(outbound_df)} righe.")

if inbound_df is not None or outbound_df is not None:
    st.subheader("Anteprima dati")
    if inbound_df is not None:
        st.markdown("**Inbound**")
        st.dataframe(inbound_df.head(30), use_container_width=True)
    if outbound_df is not None:
        st.markdown("**Outbound**")
        st.dataframe(outbound_df.head(30), use_container_width=True)

# Parametri scenari
st.subheader("Parametri scenari")
colp1, colp2, colp3, colp4 = st.columns(4)
with colp1:
    pct_threshold = st.number_input("Soglia incremento MoM (%)", min_value=0.0, value=20.0, step=1.0)
with colp2:
    eur_threshold = st.number_input("Soglia valore mese (EUR)", min_value=0.0, value=5000000.0, step=100000.0, format="%f")
with colp3:
    top_n = st.number_input("Top N Paesi (mese corrente)", min_value=1, max_value=50, value=10, step=1)
with colp4:
    delta_threshold_eur = st.number_input("Soglia Δ Amount (EUR) – Pivot", min_value=0.0, value=5000000.0, step=100000.0, format="%f")

# Policy High-Risk da UI
st.subheader("Policy High-Risk")
all_countries = collect_all_countries(inbound_df, outbound_df) if (inbound_df is not None or outbound_df is not None) else []
colhr1, colhr2 = st.columns([2,1])
with colhr1:
    policy_upload = st.file_uploader("Carica lista policy (CSV/XLSX) con colonne: Country, Risk/HighRisk", type=["csv","xlsx","xls"], key="policy")
    policy_map_uploaded = parse_policy_upload(policy_upload)
    # ---- Helper: parse_policy_upload ----
def parse_policy_upload(file) -> dict:
    """
    Legge un CSV/XLS/XLSX contenente la policy dei Paesi High Risk.
    Colonne accettate (case-insensitive):
      - 'Country' (obbligatoria)
      - 'HighRisk' (True/False)  OPPURE 'Risk' con valori testuali (High/Low/No)
    Ritorna: dict { country_name: True/False }
    """
    if file is None:
        return {}
    try:
        # Legge in base all'estensione
        name = getattr(file, "name", "").lower()
        if name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
    except Exception:
        return {}

    if df is None or df.empty:
        return {}

    # Normalizza nomi colonna
    cols_lc = {str(c).strip().lower(): c for c in df.columns}
    if "country" not in cols_lc:
        return {}

    country_col = cols_lc["country"]

    # Individua la colonna di rischio/flag
    highrisk_col = None
    if "highrisk" in cols_lc:
        highrisk_col = cols_lc["highrisk"]
        # Assicuriamoci che sia booleana
        df[highrisk_col] = df[highrisk_col].map(lambda v: bool(v))
    elif "risk" in cols_lc:
        highrisk_col = cols_lc["risk"]
        # Mappatura testuale → boolean
        # (gestisce anche varianti comuni)
        txt_map = {
            "high": True, "alto": True, "high risk": True, "alto rischio": True,
            "low": False, "basso": False, "no": False, "none": False,
            "norisk": False, "no risk": False, "ok": False, "safe": False,
            "false": False, "true": True,
        }
        df[highrisk_col] = (
            df[highrisk_col]
            .astype(str).str.strip().str.lower()
            .map(lambda s: txt_map.get(s, False))
        )
    else:
        # Nessuna colonna interpretabile come rischio
        return {}

    # Costruisce la mappa {Country: True/False}
    mapping = {}
    for _, row in df.iterrows():
        ctry = str(row.get(country_col, "")).strip()
        if not ctry:
            continue
        try:
            val = bool(row.get(highrisk_col, False))
        except Exception:
            val = False
        mapping[ctry] = val

    return mapping

    default_sel = []
    policy_sel = st.multiselect("Seleziona/override manuale i Paesi da considerare High Risk", options=all_countries, default=default_sel)
    policy_paste = st.text_area("Oppure incolla elenco paesi High Risk (separati da virgola o a capo)")
with colhr2:
    st.markdown("""
    - L'**upload** (se presente) ha priorità
    - Poi si **unisce** al multiselect e al testo
    - Se un paese non è indicato qui, si userà il **rating dai file** (se disponibile)
    """)

# Unifica le fonti policy in un unico set
policy_set = set(policy_map_uploaded.keys()) if policy_map_uploaded else set()
policy_set.update(policy_sel)

if policy_paste:
    import re
    # split su virgola, a capo e (facoltativo) punto e virgola, robusto anche a \r\n
    tokens = re.split(r"[,\r?\n;]", policy_paste)
    policy_set.update([c.strip() for c in tokens if c and c.strip()])

# Mappa finale di override **solo da UI** (usata per imporre HighRisk=True)
risk_override_map = {c: True for c in policy_set}

# Mappa HighRisk dai file (se presenti colonne di rating)
file_hr_map = derive_file_highrisk_map(inbound_df, outbound_df)

# Mappa HighRisk **effettiva** usata nei report = union (policy ∪ file)
risk_effective_map = {**{c: True for c in file_hr_map.keys()}, **risk_override_map}

# Anteprima in dashboard: paesi High Risk effettivi
st.subheader("📋 Paesi High Risk considerati nel report")
if not risk_effective_map:
    st.caption("Nessun paese marcato High Risk al momento.")
else:
    eff_list = sorted(risk_effective_map.keys())
    src = []
    for c in eff_list:
        in_policy = c in policy_set
        in_file = c in file_hr_map
        if in_policy and in_file:
            s = "Policy + File"
        elif in_policy:
            s = "Policy override"
        else:
            s = "File rating"
        src.append(s)
    st.dataframe(pd.DataFrame({"Country": eff_list, "Source": src}), use_container_width=True, height=240)

# Esecuzione scenari
run = st.button("Esegui scenari e genera Excel")
if run:
    if inbound_df is None or outbound_df is None:
        st.error("Carica entrambi i file (inbound e outbound) prima di eseguire.")
        st.stop()

    with st.spinner("Esecuzione scenari…"):
        # Calcolo TUTTI i risultati PRIMA di aprire l'Excel writer (così evitiamo NameError)
        try:
            res_a = scenario_a(inbound_df, outbound_df, pct_threshold=float(pct_threshold), eur_threshold=float(eur_threshold), top_n=int(top_n))
            log("Scenario A completato")
        except Exception as e:
            res_a = pd.DataFrame({"Errore": [str(e)]})
            log(f"Errore Scenario A: {e}")

        try:
            res_b = scenario_b(inbound_df, outbound_df)
            log("Scenario B completato")
        except Exception as e:
            res_b = pd.DataFrame({"Errore": [str(e)]})
            log(f"Errore Scenario B: {e}")

        try:
            res_risk = top_risk_countries_sheet(inbound_df, outbound_df)
            log("Top Risk Countries calcolato")
        except Exception as e:
            res_risk = pd.DataFrame({"Errore": [str(e)]})
            log(f"Errore Top Risk: {e}")

        try:
            res_a_details = scenario_a_details(inbound_df, outbound_df, pct_threshold=float(pct_threshold), eur_threshold=float(eur_threshold), top_n=int(top_n))
        except Exception as e:
            res_a_details = pd.DataFrame({"Errore":[str(e)]})
            log(f"Errore Scenario A Details: {e}")

        try:
            res_b_details = scenario_b_details(inbound_df, outbound_df)
        except Exception as e:
            res_b_details = pd.DataFrame({"Errore":[str(e)]})
            log(f"Errore Scenario B Details: {e}")

        try:
            a_byc = by_country_scenario_a(inbound_df, outbound_df, top_n=int(top_n), risk_override=risk_effective_map)
        except Exception as e:
            a_byc = pd.DataFrame({"Errore":[str(e)]})
            log(f"Errore Scenario A ByCountry: {e}")

        try:
            b_byc = by_country_scenario_b(inbound_df, outbound_df, risk_override=risk_effective_map)
        except Exception as e:
            b_byc = pd.DataFrame({"Errore":[str(e)]})
            log(f"Errore Scenario B ByCountry: {e}")

        try:
            mom_sheets = build_mom_sheets(inbound_df, outbound_df)
        except Exception as e:
            mom_sheets = {"MoM_Error": pd.DataFrame({"Errore":[str(e)]})}
            log(f"Errore MoM: {e}")

        # Ora scrittura Excel
        buffer = io.BytesIO()
        date_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"risultati_scenari_{date_tag}.xlsx"
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            if inbound_df is not None:
                inbound_df.to_excel(writer, index=False, sheet_name="Original_Inbound")
            if outbound_df is not None:
                outbound_df.to_excel(writer, index=False, sheet_name="Original_Outbound")
            # Sintesi principali
            res_a.to_excel(writer, index=False, sheet_name="Scenario_A")
            res_b.to_excel(writer, index=False, sheet_name="Scenario_B")
            res_risk.to_excel(writer, index=False, sheet_name="Top_Risk_Countries")
            # Dettagli
            res_a_details.to_excel(writer, index=False, sheet_name="Scenario_A_Details")
            res_b_details.to_excel(writer, index=False, sheet_name="Scenario_B_Details")
            # By Country
            a_byc.to_excel(writer, index=False, sheet_name="Scenario_A_ByCountry")
            b_byc.to_excel(writer, index=False, sheet_name="Scenario_B_ByCountry")
            # Role-specific country aggregations (pivot + long)
            try:
                inbound_by_creditor_pivot = role_country_inbound_pivot(inbound_df, float(delta_threshold_eur), risk_override=risk_effective_map)
            except Exception as e:
                inbound_by_creditor_pivot = pd.DataFrame({"Errore":[str(e)]})
            try:
                outbound_by_debtor_pivot = role_country_outbound_pivot(outbound_df, float(delta_threshold_eur), risk_override=risk_effective_map)
            except Exception as e:
                outbound_by_debtor_pivot = pd.DataFrame({"Errore":[str(e)]})
            # versioni long per riferimento
            try:
                inbound_by_creditor_long = role_country_inbound(inbound_df, risk_override=risk_effective_map)
            except Exception as e:
                inbound_by_creditor_long = pd.DataFrame({"Errore":[str(e)]})
            try:
                outbound_by_debtor_long = role_country_outbound(outbound_df, risk_override=risk_effective_map)
            except Exception as e:
                outbound_by_debtor_long = pd.DataFrame({"Errore":[str(e)]})
            # Scrittura: i fogli richiesti assumono layout pivot
            inbound_by_creditor_pivot.to_excel(writer, index=False, sheet_name="Inbound_By_CreditorCountry")
            outbound_by_debtor_pivot.to_excel(writer, index=False, sheet_name="Outbound_By_DebtorCountry")
            # Conditional formatting on pivot sheets
            try:
                apply_pivot_conditional_formatting(writer, "Inbound_By_CreditorCountry", inbound_by_creditor_pivot)
                apply_pivot_conditional_formatting(writer, "Outbound_By_DebtorCountry", outbound_by_debtor_pivot)
            except Exception as e:
                pass
            # E aggiungo anche i long per verifica
            inbound_by_creditor_long.to_excel(writer, index=False, sheet_name="Inbound_By_CreditorCountry_Long")
            outbound_by_debtor_long.to_excel(writer, index=False, sheet_name="Outbound_By_DebtorCountry_Long")

            # Policy sheet con la lista High-Risk da UI (se presente)
            try:
                if risk_override_map:
                    policy_df = pd.DataFrame({
                        "Country": sorted(risk_override_map.keys()),
                        "Policy HighRisk": [bool(risk_override_map[c]) for c in sorted(risk_override_map.keys())]
                    })
                else:
                    policy_df = pd.DataFrame({"Info":["Nessun override policy caricato/selezionato"]})
            except Exception as e:
                policy_df = pd.DataFrame({"Errore":[str(e)]})
            policy_df.to_excel(writer, index=False, sheet_name="Policy_HighRisk")

            # Parameters sheet
            try:
                run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                inbound_name = inbound_file.name if 'inbound_file' in globals() and inbound_file is not None else ""
                outbound_name = outbound_file.name if 'outbound_file' in globals() and outbound_file is not None else ""
                params = pd.DataFrame([
                    {"Parameter": "Pct Threshold MoM (%)", "Value": pct_threshold},
                    {"Parameter": "EUR Threshold Month (Scenario A)", "Value": eur_threshold},
                    {"Parameter": "Top N Countries (Scenario A)", "Value": top_n},
                    {"Parameter": "Δ Amount Threshold (Pivot, EUR)", "Value": delta_threshold_eur},
                    {"Parameter": "Inbound File", "Value": inbound_name},
                    {"Parameter": "Outbound File", "Value": outbound_name},
                    {"Parameter": "Run Timestamp", "Value": run_ts},
                    {"Parameter": "HighRisk Count (effective)", "Value": len(risk_effective_map) if 'risk_effective_map' in globals() else 0},
                ])
                params.to_excel(writer, index=False, sheet_name="Parameters")
                # append HighRisk list with source
                if 'risk_effective_map' in globals() and risk_effective_map:
                    eff_list = sorted(risk_effective_map.keys())
                    # try to detect source using previously built sets
                    policy_set_local = set(risk_override_map.keys()) if 'risk_override_map' in globals() else set()
                    # file_hr_map was computed before
                    file_src = set(file_hr_map.keys()) if 'file_hr_map' in globals() else set()
                    src_labels = []
                    for c in eff_list:
                        in_policy = c in policy_set_local
                        in_file = c in file_src
                        if in_policy and in_file:
                            src_labels.append("Policy + File")
                        elif in_policy:
                            src_labels.append("Policy override")
                        else:
                            src_labels.append("File rating")
                    hr_df = pd.DataFrame({"HighRisk Country": eff_list, "Source": src_labels})
                    hr_df.to_excel(writer, index=False, sheet_name="Parameters", startrow=len(params)+2)
            except Exception as e:
                pd.DataFrame({"Errore":[str(e)]}).to_excel(writer, index=False, sheet_name="Parameters")

            # MoM sheets
            for sheet, df_m in mom_sheets.items():
                sname = sheet[:31]
                df_m.to_excel(writer, index=False, sheet_name=sname)

            # MoM DETAILS: righe originali per ciascuna coppia mese vs mese
            try:
                mom_detail_sheets = build_mom_detail_sheets(inbound_df, outbound_df)
            except Exception as e:
                mom_detail_sheets = {"MoM_Detail_Error": pd.DataFrame({"Errore":[str(e)]})}
            for sheet, df_md in mom_detail_sheets.items():
                sname = sheet[:31]
                df_md.to_excel(writer, index=False, sheet_name=sname)

            # Log
            pd.DataFrame({"Log": logs}).to_excel(writer, index=False, sheet_name="Logs")

        st.success("File Excel generato.")
        st.download_button(
            "⬇️ Scarica risultati",
            data=buffer.getvalue(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("© App di supporto – pronta per integrare la logica degli scenari appena condividi le regole.")
