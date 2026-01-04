# Delphi Utility-Scale Financial Model (No Excel)
# Streamlit single-file app with: Projects + Scenarios, Macro, Timeline, Generation, Revenues,
# CAPEX, OPEX, SG&A, Unlevered Base Cash Flow, Depreciation, Compare
# All inputs in COP; outputs selectable COP/USD (USD via FX path).

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict
from datetime import date
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# -----------------------------
# Storage
# -----------------------------
PROJECTS_FILE = "delphi_projects.json"


def _load_db() -> dict:
    if os.path.exists(PROJECTS_FILE):
        try:
            with open(PROJECTS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"projects": {}}
    return {"projects": {}}


def _save_db(db: dict) -> None:
    with open(PROJECTS_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)


# -----------------------------
# Formatting helpers (thousand separators)
# -----------------------------
def _fmt_num(x: float, decimals: int = 0) -> str:
    if x is None:
        return "—"
    try:
        if isinstance(x, float) and not np.isfinite(x):
            return "—"
    except Exception:
        pass
    return f"{x:,.{decimals}f}"


def _fmt_cop(x: float) -> str:
    return f"COP {_fmt_num(float(x), 0)}"


def _fmt_usd(x: float) -> str:
    return f"USD {_fmt_num(float(x), 0)}"


def _df_format_money(df: pd.DataFrame, cols: List[str], decimals: int = 0) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda v: _fmt_num(float(v), decimals) if pd.notnull(v) else "")
    return out


def _metric_row(items):
    cols = st.columns(len(items))
    for i, (k, v) in enumerate(items):
        cols[i].metric(k, v)


# -----------------------------
# Date utilities
# -----------------------------
def _today() -> date:
    return date.today()


def _parse_date_iso(s: str) -> date:
    return date.fromisoformat(s)


def _add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    day = min(
        d.day,
        [31, 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1],
    )
    return date(y, m, day)


def _month_starts(start: date, months: int) -> List[date]:
    return [_add_months(start, i) for i in range(months)]


def build_timeline(t: "TimelineInputs") -> dict:
    start = _parse_date_iso(t.start_date) if t.start_date else _today()
    rtb = _add_months(start, int(t.dev_months))
    cod = _add_months(rtb, int(t.capex_months))
    end_op = _add_months(cod, int(t.operation_years) * 12)
    return {"start": start, "rtb": rtb, "cod": cod, "end_op": end_op}


# -----------------------------
# Inputs (dataclasses)
# -----------------------------
INDEX_CHOICES = ["Colombia CPI", "Colombia PPI", "US CPI", "Custom"]


@dataclass
class MacroInputs:
    col_cpi: float = 6.0
    col_ppi: float = 5.0
    us_cpi: float = 3.0
    custom_index_rate: float = 5.0

    fx_cop_per_usd_start: float = 4000.0
    fx_method: str = "Inflation differential (PPP approx.)"  # or "Flat"
    fx_flat: float = 4000.0


@dataclass
class TimelineInputs:
    start_date: str = str(_today())
    dev_months: int = 18
    capex_months: int = 12
    operation_years: int = 25


@dataclass
class GenerationInputs:
    mwac: float = 100.0
    mwp: float = 130.0
    p50_mwh_yr: float = 220000.0
    p75_mwh_yr: float = 210000.0
    p90_mwh_yr: float = 200000.0
    production_choice: str = "P50"
    degradation_pct: float = 0.5  # %/yr


@dataclass
class RevenueOption1PPA:
    ppa_price_cop_per_kwh: float = 320.0
    ppa_term_years: int = 12
    merchant_price_cop_per_kwh: float = 250.0
    indexation: str = "Colombia CPI"


def _default_manual_prices() -> Dict[int, float]:
    return {i: 300.0 for i in range(1, 26)}


@dataclass
class RevenueOption2Manual:
    prices_constant_cop_per_kwh: Dict[int, float] = field(default_factory=_default_manual_prices)
    indexation: str = "Colombia CPI"


CAPEX_PHASES = ["Development", "Construction", "At COD"]
CAPEX_DISTS = ["Straight-line (monthly)", "Front-loaded", "Back-loaded"]


def _default_capex_lines() -> List[Dict[str, object]]:
    return [
        {"Item": "EPC (modules + BOS + installation)", "Amount_COP": 0.0, "Phase": "Construction"},
        {"Item": "Interconnection / Substation / Line", "Amount_COP": 0.0, "Phase": "Construction"},
        {"Item": "Development costs", "Amount_COP": 0.0, "Phase": "Development"},
    ]


@dataclass
class CapexInputs:
    lines: List[Dict[str, object]] = field(default_factory=_default_capex_lines)
    distribution: str = "Straight-line (monthly)"


PHASES = ["Development", "Construction", "Operation"]


def _default_opex_other_items() -> List[Dict[str, object]]:
    return [
        {"Item": "Other OPEX item", "Amount_COP_per_year": 0.0, "Phase": "Operation", "Indexation": "Colombia CPI"},
    ]


@dataclass
class OpexInputs:
    fixed_om_cop_per_mwac_year: float = 0.0
    fixed_om_indexation: str = "Colombia CPI"

    variable_om_cop_per_mwh: float = 0.0
    variable_om_indexation: str = "Colombia CPI"

    insurance_cop_per_mwac_year: float = 0.0
    insurance_indexation: str = "Colombia CPI"

    grid_fees_cop_per_mwh: float = 0.0  # assume already in real COP (indexed by user externally, optional)

    land_hectares: float = 0.0
    land_price_dev_cop_per_ha_year: float = 0.0
    land_price_con_cop_per_ha_year: float = 0.0
    land_price_op_cop_per_ha_year: float = 0.0
    land_indexation: str = "Colombia CPI"

    ica_pct_of_revenue: float = 0.0
    gmf_pct_of_outflows: float = 0.4  # default 0.4%

    other_items: List[Dict[str, object]] = field(default_factory=_default_opex_other_items)


def _default_sga_items() -> List[Dict[str, object]]:
    return [
        {"Item": "Project management / Owner team", "Amount_COP_per_year": 0.0, "Phase": "Development", "Indexation": "Colombia CPI"},
        {"Item": "Permitting / legal / compliance", "Amount_COP_per_year": 0.0, "Phase": "Development", "Indexation": "Colombia CPI"},
        {"Item": "Corporate overhead allocation", "Amount_COP_per_year": 0.0, "Phase": "Operation", "Indexation": "Colombia CPI"},
    ]


@dataclass
class SgaInputs:
    items: List[Dict[str, object]] = field(default_factory=_default_sga_items)


@dataclass
class DepreciationInputs:
    pct_of_capex_depreciated: float = 100.0
    dep_years: int = 20  # 3–25


@dataclass
class TaxInputs:
    corporate_tax_rate_pct: float = 35.0
    allow_loss_carryforward: bool = True


@dataclass
class ScenarioInputs:
    name: str = "Base"
    macro: MacroInputs = field(default_factory=MacroInputs)
    timeline: TimelineInputs = field(default_factory=TimelineInputs)
    generation: GenerationInputs = field(default_factory=GenerationInputs)

    revenue_mode: str = "Standard PPA Parameters"
    revenue1: RevenueOption1PPA = field(default_factory=RevenueOption1PPA)
    revenue2: RevenueOption2Manual = field(default_factory=RevenueOption2Manual)

    capex: CapexInputs = field(default_factory=CapexInputs)
    opex: OpexInputs = field(default_factory=OpexInputs)
    sga: SgaInputs = field(default_factory=SgaInputs)

    depreciation: DepreciationInputs = field(default_factory=DepreciationInputs)
    tax: TaxInputs = field(default_factory=TaxInputs)


def _scenario_to_dict(s: ScenarioInputs) -> dict:
    return asdict(s)


def _scenario_from_dict(d: dict) -> ScenarioInputs:
    macro = MacroInputs(**d.get("macro", {}))
    timeline = TimelineInputs(**d.get("timeline", {}))
    generation = GenerationInputs(**d.get("generation", {}))

    revenue_mode = d.get("revenue_mode", "Standard PPA Parameters")
    revenue1 = RevenueOption1PPA(**d.get("revenue1", {}))
    revenue2 = RevenueOption2Manual(**d.get("revenue2", {}))

    capex = CapexInputs(**d.get("capex", {})) if "capex" in d else CapexInputs()
    opex = OpexInputs(**d.get("opex", {})) if "opex" in d else OpexInputs()
    sga = SgaInputs(**d.get("sga", {})) if "sga" in d else SgaInputs()

    depreciation = DepreciationInputs(**d.get("depreciation", {})) if "depreciation" in d else DepreciationInputs()
    tax = TaxInputs(**d.get("tax", {})) if "tax" in d else TaxInputs()

    return ScenarioInputs(
        name=d.get("name", "Base"),
        macro=macro,
        timeline=timeline,
        generation=generation,
        revenue_mode=revenue_mode,
        revenue1=revenue1,
        revenue2=revenue2,
        capex=capex,
        opex=opex,
        sga=sga,
        depreciation=depreciation,
        tax=tax,
    )


# -----------------------------
# Macro series
# -----------------------------
def _idx_key(choice: str) -> str:
    if choice == "Colombia CPI":
        return "col_cpi"
    if choice == "Colombia PPI":
        return "col_ppi"
    if choice == "US CPI":
        return "us_cpi"
    return "custom_index_rate"


def annual_index_series(macro: MacroInputs, base_year: int, years: List[int], index_key: str) -> pd.Series:
    r = float(getattr(macro, index_key)) / 100.0
    out = {}
    for y in years:
        n = y - base_year
        out[y] = (1.0 + r) ** n
    return pd.Series(out)


def fx_series(macro: MacroInputs, base_year: int, years: List[int]) -> pd.Series:
    fx0 = float(macro.fx_cop_per_usd_start)
    if macro.fx_method == "Flat":
        return pd.Series({y: float(macro.fx_flat or fx0) for y in years})
    # PPP-ish: grow FX with (Col CPI - US CPI)
    col = float(macro.col_cpi) / 100.0
    us = float(macro.us_cpi) / 100.0
    drift = (1.0 + col) / (1.0 + us) - 1.0
    return pd.Series({y: fx0 * ((1.0 + drift) ** (y - base_year)) for y in years})


def index_factor_for_year(macro: MacroInputs, year: int, base_year: int, index_choice: str) -> float:
    years = list(range(min(base_year, year), max(base_year, year) + 1))
    s = annual_index_series(macro, base_year, years, _idx_key(index_choice))
    return float(s.loc[year])


# -----------------------------
# Generation & Revenues (annual)
# -----------------------------
def operating_year_table(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    cod = tl["cod"]
    op_years = int(s.timeline.operation_years)

    years = list(range(cod.year, cod.year + op_years))
    base_year = cod.year

    gen = s.generation
    p_map = {"P50": gen.p50_mwh_yr, "P75": gen.p75_mwh_yr, "P90": gen.p90_mwh_yr}
    base_mwh = float(p_map.get(gen.production_choice, gen.p50_mwh_yr))
    degr = float(gen.degradation_pct) / 100.0
    mwh = [base_mwh * ((1.0 - degr) ** i) for i in range(op_years)]

    if s.revenue_mode == "Standard PPA Parameters":
        r = s.revenue1
        term = int(r.ppa_term_years)
        p0 = float(r.ppa_price_cop_per_kwh)
        pm = float(r.merchant_price_cop_per_kwh)
        index_choice = r.indexation
        price_base = [p0 if (i < term) else pm for i in range(op_years)]
    else:
        r = s.revenue2
        index_choice = r.indexation
        price_base = [float(r.prices_constant_cop_per_kwh.get(i + 1, 0.0)) for i in range(op_years)]

    idx = annual_index_series(s.macro, base_year, years, _idx_key(index_choice))
    price_indexed = [price_base[i] * float(idx.loc[years[i]]) for i in range(op_years)]

    df = pd.DataFrame({"Year": years, "Energy (MWh)": mwh})
    df["Price (COP/kWh)"] = price_indexed
    df["Revenue (COP)"] = df["Energy (MWh)"] * 1000.0 * df["Price (COP/kWh)"]
    return df


# -----------------------------
# CAPEX
# -----------------------------
def _weights(n: int, mode: str) -> np.ndarray:
    if n <= 0:
        return np.array([])
    if mode == "Straight-line (monthly)":
        w = np.ones(n)
    elif mode == "Front-loaded":
        w = np.linspace(n, 1, n)
    else:
        w = np.linspace(1, n, n)
    return w / w.sum()


def capex_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    start = tl["start"]
    rtb = tl["rtb"]
    cod = tl["cod"]

    dev_n = int(s.timeline.dev_months)
    con_n = int(s.timeline.capex_months)

    dev_months = _month_starts(start, dev_n)
    con_months = _month_starts(rtb, con_n)

    lines = s.capex.lines or []
    dev_total = sum(float(x.get("Amount_COP", 0.0) or 0.0) for x in lines if x.get("Phase") == "Development")
    con_total = sum(float(x.get("Amount_COP", 0.0) or 0.0) for x in lines if x.get("Phase") == "Construction")
    cod_total = sum(float(x.get("Amount_COP", 0.0) or 0.0) for x in lines if x.get("Phase") == "At COD")

    w_dev = _weights(dev_n, s.capex.distribution)
    w_con = _weights(con_n, s.capex.distribution)

    rows = []
    for i, m in enumerate(dev_months):
        rows.append({"Month": m, "Phase": "Development", "CAPEX (COP)": dev_total * (w_dev[i] if len(w_dev) else 0.0)})
    for i, m in enumerate(con_months):
        rows.append({"Month": m, "Phase": "Construction", "CAPEX (COP)": con_total * (w_con[i] if len(w_con) else 0.0)})
    rows.append({"Month": date(cod.year, cod.month, 1), "Phase": "At COD", "CAPEX (COP)": cod_total})

    df = pd.DataFrame(rows).sort_values("Month").reset_index(drop=True)
    df["Year"] = df["Month"].apply(lambda d: d.year)
    return df


# -----------------------------
# OPEX monthly schedule (includes Revenue + Energy + CAPEX merge)
# -----------------------------
def _phase_for_month(tl: dict, m: date) -> str:
    if m < date(tl["rtb"].year, tl["rtb"].month, 1):
        return "Development"
    if m < date(tl["cod"].year, tl["cod"].month, 1):
        return "Construction"
    return "Operation"


def opex_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    start = tl["start"]
    end_op = tl["end_op"]

    # build monthly range from start (month start) to end_op (month start, exclusive)
    months = []
    cur = date(start.year, start.month, 1)
    endm = date(end_op.year, end_op.month, 1)
    while cur < endm:
        months.append(cur)
        cur = _add_months(cur, 1)

    df = pd.DataFrame({"Month": months})
    df["Year"] = df["Month"].apply(lambda d: d.year)
    df["Phase"] = df["Month"].apply(lambda m: _phase_for_month(tl, m))

    # Revenue/Energy annual -> allocate evenly across months in operation years
    op = operating_year_table(s)
    op_map_mwh = {int(r["Year"]): float(r["Energy (MWh)"]) for _, r in op.iterrows()}
    op_map_rev = {int(r["Year"]): float(r["Revenue (COP)"]) for _, r in op.iterrows()}

    df["Energy (MWh)"] = df.apply(lambda r: (op_map_mwh.get(int(r["Year"]), 0.0) / 12.0) if r["Phase"] == "Operation" else 0.0, axis=1)
    df["Revenue (COP)"] = df.apply(lambda r: (op_map_rev.get(int(r["Year"]), 0.0) / 12.0) if r["Phase"] == "Operation" else 0.0, axis=1)

    # CAPEX monthly merge
    cap = capex_monthly_schedule(s)[["Month", "CAPEX (COP)"]].copy()
    df = df.merge(cap, on="Month", how="left")
    df["CAPEX (COP)"] = df["CAPEX (COP)"].fillna(0.0)

    # OPEX components
    mwac = float(s.generation.mwac or 0.0)

    # indexed factors by year
    base_year = tl["cod"].year  # index relative to COD year by default (practical)
    idx_fixed = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.fixed_om_indexation))
    idx_ins = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.insurance_indexation))
    idx_land = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.land_indexation))

    # fixed OM, insurance per MWac-year -> monthly
    df["Fixed O&M"] = 0.0
    df.loc[df["Phase"] == "Operation", "Fixed O&M"] = (float(s.opex.fixed_om_cop_per_mwac_year) * mwac / 12.0) * idx_fixed[df["Phase"] == "Operation"].values

    df["Insurance"] = 0.0
    df.loc[df["Phase"] == "Operation", "Insurance"] = (float(s.opex.insurance_cop_per_mwac_year) * mwac / 12.0) * idx_ins[df["Phase"] == "Operation"].values

    # variable OM per MWh
    df["Variable O&M"] = float(s.opex.variable_om_cop_per_mwh) * df["Energy (MWh)"]

    # grid fees per MWh
    df["Grid fees"] = float(s.opex.grid_fees_cop_per_mwh) * df["Energy (MWh)"]

    # land lease by phase (per ha-year -> monthly), indexed
    ha = float(s.opex.land_hectares or 0.0)
    df["Land lease"] = 0.0
    df.loc[df["Phase"] == "Development", "Land lease"] = (ha * float(s.opex.land_price_dev_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Development"].values
    df.loc[df["Phase"] == "Construction", "Land lease"] = (ha * float(s.opex.land_price_con_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Construction"].values
    df.loc[df["Phase"] == "Operation", "Land lease"] = (ha * float(s.opex.land_price_op_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Operation"].values

    # other OPEX items (dynamic)
    other = pd.DataFrame(s.opex.other_items or [])
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in other.columns:
            other[col] = "" if col != "Amount_COP_per_year" else 0.0
    other["Phase"] = other["Phase"].where(other["Phase"].isin(PHASES), "Operation")
    other["Indexation"] = other["Indexation"].where(other["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    # allocate each item to months in its phase, indexed yearly
    # (simple approach: monthly = annual / 12, applied to months where phase matches)
    for _, r in other.iterrows():
        name = str(r.get("Item", "")).strip() or "Other"
        amt = float(r.get("Amount_COP_per_year", 0.0) or 0.0)
        ph = str(r.get("Phase", "Operation"))
        idx_choice = str(r.get("Indexation", "Colombia CPI"))
        if name not in df.columns:
            df[name] = 0.0
        idx_series = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, idx_choice))
        df.loc[df["Phase"] == ph, name] = (amt / 12.0) * idx_series[df["Phase"] == ph].values

    # ICA as % revenue (operation only)
    df["ICA"] = 0.0
    df.loc[df["Phase"] == "Operation", "ICA"] = (float(s.opex.ica_pct_of_revenue) / 100.0) * df.loc[df["Phase"] == "Operation", "Revenue (COP)"]

    # OPEX subtotal (pre GMF)
    # include known components + dynamic columns except meta and CAPEX/Revenue/Energy
    meta = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)"}
    fixed_cols = ["Fixed O&M", "Insurance", "Variable O&M", "Grid fees", "Land lease", "ICA"]
    dyn_cols = [c for c in df.columns if c not in meta and c not in fixed_cols and c != "OPEX subtotal" and c != "GMF"]
    opex_cols = fixed_cols + dyn_cols
    df["OPEX subtotal"] = df[opex_cols].sum(axis=1) if opex_cols else 0.0

    # GMF (% of outgoing cash) applied to ALL outgoing cash (CAPEX + OPEX subtotal) — user-defined %
    gmf = float(s.opex.gmf_pct_of_outflows) / 100.0
    df["GMF"] = gmf * (df["CAPEX (COP)"].fillna(0.0) + df["OPEX subtotal"].fillna(0.0))

    return df


# -----------------------------
# SG&A schedules
# -----------------------------
def sga_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    start = tl["start"]
    end_op = tl["end_op"]

    months = []
    cur = date(start.year, start.month, 1)
    endm = date(end_op.year, end_op.month, 1)
    while cur < endm:
        months.append(cur)
        cur = _add_months(cur, 1)

    df = pd.DataFrame({"Month": months})
    df["Year"] = df["Month"].apply(lambda d: d.year)
    df["Phase"] = df["Month"].apply(lambda m: _phase_for_month(tl, m))

    items = pd.DataFrame(s.sga.items or [])
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in items.columns:
            items[col] = "" if col != "Amount_COP_per_year" else 0.0
    items["Phase"] = items["Phase"].where(items["Phase"].isin(PHASES), "Development")
    items["Indexation"] = items["Indexation"].where(items["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    base_year = tl["cod"].year
    for _, r in items.iterrows():
        name = str(r.get("Item", "")).strip() or "SGA"
        amt = float(r.get("Amount_COP_per_year", 0.0) or 0.0)
        ph = str(r.get("Phase", "Development"))
        idx_choice = str(r.get("Indexation", "Colombia CPI"))
        if name not in df.columns:
            df[name] = 0.0
        idx_series = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, idx_choice))
        df.loc[df["Phase"] == ph, name] = (amt / 12.0) * idx_series[df["Phase"] == ph].values

    return df.fillna(0.0)


def sga_annual_by_item(s: ScenarioInputs) -> pd.DataFrame:
    dfm = sga_monthly_schedule(s).copy()
    meta = {"Month", "Year", "Phase"}
    item_cols = [c for c in dfm.columns if c not in meta]
    annual = dfm.groupby("Year", as_index=False)[item_cols].sum() if item_cols else dfm.groupby("Year", as_index=False).size()[["Year"]]
    annual["Total SG&A (COP)"] = annual[item_cols].sum(axis=1) if item_cols else 0.0
    return annual


# -----------------------------
# Depreciation (annual)
# -----------------------------
def depreciation_annual_table(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    cod_year = tl["cod"].year

    cap_df = pd.DataFrame(s.capex.lines or [])
    total_capex = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

    pct = float(s.depreciation.pct_of_capex_depreciated or 0.0) / 100.0
    dep_years = int(s.depreciation.dep_years)

    dep_base = total_capex * pct
    annual_dep = dep_base / dep_years if dep_years > 0 else 0.0

    years = list(range(cod_year, cod_year + dep_years))
    df = pd.DataFrame(
        {"Year": years, "Depreciable CAPEX (COP)": [dep_base] * dep_years, "Depreciation (COP)": [annual_dep] * dep_years}
    )
    df["Cumulative Depreciation (COP)"] = df["Depreciation (COP)"].cumsum()
    df["Remaining Book Value (COP)"] = np.maximum(dep_base - df["Cumulative Depreciation (COP)"], 0.0)
    return df


# -----------------------------
# CASH FLOW (monthly + annual)
# -----------------------------
def _sga_monthly_total(s: ScenarioInputs) -> pd.DataFrame:
    df = sga_monthly_schedule(s).copy()
    meta = {"Month", "Year", "Phase"}
    item_cols = [c for c in df.columns if c not in meta]
    df["SG&A (COP)"] = df[item_cols].sum(axis=1) if item_cols else 0.0
    return df[["Month", "Year", "Phase", "SG&A (COP)"]]


def cashflow_monthly_table(s: ScenarioInputs) -> pd.DataFrame:
    """
    Monthly table from project start through end of operation:
    Revenue, CAPEX, Total OPEX (incl GMF), SG&A, Operating CF (EBITDA proxy),
    Unlevered CF (pre-tax), and cumulative.
    """
    base = opex_monthly_schedule(s).copy()

    # merge SG&A
    sga_m = _sga_monthly_total(s).copy()
    base = base.merge(sga_m[["Month", "SG&A (COP)"]], on="Month", how="left")
    base["SG&A (COP)"] = base["SG&A (COP)"].fillna(0.0)

    base["Total OPEX (COP)"] = base["OPEX subtotal"].fillna(0.0) + base["GMF"].fillna(0.0)

    base["Operating CF (COP)"] = (
        base["Revenue (COP)"].fillna(0.0)
        - base["Total OPEX (COP)"].fillna(0.0)
        - base["SG&A (COP)"].fillna(0.0)
    )

    base["Unlevered CF (COP)"] = base["Operating CF (COP)"].fillna(0.0) - base["CAPEX (COP)"].fillna(0.0)
    base["Cumulative Unlevered CF (COP)"] = base["Unlevered CF (COP)"].cumsum()

    cols = [
        "Month", "Year", "Phase",
        "Energy (MWh)", "Revenue (COP)",
        "CAPEX (COP)",
        "Total OPEX (COP)",
        "SG&A (COP)",
        "Operating CF (COP)",
        "Unlevered CF (COP)",
        "Cumulative Unlevered CF (COP)",
    ]
    return base[cols].copy()


def cashflow_annual_table(s: ScenarioInputs) -> pd.DataFrame:
    m = cashflow_monthly_table(s).copy()
    annual = (
        m.groupby("Year", as_index=False)[
            ["Energy (MWh)", "Revenue (COP)", "CAPEX (COP)", "Total OPEX (COP)", "SG&A (COP)", "Operating CF (COP)", "Unlevered CF (COP)"]
        ]
        .sum()
    )
    annual["Cumulative Unlevered CF (COP)"] = annual["Unlevered CF (COP)"].cumsum()
    return annual


def unlevered_base_cashflow_annual(s: ScenarioInputs) -> pd.DataFrame:
    """
    Annual unlevered base CF with depreciation + taxes (pre debt / no tax credits).
    Includes NOL carryforward logic (no taxes until accumulated losses are used).
    """
    a = cashflow_annual_table(s).copy()
    dep = depreciation_annual_table(s)[["Year", "Depreciation (COP)"]].copy()
    out = a.merge(dep, on="Year", how="left").fillna({"Depreciation (COP)": 0.0})

    out["EBITDA (COP)"] = out["Operating CF (COP)"]  # by construction (pre D&A)
    out["Taxable Income (COP)"] = out["EBITDA (COP)"] - out["Depreciation (COP)"]

    rate = float(s.tax.corporate_tax_rate_pct) / 100.0
    allow_nol = bool(s.tax.allow_loss_carryforward)

    nol = 0.0
    taxes = []
    nol_end = []
    for _, r in out.iterrows():
        ti = float(r["Taxable Income (COP)"])
        if allow_nol:
            # accumulate losses and apply to future profits
            if ti < 0:
                nol = nol + (-ti)
                taxable_after_nol = 0.0
            else:
                used = min(nol, ti)
                nol -= used
                taxable_after_nol = max(ti - used, 0.0)
        else:
            taxable_after_nol = max(ti, 0.0)

        taxes_payable = taxable_after_nol * rate
        taxes.append(taxes_payable)
        nol_end.append(nol)

    out["Loss Carryforward End (COP)"] = nol_end
    out["Taxes Payable (COP)"] = taxes

    out["Unlevered CF After Tax (COP)"] = out["Unlevered CF (COP)"] - out["Taxes Payable (COP)"]

    return out


def _irr_bisection(cashflows: List[float], low=-0.95, high=5.0, tol=1e-7, max_iter=200) -> float:
    # Returns IRR as a decimal (e.g., 0.12 = 12%). NaN if cannot solve.

    def npv(rate: float) -> float:
        denom_base = 1.0 + rate

        # Guard 1: invalid / near -100% (would blow up)
        if denom_base <= 1e-12:
            return float("inf")

        total = 0.0
        for i, cf in enumerate(cashflows):
            if i == 0:
                total += cf
                continue

            den = denom_base ** i

            # Guard 2: prevent underflow to 0.0 which causes ZeroDivisionError
            if den == 0.0:
                return float("inf")

            total += cf / den

        return total

    if len(cashflows) < 2:
        return float("nan")

    f_low = npv(low)
    f_high = npv(high)

    # If no sign change, try expanding high
    tries = 0
    while f_low * f_high > 0 and tries < 20:
        high *= 2
        f_high = npv(high)
        tries += 1

    if f_low * f_high > 0:
        return float("nan")

    for _ in range(max_iter):
        mid = (low + high) / 2
        f_mid = npv(mid)
        if abs(f_mid) < tol:
            return mid
        if f_low * f_mid <= 0:
            high = mid
            f_high = f_mid
        else:
            low = mid
            f_low = f_mid
    return mid


def _payback_months(months: List[date], unlevered_cf: List[float]) -> float:
    # Returns payback in months from Month[0]. NaN if never paid back.
    cum = 0.0
    for i, cf in enumerate(unlevered_cf):
        prev = cum
        cum += cf
        if cum >= 0 and i > 0:
            # linear interpolation within month i
            if cf == 0:
                return float(i)
            frac = (0 - prev) / cf
            return (i - 1) + frac
    return float("nan")


# -----------------------------
# APP UI
# -----------------------------
st.set_page_config(page_title="Delphi Utility-Scale Model", page_icon="⚡", layout="wide")
st.title("Delphi Utility-Scale Project Model (COP inputs, COP/USD outputs)")
st.caption(f"Projects + scenarios stored at: {PROJECTS_FILE}")

db = _load_db()

# Sidebar: project & scenario management
with st.sidebar:
    st.header("Project & Scenario")

    projects = sorted(list(db.get("projects", {}).keys()))
    project_name = st.selectbox("Project", ["(New project)"] + projects, index=0)

    if project_name == "(New project)":
        new_project = st.text_input("New project name", value="")
        if st.button("Create project", type="primary", use_container_width=True):
            if new_project.strip():
                db.setdefault("projects", {}).setdefault(new_project.strip(), {"scenarios": {}})
                _save_db(db)
                st.rerun()
            else:
                st.warning("Enter a project name.")
        st.stop()

    proj = db["projects"].setdefault(project_name, {"scenarios": {}})

    scen_names = sorted(list(proj.get("scenarios", {}).keys()))
    scenario_name = st.selectbox("Scenario", ["(New scenario)"] + scen_names, index=0)

    if scenario_name == "(New scenario)":
        new_s = st.text_input("New scenario name", value="Base")
        if st.button("Create scenario", type="primary", use_container_width=True):
            nm = new_s.strip() or "Base"
            proj["scenarios"][nm] = _scenario_to_dict(ScenarioInputs(name=nm))
            _save_db(db)
            st.rerun()
        st.stop()

    # Load scenario
    s = _scenario_from_dict(proj["scenarios"][scenario_name])
    s.name = scenario_name

    st.divider()
    cdel1, cdel2 = st.columns(2)
    with cdel1:
        if st.button("Save scenario", use_container_width=True):
            proj["scenarios"][scenario_name] = _scenario_to_dict(s)
            _save_db(db)
            st.success("Saved.")
    with cdel2:
        if st.button("Delete scenario", use_container_width=True):
            try:
                del proj["scenarios"][scenario_name]
                _save_db(db)
                st.rerun()
            except Exception:
                st.error("Could not delete.")

    st.divider()
    st.subheader("Compare")
    compare_scenarios = st.multiselect("Select scenarios", scen_names, default=[])


# Tabs (top-down)
tab_macro, tab_timeline, tab_gen, tab_rev, tab_capex, tab_opex, tab_sga, tab_dep, tab_ucf, tab_compare = st.tabs(
    [
        "A) Macroeconomic",
        "B) Timeline",
        "C) Power Generation",
        "D) Power Revenues",
        "E) CAPEX",
        "F) OPEX",
        "G) SG&A",
        "H) Depreciation",
        "I) Unlevered Base Cash Flow",
        "J) Compare",
    ]
)

# -----------------------------
# A) Macro
# -----------------------------
with tab_macro:
    st.subheader("Macroeconomic inputs (annual rates, %)")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        s.macro.col_cpi = st.number_input("Colombian CPI (%)", value=float(s.macro.col_cpi), step=0.1, format="%.2f")
    with c2:
        s.macro.col_ppi = st.number_input("Colombian PPI (%)", value=float(s.macro.col_ppi), step=0.1, format="%.2f")
    with c3:
        s.macro.us_cpi = st.number_input("US CPI (%)", value=float(s.macro.us_cpi), step=0.1, format="%.2f")
    with c4:
        s.macro.custom_index_rate = st.number_input("Custom index (%)", value=float(s.macro.custom_index_rate), step=0.1, format="%.2f")

    st.markdown("#### FX (COP per USD)")
    fx1, fx2, fx3 = st.columns([1.2, 1, 1])
    with fx1:
        s.macro.fx_cop_per_usd_start = st.number_input("Starting FX (COP/USD)", value=float(s.macro.fx_cop_per_usd_start), step=10.0, format="%.0f")
        st.caption(f"Formatted: {_fmt_cop(s.macro.fx_cop_per_usd_start)} / USD")
    with fx2:
        s.macro.fx_method = st.selectbox("FX method", ["Inflation differential (PPP approx.)", "Flat"], index=0 if s.macro.fx_method != "Flat" else 1)
    with fx3:
        s.macro.fx_flat = st.number_input("Flat FX (if selected)", value=float(s.macro.fx_flat or s.macro.fx_cop_per_usd_start), step=10.0, format="%.0f")
        st.caption(f"Formatted: {_fmt_cop(s.macro.fx_flat)} / USD")

    st.info("FX path default uses a simple PPP approximation: FX grows with (Col CPI / US CPI). You can switch to Flat FX.")


# -----------------------------
# B) Timeline
# -----------------------------
with tab_timeline:
    st.subheader("Project timeline (Development → CAPEX → Operation)")

    c1, c2, c3 = st.columns(3)
    with c1:
        s.timeline.start_date = st.date_input("Project start date", value=_parse_date_iso(s.timeline.start_date)).isoformat()
    with c2:
        s.timeline.dev_months = int(st.number_input("Development duration (months)", value=int(s.timeline.dev_months), min_value=0, step=1, format="%d"))
    with c3:
        s.timeline.capex_months = int(st.number_input("Construction/CAPEX duration (months)", value=int(s.timeline.capex_months), min_value=0, step=1, format="%d"))

    s.timeline.operation_years = int(st.number_input("Operation duration (years)", value=int(s.timeline.operation_years), min_value=1, step=1, format="%d"))

    tl = build_timeline(s.timeline)
    _metric_row(
        [
            ("Start", str(tl["start"])),
            ("RTB (end of dev)", str(tl["rtb"])),
            ("COD (start ops)", str(tl["cod"])),
            ("End of operation", str(tl["end_op"])),
        ]
    )

    # visual bar
    st.markdown("#### Timeline visual (calendar, years)")

    gantt = pd.DataFrame(
        [
            {"Stage": "Development", "Start": date(tl["start"].year, tl["start"].month, 1), "Finish": date(tl["rtb"].year, tl["rtb"].month, 1)},
            {"Stage": "Construction", "Start": date(tl["rtb"].year, tl["rtb"].month, 1), "Finish": date(tl["cod"].year, tl["cod"].month, 1)},
            {"Stage": "Operation", "Start": date(tl["cod"].year, tl["cod"].month, 1), "Finish": date(tl["end_op"].year, tl["end_op"].month, 1)},
        ]
    )

    fig = px.timeline(
        gantt,
        x_start="Start",
        x_end="Finish",
        y="Stage",
        color="Stage",
        category_orders={"Stage": ["Development", "Construction", "Operation"]},
    )
    fig.update_yaxes(autorange="reversed")     # Development on top
    fig.update_xaxes(dtick="M12", tickformat="%Y")  # show years, yearly ticks
    fig.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)



# -----------------------------
# C) Generation
# -----------------------------
with tab_gen:
    st.subheader("Power generation inputs")

    c1, c2, c3 = st.columns(3)
    with c1:
        s.generation.mwac = st.number_input("Capacity (MWac)", value=float(s.generation.mwac), step=1.0, format="%.2f")
    with c2:
        s.generation.mwp = st.number_input("Capacity (MWp)", value=float(s.generation.mwp), step=1.0, format="%.2f")
    with c3:
        s.generation.degradation_pct = st.number_input("Annual degradation (%/yr)", value=float(s.generation.degradation_pct), step=0.05, format="%.2f")

    c4, c5, c6, c7 = st.columns(4)
    with c4:
        s.generation.p50_mwh_yr = st.number_input("P50 (MWh/year)", value=float(s.generation.p50_mwh_yr), step=1000.0, format="%.0f")
    with c5:
        s.generation.p75_mwh_yr = st.number_input("P75 (MWh/year)", value=float(s.generation.p75_mwh_yr), step=1000.0, format="%.0f")
    with c6:
        s.generation.p90_mwh_yr = st.number_input("P90 (MWh/year)", value=float(s.generation.p90_mwh_yr), step=1000.0, format="%.0f")
    with c7:
        s.generation.production_choice = st.selectbox("Production choice", ["P50", "P75", "P90"], index=["P50", "P75", "P90"].index(s.generation.production_choice))

    op = operating_year_table(s)
    fig = px.line(op, x="Year", y="Energy (MWh)")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    disp = op.copy()
    disp = _df_format_money(disp, ["Energy (MWh)", "Revenue (COP)", "Price (COP/kWh)"], decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)


# -----------------------------
# D) Revenues
# -----------------------------
with tab_rev:
    st.subheader("Power revenues (indexed, annual)")

    s.revenue_mode = st.radio("Revenue mode", ["Standard PPA Parameters", "Manual annual series"], horizontal=True, index=0 if s.revenue_mode == "Standard PPA Parameters" else 1)

    if s.revenue_mode == "Standard PPA Parameters":
        r = s.revenue1
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            r.ppa_price_cop_per_kwh = st.number_input("PPA price at COD (COP/kWh, current COP)", value=float(r.ppa_price_cop_per_kwh), step=1.0, format="%.2f")
        with c2:
            r.ppa_term_years = int(st.number_input("PPA term (years)", value=int(r.ppa_term_years), min_value=1, step=1, format="%d"))
        with c3:
            r.merchant_price_cop_per_kwh = st.number_input("Post-term merchant price (COP/kWh)", value=float(r.merchant_price_cop_per_kwh), step=1.0, format="%.2f")
        with c4:
            r.indexation = st.selectbox("Indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(r.indexation) if r.indexation in INDEX_CHOICES else 0)
    else:
        r = s.revenue2
        c1, c2 = st.columns([1, 2])
        with c1:
            r.indexation = st.selectbox("Indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(r.indexation) if r.indexation in INDEX_CHOICES else 0)
        with c2:
            st.info("Enter constant COP/kWh by operating year (Year 1 = first year at COD).")

        prices_df = pd.DataFrame({"OpYear": list(range(1, int(s.timeline.operation_years) + 1))})
        prices_df["COP_per_kWh_constant"] = prices_df["OpYear"].apply(lambda i: float(r.prices_constant_cop_per_kwh.get(int(i), 0.0)))
        edited = st.data_editor(
            prices_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "OpYear": st.column_config.NumberColumn("Operating year", format="%d", disabled=True),
                "COP_per_kWh_constant": st.column_config.NumberColumn("Price (COP/kWh, constant)", step=1.0, format="%.2f"),
            },
        )
        r.prices_constant_cop_per_kwh = {int(row.OpYear): float(row.COP_per_kWh_constant) for row in edited.itertuples(index=False)}

    op = operating_year_table(s)

    # Graph: energy + revenue
    c1, c2 = st.columns(2)
    with c1:
        fig1 = px.bar(op, x="Year", y="Energy (MWh)")
        fig1.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig1, use_container_width=True)
    with c2:
        fig2 = px.line(op, x="Year", y="Revenue (COP)")
        fig2.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    disp = op.copy()
    disp = _df_format_money(disp, ["Energy (MWh)", "Price (COP/kWh)", "Revenue (COP)"], decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)


# -----------------------------
# E) CAPEX
# -----------------------------
with tab_capex:
    st.subheader("CAPEX (COP)")

    s.capex.distribution = st.selectbox("Construction spend distribution", CAPEX_DISTS, index=CAPEX_DISTS.index(s.capex.distribution) if s.capex.distribution in CAPEX_DISTS else 0)

    capex_df = pd.DataFrame(s.capex.lines or [])
    for col in ["Item", "Amount_COP", "Phase"]:
        if col not in capex_df.columns:
            capex_df[col] = "" if col != "Amount_COP" else 0.0
    capex_df = capex_df[["Item", "Amount_COP", "Phase"]].copy()
    capex_df["Phase"] = capex_df["Phase"].where(capex_df["Phase"].isin(CAPEX_PHASES), "Construction")

    edited = st.data_editor(
        capex_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Item": st.column_config.TextColumn("CAPEX line item"),
            "Amount_COP": st.column_config.NumberColumn("Amount (COP)", min_value=0.0, step=1_000_000.0, format="%.0f"),
            "Phase": st.column_config.SelectboxColumn("Phase", options=CAPEX_PHASES),
        },
    )
    s.capex.lines = edited.to_dict(orient="records")

    total_capex = float(edited["Amount_COP"].fillna(0).sum())
    mwac = float(s.generation.mwac or 0.0)
    mwp = float(s.generation.mwp or 0.0)
    capex_per_mwac = (total_capex / mwac) if mwac > 0 else np.nan
    capex_per_mwp = (total_capex / mwp) if mwp > 0 else np.nan

    c1, c2, c3 = st.columns(3)
    c1.metric("Total CAPEX (COP)", _fmt_cop(total_capex))
    c2.metric("CAPEX / MWac (COP)", _fmt_cop(capex_per_mwac) if np.isfinite(capex_per_mwac) else "—")
    c3.metric("CAPEX / MWp (COP)", _fmt_cop(capex_per_mwp) if np.isfinite(capex_per_mwp) else "—")

    st.markdown("#### CAPEX breakdown (by line item)")

    capex_pie = edited.copy()
    capex_pie["Item"] = capex_pie["Item"].fillna("").astype(str).str.strip()
    capex_pie["Amount_COP"] = pd.to_numeric(capex_pie["Amount_COP"], errors="coerce").fillna(0.0)

    capex_pie = capex_pie[capex_pie["Amount_COP"] > 0].copy()

    if capex_pie.empty:
        st.info("Enter CAPEX amounts to see the breakdown chart.")
    else:
        # Optional: group very small items into "Other" to keep the chart readable
        share = capex_pie["Amount_COP"] / capex_pie["Amount_COP"].sum()
        small = share < 0.03  # <3%
        if small.any() and (~small).any():
            other_amt = float(capex_pie.loc[small, "Amount_COP"].sum())
            capex_pie = capex_pie.loc[~small, ["Item", "Amount_COP"]]
            capex_pie = pd.concat(
                [capex_pie, pd.DataFrame([{"Item": "Other (<3% each)", "Amount_COP": other_amt}])],
                ignore_index=True
            )

        fig_pie = px.pie(
            capex_pie,
            names="Item",
            values="Amount_COP",
            hole=0.45,  # donut style looks cleaner
        )
        fig_pie.update_traces(textinfo="percent+label")
        fig_pie.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
        st.plotly_chart(fig_pie, use_container_width=True)


    st.markdown("#### CAPEX schedule (monthly, aligned to timeline)")
    sched = capex_monthly_schedule(s)
    sched_disp = sched.copy()
    sched_disp = _df_format_money(sched_disp, ["CAPEX (COP)"], decimals=0)
    st.dataframe(sched_disp[["Month", "Phase", "CAPEX (COP)"]], use_container_width=True, hide_index=True)

    fig = px.bar(sched, x="Month", y="CAPEX (COP)", color="Phase")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Annual CAPEX (calendar years)")
    ann = sched.groupby("Year", as_index=False)["CAPEX (COP)"].sum()
    ann_disp = _df_format_money(ann, ["CAPEX (COP)"], decimals=0)
    st.dataframe(ann_disp, use_container_width=True, hide_index=True)


# -----------------------------
# F) OPEX
# -----------------------------
with tab_opex:
    st.subheader("OPEX (COP) — operating costs, land lease, taxes & levies")

    c1, c2, c3 = st.columns(3)
    with c1:
        s.opex.fixed_om_cop_per_mwac_year = st.number_input("Fixed O&M (COP/MWac-year)", value=float(s.opex.fixed_om_cop_per_mwac_year), step=1_000_000.0, format="%.0f")
        s.opex.fixed_om_indexation = st.selectbox("Fixed O&M indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(s.opex.fixed_om_indexation) if s.opex.fixed_om_indexation in INDEX_CHOICES else 0)
    with c2:
        s.opex.variable_om_cop_per_mwh = st.number_input("Variable O&M (COP/MWh)", value=float(s.opex.variable_om_cop_per_mwh), step=1_000.0, format="%.0f")
        s.opex.grid_fees_cop_per_mwh = st.number_input("Grid fees (COP/MWh)", value=float(s.opex.grid_fees_cop_per_mwh), step=1_000.0, format="%.0f")
    with c3:
        s.opex.insurance_cop_per_mwac_year = st.number_input("Insurance (COP/MWac-year)", value=float(s.opex.insurance_cop_per_mwac_year), step=1_000_000.0, format="%.0f")
        s.opex.insurance_indexation = st.selectbox("Insurance indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(s.opex.insurance_indexation) if s.opex.insurance_indexation in INDEX_CHOICES else 0)

    st.markdown("#### Land lease")
    lc1, lc2, lc3, lc4, lc5 = st.columns([1, 1, 1, 1, 1.2])
    with lc1:
        s.opex.land_hectares = st.number_input("Hectares leased (Ha)", value=float(s.opex.land_hectares), step=1.0, format="%.2f")
    with lc2:
        s.opex.land_price_dev_cop_per_ha_year = st.number_input("Dev lease (COP/Ha-year)", value=float(s.opex.land_price_dev_cop_per_ha_year), step=1_000_000.0, format="%.0f")
    with lc3:
        s.opex.land_price_con_cop_per_ha_year = st.number_input("Const lease (COP/Ha-year)", value=float(s.opex.land_price_con_cop_per_ha_year), step=1_000_000.0, format="%.0f")
    with lc4:
        s.opex.land_price_op_cop_per_ha_year = st.number_input("Op lease (COP/Ha-year)", value=float(s.opex.land_price_op_cop_per_ha_year), step=1_000_000.0, format="%.0f")
    with lc5:
        s.opex.land_indexation = st.selectbox("Land lease indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(s.opex.land_indexation) if s.opex.land_indexation in INDEX_CHOICES else 0)

    st.markdown("#### Taxes & levies")
    tc1, tc2 = st.columns(2)
    with tc1:
        s.opex.ica_pct_of_revenue = st.number_input("ICA (% of revenue)", value=float(s.opex.ica_pct_of_revenue), step=0.01, format="%.3f")
    with tc2:
        s.opex.gmf_pct_of_outflows = st.number_input("GMF (% of outgoing cash)", value=float(s.opex.gmf_pct_of_outflows), step=0.01, format="%.3f")

    st.markdown("#### Other OPEX items (dynamic)")
    o_df = pd.DataFrame(s.opex.other_items or [])
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in o_df.columns:
            o_df[col] = "" if col != "Amount_COP_per_year" else 0.0
    o_df = o_df[["Item", "Amount_COP_per_year", "Phase", "Indexation"]].copy()
    o_df["Phase"] = o_df["Phase"].where(o_df["Phase"].isin(PHASES), "Operation")
    o_df["Indexation"] = o_df["Indexation"].where(o_df["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    o_edited = st.data_editor(
        o_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Item": st.column_config.TextColumn("OPEX item"),
            "Amount_COP_per_year": st.column_config.NumberColumn("Amount (COP/year)", min_value=0.0, step=1_000_000.0, format="%.0f"),
            "Phase": st.column_config.SelectboxColumn("Phase", options=PHASES),
            "Indexation": st.column_config.SelectboxColumn("Indexation", options=INDEX_CHOICES),
        },
    )
    s.opex.other_items = o_edited.to_dict(orient="records")

    st.divider()
    st.markdown("## Outputs")

    om = opex_monthly_schedule(s).copy()
    annual = om.groupby("Year", as_index=False)[["OPEX subtotal", "GMF"]].sum()
    annual["Total OPEX (COP)"] = annual["OPEX subtotal"] + annual["GMF"]

    # OPEX per MWh (operation years)
    op = operating_year_table(s)[["Year", "Energy (MWh)"]].copy()
    annual = annual.merge(op, on="Year", how="left").fillna({"Energy (MWh)": 0.0})
    annual["OPEX per MWh (COP/MWh)"] = np.where(annual["Energy (MWh)"] > 0, annual["Total OPEX (COP)"] / annual["Energy (MWh)"], 0.0)

    # Stacked OPEX by category (uses the monthly schedule columns)
    om_full = opex_monthly_schedule(s).copy()

    # Identify the OPEX item columns that make up "OPEX subtotal"
    meta_cols = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)", "OPEX subtotal", "GMF"}
    fixed_cols = {"Fixed O&M", "Insurance", "Variable O&M", "Grid fees", "Land lease", "ICA"}
    dyn_cols = [c for c in om_full.columns if c not in meta_cols and c not in fixed_cols]
    opex_item_cols = list(fixed_cols) + dyn_cols

    annual_items = om_full.groupby("Year", as_index=False)[opex_item_cols + ["GMF"]].sum()
    annual_items["Total OPEX (COP)"] = annual_items[opex_item_cols].sum(axis=1) + annual_items["GMF"]

    # Convert to long form for stacked bar
    long = annual_items.melt(id_vars=["Year"], value_vars=opex_item_cols + ["GMF"], var_name="Item", value_name="OPEX (COP)")

    fig = px.bar(long, x="Year", y="OPEX (COP)", color="Item", barmode="stack")
    fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)


    disp = annual.copy()
    disp = _df_format_money(disp, ["OPEX subtotal", "GMF", "Total OPEX (COP)", "OPEX per MWh (COP/MWh)", "Energy (MWh)"], decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)


# -----------------------------
# G) SG&A
# -----------------------------
with tab_sga:
    st.subheader("SG&A (COP) — Development, Construction, and Operation")

    sga_df = pd.DataFrame(s.sga.items or [])
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in sga_df.columns:
            sga_df[col] = "" if col != "Amount_COP_per_year" else 0.0
    sga_df = sga_df[["Item", "Amount_COP_per_year", "Phase", "Indexation"]].copy()
    sga_df["Phase"] = sga_df["Phase"].where(sga_df["Phase"].isin(PHASES), "Development")
    sga_df["Indexation"] = sga_df["Indexation"].where(sga_df["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    sga_edited = st.data_editor(
        sga_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Item": st.column_config.TextColumn("SG&A item"),
            "Amount_COP_per_year": st.column_config.NumberColumn("Amount (COP/year)", min_value=0.0, step=1_000_000.0, format="%.0f"),
            "Phase": st.column_config.SelectboxColumn("Phase", options=PHASES),
            "Indexation": st.column_config.SelectboxColumn("Indexation", options=INDEX_CHOICES),
        },
    )
    s.sga.items = sga_edited.to_dict(orient="records")

    st.divider()
    st.markdown("## Outputs")

    annual_sga = sga_annual_by_item(s)
    item_cols = [c for c in annual_sga.columns if c not in ["Year", "Total SG&A (COP)"]]
    if item_cols:
        annual_long = annual_sga.melt(id_vars=["Year"], value_vars=item_cols, var_name="Item", value_name="SG&A (COP)")
        fig = px.bar(annual_long, x="Year", y="SG&A (COP)", color="Item", barmode="stack")
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Add SG&A line items to see the chart.")

    annual_disp = annual_sga.copy()
    money_cols = [c for c in annual_disp.columns if c != "Year"]
    annual_disp = _df_format_money(annual_disp, money_cols, decimals=0)
    st.dataframe(annual_disp, use_container_width=True, hide_index=True)

# -----------------------------
# I) Depreciation
# -----------------------------
with tab_dep:
    st.subheader("Depreciation (linear, starting at COD)")

    c1, c2 = st.columns(2)
    with c1:
        s.depreciation.pct_of_capex_depreciated = st.number_input(
            "% of CAPEX depreciated",
            value=float(s.depreciation.pct_of_capex_depreciated),
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            format="%.1f",
        )
    with c2:
        s.depreciation.dep_years = int(
            st.number_input(
                "Depreciation period (years after COD)",
                value=int(s.depreciation.dep_years),
                min_value=3,
                max_value=25,
                step=1,
                format="%d",
            )
        )

    dep = depreciation_annual_table(s)

    fig = px.bar(dep, x="Year", y="Depreciation (COP)")
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    dep_disp = dep.copy()
    money_cols = [c for c in dep_disp.columns if c != "Year"]
    dep_disp = _df_format_money(dep_disp, money_cols, decimals=0)
    st.dataframe(dep_disp, use_container_width=True, hide_index=True)

# -----------------------------
# H) Unlevered Base Cash Flow
# -----------------------------
with tab_ucf:
    st.subheader("Unlevered Base Cash Flow (pre debt, no tax credits)")

    # Tax inputs
    tx1, tx2 = st.columns([1, 2])
    with tx1:
        s.tax.corporate_tax_rate_pct = st.number_input("Corporate income tax rate (%)", value=float(s.tax.corporate_tax_rate_pct), step=0.5, format="%.2f")
    with tx2:
        s.tax.allow_loss_carryforward = st.checkbox("Apply loss carryforward (NOL) so taxes are zero until losses are used", value=bool(s.tax.allow_loss_carryforward))

    currency = st.radio("Display currency", ["COP", "USD"], horizontal=True, index=0)

    m = cashflow_monthly_table(s)
    a = unlevered_base_cashflow_annual(s)

    # FX for annual conversion (USD view)
    tl = build_timeline(s.timeline)
    years = list(a["Year"].astype(int).tolist())
    fx = fx_series(s.macro, tl["cod"].year, years)  # COP per USD

    def _conv(series: pd.Series, year_col: pd.Series) -> pd.Series:
        if currency == "COP":
            return series
        # USD = COP / FX
        return series / year_col.map(lambda y: float(fx.loc[int(y)]) if int(y) in fx.index else float(s.macro.fx_cop_per_usd_start))

    # headline metrics
    total_rev = float(m["Revenue (COP)"].sum())
    total_ebitda = float(a["EBITDA (COP)"].sum())
    total_tax = float(a["Taxes Payable (COP)"].sum())
    end_cum_pre_tax = float(m["Cumulative Unlevered CF (COP)"].iloc[-1]) if len(m) else 0.0

    if currency == "COP":
        _metric_row(
            [
                ("Total Revenue", _fmt_cop(total_rev)),
                ("Total EBITDA", _fmt_cop(total_ebitda)),
                ("Total Taxes Payable", _fmt_cop(total_tax)),
                ("End Cumulative Unlevered CF (pre-tax)", _fmt_cop(end_cum_pre_tax)),
            ]
        )
    else:
        # rough conversions for headline metrics using start FX (fast + stable)
        fx0 = float(s.macro.fx_cop_per_usd_start)
        _metric_row(
            [
                ("Total Revenue", _fmt_usd(total_rev / fx0)),
                ("Total EBITDA", _fmt_usd(total_ebitda / fx0)),
                ("Total Taxes Payable", _fmt_usd(total_tax / fx0)),
                ("End Cumulative Unlevered CF (pre-tax)", _fmt_usd(end_cum_pre_tax / fx0)),
            ]
        )

    st.divider()

    # --- KPIs (pre-debt) ---
    m = cashflow_monthly_table(s).copy()

    # Ensure numeric
    for col in ["CAPEX (COP)", "Unlevered CF (COP)", "Cumulative Unlevered CF (COP)"]:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors="coerce").fillna(0.0)

    cap_total = float(m["CAPEX (COP)"].sum()) if "CAPEX (COP)" in m.columns else 0.0

    monthly_cf = m["Unlevered CF (COP)"].astype(float).tolist() if "Unlevered CF (COP)" in m.columns else []
    monthly_dates = m["Month"].tolist() if "Month" in m.columns else []

    # IRR only if there is at least one negative and one positive cashflow
    has_pos = any(cf > 0 for cf in monthly_cf)
    has_neg = any(cf < 0 for cf in monthly_cf)
    irr_m = _irr_bisection(monthly_cf) if (has_pos and has_neg) else float("nan")
    irr_annual_equiv = (1.0 + irr_m) ** 12 - 1.0 if np.isfinite(irr_m) else float("nan")

    # Payback: only meaningful if cumulative ever goes from negative to >=0
    payback_m = _payback_months(monthly_dates, monthly_cf) if monthly_cf else float("nan")
    payback_years = payback_m / 12.0 if np.isfinite(payback_m) else float("nan")

    cum = m["Cumulative Unlevered CF (COP)"].astype(float) if "Cumulative Unlevered CF (COP)" in m.columns else pd.Series([0.0])
    min_cum = float(cum.min()) if len(cum) else 0.0
    peak_funding = max(0.0, -min_cum)  # positive number; never show -0

    _metric_row([
        ("Total Investment (CAPEX)", _fmt_cop(cap_total)),
        ("Unlevered IRR (annualized, pre-tax)", f"{irr_annual_equiv*100:,.2f}%" if np.isfinite(irr_annual_equiv) else "—"),
        ("Payback (years, pre-tax)", f"{payback_years:,.2f}" if np.isfinite(payback_years) else "—"),
        ("Peak Funding Need", _fmt_cop(peak_funding)),
    ])


    # Annual table (with currency switch)
    annual_view = a.copy()
    money_cols = [
        "Revenue (COP)", "Total OPEX (COP)", "SG&A (COP)", "EBITDA (COP)", "Depreciation (COP)",
        "Taxable Income (COP)", "Taxes Payable (COP)", "CAPEX (COP)", "Unlevered CF (COP)", "Unlevered CF After Tax (COP)",
        "Loss Carryforward End (COP)", "Cumulative Unlevered CF (COP)",
    ]
    for c in money_cols:
        if c in annual_view.columns:
            annual_view[c] = _conv(annual_view[c], annual_view["Year"])

    # relabel currency in column names for display
    if currency == "USD":
        ren = {c: c.replace("(COP)", "(USD)") for c in annual_view.columns if "(COP)" in c}
        annual_view = annual_view.rename(columns=ren)

    st.markdown("### Annual summary (calendar years)")
    display_cols = [
        "Year",
        "Revenue (COP)" if currency == "COP" else "Revenue (USD)",
        "Total OPEX (COP)" if currency == "COP" else "Total OPEX (USD)",
        "SG&A (COP)" if currency == "COP" else "SG&A (USD)",
        "EBITDA (COP)" if currency == "COP" else "EBITDA (USD)",
        "Depreciation (COP)" if currency == "COP" else "Depreciation (USD)",
        "Taxable Income (COP)" if currency == "COP" else "Taxable Income (USD)",
        "Taxes Payable (COP)" if currency == "COP" else "Taxes Payable (USD)",
        "Unlevered CF (COP)" if currency == "COP" else "Unlevered CF (USD)",
        "Unlevered CF After Tax (COP)" if currency == "COP" else "Unlevered CF After Tax (USD)",
        "Loss Carryforward End (COP)" if currency == "COP" else "Loss Carryforward End (USD)",
    ]
    display_cols = [c for c in display_cols if c in annual_view.columns]
    disp = annual_view[display_cols].copy()

    # format: only numeric columns (exclude Year)
    fmt_cols = [c for c in disp.columns if c != "Year"]
    disp = _df_format_money(disp, fmt_cols, decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    # chart: after-tax unlevered CF
    y_after = "Unlevered CF After Tax (COP)" if currency == "COP" else "Unlevered CF After Tax (USD)"
    fig = px.bar(annual_view, x="Year", y=y_after)
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Monthly cash flow (detailed, pre-tax)")
    m_disp = m.copy()
    m_money = [c for c in m_disp.columns if c not in ["Month", "Year", "Phase"]]
    m_disp = _df_format_money(m_disp, m_money, decimals=0)
    st.dataframe(m_disp, use_container_width=True, hide_index=True)


# -----------------------------
# J) Compare
# -----------------------------
with tab_compare:
    st.subheader("Scenario comparison")

    if not compare_scenarios:
        st.warning("Select scenarios to compare in the sidebar.")
    else:
        rows = []
        for nm in compare_scenarios:
            sd = _scenario_from_dict(proj["scenarios"][nm])

            rev = operating_year_table(sd)
            cap_df = pd.DataFrame(sd.capex.lines or [])
            total_cap = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

            om = opex_monthly_schedule(sd)
            annual_opex = om.groupby("Year", as_index=False)[["OPEX subtotal", "GMF"]].sum()
            annual_opex["Total OPEX (COP)"] = annual_opex["OPEX subtotal"] + annual_opex["GMF"]

            total_opex_op = float(annual_opex["Total OPEX (COP)"].sum())

            rows.append(
                {
                    "Scenario": nm,
                    "Total CAPEX (COP)": total_cap,
                    "CAPEX/MWac (COP)": total_cap / float(sd.generation.mwac) if float(sd.generation.mwac) > 0 else np.nan,
                    "Total OPEX (COP)": total_opex_op,
                    "Total Revenue (COP)": float(rev["Revenue (COP)"].sum()),
                }
            )

        summary = pd.DataFrame(rows).sort_values("Scenario")
        disp = summary.copy()
        for col in ["Total CAPEX (COP)", "CAPEX/MWac (COP)", "Total OPEX (COP)", "Total Revenue (COP)"]:
            if col in disp.columns:
                disp[col] = disp[col].apply(lambda v: _fmt_num(float(v), 0) if pd.notnull(v) else "")
        st.dataframe(disp, use_container_width=True, hide_index=True)


# Persist scenario on each run
proj["scenarios"][scenario_name] = _scenario_to_dict(s)
_save_db(db)
