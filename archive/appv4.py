# Delphi Utility-Scale Financial Model (No Excel)
# Streamlit single-file app with: Projects + Scenarios, Macro, Timeline, Generation, Revenues, CAPEX, OPEX, Comparison
# All inputs in COP; outputs selectable COP/USD (USD via FX path).
#
# Files saved to:
#   G:\My Drive\Delphi Financial Models\data\projects.json
#
# Run:
#   py -m pip install -r requirements.txt
#   py -m streamlit run app.py

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict, field
from datetime import date
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------
# Storage (Google Drive folder)
# -----------------------------
DATA_DIR = r"G:\My Drive\Delphi Financial Models\data"
PROJECTS_FILE = os.path.join(DATA_DIR, "projects.json")


def _ensure_storage():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(PROJECTS_FILE):
        with open(PROJECTS_FILE, "w", encoding="utf-8") as f:
            json.dump({"projects": {}}, f, indent=2)


def _load_db() -> dict:
    _ensure_storage()
    with open(PROJECTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_db(db: dict) -> None:
    _ensure_storage()
    with open(PROJECTS_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2)


def _today() -> date:
    return date.today()


# -----------------------------
# Constants
# -----------------------------
INDEX_CHOICES = ["Colombia CPI", "Colombia PPI", "US CPI", "Custom"]

PHASES = ["Development", "Construction", "Operation"]
CAPEX_PHASES = ["Development", "Construction", "At COD"]
CAPEX_DISTRIBUTIONS = ["Straight-line (monthly)", "Front-loaded", "Back-loaded"]


# -----------------------------
# Model structures
# -----------------------------
@dataclass
class MacroInputs:
    col_cpi: float = 6.0
    col_ppi: float = 5.0
    us_cpi: float = 3.0
    fx_cop_per_usd_start: float = 4200.0
    fx_method: str = "Inflation differential (PPP approx.)"  # or "Flat"
    fx_flat: Optional[float] = None
    custom_index_rate: float = 6.0
    yearly_overrides: Dict[str, Dict[int, float]] = field(
        default_factory=lambda: {"col_cpi": {}, "col_ppi": {}, "us_cpi": {}, "custom": {}}
    )


@dataclass
class TimelineInputs:
    start_date: str = ""
    dev_months: int = 18
    capex_months: int = 12
    operation_years: int = 25
    bank_guarantee_month: Optional[int] = None


@dataclass
class GenerationInputs:
    mwac: float = 80.0
    mwp: float = 100.0
    p50_mwh_yr: float = 200000.0
    p75_mwh_yr: float = 190000.0
    p90_mwh_yr: float = 180000.0
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


def _default_capex_lines() -> List[Dict[str, object]]:
    return [
        {"Item": "EPC (modules + BOS + installation)", "Amount_COP": 0.0, "Phase": "Construction"},
        {"Item": "Interconnection / Substation / Line", "Amount_COP": 0.0, "Phase": "Construction"},
        {"Item": "Development costs", "Amount_COP": 0.0, "Phase": "Development"},
        {"Item": "Owner's costs", "Amount_COP": 0.0, "Phase": "Construction"},
        {"Item": "Contingency", "Amount_COP": 0.0, "Phase": "Construction"},
    ]


@dataclass
class CapexInputs:
    lines: List[Dict[str, object]] = field(default_factory=_default_capex_lines)
    distribution: str = "Straight-line (monthly)"


def _default_opex_other_items() -> List[Dict[str, object]]:
    return [
        {"Item": "Security / Monitoring", "Amount_COP_per_year": 0.0, "Phase": "Operation", "Indexation": "Colombia CPI"},
        {"Item": "Admin / Back office", "Amount_COP_per_year": 0.0, "Phase": "Operation", "Indexation": "Colombia CPI"},
    ]


@dataclass
class OpexInputs:
    fixed_om_cop_per_mwac_year: float = 0.0
    fixed_om_indexation: str = "Colombia CPI"

    variable_om_cop_per_mwh: float = 0.0
    variable_om_indexation: str = "Colombia CPI"  # reserved

    insurance_cop_per_mwac_year: float = 0.0
    insurance_indexation: str = "Colombia CPI"

    grid_fees_cop_per_mwh: float = 0.0

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
    pct_of_capex_depreciated: float = 100.0   # default 100%
    dep_years: int = 20                       # user can pick 3–25


@dataclass
class ScenarioInputs:
    name: str = "Base"
    macro: MacroInputs = field(default_factory=MacroInputs)
    timeline: TimelineInputs = field(default_factory=TimelineInputs)
    generation: GenerationInputs = field(default_factory=GenerationInputs)
    capex: CapexInputs = field(default_factory=CapexInputs)
    opex: OpexInputs = field(default_factory=OpexInputs)
    sga: SgaInputs = field(default_factory=SgaInputs)
    depreciation: DepreciationInputs = field(default_factory=DepreciationInputs)
    revenue_mode: str = "Standard PPA Parameters"
    revenue1: RevenueOption1PPA = field(default_factory=RevenueOption1PPA)
    revenue2: RevenueOption2Manual = field(default_factory=RevenueOption2Manual)


def _scenario_to_dict(s: ScenarioInputs) -> dict:
    return asdict(s)


def _scenario_from_dict(d: dict) -> ScenarioInputs:
    macro = MacroInputs(**d.get("macro", {}))
    timeline = TimelineInputs(**d.get("timeline", {}))
    generation = GenerationInputs(**d.get("generation", {}))
    capex = CapexInputs(**d.get("capex", {})) if "capex" in d else CapexInputs()
    opex = OpexInputs(**d.get("opex", {})) if "opex" in d else OpexInputs()
    sga = SgaInputs(**d.get("sga", {})) if "sga" in d else SgaInputs()
    depreciation = DepreciationInputs(**d.get("depreciation", {})) if "depreciation" in d else DepreciationInputs()
    revenue1 = RevenueOption1PPA(**d.get("revenue1", {}))
    revenue2 = RevenueOption2Manual(**d.get("revenue2", {}))
    return ScenarioInputs(
        name=d.get("name", "Base"),
        macro=macro,
        timeline=timeline,
        generation=generation,
        capex=capex,
        opex=opex,
        sga=sga,
        depreciation=depreciation,
        revenue_mode=d.get("revenue_mode", "Standard PPA Parameters"),
        revenue1=revenue1,
        revenue2=revenue2,
    )


# -----------------------------
# Helpers: formatting (commas)
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
    return f"COP {_fmt_num(x, 0)}"


def _fmt_usd(x: float) -> str:
    return f"USD {_fmt_num(x, 0)}"


def _df_format_money(df: pd.DataFrame, cols: List[str], decimals: int = 0) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda v: _fmt_num(float(v), decimals) if pd.notnull(v) else "")
    return out


# -----------------------------
# Date utilities
# -----------------------------
def _parse_date_iso(s: str) -> date:
    return date.fromisoformat(s)


def _add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    day = min(
        d.day,
        [
            31,
            29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28,
            31,
            30,
            31,
            30,
            31,
            31,
            30,
            31,
            30,
            31,
        ][m - 1],
    )
    return date(y, m, day)


def build_timeline(t: TimelineInputs) -> dict:
    start = _parse_date_iso(t.start_date) if t.start_date else _today()
    rtb = _add_months(start, int(t.dev_months))
    cod = _add_months(rtb, int(t.capex_months))
    end_op = _add_months(cod, int(t.operation_years) * 12)
    return {"start": start, "rtb": rtb, "cod": cod, "end_op": end_op}


def _month_starts(start: date, months: int) -> List[date]:
    return [_add_months(start, i) for i in range(months)]


def _months_between_inclusive(start: date, end: date) -> List[date]:
    months = []
    cur = date(start.year, start.month, 1)
    endm = date(end.year, end.month, 1)
    while cur <= endm:
        months.append(cur)
        cur = _add_months(cur, 1)
    return months


def _phase_for_month(m: date, tl: dict) -> str:
    rtb_m = date(tl["rtb"].year, tl["rtb"].month, 1)
    cod_m = date(tl["cod"].year, tl["cod"].month, 1)
    if m < rtb_m:
        return "Development"
    if m < cod_m:
        return "Construction"
    return "Operation"


# -----------------------------
# Indexing / FX
# -----------------------------
def annual_index_series(macro: MacroInputs, base_year: int, years: List[int], which: str) -> pd.Series:
    base_rate = {
        "col_cpi": macro.col_cpi,
        "col_ppi": macro.col_ppi,
        "us_cpi": macro.us_cpi,
        "custom": macro.custom_index_rate,
    }[which]
    overrides = macro.yearly_overrides.get(which, {})

    idx: Dict[int, float] = {}
    for y in years:
        if y == base_year:
            idx[y] = 1.0
        elif y > base_year:
            prev = y - 1
            prev_level = idx.get(prev, 1.0)
            rate = overrides.get(y, base_rate) / 100.0
            idx[y] = prev_level * (1.0 + rate)
        else:
            nxt = y + 1
            nxt_level = idx.get(nxt, 1.0)
            rate = overrides.get(nxt, base_rate) / 100.0
            idx[y] = nxt_level / (1.0 + rate)

    return pd.Series(idx).reindex(years).astype(float)


def fx_path_cop_per_usd(macro: MacroInputs, base_year: int, years: List[int]) -> pd.Series:
    fx0 = float(macro.fx_flat) if (macro.fx_method == "Flat" and macro.fx_flat) else float(macro.fx_cop_per_usd_start)
    if macro.fx_method == "Flat":
        return pd.Series({y: fx0 for y in years}).reindex(years).astype(float)

    col = annual_index_series(macro, base_year, years, "col_cpi")
    us = annual_index_series(macro, base_year, years, "us_cpi")
    return fx0 * (col / us)


def _idx_key(index_choice: str) -> str:
    return {"Colombia CPI": "col_cpi", "Colombia PPI": "col_ppi", "US CPI": "us_cpi", "Custom": "custom"}[index_choice]


def _index_factor_for_year(macro: MacroInputs, year: int, base_year: int, index_choice: str) -> float:
    years = list(range(min(base_year, year), max(base_year, year) + 1))
    s = annual_index_series(macro, base_year, years, _idx_key(index_choice))
    return float(s.loc[year])


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
        price_base = [
            float(r.prices_constant_cop_per_kwh.get(i + 1, list(r.prices_constant_cop_per_kwh.values())[-1]))
            for i in range(op_years)
        ]

    idx = annual_index_series(s.macro, base_year, years, _idx_key(index_choice))
    price_cop_per_kwh = np.array(price_base) * idx.values

    kwh = np.array(mwh) * 1000.0
    revenue_cop = kwh * price_cop_per_kwh

    fx = fx_path_cop_per_usd(s.macro, base_year, years)
    revenue_usd = revenue_cop / fx.values

    return pd.DataFrame(
        {
            "Year": years,
            "Operating Year #": list(range(1, op_years + 1)),
            "Energy (MWh)": mwh,
            "Price (COP/kWh)": price_cop_per_kwh,
            "Revenue (COP)": revenue_cop,
            "FX (COP/USD)": fx.values,
            "Revenue (USD)": revenue_usd,
        }
    )


def _monthly_operation_energy_and_revenue(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    cod = date(tl["cod"].year, tl["cod"].month, 1)
    end_op = date(tl["end_op"].year, tl["end_op"].month, 1)

    op_annual = operating_year_table(s)[["Year", "Energy (MWh)", "Revenue (COP)"]].copy()
    months = _months_between_inclusive(cod, end_op)

    rows = []
    for m in months:
        y = m.year
        if y not in set(op_annual["Year"]):
            continue
        yr_row = op_annual.loc[op_annual["Year"] == y].iloc[0]
        rows.append(
            {
                "Month": m,
                "Year": y,
                "Energy (MWh)": float(yr_row["Energy (MWh)"]) / 12.0,
                "Revenue (COP)": float(yr_row["Revenue (COP)"]) / 12.0,
            }
        )
    return pd.DataFrame(rows)


# -----------------------------
# OPEX (monthly + annual)
# -----------------------------
def opex_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    start = date(tl["start"].year, tl["start"].month, 1)
    end_op = date(tl["end_op"].year, tl["end_op"].month, 1)
    months = _months_between_inclusive(start, end_op)

    op_monthly = _monthly_operation_energy_and_revenue(s)
    op_monthly = op_monthly.set_index("Month") if not op_monthly.empty else pd.DataFrame()

    base_year_dev = tl["start"].year
    base_year_con = tl["rtb"].year
    base_year_op = tl["cod"].year

    mwac = float(s.generation.mwac or 0.0)
    o = s.opex

    rows = []
    for m in months:
        phase = _phase_for_month(m, tl)
        y = m.year

        # Land lease
        if phase == "Development":
            land_rate = float(o.land_price_dev_cop_per_ha_year or 0.0)
            land_base_year = base_year_dev
        elif phase == "Construction":
            land_rate = float(o.land_price_con_cop_per_ha_year or 0.0)
            land_base_year = base_year_con
        else:
            land_rate = float(o.land_price_op_cop_per_ha_year or 0.0)
            land_base_year = base_year_op

        land_idx = _index_factor_for_year(s.macro, y, land_base_year, o.land_indexation)
        land_month = float(o.land_hectares or 0.0) * land_rate * land_idx / 12.0

        if phase == "Operation":
            # Energy / revenue
            if not op_monthly.empty and m in op_monthly.index:
                energy_mwh = float(op_monthly.loc[m, "Energy (MWh)"])
                revenue_cop = float(op_monthly.loc[m, "Revenue (COP)"])
            else:
                energy_mwh = 0.0
                revenue_cop = 0.0

            # Fixed O&M
            fom_idx = _index_factor_for_year(s.macro, y, base_year_op, o.fixed_om_indexation)
            fixed_om_month = (float(o.fixed_om_cop_per_mwac_year or 0.0) * mwac * fom_idx) / 12.0

            # Insurance
            ins_idx = _index_factor_for_year(s.macro, y, base_year_op, o.insurance_indexation)
            insurance_month = (float(o.insurance_cop_per_mwac_year or 0.0) * mwac * ins_idx) / 12.0

            # Variable O&M + grid fees
            variable_om_month = float(o.variable_om_cop_per_mwh or 0.0) * energy_mwh
            grid_fees_month = float(o.grid_fees_cop_per_mwh or 0.0) * energy_mwh

            # ICA
            ica_month = (float(o.ica_pct_of_revenue or 0.0) / 100.0) * revenue_cop
        else:
            energy_mwh = 0.0
            revenue_cop = 0.0
            fixed_om_month = 0.0
            insurance_month = 0.0
            variable_om_month = 0.0
            grid_fees_month = 0.0
            ica_month = 0.0

        # Other items (COP/year), by phase
        other_costs: Dict[str, float] = {}
        for item in (o.other_items or []):
            it_name = str(item.get("Item", "Other")).strip() or "Other"
            it_phase = item.get("Phase", "Operation")
            it_amount = float(item.get("Amount_COP_per_year", 0.0) or 0.0)
            it_index = item.get("Indexation", "Colombia CPI")

            if it_phase != phase:
                continue

            base_year = base_year_op if phase == "Operation" else (base_year_con if phase == "Construction" else base_year_dev)
            idx = _index_factor_for_year(s.macro, y, base_year, it_index)
            other_costs[it_name] = other_costs.get(it_name, 0.0) + (it_amount * idx / 12.0)

        row = {
            "Month": m,
            "Year": y,
            "Phase": phase,
            "Energy (MWh)": energy_mwh,
            "Revenue (COP)": revenue_cop,
            "Land lease": land_month,
            "Fixed O&M": fixed_om_month,
            "Variable O&M": variable_om_month,
            "Insurance": insurance_month,
            "Grid fees": grid_fees_month,
            "ICA": ica_month,
        }
        for k, v in other_costs.items():
            row[k] = v

        rows.append(row)

    df = pd.DataFrame(rows).fillna(0.0)

    # Merge CAPEX monthly (for GMF)
    cap = capex_monthly_schedule(s)[["Month", "CAPEX (COP)"]].copy()
    cap["Month"] = cap["Month"].apply(lambda d: date(d.year, d.month, 1))
    df = df.merge(cap, on="Month", how="left")
    df["CAPEX (COP)"] = df["CAPEX (COP)"].fillna(0.0)

    meta_cols = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)"}
    cost_cols = [c for c in df.columns if c not in meta_cols]
    df["OPEX subtotal"] = df[cost_cols].sum(axis=1)

    gmf_rate = float(s.opex.gmf_pct_of_outflows or 0.0) / 100.0
    df["GMF"] = gmf_rate * (df["CAPEX (COP)"] + df["OPEX subtotal"])

    return df


def opex_annual_by_item(s: ScenarioInputs) -> Tuple[pd.DataFrame, pd.DataFrame]:
    dfm = opex_monthly_schedule(s).copy()

    meta_cols = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)", "OPEX subtotal"}
    item_cols = [c for c in dfm.columns if c not in meta_cols]

    annual = dfm.groupby("Year", as_index=False)[item_cols].sum()
    annual["Total OPEX (COP)"] = annual[item_cols].sum(axis=1)  # <-- requested total per year

    op = operating_year_table(s)[["Year", "Energy (MWh)"]].copy()
    df_op = dfm[dfm["Phase"] == "Operation"].copy()
    annual_op = df_op.groupby("Year", as_index=False)[item_cols].sum()
    op = op.merge(annual_op, on="Year", how="left").fillna(0.0)

    op["Total OPEX (COP)"] = op[item_cols].sum(axis=1)
    op["OPEX per MWh (COP/MWh)"] = np.where(op["Energy (MWh)"] > 0, op["Total OPEX (COP)"] / op["Energy (MWh)"], np.nan)

    return annual, op

def sga_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    """
    SG&A modeled as COP/year items assigned to a Phase.
    Each item is indexed annually, then allocated evenly across the months that exist in that Phase for that calendar year.
    """
    tl = build_timeline(s.timeline)
    start = date(tl["start"].year, tl["start"].month, 1)
    end_op = date(tl["end_op"].year, tl["end_op"].month, 1)
    months = _months_between_inclusive(start, end_op)

    base_year_dev = tl["start"].year
    base_year_con = tl["rtb"].year
    base_year_op = tl["cod"].year

    # Count how many months per (year, phase) exist so we can allocate per-year-per-phase properly
    phase_month_counts: Dict[Tuple[int, str], int] = {}
    for m in months:
        ph = _phase_for_month(m, tl)
        phase_month_counts[(m.year, ph)] = phase_month_counts.get((m.year, ph), 0) + 1

    rows = []
    for m in months:
        y = m.year
        ph = _phase_for_month(m, tl)

        row = {"Month": m, "Year": y, "Phase": ph}

        for it in (s.sga.items or []):
            name = str(it.get("Item", "SG&A")).strip() or "SG&A"
            item_phase = it.get("Phase", "Development")
            amt_year = float(it.get("Amount_COP_per_year", 0.0) or 0.0)
            idx_choice = it.get("Indexation", "Colombia CPI")

            if item_phase != ph:
                continue

            base_y = base_year_op
            if ph == "Development":
                base_y = base_year_dev
            elif ph == "Construction":
                base_y = base_year_con

            idx = _index_factor_for_year(s.macro, y, base_y, idx_choice)

            months_in_year_phase = phase_month_counts.get((y, ph), 0)
            if months_in_year_phase <= 0:
                continue

            row[name] = row.get(name, 0.0) + (amt_year * idx) / months_in_year_phase

        rows.append(row)

    return pd.DataFrame(rows).fillna(0.0)


def sga_annual_by_item(s: ScenarioInputs) -> pd.DataFrame:
    dfm = sga_monthly_schedule(s).copy()
    meta = {"Month", "Year", "Phase"}
    item_cols = [c for c in dfm.columns if c not in meta]

    annual = dfm.groupby("Year", as_index=False)[item_cols].sum() if item_cols else dfm.groupby("Year", as_index=False).size()[["Year"]]
    if item_cols:
        annual["Total SG&A (COP)"] = annual[item_cols].sum(axis=1)
    else:
        annual["Total SG&A (COP)"] = 0.0
    return annual

# -----------------------------
# CASH FLOW (Operating CF pre D&A / taxes / financing)
# -----------------------------
def _sga_monthly_total(s: ScenarioInputs) -> pd.DataFrame:
    df = sga_monthly_schedule(s).copy()
    meta = {"Month", "Year", "Phase"}
    item_cols = [c for c in df.columns if c not in meta]
    df["SG&A (COP)"] = df[item_cols].sum(axis=1) if item_cols else 0.0
    return df[["Month", "Year", "Phase", "SG&A (COP)"]]


def cashflow_monthly_table(s: ScenarioInputs) -> pd.DataFrame:
    """
    Returns a monthly table from project start through end of operation with:
    Revenue, CAPEX, OPEX (incl. GMF), SG&A, Operating CF (pre D&A/tax), Unlevered CF, and cumulative.
    """
    # Use OPEX monthly as the base because it already contains:
    # Month/Year/Phase + Energy + Revenue + CAPEX + OPEX subtotal + GMF
    base = opex_monthly_schedule(s).copy()

    # Merge SG&A monthly total
    sga_m = _sga_monthly_total(s).copy()
    base = base.merge(sga_m[["Month", "SG&A (COP)"]], on="Month", how="left")
    base["SG&A (COP)"] = base["SG&A (COP)"].fillna(0.0)

    # Totals
    base["Total OPEX (COP)"] = base["OPEX subtotal"].fillna(0.0) + base["GMF"].fillna(0.0)

    # Operating CF (pre depreciation, taxes, financing)
    base["Operating CF (COP)"] = (
        base["Revenue (COP)"].fillna(0.0)
        - base["Total OPEX (COP)"].fillna(0.0)
        - base["SG&A (COP)"].fillna(0.0)
    )

    # Unlevered CF (pre financing) for now: Operating CF - CAPEX
    base["Unlevered CF (COP)"] = base["Operating CF (COP)"].fillna(0.0) - base["CAPEX (COP)"].fillna(0.0)
    base["Cumulative Unlevered CF (COP)"] = base["Unlevered CF (COP)"].cumsum()

    # Keep a clean set of columns for display
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
# -----------------------------
# Depreciation
# -----------------------------
def depreciation_annual_table(s: ScenarioInputs) -> pd.DataFrame:
    """
    Linear depreciation starting at COD year.
    Depreciable base = % of total CAPEX.
    """
    tl = build_timeline(s.timeline)
    cod_year = tl["cod"].year

    # Total CAPEX from lines
    cap_df = pd.DataFrame(s.capex.lines or [])
    total_capex = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

    pct = float(s.depreciation.pct_of_capex_depreciated or 0.0) / 100.0
    dep_years = int(s.depreciation.dep_years)

    dep_base = total_capex * pct
    annual_dep = dep_base / dep_years if dep_years > 0 else 0.0

    years = list(range(cod_year, cod_year + dep_years))
    df = pd.DataFrame(
        {
            "Year": years,
            "Depreciable CAPEX (COP)": [dep_base] * dep_years,
            "Depreciation (COP)": [annual_dep] * dep_years,
        }
    )
    df["Cumulative Depreciation (COP)"] = df["Depreciation (COP)"].cumsum()
    df["Remaining Book Value (COP)"] = np.maximum(dep_base - df["Cumulative Depreciation (COP)"], 0.0)
    return df


# -----------------------------
# UI helpers
# -----------------------------
def _metric_row(cols):
    c = st.columns(len(cols))
    for i, (label, val) in enumerate(cols):
        c[i].metric(label, val)


# -----------------------------
# App
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
        new_project = st.text_input("New project name", value="Delphi - Utility Scale")
        create_project = st.button("Create project", use_container_width=True)
        if create_project:
            if new_project.strip() == "":
                st.error("Project name cannot be empty.")
            else:
                db["projects"].setdefault(new_project, {"scenarios": {}})
                if "Base" not in db["projects"][new_project]["scenarios"]:
                    s0 = ScenarioInputs(name="Base")
                    s0.timeline.start_date = _today().isoformat()
                    db["projects"][new_project]["scenarios"]["Base"] = _scenario_to_dict(s0)
                _save_db(db)
                st.success("Project created. Select it from the dropdown.")
        st.stop()

    proj = db["projects"].setdefault(project_name, {"scenarios": {}})
    scenarios = sorted(list(proj.get("scenarios", {}).keys()))
    if not scenarios:
        s0 = ScenarioInputs(name="Base")
        s0.timeline.start_date = _today().isoformat()
        proj["scenarios"]["Base"] = _scenario_to_dict(s0)
        _save_db(db)
        scenarios = ["Base"]

    scenario_name = st.selectbox("Scenario", scenarios, index=0)

    colA, colB = st.columns(2)
    with colA:
        new_scen_name = st.text_input("New scenario name", value="Scenario 2")
    with colB:
        clone_from = st.selectbox("Clone from", scenarios, index=0)

    if st.button("Create / Clone scenario", use_container_width=True):
        nm = new_scen_name.strip()
        if nm == "":
            st.error("Scenario name cannot be empty.")
        elif nm in proj["scenarios"]:
            st.error("Scenario already exists.")
        else:
            base = proj["scenarios"][clone_from]
            proj["scenarios"][nm] = json.loads(json.dumps(base))
            proj["scenarios"][nm]["name"] = nm
            _save_db(db)
            st.success("Scenario created.")
            st.rerun()

    del_col1, del_col2 = st.columns(2)
    with del_col1:
        if st.button("Save scenario", type="primary", use_container_width=True):
            st.success("Saved.")
    with del_col2:
        if st.button("Delete scenario", use_container_width=True, disabled=(len(scenarios) == 1)):
            if len(scenarios) == 1:
                st.warning("Cannot delete the only scenario.")
            else:
                del proj["scenarios"][scenario_name]
                _save_db(db)
                st.success("Scenario deleted.")
                st.rerun()

    st.divider()
    compare_scenarios = st.multiselect("Compare scenarios", scenarios, default=[scenario_name])

# Load scenario
key = f"{project_name}::{scenario_name}"
if "loaded_key" not in st.session_state or st.session_state.loaded_key != key:
    s0 = _scenario_from_dict(proj["scenarios"][scenario_name])
    st.session_state.loaded_key = key
    st.session_state.scenario = s0

s: ScenarioInputs = st.session_state.scenario

# TAB ORDER: Revenues right after Generation (requested)
tab_macro, tab_timeline, tab_gen, tab_rev, tab_capex, tab_opex, tab_sga, tab_cf, tab_dep, tab_compare = st.tabs(
    ["A) Macroeconomic", "B) Timeline", "C) Power Generation", "D) Power Revenues", "E) CAPEX", "F) OPEX", "G) SG&A", "H) Cash Flow", "I) Depreciation", "Compare"]
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

    t1, t2, t3, t4 = st.columns([1.2, 1, 1, 1])
    with t1:
        if not s.timeline.start_date:
            s.timeline.start_date = _today().isoformat()
        s.timeline.start_date = st.date_input("Project start date", value=_parse_date_iso(s.timeline.start_date)).isoformat()
    with t2:
        s.timeline.dev_months = int(st.number_input("Development (months)", value=int(s.timeline.dev_months), step=1, format="%d"))
    with t3:
        s.timeline.capex_months = int(st.number_input("CAPEX / Construction (months)", value=int(s.timeline.capex_months), step=1, format="%d"))
    with t4:
        s.timeline.operation_years = int(st.number_input("Operation (years)", value=int(s.timeline.operation_years), step=1, format="%d"))

    tl = build_timeline(s.timeline)
    _metric_row([("RTB", tl["rtb"].isoformat()), ("COD", tl["cod"].isoformat()), ("End of Operation", tl["end_op"].isoformat())])

    st.markdown("#### Visual guideline")
    gantt = pd.DataFrame(
        [
            {"Stage": "Development", "Start": tl["start"], "Finish": tl["rtb"]},
            {"Stage": "CAPEX / Construction", "Start": tl["rtb"], "Finish": tl["cod"]},
            {"Stage": "Operation", "Start": tl["cod"], "Finish": tl["end_op"]},
        ]
    )
    fig = px.timeline(gantt, x_start="Start", x_end="Finish", y="Stage")
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# C) Generation
# -----------------------------
with tab_gen:
    st.subheader("Power generation")

    g1, g2, g3, g4 = st.columns(4)
    with g1:
        s.generation.mwac = st.number_input("Capacity (MWac)", value=float(s.generation.mwac), step=0.1, format="%.2f")
    with g2:
        s.generation.mwp = st.number_input("Capacity (MWp)", value=float(s.generation.mwp), step=0.1, format="%.2f")
    with g3:
        s.generation.degradation_pct = st.number_input("Annual degradation (%)", value=float(s.generation.degradation_pct), step=0.05, format="%.3f")
    with g4:
        s.generation.production_choice = st.selectbox("Production choice", ["P50", "P75", "P90"], index=["P50", "P75", "P90"].index(s.generation.production_choice))

    p1, p2, p3 = st.columns(3)
    with p1:
        s.generation.p50_mwh_yr = st.number_input("P50 production (MWh/yr)", value=float(s.generation.p50_mwh_yr), step=1000.0, format="%.0f")
        st.caption(f"Formatted: {_fmt_num(s.generation.p50_mwh_yr, 0)}")
    with p2:
        s.generation.p75_mwh_yr = st.number_input("P75 production (MWh/yr)", value=float(s.generation.p75_mwh_yr), step=1000.0, format="%.0f")
        st.caption(f"Formatted: {_fmt_num(s.generation.p75_mwh_yr, 0)}")
    with p3:
        s.generation.p90_mwh_yr = st.number_input("P90 production (MWh/yr)", value=float(s.generation.p90_mwh_yr), step=1000.0, format="%.0f")
        st.caption(f"Formatted: {_fmt_num(s.generation.p90_mwh_yr, 0)}")

    op_years = int(s.timeline.operation_years)
    base = {"P50": s.generation.p50_mwh_yr, "P75": s.generation.p75_mwh_yr, "P90": s.generation.p90_mwh_yr}[s.generation.production_choice]
    degr = float(s.generation.degradation_pct) / 100.0
    years = list(range(1, op_years + 1))
    mwh = [base * ((1.0 - degr) ** (i - 1)) for i in years]
    df_deg = pd.DataFrame({"Operating Year #": years, "Energy (MWh)": mwh})

    st.markdown("#### Degradation curve (selected probability)")
    fig = px.line(df_deg, x="Operating Year #", y="Energy (MWh)")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# D) Revenues (moved right after Generation)
# -----------------------------
with tab_rev:
    st.subheader("Power revenues")

    mode = st.radio(
        "Revenue method",
        ["Standard PPA Parameters", "Manual Annual Series"],
        index=0 if s.revenue_mode == "Standard PPA Parameters" else 1,
        horizontal=True,
    )
    s.revenue_mode = mode

    if mode == "Standard PPA Parameters":
        r = s.revenue1
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            r.ppa_price_cop_per_kwh = st.number_input("PPA price at COD (COP/kWh)", value=float(r.ppa_price_cop_per_kwh), step=1.0, format="%.2f")
        with c2:
            r.ppa_term_years = int(st.number_input("PPA term (years)", value=int(r.ppa_term_years), step=1, format="%d"))
        with c3:
            r.merchant_price_cop_per_kwh = st.number_input("Post-term / merchant (COP/kWh)", value=float(r.merchant_price_cop_per_kwh), step=1.0, format="%.2f")
        with c4:
            r.indexation = st.selectbox("Indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(r.indexation) if r.indexation in INDEX_CHOICES else 0)
    else:
        r = s.revenue2
        c1, c2 = st.columns([1, 1])
        with c1:
            r.indexation = st.selectbox("Indexation", INDEX_CHOICES, index=INDEX_CHOICES.index(r.indexation) if r.indexation in INDEX_CHOICES else 0)
        with c2:
            st.caption("Enter constant COP/kWh for each operating year. Indexation converts to nominal.")

        op_years = int(s.timeline.operation_years)
        base_tbl = pd.DataFrame(
            {
                "Operating Year #": list(range(1, op_years + 1)),
                "Price (constant COP/kWh)": [float(r.prices_constant_cop_per_kwh.get(i, 300.0)) for i in range(1, op_years + 1)],
            }
        )
        edited2 = st.data_editor(base_tbl, use_container_width=True, hide_index=True, num_rows="fixed")
        r.prices_constant_cop_per_kwh = {int(row["Operating Year #"]): float(row["Price (constant COP/kWh)"]) for _, row in edited2.iterrows()}

    df = operating_year_table(s)
    out_currency = st.selectbox("Output currency", ["COP", "USD"], index=0)

    if out_currency == "COP":
        df_show = df[["Year", "Operating Year #", "Energy (MWh)", "Price (COP/kWh)", "Revenue (COP)"]].copy()
        df_disp = df_show.copy()
        df_disp["Energy (MWh)"] = df_disp["Energy (MWh)"].round(0)
        df_disp["Price (COP/kWh)"] = df_disp["Price (COP/kWh)"].round(2)
        df_disp["Revenue (COP)"] = df_disp["Revenue (COP)"].round(0)
        df_disp = _df_format_money(df_disp, ["Energy (MWh)", "Revenue (COP)"], decimals=0)
        st.dataframe(df_disp, use_container_width=True, hide_index=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=df["Year"], y=df["Energy (MWh)"], name="Energy (MWh)"))
        fig.add_trace(go.Scatter(x=df["Year"], y=df["Revenue (COP)"], name="Revenue (COP)", yaxis="y2"))
        fig.update_layout(
            height=360,
            yaxis=dict(title="Energy (MWh)"),
            yaxis2=dict(title="Revenue (COP)", overlaying="y", side="right"),
            margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h"),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        df_show = df[["Year", "Operating Year #", "Energy (MWh)", "FX (COP/USD)", "Revenue (USD)"]].copy()
        df_disp = df_show.copy()
        df_disp["Energy (MWh)"] = df_disp["Energy (MWh)"].round(0)
        df_disp["FX (COP/USD)"] = df_disp["FX (COP/USD)"].round(0)
        df_disp["Revenue (USD)"] = df_disp["Revenue (USD)"].round(0)
        df_disp = _df_format_money(df_disp, ["Energy (MWh)", "FX (COP/USD)", "Revenue (USD)"], decimals=0)
        st.dataframe(df_disp, use_container_width=True, hide_index=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=df["Year"], y=df["Energy (MWh)"], name="Energy (MWh)"))
        fig.add_trace(go.Scatter(x=df["Year"], y=df["Revenue (USD)"], name="Revenue (USD)", yaxis="y2"))
        fig.update_layout(
            height=360,
            yaxis=dict(title="Energy (MWh)"),
            yaxis2=dict(title="Revenue (USD)", overlaying="y", side="right"),
            margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h"),
        )
        st.plotly_chart(fig, use_container_width=True)

    total_rev_cop = float(df["Revenue (COP)"].sum())
    total_rev_usd = float(df["Revenue (USD)"].sum())
    _metric_row(
        [
            ("Total operating revenue (COP)", _fmt_cop(total_rev_cop)),
            ("Total operating revenue (USD)", _fmt_usd(total_rev_usd)),
            ("COD year", str(build_timeline(s.timeline)["cod"].year)),
        ]
    )

# -----------------------------
# E) CAPEX
# -----------------------------
with tab_capex:
    st.subheader("CAPEX (COP) — line items, schedule, and unit economics")

    s.capex.distribution = st.selectbox(
        "Distribution for Development/Construction phases",
        CAPEX_DISTRIBUTIONS,
        index=CAPEX_DISTRIBUTIONS.index(s.capex.distribution) if s.capex.distribution in CAPEX_DISTRIBUTIONS else 0,
    )

    st.markdown("#### CAPEX line items (add rows as needed)")
    capex_df = pd.DataFrame(s.capex.lines or _default_capex_lines())
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

    st.markdown("#### CAPEX schedule (monthly, aligned to timeline)")
    sched = capex_monthly_schedule(s)
    sched_disp = sched.copy()
    sched_disp["CAPEX (COP)"] = sched_disp["CAPEX (COP)"].round(0)
    sched_disp = _df_format_money(sched_disp, ["CAPEX (COP)"], decimals=0)
    st.dataframe(sched_disp[["Month", "Phase", "CAPEX (COP)"]], use_container_width=True, hide_index=True)

    fig = px.bar(sched, x="Month", y="CAPEX (COP)", color="Phase")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Annual CAPEX (calendar years)")
    ann = sched.groupby("Year", as_index=False)["CAPEX (COP)"].sum()
    ann_disp = ann.copy()
    ann_disp["CAPEX (COP)"] = ann_disp["CAPEX (COP)"].round(0)
    ann_disp = _df_format_money(ann_disp, ["CAPEX (COP)"], decimals=0)
    st.dataframe(ann_disp, use_container_width=True, hide_index=True)

# -----------------------------
# F) OPEX
# -----------------------------
with tab_opex:
    st.subheader("OPEX (COP) — operating costs, land lease, taxes & levies")

    st.markdown("### Core operating costs")

    a1, a2, a3, a4 = st.columns(4)
    with a1:
        s.opex.fixed_om_cop_per_mwac_year = st.number_input(
            "Fixed O&M (COP/MWac-year)",
            value=float(s.opex.fixed_om_cop_per_mwac_year),
            step=1_000_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: {_fmt_cop(s.opex.fixed_om_cop_per_mwac_year)} / MWac-year")
    with a2:
        s.opex.fixed_om_indexation = st.selectbox(
            "Fixed O&M indexation",
            INDEX_CHOICES,
            index=INDEX_CHOICES.index(s.opex.fixed_om_indexation) if s.opex.fixed_om_indexation in INDEX_CHOICES else 0,
        )
    with a3:
        s.opex.variable_om_cop_per_mwh = st.number_input(
            "Variable O&M (COP/MWh)",
            value=float(s.opex.variable_om_cop_per_mwh),
            step=1_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: COP {_fmt_num(s.opex.variable_om_cop_per_mwh,0)} / MWh")
    with a4:
        s.opex.grid_fees_cop_per_mwh = st.number_input(
            "Grid fees (COP/MWh)",
            value=float(s.opex.grid_fees_cop_per_mwh),
            step=1_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: COP {_fmt_num(s.opex.grid_fees_cop_per_mwh,0)} / MWh")

    b1, b2 = st.columns(2)
    with b1:
        s.opex.insurance_cop_per_mwac_year = st.number_input(
            "Insurance (COP/MWac-year)",
            value=float(s.opex.insurance_cop_per_mwac_year),
            step=1_000_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: {_fmt_cop(s.opex.insurance_cop_per_mwac_year)} / MWac-year")
    with b2:
        s.opex.insurance_indexation = st.selectbox(
            "Insurance indexation",
            INDEX_CHOICES,
            index=INDEX_CHOICES.index(s.opex.insurance_indexation) if s.opex.insurance_indexation in INDEX_CHOICES else 0,
        )

    st.markdown("### Land lease")
    l1, l2, l3, l4 = st.columns(4)
    with l1:
        s.opex.land_hectares = st.number_input("Hectares leased (Ha)", value=float(s.opex.land_hectares), step=1.0, format="%.2f")
    with l2:
        s.opex.land_indexation = st.selectbox(
            "Land lease indexation",
            INDEX_CHOICES,
            index=INDEX_CHOICES.index(s.opex.land_indexation) if s.opex.land_indexation in INDEX_CHOICES else 0,
        )
    with l3:
        s.opex.land_price_dev_cop_per_ha_year = st.number_input(
            "Lease price Dev (COP/Ha-year)",
            value=float(s.opex.land_price_dev_cop_per_ha_year),
            step=100_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: COP {_fmt_num(s.opex.land_price_dev_cop_per_ha_year,0)} / Ha-year")
    with l4:
        s.opex.land_price_con_cop_per_ha_year = st.number_input(
            "Lease price Constr (COP/Ha-year)",
            value=float(s.opex.land_price_con_cop_per_ha_year),
            step=100_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: COP {_fmt_num(s.opex.land_price_con_cop_per_ha_year,0)} / Ha-year")

    l5, _ = st.columns([1, 3])
    with l5:
        s.opex.land_price_op_cop_per_ha_year = st.number_input(
            "Lease price Oper (COP/Ha-year)",
            value=float(s.opex.land_price_op_cop_per_ha_year),
            step=100_000.0,
            format="%.0f",
        )
        st.caption(f"Formatted: COP {_fmt_num(s.opex.land_price_op_cop_per_ha_year,0)} / Ha-year")

    st.markdown("### Taxes & levies")
    t1, t2 = st.columns(2)
    with t1:
        s.opex.ica_pct_of_revenue = st.number_input(
            "ICA (% of revenue)",
            value=float(s.opex.ica_pct_of_revenue),
            step=0.01,
            format="%.4f",
            help="Regional municipal tax (varies by location). Applied to operating revenues.",
        )
    with t2:
        s.opex.gmf_pct_of_outflows = st.number_input(
            "GMF (% of outgoing cash)",
            value=float(s.opex.gmf_pct_of_outflows),
            step=0.01,
            format="%.4f",
            help="Financial transactions tax on cash outflows (default 0.4%). Applied to (CAPEX + OPEX) outflows in this V1.",
        )

    st.markdown("### Other OPEX items (add rows)")
    other_df = pd.DataFrame(s.opex.other_items or _default_opex_other_items())
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in other_df.columns:
            other_df[col] = "" if col not in ["Amount_COP_per_year"] else 0.0
    other_df = other_df[["Item", "Amount_COP_per_year", "Phase", "Indexation"]].copy()
    other_df["Phase"] = other_df["Phase"].where(other_df["Phase"].isin(PHASES), "Operation")
    other_df["Indexation"] = other_df["Indexation"].where(other_df["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    other_edited = st.data_editor(
        other_df,
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
    s.opex.other_items = other_edited.to_dict(orient="records")

    st.divider()
    st.markdown("## Outputs")

    annual_items, op_metrics = opex_annual_by_item(s)

    # Annual stacked chart
    st.markdown("### Annual OPEX by item (calendar years)")
    item_cols = [c for c in annual_items.columns if c not in ["Year"]]
    annual_long = annual_items.melt(id_vars=["Year"], value_vars=[c for c in item_cols if c != "Total OPEX (COP)"], var_name="Item", value_name="OPEX (COP)")
    fig = px.bar(annual_long, x="Year", y="OPEX (COP)", color="Item", barmode="stack")
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    # Annual table with Total OPEX (formatted with separators)
    annual_disp = annual_items.copy()
    money_cols = [c for c in annual_disp.columns if c != "Year"]
    annual_disp = _df_format_money(annual_disp, money_cols, decimals=0)
    st.dataframe(annual_disp, use_container_width=True, hide_index=True)

    st.markdown("### OPEX per MWh (Operating years only)")
    fig2 = px.line(op_metrics, x="Year", y="OPEX per MWh (COP/MWh)")
    fig2.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig2, use_container_width=True)

    show_cols = ["Year", "Energy (MWh)", "Total OPEX (COP)", "OPEX per MWh (COP/MWh)"]
    op_disp = op_metrics.copy()
    op_disp = _df_format_money(op_disp, ["Energy (MWh)", "Total OPEX (COP)", "OPEX per MWh (COP/MWh)"], decimals=0)
    st.dataframe(op_disp[show_cols], use_container_width=True, hide_index=True)

# -----------------------------
# SGA
# -----------------------------
with tab_sga:
    st.subheader("SG&A (COP) — Development, Construction, and Operation")

    st.markdown("### SG&A items (add rows as needed)")
    sga_df = pd.DataFrame(s.sga.items or _default_sga_items())
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

    # Chart (stacked by item; excludes Total column)
    item_cols = [c for c in annual_sga.columns if c not in ["Year", "Total SG&A (COP)"]]
    if item_cols:
        annual_long = annual_sga.melt(id_vars=["Year"], value_vars=item_cols, var_name="Item", value_name="SG&A (COP)")
        fig = px.bar(annual_long, x="Year", y="SG&A (COP)", color="Item", barmode="stack")
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Add SG&A line items to see the chart.")

    # Table with separators
    annual_disp = annual_sga.copy()
    money_cols = [c for c in annual_disp.columns if c != "Year"]
    annual_disp = _df_format_money(annual_disp, money_cols, decimals=0)
    st.dataframe(annual_disp, use_container_width=True, hide_index=True)

# -----------------------------
# H) Cash Flow
# -----------------------------
with tab_cf:
    st.subheader("Operating Cash Flow (pre depreciation, taxes, financing)")

    m = cashflow_monthly_table(s)
    a = cashflow_annual_table(s)

    # Key metrics (lifetime totals)
    total_rev = float(m["Revenue (COP)"].sum())
    total_opcf = float(m["Operating CF (COP)"].sum())
    total_capex = float(m["CAPEX (COP)"].sum())
    end_cum = float(m["Cumulative Unlevered CF (COP)"].iloc[-1]) if len(m) else 0.0

    _metric_row(
        [
            ("Total Revenue (COP)", _fmt_cop(total_rev)),
            ("Total Operating CF (COP)", _fmt_cop(total_opcf)),
            ("Total CAPEX (COP)", _fmt_cop(total_capex)),
            ("End Cumulative Unlevered CF (COP)", _fmt_cop(end_cum)),
        ]
    )

    st.divider()

    # Charts
    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = px.line(m, x="Month", y="Cumulative Unlevered CF (COP)")
        fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig2 = px.bar(a, x="Year", y="Operating CF (COP)")
        fig2.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### Annual summary (calendar years)")
    a_disp = a.copy()
    money_cols = [c for c in a_disp.columns if c not in ["Year"]]
    a_disp = _df_format_money(a_disp, money_cols, decimals=0)
    st.dataframe(a_disp, use_container_width=True, hide_index=True)

    st.markdown("### Monthly cash flow (detailed)")
    m_disp = m.copy()
    money_cols_m = [c for c in m_disp.columns if c not in ["Month", "Year", "Phase"]]
    m_disp = _df_format_money(m_disp, money_cols_m, decimals=0)
    st.dataframe(m_disp, use_container_width=True, hide_index=True)

# -----------------------------
# Depreciation
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

    # Graph: annual depreciation
    fig = px.bar(dep, x="Year", y="Depreciation (COP)")
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    # Table with separators
    dep_disp = dep.copy()
    money_cols = [c for c in dep_disp.columns if c != "Year"]
    dep_disp = _df_format_money(dep_disp, money_cols, decimals=0)
    st.dataframe(dep_disp, use_container_width=True, hide_index=True)


# -----------------------------
# Compare
# -----------------------------
with tab_compare:
    st.subheader("Scenario comparison (CAPEX + OPEX + Revenues)")

    if not compare_scenarios:
        st.warning("Select scenarios to compare in the sidebar.")
    else:
        rows = []
        for nm in compare_scenarios:
            sd = _scenario_from_dict(proj["scenarios"][nm])

            rev = operating_year_table(sd)
            cap_df = pd.DataFrame(sd.capex.lines or [])
            total_cap = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

            annual_items_x, opm = opex_annual_by_item(sd)
            total_opex_op = float(opm["Total OPEX (COP)"].sum())

            mwac_x = float(sd.generation.mwac or 0.0)
            mwp_x = float(sd.generation.mwp or 0.0)

            rows.append(
                {
                    "Scenario": nm,
                    "P-Choice": sd.generation.production_choice,
                    "COD": build_timeline(sd.timeline)["cod"].isoformat(),
                    "Total CAPEX (COP)": total_cap,
                    "CAPEX/MWac (COP)": (total_cap / mwac_x) if mwac_x > 0 else np.nan,
                    "CAPEX/MWp (COP)": (total_cap / mwp_x) if mwp_x > 0 else np.nan,
                    "Total OPEX (COP, operating years)": total_opex_op,
                    "Total Revenue (COP)": float(rev["Revenue (COP)"].sum()),
                }
            )

        summary = pd.DataFrame(rows).sort_values("Scenario")

        # Format summary for readability
        sum_disp = summary.copy()
        for col in ["Total CAPEX (COP)", "CAPEX/MWac (COP)", "CAPEX/MWp (COP)", "Total OPEX (COP, operating years)", "Total Revenue (COP)"]:
            if col in sum_disp.columns:
                sum_disp[col] = sum_disp[col].apply(lambda v: _fmt_num(float(v), 0) if pd.notnull(v) else "")
        st.dataframe(sum_disp, use_container_width=True, hide_index=True)

# Persist scenario on each run
proj["scenarios"][scenario_name] = _scenario_to_dict(s)
_save_db(db)
