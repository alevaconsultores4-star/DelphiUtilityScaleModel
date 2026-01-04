# Delphi Utility-Scale Financial Model (No Excel)
# Streamlit single-file app with: Projects + Scenarios, Macro, Timeline, Generation, Revenues,
# CAPEX, OPEX, SG&A, Depreciation, Debt & Covenants, Unlevered Base Cash Flow, Compare
# All inputs in COP; outputs selectable COP/USD (USD via FX path).

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict
from datetime import date
from typing import Dict, List

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
    pct_of_capex_depreciated: float = 100.0
    dep_years: int = 20  # 3–25


@dataclass
class TaxInputs:
    corporate_tax_rate_pct: float = 35.0
    allow_loss_carryforward: bool = True

@dataclass
class RenewableIncentivesInputs:
    # Special deduction (Ley 1715 / 2099 style): up to 50% of eligible investment
    # usable over up to 15 years, but each year capped at 50% of taxable income.
    enable_special_deduction: bool = True
    special_deduction_pct_of_capex: float = 50.0          # 0–50 (%)
    special_deduction_years: int = 15                     # typically 15
    special_deduction_max_pct_of_taxable_income: float = 50.0  # cap per year

    # VAT: model as either excluded (default) or refunded as a one-time cash inflow
    vat_mode: str = "Excluded"                            # "Excluded" or "Refund"
    vat_pct_of_capex: float = 0.0                         # if you prefer % input
    vat_fixed_cop: float = 0.0                            # or fixed COP input
    vat_refund_year_index: int = 1                        # refund in op Year 1 by default

@dataclass
class WorkingCapitalInputs:
    ar_days: int = 90   # revenue collection lag
    ap_days: int = 60   # expense payment lag
    apply_ap_to_opex: bool = True
    apply_ap_to_sga: bool = True

@dataclass
class DebtInputs:
    enabled: bool = False
    debt_pct_of_capex: float = 70.0  # %
    tenor_years: int = 7            # allow 5–10
    grace_years: int = 0            # 0–(tenor-1)

    # Pricing (Natural COP)
    base_rate_pct: float = 9.0      # e.g., IBR
    margin_pct: float = 7.0         # spread
    # Fees
    upfront_fee_bps: float = 175.0  # bps of total debt, paid at COD/first draw
    commitment_fee_pct_of_margin: float = 30.0  # % of margin, on undrawn during construction

    # Covenants (for indicators)
    target_dscr: float = 1.20
    min_dscr_covenant: float = 1.20
    lockup_dscr: float = 1.15

    amortization_type: str = "Sculpted to DSCR"  # or "Equal principal"
    balloon_pct: float = 0.0   # NEW: % of original debt left as balloon (0 = fully amortizing)

@dataclass
class RenewableTaxInputs:
    enable_special_deduction: bool = True
    special_deduction_pct_of_capex: float = 50.0   # % of CAPEX
    special_deduction_years: int = 15               # max carryforward

    enable_vat_refund: bool = True
    vat_refund_mode: str = "percent"                # "percent" or "fixed"
    vat_pct_of_capex: float = 19.0                  # if percent
    vat_fixed_cop: float = 0.0                      # if fixed
    vat_refund_year: int = 1                         # years after COD


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
    wc: WorkingCapitalInputs = field(default_factory=WorkingCapitalInputs)

    debt: DebtInputs = field(default_factory=DebtInputs)
    renewable_tax: RenewableTaxInputs = field(default_factory=RenewableTaxInputs)
    incentives: RenewableIncentivesInputs = field(default_factory=RenewableIncentivesInputs)



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
    wc = WorkingCapitalInputs(**d.get("wc", {})) if "wc" in d else WorkingCapitalInputs()
    debt = DebtInputs(**d.get("debt", {})) if "debt" in d else DebtInputs()
    incentives = RenewableIncentivesInputs(**d.get("incentives", {}))

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
        wc=wc,
        debt=debt,
        incentives=incentives,
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
    end_op = tl["end_op"]
    op_years = int(s.timeline.operation_years)

    # Get all calendar years from COD to end of operation
    years = list(range(cod.year, end_op.year + 1))
    base_year = cod.year

    gen = s.generation
    p_map = {"P50": gen.p50_mwh_yr, "P75": gen.p75_mwh_yr, "P90": gen.p90_mwh_yr}
    base_mwh = float(p_map.get(gen.production_choice, gen.p50_mwh_yr))
    degr = float(gen.degradation_pct) / 100.0
    
    # Calculate prorated generation for each calendar year
    mwh_list = []
    price_base_list = []
    
    for year in years:
        # Determine start and end dates for this calendar year
        if year == cod.year:
            # First year: from COD date to end of year
            year_start = cod
            year_end = date(year, 12, 31)
        elif year == end_op.year:
            # Last year: from start of year to end_op date
            year_start = date(year, 1, 1)
            year_end = end_op
        else:
            # Full year
            year_start = date(year, 1, 1)
            year_end = date(year, 12, 31)
        
        # Calculate months of operation in this calendar year
        months_in_year = (year_end.year - year_start.year) * 12 + (year_end.month - year_start.month) + 1
        fraction_of_year = months_in_year / 12.0
        
        # Calculate which operating year the middle of this calendar year falls into
        # Use the midpoint of the calendar year to determine operating year
        if year == cod.year:
            mid_month = (cod.month + 12) / 2.0
        elif year == end_op.year:
            mid_month = (1 + end_op.month) / 2.0
        else:
            mid_month = 6.5  # Middle of year
        
        # Calculate months from COD to midpoint of this calendar year
        months_from_cod = (year - cod.year) * 12 + (mid_month - cod.month)
        operating_year_num = int(months_from_cod // 12)  # 0-indexed operating year
        
        # Calculate annual generation for this operating year (with degradation)
        annual_mwh = base_mwh * ((1.0 - degr) ** operating_year_num)
        
        # Prorate by fraction of year
        prorated_mwh = annual_mwh * fraction_of_year
        
        # Calculate price based on operating year
        if s.revenue_mode == "Standard PPA Parameters":
            r = s.revenue1
            term = int(r.ppa_term_years)
            p0 = float(r.ppa_price_cop_per_kwh)
            pm = float(r.merchant_price_cop_per_kwh)
            price = p0 if (operating_year_num < term) else pm
        else:
            r = s.revenue2
            price = float(r.prices_constant_cop_per_kwh.get(operating_year_num + 1, 0.0))
        
        mwh_list.append(prorated_mwh)
        price_base_list.append(price)
    
    # Apply indexation
    index_choice = s.revenue1.indexation if s.revenue_mode == "Standard PPA Parameters" else s.revenue2.indexation
    idx = annual_index_series(s.macro, base_year, years, _idx_key(index_choice))
    price_indexed = [price_base_list[i] * float(idx.loc[years[i]]) for i in range(len(years))]

    df = pd.DataFrame({"Year": years, "Energy (MWh)": mwh_list})
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

    months = []
    cur = date(start.year, start.month, 1)
    endm = date(end_op.year, end_op.month, 1)
    while cur < endm:
        months.append(cur)
        cur = _add_months(cur, 1)

    df = pd.DataFrame({"Month": months})
    df["Year"] = df["Month"].apply(lambda d: d.year)
    df["Phase"] = df["Month"].apply(lambda m: _phase_for_month(tl, m))

    op = operating_year_table(s)
    op_map_mwh = {int(r["Year"]): float(r["Energy (MWh)"]) for _, r in op.iterrows()}
    op_map_rev = {int(r["Year"]): float(r["Revenue (COP)"]) for _, r in op.iterrows()}
    
    # Calculate operating months per calendar year for proper proration
    cod = tl["cod"]
    operating_months_per_year = {}
    for year in op["Year"].unique():
        year_int = int(year)
        if year_int == cod.year:
            # First year: from COD month to end of year
            operating_months_per_year[year_int] = 13 - cod.month
        elif year_int == end_op.year:
            # Last year: from start of year to end_op month
            operating_months_per_year[year_int] = end_op.month
        else:
            # Full year
            operating_months_per_year[year_int] = 12

    def get_monthly_value(row, annual_map, months_map):
        if row["Phase"] != "Operation":
            return 0.0
        year = int(row["Year"])
        annual_val = annual_map.get(year, 0.0)
        months = months_map.get(year, 12)
        return annual_val / months if months > 0 else 0.0

    df["Energy (MWh)"] = df.apply(lambda r: get_monthly_value(r, op_map_mwh, operating_months_per_year), axis=1)
    df["Revenue (COP)"] = df.apply(lambda r: get_monthly_value(r, op_map_rev, operating_months_per_year), axis=1)

    cap = capex_monthly_schedule(s)[["Month", "CAPEX (COP)"]].copy()
    df = df.merge(cap, on="Month", how="left")
    df["CAPEX (COP)"] = df["CAPEX (COP)"].fillna(0.0)

    mwac = float(s.generation.mwac or 0.0)

    base_year = tl["cod"].year
    idx_fixed = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.fixed_om_indexation))
    idx_ins = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.insurance_indexation))
    idx_land = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, s.opex.land_indexation))

    df["Fixed O&M"] = 0.0
    df.loc[df["Phase"] == "Operation", "Fixed O&M"] = (float(s.opex.fixed_om_cop_per_mwac_year) * mwac / 12.0) * idx_fixed[df["Phase"] == "Operation"].values

    df["Insurance"] = 0.0
    df.loc[df["Phase"] == "Operation", "Insurance"] = (float(s.opex.insurance_cop_per_mwac_year) * mwac / 12.0) * idx_ins[df["Phase"] == "Operation"].values

    df["Variable O&M"] = float(s.opex.variable_om_cop_per_mwh) * df["Energy (MWh)"]
    df["Grid fees"] = float(s.opex.grid_fees_cop_per_mwh) * df["Energy (MWh)"]

    ha = float(s.opex.land_hectares or 0.0)
    df["Land lease"] = 0.0
    df.loc[df["Phase"] == "Development", "Land lease"] = (ha * float(s.opex.land_price_dev_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Development"].values
    df.loc[df["Phase"] == "Construction", "Land lease"] = (ha * float(s.opex.land_price_con_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Construction"].values
    df.loc[df["Phase"] == "Operation", "Land lease"] = (ha * float(s.opex.land_price_op_cop_per_ha_year) / 12.0) * idx_land[df["Phase"] == "Operation"].values

    other = pd.DataFrame(s.opex.other_items or [])
    for col in ["Item", "Amount_COP_per_year", "Phase", "Indexation"]:
        if col not in other.columns:
            other[col] = "" if col != "Amount_COP_per_year" else 0.0
    other["Phase"] = other["Phase"].where(other["Phase"].isin(PHASES), "Operation")
    other["Indexation"] = other["Indexation"].where(other["Indexation"].isin(INDEX_CHOICES), "Colombia CPI")

    for _, r in other.iterrows():
        name = str(r.get("Item", "")).strip() or "Other"
        amt = float(r.get("Amount_COP_per_year", 0.0) or 0.0)
        ph = str(r.get("Phase", "Operation"))
        idx_choice = str(r.get("Indexation", "Colombia CPI"))
        if name not in df.columns:
            df[name] = 0.0
        idx_series = df["Year"].apply(lambda y: index_factor_for_year(s.macro, int(y), base_year, idx_choice))
        df.loc[df["Phase"] == ph, name] = (amt / 12.0) * idx_series[df["Phase"] == ph].values

    df["ICA"] = 0.0
    df.loc[df["Phase"] == "Operation", "ICA"] = (float(s.opex.ica_pct_of_revenue) / 100.0) * df.loc[df["Phase"] == "Operation", "Revenue (COP)"]

    meta = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)"}
    fixed_cols = ["Fixed O&M", "Insurance", "Variable O&M", "Grid fees", "Land lease", "ICA"]
    dyn_cols = [c for c in df.columns if c not in meta and c not in fixed_cols and c not in {"OPEX subtotal", "GMF"}]
    opex_cols = fixed_cols + dyn_cols
    df["OPEX subtotal"] = df[opex_cols].sum(axis=1) if opex_cols else 0.0

    # GMF on outgoing cash: CAPEX + OPEX subtotal
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
    cod = tl["cod"]
    cod_year = cod.year

    cap_df = pd.DataFrame(s.capex.lines or [])
    total_capex = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

    pct = float(s.depreciation.pct_of_capex_depreciated or 0.0) / 100.0
    dep_years = int(s.depreciation.dep_years)
    dep_months = dep_years * 12  # Total months of depreciation

    dep_base = total_capex * pct
    monthly_dep = dep_base / dep_months if dep_months > 0 else 0.0
    annual_dep = monthly_dep * 12.0  # Full year equivalent

    # Calculate end date of depreciation period
    dep_end = _add_months(cod, dep_months)
    
    # Get all calendar years from COD to end of depreciation
    years = list(range(cod_year, dep_end.year + 1))
    
    # Calculate prorated depreciation for each calendar year
    dep_list = []
    cumulative_dep = 0.0
    
    for year in years:
        # Determine start and end dates for this calendar year
        if year == cod_year:
            # First year: from COD date to end of year
            year_start = cod
            year_end = date(year, 12, 31)
        elif year == dep_end.year:
            # Last year: from start of year to dep_end date
            year_start = date(year, 1, 1)
            year_end = dep_end
        else:
            # Full year
            year_start = date(year, 1, 1)
            year_end = date(year, 12, 31)
        
        # Calculate months of depreciation in this calendar year
        months_in_year = (year_end.year - year_start.year) * 12 + (year_end.month - year_start.month) + 1
        
        # Calculate depreciation for this year
        year_dep = monthly_dep * months_in_year
        dep_list.append(year_dep)
        cumulative_dep += year_dep
    
    df = pd.DataFrame({
        "Year": years,
        "Depreciable CAPEX (COP)": [dep_base] * len(years),
        "Depreciation (COP)": dep_list
    })
    df["Cumulative Depreciation (COP)"] = df["Depreciation (COP)"].cumsum()
    df["Remaining Book Value (COP)"] = np.maximum(dep_base - df["Cumulative Depreciation (COP)"], 0.0)
    return df

# -----------------------------
# CAPEX deduction for renewable incentives
# -----------------------------

def renewable_tax_annual_table(s: ScenarioInputs) -> pd.DataFrame:
    tl = build_timeline(s.timeline)
    cod_year = tl["cod"].year

    cf = cashflow_annual_table(s).copy()

    # Base taxable income BEFORE incentives
    cf["Taxable Income (Pre-Incentives)"] = (
        cf["Operating CF (COP)"]
        - cf.get("Depreciation (COP)", 0.0)
    )

    # -------------------------
    # Special CAPEX deduction
    # -------------------------
    eligible_capex = _eligible_capex_for_tax(s)
    deduction_pool = eligible_capex
    max_years = int(s.renewable_tax.special_deduction_years)

    deductions = []
    remaining = deduction_pool

    for _, r in cf.iterrows():
        year = int(r["Year"])
        ti = max(0.0, float(r["Taxable Income (Pre-Incentives)"]))

        if (
            not s.renewable_tax.enable_special_deduction
            or remaining <= 0
            or year < cod_year
            or year >= cod_year + max_years
        ):
            used = 0.0
        else:
            used = min(remaining, 0.5 * ti)

        remaining -= used
        deductions.append(used)

    cf["Special Deduction Used (COP)"] = deductions
    cf["Remaining Deduction Balance (COP)"] = (
        deduction_pool - pd.Series(deductions).cumsum()
    ).clip(lower=0.0)

    # -------------------------
    # Final taxable income
    # -------------------------
    cf["Taxable Income (After Incentives)"] = (
        cf["Taxable Income (Pre-Incentives)"]
        - cf["Special Deduction Used (COP)"]
    ).clip(lower=0.0)

    tax_rate = float(s.tax.corporate_tax_rate_pct) / 100.0
    cf["Income Tax (COP)"] = cf["Taxable Income (After Incentives)"] * tax_rate

    # -------------------------
    # VAT refund (cash inflow)
    # -------------------------
    cf["VAT Refund (COP)"] = 0.0

    if s.renewable_tax.enable_vat_refund:
        if s.renewable_tax.vat_refund_mode == "percent":
            vat_amt = (
                _total_capex_from_lines(s)
                * float(s.renewable_tax.vat_pct_of_capex)
                / 100.0
            )
        else:
            vat_amt = float(s.renewable_tax.vat_fixed_cop)

        refund_year = cod_year + int(s.renewable_tax.vat_refund_year)
        cf.loc[cf["Year"] == refund_year, "VAT Refund (COP)"] = vat_amt

    return cf


# -----------------------------
# CASH FLOW (monthly + annual)
# -----------------------------
def _sga_monthly_total(s: ScenarioInputs) -> pd.DataFrame:
    df = sga_monthly_schedule(s).copy()
    meta = {"Month", "Year", "Phase"}
    item_cols = [c for c in df.columns if c not in meta]
    df["SG&A (COP)"] = df[item_cols].sum(axis=1) if item_cols else 0.0
    return df[["Month", "Year", "Phase", "SG&A (COP)"]]

def _shift_by_months(values: pd.Series, months: int) -> pd.Series:
    """Shift a monthly series forward by N months (cash received/paid later).
    Missing months are filled with 0.0."""
    months = int(max(0, months))
    if months == 0:
        return values.astype(float)
    return values.shift(months, fill_value=0.0).astype(float)

def cashflow_monthly_table(s: ScenarioInputs) -> pd.DataFrame:
    base = opex_monthly_schedule(s).copy()

    sga_m = _sga_monthly_total(s).copy()
    base = base.merge(sga_m[["Month", "SG&A (COP)"]], on="Month", how="left")
    base["SG&A (COP)"] = base["SG&A (COP)"].fillna(0.0)

    base["Total OPEX (COP)"] = base["OPEX subtotal"].fillna(0.0) + base["GMF"].fillna(0.0)

    # --- Accrual (non-cash timing) operating CF + unlevered CF (this is what debt sizing will use initially) ---
    base["Operating CF (COP)"] = (
        base["Revenue (COP)"].fillna(0.0)
        - base["Total OPEX (COP)"].fillna(0.0)
        - base["SG&A (COP)"].fillna(0.0)
    )

    base["Unlevered CF (COP)"] = base["Operating CF (COP)"] - base["CAPEX (COP)"].fillna(0.0)
    base["Cumulative Unlevered CF (COP)"] = base["Unlevered CF (COP)"].cumsum()


    # -----------------------------
    # Working Capital (simple AR/AP lags) + Cash CF
    # -----------------------------
    # You said you use s.wc (not s.working_capital)
    ar_days = float(getattr(s.wc, "ar_days", 90.0))   # revenue collected in ~90 days
    ap_days = float(getattr(s.wc, "ap_days", 60.0))   # expenses paid in ~60 days

    # Convert days to whole months (simple approximation)
    ar_lag_m = int(round(ar_days / 30.0))
    ap_lag_m = int(round(ap_days / 30.0))

    # What "cash paid expenses" means here: OPEX + SG&A (exclude CAPEX)
    base["Total OpEx+SGA (COP)"] = base["Total OPEX (COP)"].fillna(0.0) + base["SG&A (COP)"].fillna(0.0)

    # Cash timing via shifts (simple, fast, stable)
    base["Cash Collected (COP)"]   = base["Revenue (COP)"].fillna(0.0).shift(ar_lag_m, fill_value=0.0)
    base["Cash Paid OPEX (COP)"]   = base["Total OPEX (COP)"].fillna(0.0).shift(ap_lag_m, fill_value=0.0)
    base["Cash Paid SG&A (COP)"]   = base["SG&A (COP)"].fillna(0.0).shift(ap_lag_m, fill_value=0.0)

    # AR/AP balances implied by the lag model
    # AR(t) = AR(t-1) + Revenue(t) - CashCollected(t)
    # AP(t) = AP(t-1) + Expenses(t) - CashPaid(t)
    ar = []
    ap = []
    ar_prev = 0.0
    ap_prev = 0.0

    rev = base["Revenue (COP)"].fillna(0.0).to_numpy()
    cash_in = base["Cash Collected (COP)"].fillna(0.0).to_numpy()

    opex = base["Total OPEX (COP)"].fillna(0.0).to_numpy()
    sga  = base["SG&A (COP)"].fillna(0.0).to_numpy()
    cash_opex = base["Cash Paid OPEX (COP)"].fillna(0.0).to_numpy()
    cash_sga  = base["Cash Paid SG&A (COP)"].fillna(0.0).to_numpy()

    for i in range(len(base)):
        ar_now = ar_prev + rev[i] - cash_in[i]
        ap_now = ap_prev + (opex[i] + sga[i]) - (cash_opex[i] + cash_sga[i])
        ar.append(ar_now)
        ap.append(ap_now)
        ar_prev, ap_prev = ar_now, ap_now

    base["AR Balance (COP)"] = ar
    base["AP Balance (COP)"] = ap

    base["Net Working Capital (COP)"] = base["AR Balance (COP)"] - base["AP Balance (COP)"]
    base["ΔNWC (COP)"] = base["Net Working Capital (COP)"].diff().fillna(base["Net Working Capital (COP)"])

    # Cash operating CF and cash unlevered CF
    base["Operating CF (Cash, COP)"] = (
        base["Cash Collected (COP)"].fillna(0.0)
        - base["Cash Paid OPEX (COP)"].fillna(0.0)
        - base["Cash Paid SG&A (COP)"].fillna(0.0)
    )

    # CAPEX assumed paid when incurred (no lag)
    base["Unlevered CF (Cash, COP)"] = base["Operating CF (Cash, COP)"] - base["CAPEX (COP)"].fillna(0.0)
    base["Cumulative Unlevered CF (Cash, COP)"] = base["Unlevered CF (Cash, COP)"].cumsum()

    # -----------------------------
    # Debt fees (cash) — only when there is debt
    # -----------------------------
    base["Debt Fees (COP)"] = 0.0

    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * _total_capex_from_lines(s)
    debt_amt = max(debt_amt, 0.0)

    if debt_amt > 0:
        # 1) Commitment fee (annual -> spread monthly)
        cf_ann = debt_commitment_fee_annual(s).copy()  # expects columns: Year, Commitment Fee (COP)
        if (not cf_ann.empty) and ("Year" in cf_ann.columns) and ("Commitment Fee (COP)" in cf_ann.columns):
            fee_map = {int(r["Year"]): float(r["Commitment Fee (COP)"]) for _, r in cf_ann.iterrows()}
            base["Debt Fees (COP)"] += base["Year"].map(lambda y: fee_map.get(int(y), 0.0) / 12.0)

        # 2) Upfront fee at COD month (if you have upfront_fee_pct in s.debt)
        tl = build_timeline(s.timeline)
        cod_m = date(tl["cod"].year, tl["cod"].month, 1)
        upfront_pct = float(getattr(s.debt, "upfront_fee_pct", 0.0)) / 100.0
        if upfront_pct > 0:
            base.loc[base["Month"] == cod_m, "Debt Fees (COP)"] += debt_amt * upfront_pct

    # Apply fees to CASH unlevered CF (equity cash)
    base["Unlevered CF (Cash, COP)"] = base["Unlevered CF (Cash, COP)"].fillna(0.0) - base["Debt Fees (COP)"].fillna(0.0)
    base["Cumulative Unlevered CF (Cash, COP)"] = base["Unlevered CF (Cash, COP)"].cumsum()


    cols = [
        "Month", "Year", "Phase",
        "Energy (MWh)", "Revenue (COP)", "Cash Collected (COP)",
        "CAPEX (COP)", "Debt Fees (COP)",
        "Total OPEX (COP)", "Cash Paid OPEX (COP)",
        "SG&A (COP)", "Cash Paid SG&A (COP)",
        "Operating CF (COP)", "Operating CF (Cash, COP)",
        "AR Balance (COP)", "AP Balance (COP)", "Net Working Capital (COP)", "ΔNWC (COP)",
        "Unlevered CF (COP)", "Unlevered CF (Cash, COP)",
        "Cumulative Unlevered CF (COP)", "Cumulative Unlevered CF (Cash, COP)",
    ]

    # IMPORTANT: prevents KeyError if some columns are not created yet
    cols = [c for c in cols if c in base.columns]

    return base.loc[:, cols].copy()


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
    a = cashflow_annual_table(s).copy()
    dep = depreciation_annual_table(s)[["Year", "Depreciation (COP)"]].copy()
    out = a.merge(dep, on="Year", how="left").fillna({"Depreciation (COP)": 0.0})

    out["EBITDA (COP)"] = out["Operating CF (COP)"]
    out["Taxable Income (COP)"] = out["EBITDA (COP)"] - out["Depreciation (COP)"]

    rate = float(s.tax.corporate_tax_rate_pct) / 100.0
    allow_nol = bool(s.tax.allow_loss_carryforward)

    # ---- Renewable incentive (Law 1715 / 2099): special deduction ----
    # Pool = 50% of total investment (CAPEX). Usable up to 15 years starting year after COD.
    # Annual cap: deduction <= 50% of renta líquida (before this deduction).

    tl = build_timeline(s.timeline)
    cod_year = int(tl["cod"].year)

    total_capex = float(_total_capex_from_lines(s))  # you already have this helper for debt
    special_pool_total = 0.50 * total_capex          # 50% of total investment
    special_pool_remaining = special_pool_total

    special_years = 15
    special_start_year = cod_year + 1
    special_end_year = special_start_year + special_years - 1

    # ---- Incentives + tax calc ----
    inc = getattr(s, "incentives", RenewableIncentivesInputs())

    # Total eligible CAPEX (same base you already use for depreciation)
    cap_df = pd.DataFrame(s.capex.lines or [])
    total_capex = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0

    # Special deduction pool
    pool_total = (float(inc.special_deduction_pct_of_capex) / 100.0) * total_capex if bool(inc.enable_special_deduction) else 0.0
    pool_remaining = pool_total

    # VAT refund (cash only)
    vat_amount = (float(inc.vat_pct_of_capex) / 100.0) * total_capex + float(inc.vat_fixed_cop)
    vat_amount = max(0.0, vat_amount) if inc.vat_mode == "Refund" else 0.0

    nol = 0.0
    taxes = []
    nol_end = []

    ded_used_list = []
    ded_rem_list = []
    taxable_after_all_list = []
    vat_refund_list = []

    # Operating-year index so "refund in op Year #1" maps correctly
    # Your annual table starts at COD year; that is operating year 1.
    op_year_counter = 0
    max_years = int(max(1, inc.special_deduction_years))

    for _, r in out.iterrows():
        op_year_counter += 1

        ti = float(r["Taxable Income (COP)"])

        # Apply NOL first (your existing logic)
        if allow_nol:
            if ti < 0:
                nol = nol + (-ti)
                taxable_after_nol = 0.0
            else:
                used = min(nol, ti)
                nol -= used
                taxable_after_nol = max(ti - used, 0.0)
        else:
            taxable_after_nol = max(ti, 0.0)

        # Special deduction: only within window, only if taxable income exists
        ded_used = 0.0
        if bool(inc.enable_special_deduction) and op_year_counter <= max_years and taxable_after_nol > 0 and pool_remaining > 0:
            annual_cap = (float(inc.special_deduction_max_pct_of_taxable_income) / 100.0) * taxable_after_nol
            ded_used = min(pool_remaining, annual_cap)
            pool_remaining -= ded_used

        taxable_after_all = max(taxable_after_nol - ded_used, 0.0)

        taxes_payable = taxable_after_all * rate
        taxes.append(taxes_payable)
        nol_end.append(nol)

        ded_used_list.append(ded_used)
        ded_rem_list.append(pool_remaining)
        taxable_after_all_list.append(taxable_after_all)

        # VAT refund: one-time cash inflow in operating year N (default year 1)
        vat_refund = vat_amount if (vat_amount > 0 and op_year_counter == int(inc.vat_refund_year_index)) else 0.0
        vat_refund_list.append(vat_refund)

    out["Loss Carryforward End (COP)"] = nol_end
    out["Special Deduction Used (COP)"] = ded_used_list
    out["Special Deduction Remaining (COP)"] = ded_rem_list
    out["Taxable Income After Incentives (COP)"] = taxable_after_all_list

    out["Taxes Payable (COP)"] = taxes

    # After-tax CF: taxes reduce cash; VAT refund adds cash (does not affect taxes)
    out["VAT Refund (COP)"] = vat_refund_list
    out["Unlevered CF After Tax (COP)"] = out["Unlevered CF (COP)"] - out["Taxes Payable (COP)"] + out["VAT Refund (COP)"]

    return out


# -----------------------------
# IRR / Payback helpers
# -----------------------------
def _irr_bisection(cashflows: List[float], low=-0.95, high=5.0, tol=1e-7, max_iter=200) -> float:
    def npv(rate: float) -> float:
        denom_base = 1.0 + rate
        if denom_base <= 1e-12:
            return float("inf")
        total = 0.0
        for i, cf in enumerate(cashflows):
            if i == 0:
                total += cf
                continue
            den = denom_base ** i
            if den == 0.0:
                return float("inf")
            total += cf / den
        return total

    if len(cashflows) < 2:
        return float("nan")

    f_low = npv(low)
    f_high = npv(high)

    tries = 0
    while f_low * f_high > 0 and tries < 20:
        high *= 2
        f_high = npv(high)
        tries += 1

    if f_low * f_high > 0:
        return float("nan")

    mid = float("nan")
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
    cum = 0.0
    for i, cf in enumerate(unlevered_cf):
        prev = cum
        cum += cf
        if cum >= 0 and i > 0:
            if cf == 0:
                return float(i)
            frac = (0 - prev) / cf
            return (i - 1) + frac
    return float("nan")


# -----------------------------
# Debt engine (annual sculpting + fees)
# -----------------------------
def _total_capex_from_lines(s: ScenarioInputs) -> float:
    cap_df = pd.DataFrame(s.capex.lines or [])
    if cap_df.empty or "Amount_COP" not in cap_df.columns:
        return 0.0
    return float(pd.to_numeric(cap_df["Amount_COP"], errors="coerce").fillna(0.0).sum())

def _eligible_capex_for_tax(s: ScenarioInputs) -> float:
    total_capex = _total_capex_from_lines(s)
    pct = max(0.0, min(100.0, s.renewable_tax.special_deduction_pct_of_capex))
    return total_capex * pct / 100.0

def _month_start(d: date) -> date:
    return date(d.year, d.month, 1)

def debt_commitment_fee_annual(s: ScenarioInputs) -> pd.DataFrame:
    """
    Commitment fee on undrawn debt during construction:
    Commitment rate = margin * commitment_fee_pct_of_margin
    Applied monthly on undrawn = TotalDebt - CumulativeDrawn
    """
    tl = build_timeline(s.timeline)
    cod_m = date(tl["cod"].year, tl["cod"].month, 1)

    total_capex = _total_capex_from_lines(s)
    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex
    debt_amt = max(debt_amt, 0.0)

    # Monthly CAPEX schedule up to COD month (exclude COD month)
    capm = capex_monthly_schedule(s).copy()
    capm["Month"] = pd.to_datetime(capm["Month"])
    capm = capm[capm["Month"] < pd.to_datetime(cod_m)].copy()
    if capm.empty or debt_amt <= 0:
        return pd.DataFrame({"Year": [], "Commitment Fee (COP)": []})

    # Assume debt draws pro-rata with CAPEX spend during dev+construction (pre-COD)
    draw_m = capm.copy()
    draw_m["Draw (COP)"] = (float(s.debt.debt_pct_of_capex) / 100.0) * pd.to_numeric(draw_m["CAPEX (COP)"], errors="coerce").fillna(0.0)

    draw_m["Cum Draw (COP)"] = draw_m["Draw (COP)"].cumsum()
    draw_m["Undrawn (COP)"] = np.maximum(debt_amt - draw_m["Cum Draw (COP)"], 0.0)

    margin = float(s.debt.margin_pct) / 100.0
    commit_mult = float(s.debt.commitment_fee_pct_of_margin) / 100.0
    commit_rate_annual = margin * commit_mult
    commit_rate_monthly = commit_rate_annual / 12.0

    draw_m["Commitment Fee (COP)"] = draw_m["Undrawn (COP)"] * commit_rate_monthly
    draw_m["Year"] = draw_m["Month"].dt.year

    ann = draw_m.groupby("Year", as_index=False)[["Commitment Fee (COP)"]].sum()
    return ann


def debt_schedule_annual(s: ScenarioInputs) -> pd.DataFrame:
    """
    Annual debt schedule starting at COD year for tenor years.
    Amortization:
      - Sculpted to Target DSCR, or
      - Equal principal
    DSCR calculated as: Operating CF / Debt Service
    """
    tl = build_timeline(s.timeline)
    cod_year = tl["cod"].year

    total_capex = _total_capex_from_lines(s)
    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex
    debt_amt = max(debt_amt, 0.0)

    tenor = int(s.debt.tenor_years)
    tenor = max(5, min(10, tenor))
    grace = int(s.debt.grace_years)
    grace = max(0, min(tenor - 1, grace))

    years = list(range(cod_year, cod_year + tenor))

    # Operating CF annual from model (EBITDA proxy)
    a = cashflow_annual_table(s).copy()
    opcf_map = {int(r["Year"]): float(r["Operating CF (COP)"]) for _, r in a.iterrows()}
    cf = [float(opcf_map.get(y, 0.0)) for y in years]

    all_in = (float(s.debt.base_rate_pct) + float(s.debt.margin_pct)) / 100.0
    target = float(s.debt.target_dscr) if float(s.debt.target_dscr) > 0 else 1.20

    # NEW: balloon target (COP). % of original debt you want left outstanding at maturity.
    balloon_pct = float(getattr(s.debt, "balloon_pct", 0.0)) / 100.0
    balloon_target = debt_amt * balloon_pct
    balloon_target = max(0.0, min(balloon_target, debt_amt))

    outstanding = debt_amt
    rows = []

    for i, y in enumerate(years, start=1):
        operating_cf = float(cf[i - 1])
        interest = outstanding * all_in

        if outstanding <= 1e-9:
            principal = 0.0
            debt_service = 0.0
            dscr = float("nan")
            rows.append({
                "Year": y,
                "Operating CF (COP)": operating_cf,
                "Interest (COP)": interest,
                "Principal (COP)": principal,
                "Debt Service (COP)": debt_service,
                "DSCR": dscr,
                "Outstanding End (COP)": 0.0,
            })
            outstanding = 0.0
            continue

        if i <= grace:
            principal = 0.0
            # During grace, interest is still due (cash outflow)
            # If cash flow is insufficient, interest is still paid (may create negative levered CF)
        else:
            # Calculate remaining amortization years (tenor minus grace, minus years already amortized)
            # Years already amortized = (i - 1 - grace) since grace years don't count
            years_amortized = max(0, i - 1 - grace)
            remaining_years = max(tenor - grace - years_amortized, 1)

            if s.debt.amortization_type == "Equal principal":
                # Amortize evenly down to balloon_target over remaining years
                principal = max(0.0, (outstanding - balloon_target) / remaining_years)

                # Final year: force ending outstanding = balloon_target
                if i == tenor:
                    principal = max(0.0, outstanding - balloon_target)

            else:
                # Sculpt to target DSCR (CFADS proxy = Operating CF)
                max_ds = operating_cf / target if target > 0 else 0.0
                principal = max(0.0, max_ds - interest)

                # Constrain: don't pay more than what would amortize evenly over remaining years
                # This ensures we respect the tenor (don't pay off too early)
                max_principal_by_tenor = max(0.0, (outstanding - balloon_target) / remaining_years) if remaining_years > 0 else outstanding - balloon_target
                principal = min(principal, max(0.0, max_principal_by_tenor))

                # Do not amortize below balloon target (unless final year)
                if i < tenor:
                    principal = min(principal, max(0.0, outstanding - balloon_target))
                else:
                    principal = max(0.0, outstanding - balloon_target)

        principal = min(principal, outstanding)
        debt_service = interest + principal
        dscr = (operating_cf / debt_service) if debt_service > 0 else float("nan")

        outstanding = outstanding - principal

        rows.append({
            "Year": y,
            "Operating CF (COP)": operating_cf,
            "Interest (COP)": interest,
            "Principal (COP)": principal,
            "Debt Service (COP)": debt_service,
            "DSCR": dscr,
            "Outstanding End (COP)": outstanding,
        })

    df = pd.DataFrame(rows)

    balloon = float(df["Outstanding End (COP)"].iloc[-1]) if len(df) else 0.0
    df["Balloon at Maturity (COP)"] = 0.0
    if len(df):
        df.loc[df.index[-1], "Balloon at Maturity (COP)"] = balloon

    # Upfront fee (paid at COD, show in first year)
    upfront_fee = debt_amt * (float(s.debt.upfront_fee_bps) / 10_000.0)
    df["Upfront Fee (COP)"] = 0.0
    if len(df):
        df.loc[df.index[0], "Upfront Fee (COP)"] = upfront_fee

    # Covenant / lock-up flags
    min_covenant = float(s.debt.min_dscr_covenant)
    lockup = float(s.debt.lockup_dscr)
    df["Lock-up? (DSCR < lock-up)"] = df["DSCR"].apply(lambda x: bool(np.isfinite(x) and x < lockup))
    df["Breach? (DSCR < min)"] = df["DSCR"].apply(lambda x: bool(np.isfinite(x) and x < min_covenant))

    return df


# -----------------------------
# Levered Cash Flow (Equity Cash Flow)
# -----------------------------
def levered_cashflow_annual(s: ScenarioInputs) -> pd.DataFrame:
    """
    Calculate levered (equity) cash flow after-tax.
    Starts with unlevered after-tax CF, adds debt draws, subtracts debt service and fees.
    If debt is disabled or not enabled, assumes no debt (returns unlevered CF).
    """
    # Start with unlevered after-tax cash flow
    unlevered = unlevered_base_cashflow_annual(s).copy()
    
    # Check if debt is enabled
    debt_enabled = bool(getattr(s.debt, "enabled", False))
    total_capex = _total_capex_from_lines(s)
    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex if debt_enabled else 0.0
    debt_amt = max(debt_amt, 0.0)
    
    # Initialize debt-related columns
    unlevered["Debt Draw (COP)"] = 0.0
    unlevered["Interest (COP)"] = 0.0
    unlevered["Principal (COP)"] = 0.0
    unlevered["Debt Service (COP)"] = 0.0
    unlevered["Debt Fees (COP)"] = 0.0
    
    if debt_enabled and debt_amt > 0:
        # Get debt schedule (annual interest, principal, upfront fee)
        ds = debt_schedule_annual(s).copy()
        
        if not ds.empty and "Year" in ds.columns:
            # Merge debt schedule - drop initialized columns first to avoid suffix conflicts
            merge_cols = ["Year"]
            for col in ["Interest (COP)", "Principal (COP)", "Upfront Fee (COP)"]:
                if col in ds.columns:
                    merge_cols.append(col)
                    # Drop initialized column if it exists to avoid merge conflicts
                    if col in unlevered.columns:
                        unlevered = unlevered.drop(columns=[col])
            
            if len(merge_cols) > 1:  # More than just "Year"
                unlevered = unlevered.merge(ds[merge_cols], on="Year", how="left")
                
                # Ensure columns exist and fill NaN values (after merge, columns should exist)
                for col in ["Interest (COP)", "Principal (COP)"]:
                    if col in unlevered.columns:
                        unlevered[col] = unlevered[col].fillna(0.0)
                    else:
                        unlevered[col] = 0.0
                
                unlevered["Debt Service (COP)"] = unlevered["Interest (COP)"] + unlevered["Principal (COP)"]
                
                # Upfront fee
                if "Upfront Fee (COP)" in unlevered.columns:
                    unlevered["Debt Fees (COP)"] = unlevered["Debt Fees (COP)"] + unlevered["Upfront Fee (COP)"].fillna(0.0)
            
            # Debt draws: pro-rata with CAPEX during construction (up to and including COD)
            tl = build_timeline(s.timeline)
            cod_year = tl["cod"].year
            cod_m = date(tl["cod"].year, tl["cod"].month, 1)
            
            # Get monthly CAPEX schedule
            capm = capex_monthly_schedule(s).copy()
            capm["Year"] = capm["Month"].apply(lambda d: d.year)
            
            # Calculate debt draws (pro-rata with CAPEX, up to and including COD month)
            if not capm.empty:
                # Include CAPEX up to and including COD month
                up_to_cod = capm[capm["Month"] <= cod_m].copy()
                if not up_to_cod.empty:
                    up_to_cod["Debt Draw (COP)"] = (float(s.debt.debt_pct_of_capex) / 100.0) * pd.to_numeric(up_to_cod["CAPEX (COP)"], errors="coerce").fillna(0.0)
                    
                    # Ensure total debt draws equal debt amount (adjust last draw if needed)
                    total_draws = up_to_cod["Debt Draw (COP)"].sum()
                    if total_draws > 0 and abs(total_draws - debt_amt) > 1.0:  # Allow small rounding differences
                        # Adjust proportionally or add remainder to COD month
                        adjustment = debt_amt - total_draws
                        if cod_m in up_to_cod["Month"].values:
                            # Add adjustment to COD month
                            cod_idx = up_to_cod[up_to_cod["Month"] == cod_m].index[0]
                            up_to_cod.loc[cod_idx, "Debt Draw (COP)"] += adjustment
                        else:
                            # Scale proportionally
                            if total_draws > 0:
                                scale_factor = debt_amt / total_draws
                                up_to_cod["Debt Draw (COP)"] = up_to_cod["Debt Draw (COP)"] * scale_factor
                    
                    # Aggregate to annual
                    annual_draws = up_to_cod.groupby("Year", as_index=False)["Debt Draw (COP)"].sum()
                    # Merge into unlevered (drop the initialized column first to avoid suffix conflicts)
                    if "Debt Draw (COP)" in unlevered.columns:
                        unlevered = unlevered.drop(columns=["Debt Draw (COP)"])
                    unlevered = unlevered.merge(annual_draws, on="Year", how="left")
                    # Ensure column exists and fill NaN
                    if "Debt Draw (COP)" in unlevered.columns:
                        unlevered["Debt Draw (COP)"] = unlevered["Debt Draw (COP)"].fillna(0.0)
                    else:
                        unlevered["Debt Draw (COP)"] = 0.0
                # If no CAPEX up to COD, Debt Draw stays at initialized 0.0
            
            # Commitment fees (during construction)
            cf_ann = debt_commitment_fee_annual(s).copy()
            if not cf_ann.empty and "Year" in cf_ann.columns and "Commitment Fee (COP)" in cf_ann.columns:
                unlevered = unlevered.merge(cf_ann[["Year", "Commitment Fee (COP)"]], on="Year", how="left")
                unlevered["Debt Fees (COP)"] = unlevered["Debt Fees (COP)"] + unlevered["Commitment Fee (COP)"].fillna(0.0)
    
    # Calculate levered (equity) cash flow
    # Equity CF = Unlevered After-Tax CF + Debt Draws - Debt Service - Debt Fees
    unlevered["Levered CF (After-tax, COP)"] = (
        unlevered["Unlevered CF After Tax (COP)"].fillna(0.0)
        + unlevered["Debt Draw (COP)"].fillna(0.0)
        - unlevered["Debt Service (COP)"].fillna(0.0)
        - unlevered["Debt Fees (COP)"].fillna(0.0)
    )
    
    unlevered["Cumulative Levered CF (COP)"] = unlevered["Levered CF (After-tax, COP)"].cumsum()
    
    return unlevered


def levered_cashflow_monthly(s: ScenarioInputs) -> pd.DataFrame:
    """
    Calculate monthly levered (equity) cash flow after-tax.
    Starts with monthly unlevered CF, adds debt draws, subtracts debt service and fees.
    """
    # Start with monthly cash flow table
    monthly = cashflow_monthly_table(s).copy()
    
    # Get annual levered CF to extract debt components
    annual_levered = levered_cashflow_annual(s).copy()
    
    # Initialize debt columns
    monthly["Debt Draw (COP)"] = 0.0
    monthly["Interest (COP)"] = 0.0
    monthly["Principal (COP)"] = 0.0
    monthly["Debt Service (COP)"] = 0.0
    monthly["Debt Fees (COP)"] = 0.0
    
    # Check if debt is enabled
    debt_enabled = bool(getattr(s.debt, "enabled", False))
    total_capex = _total_capex_from_lines(s)
    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex if debt_enabled else 0.0
    debt_amt = max(debt_amt, 0.0)
    
    if debt_enabled and debt_amt > 0:
        tl = build_timeline(s.timeline)
        cod_m = date(tl["cod"].year, tl["cod"].month, 1)
        
        # Debt draws: pro-rata with CAPEX during construction (pre-COD)
        capm = capex_monthly_schedule(s).copy()
        if not capm.empty:
            capm["Debt Draw (COP)"] = (float(s.debt.debt_pct_of_capex) / 100.0) * pd.to_numeric(capm["CAPEX (COP)"], errors="coerce").fillna(0.0)
            # Only draws before COD
            capm.loc[capm["Month"] >= cod_m, "Debt Draw (COP)"] = 0.0
            # Merge into monthly (drop initialized column first to avoid suffix conflicts)
            if "Debt Draw (COP)" in monthly.columns:
                monthly = monthly.drop(columns=["Debt Draw (COP)"])
            monthly = monthly.merge(capm[["Month", "Debt Draw (COP)"]], on="Month", how="left")
            # Ensure column exists and fill NaN
            if "Debt Draw (COP)" in monthly.columns:
                monthly["Debt Draw (COP)"] = monthly["Debt Draw (COP)"].fillna(0.0)
            else:
                monthly["Debt Draw (COP)"] = 0.0
        
        # Debt service: spread annual interest and principal evenly across months
        if not annual_levered.empty:
            # Create a map of annual debt service components
            interest_map = {int(r["Year"]): float(r["Interest (COP)"]) for _, r in annual_levered.iterrows()}
            principal_map = {int(r["Year"]): float(r["Principal (COP)"]) for _, r in annual_levered.iterrows()}
            
            monthly["Interest (COP)"] = monthly["Year"].map(lambda y: interest_map.get(int(y), 0.0) / 12.0)
            monthly["Principal (COP)"] = monthly["Year"].map(lambda y: principal_map.get(int(y), 0.0) / 12.0)
            monthly["Debt Service (COP)"] = monthly["Interest (COP)"] + monthly["Principal (COP)"]
            
            # Only apply debt service during operation (after COD)
            monthly.loc[monthly["Month"] < cod_m, "Interest (COP)"] = 0.0
            monthly.loc[monthly["Month"] < cod_m, "Principal (COP)"] = 0.0
            monthly.loc[monthly["Month"] < cod_m, "Debt Service (COP)"] = 0.0
        
        # Debt fees: upfront fee at COD, commitment fees during construction
        # Initialize Debt Fees column if not exists
        if "Debt Fees (COP)" not in monthly.columns:
            monthly["Debt Fees (COP)"] = 0.0
        
        # Upfront fee
        upfront_fee = debt_amt * (float(s.debt.upfront_fee_bps) / 10_000.0)
        monthly.loc[monthly["Month"] == cod_m, "Debt Fees (COP)"] = upfront_fee
        
        # Commitment fees: get from annual and spread
        cf_ann = debt_commitment_fee_annual(s).copy()
        if not cf_ann.empty and "Year" in cf_ann.columns:
            cf_map = {int(r["Year"]): float(r["Commitment Fee (COP)"]) for _, r in cf_ann.iterrows()}
            monthly["Commitment Fee (COP)"] = monthly["Year"].map(lambda y: cf_map.get(int(y), 0.0) / 12.0)
            monthly["Debt Fees (COP)"] = monthly["Debt Fees (COP)"].fillna(0.0) + monthly["Commitment Fee (COP)"].fillna(0.0)
        
        monthly["Debt Fees (COP)"] = monthly["Debt Fees (COP)"].fillna(0.0)
    
    # Calculate levered CF (need to estimate after-tax CF monthly)
    # For monthly, we'll approximate: use unlevered CF and apply annual tax rate proportionally
    # Better approach: use annual after-tax CF and spread proportionally
    if not annual_levered.empty:
        # Map annual after-tax unlevered CF to monthly (proportional)
        unlevered_after_tax_map = {int(r["Year"]): float(r["Unlevered CF After Tax (COP)"]) for _, r in annual_levered.iterrows()}
        unlevered_cf_map = {int(r["Year"]): float(r["Unlevered CF (COP)"]) for _, r in annual_levered.iterrows()}
        
        monthly["Unlevered CF After Tax (Monthly, COP)"] = 0.0
        for year in monthly["Year"].unique():
            year_unlevered = unlevered_cf_map.get(int(year), 0.0)
            year_after_tax = unlevered_after_tax_map.get(int(year), 0.0)
            if abs(year_unlevered) > 1e-6:
                tax_factor = year_after_tax / year_unlevered
            else:
                tax_factor = 1.0
            
            year_mask = monthly["Year"] == year
            monthly.loc[year_mask, "Unlevered CF After Tax (Monthly, COP)"] = (
                monthly.loc[year_mask, "Unlevered CF (COP)"] * tax_factor
            )
    else:
        monthly["Unlevered CF After Tax (Monthly, COP)"] = monthly["Unlevered CF (COP)"]
    
    # Levered CF = Unlevered After-Tax CF + Debt Draws - Debt Service - Debt Fees
    monthly["Levered CF (After-tax, COP)"] = (
        monthly["Unlevered CF After Tax (Monthly, COP)"].fillna(0.0)
        + monthly["Debt Draw (COP)"].fillna(0.0)
        - monthly["Debt Service (COP)"].fillna(0.0)
        - monthly["Debt Fees (COP)"].fillna(0.0)
    )
    
    monthly["Cumulative Levered CF (COP)"] = monthly["Levered CF (After-tax, COP)"].cumsum()
    
    return monthly


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
tab_macro, tab_timeline, tab_gen, tab_rev, tab_capex, tab_opex, tab_sga, tab_dep, tab_incent, tab_ucf, tab_debt, tab_levered, tab_compare = st.tabs(
    [
        "A) Macroeconomic",
        "B) Timeline",
        "C) Power Generation",
        "D) Power Revenues",
        "E) CAPEX",
        "F) OPEX",
        "G) SG&A",
        "H) Depreciation",
        "I) Renewable tax benefits",
        "J) Unlevered Base Cash Flow",
        "K) Debt & Covenants",
        "L) Levered Cash Flow",
        "M) Compare",
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
    fig.update_yaxes(autorange="reversed")
    fig.update_xaxes(dtick="M12", tickformat="%Y")
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

    total_capex = float(pd.to_numeric(edited["Amount_COP"], errors="coerce").fillna(0.0).sum())
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
        share = capex_pie["Amount_COP"] / capex_pie["Amount_COP"].sum()
        small = share < 0.03
        if small.any() and (~small).any():
            other_amt = float(capex_pie.loc[small, "Amount_COP"].sum())
            capex_pie = capex_pie.loc[~small, ["Item", "Amount_COP"]]
            capex_pie = pd.concat(
                [capex_pie, pd.DataFrame([{"Item": "Other (<3% each)", "Amount_COP": other_amt}])],
                ignore_index=True
            )
        fig_pie = px.pie(capex_pie, names="Item", values="Amount_COP", hole=0.45)
        fig_pie.update_traces(textinfo="percent+label")
        fig_pie.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("#### CAPEX schedule (monthly, aligned to timeline)")
    sched = capex_monthly_schedule(s)
    sched_disp = _df_format_money(sched.copy(), ["CAPEX (COP)"], decimals=0)
    st.dataframe(sched_disp[["Month", "Phase", "CAPEX (COP)"]], use_container_width=True, hide_index=True)

    fig = px.bar(sched, x="Month", y="CAPEX (COP)", color="Phase")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Annual CAPEX (calendar years)")
    ann = sched.groupby("Year", as_index=False)["CAPEX (COP)"].sum()
    ann_disp = _df_format_money(ann.copy(), ["CAPEX (COP)"], decimals=0)
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

    om_full = opex_monthly_schedule(s).copy()

    meta_cols = {"Month", "Year", "Phase", "Energy (MWh)", "Revenue (COP)", "CAPEX (COP)", "OPEX subtotal", "GMF"}
    fixed_cols = {"Fixed O&M", "Insurance", "Variable O&M", "Grid fees", "Land lease", "ICA"}
    dyn_cols = [c for c in om_full.columns if c not in meta_cols and c not in fixed_cols]
    opex_item_cols = list(fixed_cols) + dyn_cols

    annual_items = om_full.groupby("Year", as_index=False)[opex_item_cols + ["GMF"]].sum()
    annual_items["Total OPEX (COP)"] = annual_items[opex_item_cols].sum(axis=1) + annual_items["GMF"]

    long = annual_items.melt(id_vars=["Year"], value_vars=opex_item_cols + ["GMF"], var_name="Item", value_name="OPEX (COP)")
    fig = px.bar(long, x="Year", y="OPEX (COP)", color="Item", barmode="stack")
    fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)

    annual = om_full.groupby("Year", as_index=False)[["OPEX subtotal", "GMF"]].sum()
    annual["Total OPEX (COP)"] = annual["OPEX subtotal"] + annual["GMF"]

    op = operating_year_table(s)[["Year", "Energy (MWh)"]].copy()
    annual = annual.merge(op, on="Year", how="left").fillna({"Energy (MWh)": 0.0})
    annual["OPEX per MWh (COP/MWh)"] = np.where(annual["Energy (MWh)"] > 0, annual["Total OPEX (COP)"] / annual["Energy (MWh)"], 0.0)

    disp = _df_format_money(annual.copy(), ["OPEX subtotal", "GMF", "Total OPEX (COP)", "OPEX per MWh (COP/MWh)", "Energy (MWh)"], decimals=0)
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

    annual_disp = _df_format_money(annual_sga.copy(), [c for c in annual_sga.columns if c != "Year"], decimals=0)
    st.dataframe(annual_disp, use_container_width=True, hide_index=True)


# -----------------------------
# H) Depreciation
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

    dep_disp = _df_format_money(dep.copy(), [c for c in dep.columns if c != "Year"], decimals=0)
    st.dataframe(dep_disp, use_container_width=True, hide_index=True)


# -----------------------------
# K) Debt & Covenants
# -----------------------------
with tab_debt:
    st.subheader("Debt & Covenants (tenor 5–10 years, sculpted amortization)")

    s.debt.enabled = st.checkbox("Enable debt", value=bool(s.debt.enabled))
    s.debt.balloon_pct = st.slider("Balloon (% of original debt at maturity)", 0.0, 50.0, float(getattr(s.debt, "balloon_pct", 0.0)), 1.0)

    if not s.debt.enabled:
        st.info("Debt is disabled. Enable it to see debt sizing, amortization, DSCR, and covenant indicators.")
    else:
        total_capex = _total_capex_from_lines(s)
        s.debt.debt_pct_of_capex = st.slider("Debt as % of CAPEX", min_value=0.0, max_value=90.0, value=float(s.debt.debt_pct_of_capex), step=1.0)
        debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex

        c1, c2, c3 = st.columns(3)
        with c1:
            s.debt.tenor_years = int(st.number_input("Tenor (years)", value=int(s.debt.tenor_years), min_value=5, max_value=10, step=1, format="%d"))
        with c2:
            s.debt.grace_years = int(st.number_input("Grace period (years)", value=int(s.debt.grace_years), min_value=0, max_value=max(0, int(s.debt.tenor_years) - 1), step=1, format="%d"))
        with c3:
            s.debt.amortization_type = st.selectbox("Amortization", ["Sculpted to DSCR", "Equal principal"], index=0 if s.debt.amortization_type != "Equal principal" else 1)

        st.markdown("#### Pricing (Natural COP)")
        p1, p2, p3 = st.columns(3)
        with p1:
            s.debt.base_rate_pct = st.number_input("Base rate (e.g., IBR) %", value=float(s.debt.base_rate_pct), step=0.25, format="%.2f")
        with p2:
            s.debt.margin_pct = st.number_input("Margin (spread) %", value=float(s.debt.margin_pct), step=0.25, format="%.2f")
        with p3:
            all_in = float(s.debt.base_rate_pct) + float(s.debt.margin_pct)
            st.metric("All-in interest rate", f"{all_in:,.2f}%")

        st.markdown("#### Fees")
        f1, f2 = st.columns(2)
        with f1:
            s.debt.upfront_fee_bps = st.number_input("Upfront / structuring fee (bps of debt)", value=float(s.debt.upfront_fee_bps), step=5.0, format="%.0f")
        with f2:
            s.debt.commitment_fee_pct_of_margin = st.number_input("Commitment fee (% of margin) on undrawn", value=float(s.debt.commitment_fee_pct_of_margin), step=1.0, format="%.0f")

        st.markdown("#### Covenants")
        k1, k2, k3 = st.columns(3)
        with k1:
            s.debt.target_dscr = st.number_input("Target DSCR (for sculpting)", value=float(s.debt.target_dscr), step=0.01, format="%.2f")
        with k2:
            s.debt.min_dscr_covenant = st.number_input("Minimum DSCR covenant", value=float(s.debt.min_dscr_covenant), step=0.01, format="%.2f")
        with k3:
            s.debt.lockup_dscr = st.number_input("Lock-up DSCR threshold", value=float(s.debt.lockup_dscr), step=0.01, format="%.2f")

        _metric_row([
            ("Total CAPEX", _fmt_cop(total_capex)),
            ("Debt amount", _fmt_cop(debt_amt)),
            ("Equity (implied)", _fmt_cop(max(total_capex - debt_amt, 0.0))),
        ])

        st.divider()

        ds = debt_schedule_annual(s)
        com = debt_commitment_fee_annual(s)

        # KPIs over debt life where debt service > 0
        ds_valid = ds[pd.to_numeric(ds["Debt Service (COP)"], errors="coerce").fillna(0.0) > 0].copy()
        min_dscr = float(ds_valid["DSCR"].min()) if not ds_valid.empty else float("nan")
        lockup_years = int(ds["Lock-up? (DSCR < lock-up)"].sum()) if "Lock-up? (DSCR < lock-up)" in ds.columns else 0
        breach_years = int(ds["Breach? (DSCR < min)"].sum()) if "Breach? (DSCR < min)" in ds.columns else 0
        balloon = float(ds["Balloon at Maturity (COP)"].sum()) if "Balloon at Maturity (COP)" in ds.columns else 0.0
        upfront_fee = float(ds["Upfront Fee (COP)"].sum()) if "Upfront Fee (COP)" in ds.columns else 0.0
        commit_total = float(com["Commitment Fee (COP)"].sum()) if (not com.empty and "Commitment Fee (COP)" in com.columns) else 0.0

        _metric_row([
            ("Min DSCR (debt life)", f"{min_dscr:,.2f}x" if np.isfinite(min_dscr) else "—"),
            ("Lock-up years", str(lockup_years)),
            ("Breach years", str(breach_years)),
            ("Balloon", _fmt_cop(balloon)),
        ])
        _metric_row([
            ("Upfront fee", _fmt_cop(upfront_fee)),
            ("Commitment fees (total)", _fmt_cop(commit_total)),
            ("Tenor", f"{int(s.debt.tenor_years)} years"),
            ("Grace", f"{int(s.debt.grace_years)} years"),
        ])

        st.markdown("### Debt schedule (annual)")
        disp = ds.copy()
        money_cols = ["Operating CF (COP)", "Interest (COP)", "Principal (COP)", "Debt Service (COP)", "Outstanding End (COP)", "Balloon at Maturity (COP)", "Upfront Fee (COP)"]
        disp = _df_format_money(disp, money_cols, decimals=0)
        st.dataframe(disp, use_container_width=True, hide_index=True)

        st.markdown("### DSCR vs covenant thresholds")
        ds_plot = ds.copy()
        ds_plot["Min covenant"] = float(s.debt.min_dscr_covenant)
        ds_plot["Lock-up"] = float(s.debt.lockup_dscr)
        ds_long = ds_plot.melt(id_vars=["Year"], value_vars=["DSCR", "Min covenant", "Lock-up"], var_name="Line", value_name="Value")
        fig = px.line(ds_long, x="Year", y="Value", color="Line")
        fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Commitment fees (during construction)")
        if com.empty:
            st.info("No commitment fees computed (likely no construction months or debt amount is 0).")
        else:
            com_disp = _df_format_money(com.copy(), ["Commitment Fee (COP)"], decimals=0)
            st.dataframe(com_disp, use_container_width=True, hide_index=True)
            figc = px.bar(com, x="Year", y="Commitment Fee (COP)")
            figc.update_layout(height=260, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(figc, use_container_width=True)


# -----------------------------
# I) Renewable tax benefits (moved before Unlevered CF)
# -----------------------------
with tab_incent:
    st.subheader("Renewable tax benefits (Colombia)")

    inc = s.incentives

    st.markdown("### Special deduction (up to 50% of eligible CAPEX over up to 15 years)")
    c1, c2, c3 = st.columns(3)
    with c1:
        inc.enable_special_deduction = st.checkbox("Enable special deduction", value=bool(inc.enable_special_deduction))
    with c2:
        inc.special_deduction_pct_of_capex = st.number_input(
            "Deduction pool (% of CAPEX, max 50%)",
            value=float(inc.special_deduction_pct_of_capex),
            min_value=0.0, max_value=50.0, step=1.0, format="%.1f"
        )
    with c3:
        inc.special_deduction_years = int(st.number_input(
            "Carryforward window (years)",
            value=int(inc.special_deduction_years),
            min_value=1, max_value=30, step=1, format="%d"
        ))

    inc.special_deduction_max_pct_of_taxable_income = st.number_input(
        "Annual cap (% of taxable income, typically 50%)",
        value=float(inc.special_deduction_max_pct_of_taxable_income),
        min_value=0.0, max_value=100.0, step=1.0, format="%.1f"
    )

    st.divider()
    st.markdown("### VAT treatment")
    inc.vat_mode = st.selectbox("VAT mode", ["Excluded", "Refund"], index=0 if inc.vat_mode != "Refund" else 1)

    c4, c5, c6 = st.columns(3)
    with c4:
        inc.vat_pct_of_capex = st.number_input(
            "VAT as % of CAPEX (optional)",
            value=float(inc.vat_pct_of_capex),
            min_value=0.0, max_value=100.0, step=0.5, format="%.2f"
        )
    with c5:
        inc.vat_fixed_cop = st.number_input(
            "VAT fixed COP (optional)",
            value=float(inc.vat_fixed_cop),
            min_value=0.0, step=1_000_000.0, format="%.0f"
        )
    with c6:
        inc.vat_refund_year_index = int(st.number_input(
            "Refund in operating year #",
            value=int(inc.vat_refund_year_index),
            min_value=1, max_value=10, step=1, format="%d"
        ))

    st.caption("If VAT is Excluded, leave VAT inputs at 0. If Refund, you can use % or fixed; both will be added (so use one).")
    st.info("ℹ️ These tax benefits are automatically applied in the Unlevered Base Cash Flow calculation below.")


# -----------------------------
# J) Unlevered Base Cash Flow
# -----------------------------
with tab_ucf:
    st.subheader("Unlevered Base Cash Flow (includes tax benefits, pre-debt)")

    st.info("This tab calculates unlevered cash flow before and after tax. Tax benefits from the Renewable tax benefits tab are automatically included. This is the base for calculating debt capacity and equity returns.")

    tx1, tx2 = st.columns([1, 2])
    with tx1:
        s.tax.corporate_tax_rate_pct = st.number_input("Corporate income tax rate (%)", value=float(s.tax.corporate_tax_rate_pct), step=0.5, format="%.2f")
    with tx2:
        s.tax.allow_loss_carryforward = st.checkbox("Apply loss carryforward (NOL) so taxes are zero until losses are used", value=bool(s.tax.allow_loss_carryforward))

    st.markdown("### Working Capital (timing)")
    wc1, wc2, wc3, wc4 = st.columns([1, 1, 1, 1])
    with wc1:
        s.wc.ar_days = int(st.number_input("AR days (revenue collection)", value=int(s.wc.ar_days), min_value=0, step=15, format="%d"))
    with wc2:
        s.wc.ap_days = int(st.number_input("AP days (expense payment)", value=int(s.wc.ap_days), min_value=0, step=15, format="%d"))
    with wc3:
        s.wc.apply_ap_to_opex = st.checkbox("Apply AP lag to OPEX", value=bool(s.wc.apply_ap_to_opex))
    with wc4:
        s.wc.apply_ap_to_sga = st.checkbox("Apply AP lag to SG&A", value=bool(s.wc.apply_ap_to_sga))

    currency = st.radio("Display currency", ["COP", "USD"], horizontal=True, index=0, key="currency_unlevered")

    m = cashflow_monthly_table(s)
    a = unlevered_base_cashflow_annual(s)

    tl = build_timeline(s.timeline)
    years = list(a["Year"].astype(int).tolist())
    fx = fx_series(s.macro, tl["cod"].year, years)

    def _conv(series: pd.Series, year_col: pd.Series) -> pd.Series:
        if currency == "COP":
            return series
        return series / year_col.map(lambda y: float(fx.loc[int(y)]) if int(y) in fx.index else float(s.macro.fx_cop_per_usd_start))

    st.divider()

    # --- KPIs ---
    mm = cashflow_monthly_table(s).copy()
    for col in ["CAPEX (COP)", "Unlevered CF (COP)", "Cumulative Unlevered CF (COP)"]:
        if col in mm.columns:
            mm[col] = pd.to_numeric(mm[col], errors="coerce").fillna(0.0)

    cap_total = float(pd.to_numeric(mm["CAPEX (COP)"], errors="coerce").fillna(0.0).sum()) if "CAPEX (COP)" in mm.columns else 0.0
    
    # Pre-tax cash flows
    monthly_cf_pre_tax = mm["Unlevered CF (COP)"].astype(float).tolist() if "Unlevered CF (COP)" in mm.columns else []
    monthly_dates = mm["Month"].tolist() if "Month" in mm.columns else []

    # After-tax cash flows: calculate monthly from annual by prorating
    # Create a map of annual after-tax CF by year
    annual_after_tax_map = {}
    annual_pre_tax_map = {}
    for _, row in a.iterrows():
        year = int(row["Year"])
        annual_after_tax_map[year] = float(row.get("Unlevered CF After Tax (COP)", 0.0))
        annual_pre_tax_map[year] = float(row.get("Unlevered CF (COP)", 0.0))
    
    # Calculate operating months per year for proration
    cod = tl["cod"]
    end_op = tl["end_op"]
    operating_months_map = {}
    for year in annual_after_tax_map.keys():
        if year == cod.year:
            operating_months_map[year] = 13 - cod.month
        elif year == end_op.year:
            operating_months_map[year] = end_op.month
        else:
            operating_months_map[year] = 12
    
    # Create monthly after-tax CF
    monthly_cf_after_tax = []
    for i, month_date in enumerate(monthly_dates):
        year = month_date.year
        phase = _phase_for_month(tl, month_date)
        
        if phase == "Operation" and year in annual_after_tax_map:
            # Prorate annual after-tax CF to monthly
            operating_months = operating_months_map.get(year, 12)
            if operating_months > 0:
                monthly_after_tax = annual_after_tax_map[year] / operating_months
            else:
                monthly_after_tax = 0.0
            monthly_cf_after_tax.append(monthly_after_tax)
        else:
            # For non-operation months, use pre-tax CF (no tax impact)
            monthly_cf_after_tax.append(monthly_cf_pre_tax[i] if i < len(monthly_cf_pre_tax) else 0.0)

    # Pre-tax IRR
    has_pos_pre = any(cf > 0 for cf in monthly_cf_pre_tax)
    has_neg_pre = any(cf < 0 for cf in monthly_cf_pre_tax)
    irr_m_pre = _irr_bisection(monthly_cf_pre_tax) if (has_pos_pre and has_neg_pre) else float("nan")
    irr_annual_pre = (1.0 + irr_m_pre) ** 12 - 1.0 if np.isfinite(irr_m_pre) else float("nan")

    # After-tax IRR
    has_pos_after = any(cf > 0 for cf in monthly_cf_after_tax)
    has_neg_after = any(cf < 0 for cf in monthly_cf_after_tax)
    irr_m_after = _irr_bisection(monthly_cf_after_tax) if (has_pos_after and has_neg_after) else float("nan")
    irr_annual_after = (1.0 + irr_m_after) ** 12 - 1.0 if np.isfinite(irr_m_after) else float("nan")

    # Payback (after-tax)
    payback_m_after = _payback_months(monthly_dates, monthly_cf_after_tax) if monthly_cf_after_tax else float("nan")
    payback_years_after = payback_m_after / 12.0 if np.isfinite(payback_m_after) else float("nan")

    # Peak funding (from pre-tax cumulative CF)
    cum = mm["Cumulative Unlevered CF (COP)"].astype(float) if "Cumulative Unlevered CF (COP)" in mm.columns else pd.Series([0.0])
    min_cum = float(cum.min()) if len(cum) else 0.0
    peak_funding = float(max(0.0, -min_cum)) if np.isfinite(min_cum) else 0.0

    if currency == "COP":
        _metric_row([
            ("Total Investment (CAPEX)", _fmt_cop(cap_total)),
            ("Unlevered IRR (annualized, pre-tax)", f"{irr_annual_pre*100:,.2f}%" if np.isfinite(irr_annual_pre) else "—"),
            ("Unlevered IRR (annualized, after-tax)", f"{irr_annual_after*100:,.2f}%" if np.isfinite(irr_annual_after) else "—"),
            ("Payback (years, after-tax)", f"{payback_years_after:,.2f}" if np.isfinite(payback_years_after) else "—"),
        ])
        _metric_row([
            ("Peak Funding Need", _fmt_cop(peak_funding)),
            ("", ""),
            ("", ""),
            ("", ""),
        ])
    else:
        fx0 = float(s.macro.fx_cop_per_usd_start)
        _metric_row([
            ("Total Investment (CAPEX)", _fmt_usd(cap_total / fx0)),
            ("Unlevered IRR (annualized, pre-tax)", f"{irr_annual_pre*100:,.2f}%" if np.isfinite(irr_annual_pre) else "—"),
            ("Unlevered IRR (annualized, after-tax)", f"{irr_annual_after*100:,.2f}%" if np.isfinite(irr_annual_after) else "—"),
            ("Payback (years, after-tax)", f"{payback_years_after:,.2f}" if np.isfinite(payback_years_after) else "—"),
        ])
        _metric_row([
            ("Peak Funding Need", _fmt_usd(peak_funding / fx0)),
            ("", ""),
            ("", ""),
            ("", ""),
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
    disp = _df_format_money(disp, [c for c in disp.columns if c != "Year"], decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    y_after = "Unlevered CF After Tax (COP)" if currency == "COP" else "Unlevered CF After Tax (USD)"
    fig = px.bar(annual_view, x="Year", y=y_after)
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Monthly cash flow (detailed, pre-tax)")
    m_disp = mm.copy()
    m_money = [c for c in m_disp.columns if c not in ["Month", "Year", "Phase"]]
    m_disp = _df_format_money(m_disp, m_money, decimals=0)
    st.dataframe(m_disp, use_container_width=True, hide_index=True)


# -----------------------------
# K) Debt & Covenants
# -----------------------------
with tab_levered:
    st.subheader("Levered Cash Flow (Equity Cash Flow After-Tax)")
    
    st.info("This tab calculates the levered after-tax cash flow available to equity investors. "
            "If debt is disabled or not enabled, the model assumes no debt and returns unlevered cash flow. "
            "This cash flow will be used to calculate Investor Equity IRR in the summary tab.")
    
    currency = st.radio("Display currency", ["COP", "USD"], horizontal=True, index=0, key="currency_levered")
    
    # Calculate levered cash flows
    annual_levered = levered_cashflow_annual(s)
    monthly_levered = levered_cashflow_monthly(s)
    
    # Check debt status
    debt_enabled = bool(getattr(s.debt, "enabled", False))
    total_capex = _total_capex_from_lines(s)
    debt_amt = (float(s.debt.debt_pct_of_capex) / 100.0) * total_capex if debt_enabled else 0.0
    
    if not debt_enabled or debt_amt <= 0:
        st.warning("⚠️ Debt is not enabled or debt amount is zero. Showing unlevered cash flow (no debt impact).")
    
    # FX conversion helper
    tl = build_timeline(s.timeline)
    years = list(annual_levered["Year"].astype(int).tolist())
    fx = fx_series(s.macro, tl["cod"].year, years)
    
    def _conv(series: pd.Series, year_col: pd.Series) -> pd.Series:
        if currency == "COP":
            return series
        return series / year_col.map(lambda y: float(fx.loc[int(y)]) if int(y) in fx.index else float(s.macro.fx_cop_per_usd_start))
    
    # Key metrics
    total_unlevered_after_tax = float(annual_levered["Unlevered CF After Tax (COP)"].sum())
    total_debt_draws = float(annual_levered["Debt Draw (COP)"].sum())
    total_debt_service = float(annual_levered["Debt Service (COP)"].sum())
    total_debt_fees = float(annual_levered["Debt Fees (COP)"].sum())
    total_levered_cf = float(annual_levered["Levered CF (After-tax, COP)"].sum())
    end_cum_levered = float(annual_levered["Cumulative Levered CF (COP)"].iloc[-1]) if len(annual_levered) else 0.0
    
    if currency == "COP":
        _metric_row([
            ("Total Unlevered CF After-Tax", _fmt_cop(total_unlevered_after_tax)),
            ("Total Debt Draws", _fmt_cop(total_debt_draws)),
            ("Total Debt Service", _fmt_cop(total_debt_service)),
            ("Total Debt Fees", _fmt_cop(total_debt_fees)),
        ])
        _metric_row([
            ("Total Levered CF (Equity)", _fmt_cop(total_levered_cf)),
            ("End Cumulative Levered CF", _fmt_cop(end_cum_levered)),
            ("Debt Status", "Enabled" if debt_enabled and debt_amt > 0 else "No Debt"),
            ("Debt Amount", _fmt_cop(debt_amt) if debt_enabled else "—"),
        ])
    else:
        fx0 = float(s.macro.fx_cop_per_usd_start)
        _metric_row([
            ("Total Unlevered CF After-Tax", _fmt_usd(total_unlevered_after_tax / fx0)),
            ("Total Debt Draws", _fmt_usd(total_debt_draws / fx0)),
            ("Total Debt Service", _fmt_usd(total_debt_service / fx0)),
            ("Total Debt Fees", _fmt_usd(total_debt_fees / fx0)),
        ])
        _metric_row([
            ("Total Levered CF (Equity)", _fmt_usd(total_levered_cf / fx0)),
            ("End Cumulative Levered CF", _fmt_usd(end_cum_levered / fx0)),
            ("Debt Status", "Enabled" if debt_enabled and debt_amt > 0 else "No Debt"),
            ("Debt Amount", _fmt_usd(debt_amt / fx0) if debt_enabled else "—"),
        ])
    
    st.divider()
    
    # Calculate Equity IRR
    monthly_cf = monthly_levered["Levered CF (After-tax, COP)"].astype(float).tolist() if "Levered CF (After-tax, COP)" in monthly_levered.columns else []
    has_pos = any(cf > 0 for cf in monthly_cf)
    has_neg = any(cf < 0 for cf in monthly_cf)
    irr_m = _irr_bisection(monthly_cf) if (has_pos and has_neg) else float("nan")
    irr_annual_equiv = (1.0 + irr_m) ** 12 - 1.0 if np.isfinite(irr_m) else float("nan")
    
    # Payback
    monthly_dates = monthly_levered["Month"].tolist() if "Month" in monthly_levered.columns else []
    payback_m = _payback_months(monthly_dates, monthly_cf) if monthly_cf else float("nan")
    payback_years = payback_m / 12.0 if np.isfinite(payback_m) else float("nan")
    
    # Peak equity funding
    cum = monthly_levered["Cumulative Levered CF (COP)"].astype(float) if "Cumulative Levered CF (COP)" in monthly_levered.columns else pd.Series([0.0])
    min_cum = float(cum.min()) if len(cum) else 0.0
    peak_equity_funding = float(max(0.0, -min_cum)) if np.isfinite(min_cum) else 0.0
    
    _metric_row([
        ("Equity IRR (annualized, after-tax)", f"{irr_annual_equiv*100:,.2f}%" if np.isfinite(irr_annual_equiv) else "—"),
        ("Payback (years, after-tax)", f"{payback_years:,.2f}" if np.isfinite(payback_years) else "—"),
        ("Peak Equity Funding Need", _fmt_cop(peak_equity_funding) if currency == "COP" else _fmt_usd(peak_equity_funding / fx0)),
    ])
    
    st.divider()
    
    # Annual table
    annual_view = annual_levered.copy()
    money_cols = [
        "Unlevered CF After Tax (COP)", "Debt Draw (COP)", "Interest (COP)", "Principal (COP)",
        "Debt Service (COP)", "Debt Fees (COP)", "Levered CF (After-tax, COP)", "Cumulative Levered CF (COP)",
    ]
    for c in money_cols:
        if c in annual_view.columns:
            annual_view[c] = _conv(annual_view[c], annual_view["Year"])
    
    if currency == "USD":
        ren = {c: c.replace("(COP)", "(USD)") for c in annual_view.columns if "(COP)" in c}
        annual_view = annual_view.rename(columns=ren)
    
    st.markdown("### Annual levered cash flow (calendar years)")
    display_cols = [
        "Year",
        "Unlevered CF After Tax (COP)" if currency == "COP" else "Unlevered CF After Tax (USD)",
        "Debt Draw (COP)" if currency == "COP" else "Debt Draw (USD)",
        "Interest (COP)" if currency == "COP" else "Interest (USD)",
        "Principal (COP)" if currency == "COP" else "Principal (USD)",
        "Debt Service (COP)" if currency == "COP" else "Debt Service (USD)",
        "Debt Fees (COP)" if currency == "COP" else "Debt Fees (USD)",
        "Levered CF (After-tax, COP)" if currency == "COP" else "Levered CF (After-tax, USD)",
        "Cumulative Levered CF (COP)" if currency == "COP" else "Cumulative Levered CF (USD)",
    ]
    display_cols = [c for c in display_cols if c in annual_view.columns]
    disp = annual_view[display_cols].copy()
    disp = _df_format_money(disp, [c for c in disp.columns if c != "Year"], decimals=0)
    st.dataframe(disp, use_container_width=True, hide_index=True)
    
    # Charts
    c1, c2 = st.columns(2)
    with c1:
        y_levered = "Levered CF (After-tax, COP)" if currency == "COP" else "Levered CF (After-tax, USD)"
        fig1 = px.bar(annual_view, x="Year", y=y_levered)
        fig1.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), title="Levered CF (After-tax)")
        st.plotly_chart(fig1, use_container_width=True)
    
    with c2:
        y_cum = "Cumulative Levered CF (COP)" if currency == "COP" else "Cumulative Levered CF (USD)"
        fig2 = px.line(annual_view, x="Year", y=y_cum)
        fig2.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), title="Cumulative Levered CF")
        st.plotly_chart(fig2, use_container_width=True)
    
    # Comparison chart: Unlevered vs Levered
    st.markdown("### Unlevered vs Levered Cash Flow Comparison")
    compare_df = annual_view[["Year"]].copy()
    compare_df["Unlevered After-Tax"] = annual_view["Unlevered CF After Tax (COP)" if currency == "COP" else "Unlevered CF After Tax (USD)"]
    compare_df["Levered (Equity)"] = annual_view[y_levered]
    compare_long = compare_df.melt(id_vars=["Year"], value_vars=["Unlevered After-Tax", "Levered (Equity)"], var_name="Type", value_name="Cash Flow")
    fig3 = px.bar(compare_long, x="Year", y="Cash Flow", color="Type", barmode="group")
    fig3.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig3, use_container_width=True)
    
    # Monthly table (optional, collapsed)
    with st.expander("Monthly levered cash flow (detailed)"):
        m_disp = monthly_levered.copy()
        m_money = [c for c in m_disp.columns if c not in ["Month", "Year", "Phase"] and "(COP)" in c]
        for c in m_money:
            if c in m_disp.columns:
                m_disp[c] = _conv(m_disp[c], m_disp["Year"])
                if currency == "USD":
                    new_name = c.replace("(COP)", "(USD)")
                    m_disp = m_disp.rename(columns={c: new_name})
        
        m_disp = _df_format_money(m_disp, [c for c in m_disp.columns if c not in ["Month", "Year", "Phase"]], decimals=0)
        st.dataframe(m_disp, use_container_width=True, hide_index=True)


# -----------------------------
# M) Compare
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
            total_cap = _total_capex_from_lines(sd)

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
