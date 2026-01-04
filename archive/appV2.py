# Delphi Utility-Scale Financial Model (No Excel)
# Streamlit single-file app with: Projects + Scenarios, Macro, Timeline, Generation, CAPEX, Revenues, Comparison
# All inputs in COP; outputs selectable COP/USD (USD via FX path).
#
# Run:
#   py -m pip install -r requirements.txt
#   py -m streamlit run app.py

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict, field
from datetime import date
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------
# Storage (SAVE EVERYTHING IN YOUR DRIVE FOLDER)
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
# Model Structures
# -----------------------------
INDEX_CHOICES = ["Colombia CPI", "Colombia PPI", "US CPI", "Custom"]

CAPEX_PHASES = ["Development", "Construction", "At COD"]
CAPEX_DISTRIBUTIONS = ["Straight-line (monthly)", "Front-loaded", "Back-loaded"]


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
    start_date: str = ""  # ISO
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
    degradation_pct: float = 0.5


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
    # User can add rows in the UI (dynamic table).
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


@dataclass
class ScenarioInputs:
    name: str = "Base"
    macro: MacroInputs = field(default_factory=MacroInputs)
    timeline: TimelineInputs = field(default_factory=TimelineInputs)
    generation: GenerationInputs = field(default_factory=GenerationInputs)
    capex: CapexInputs = field(default_factory=CapexInputs)
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
    revenue1 = RevenueOption1PPA(**d.get("revenue1", {}))
    revenue2 = RevenueOption2Manual(**d.get("revenue2", {}))
    return ScenarioInputs(
        name=d.get("name", "Base"),
        macro=macro,
        timeline=timeline,
        generation=generation,
        capex=capex,
        revenue_mode=d.get("revenue_mode", "Standard PPA Parameters"),
        revenue1=revenue1,
        revenue2=revenue2,
    )


# -----------------------------
# Core Calculations
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


def _weights(n: int, mode: str) -> np.ndarray:
    if n <= 0:
        return np.array([])
    if mode == "Straight-line (monthly)":
        w = np.ones(n)
    elif mode == "Front-loaded":
        w = np.linspace(n, 1, n)  # heavier earlier
    else:  # Back-loaded
        w = np.linspace(1, n, n)  # heavier later
    return w / w.sum()


def capex_monthly_schedule(s: ScenarioInputs) -> pd.DataFrame:
    """
    CAPEX cash outflows:
      - Development: spread from start -> RTB over dev_months
      - Construction: spread from RTB -> COD over capex_months
      - At COD: booked once at COD
    COP only (no escalation in V1).
    """
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
    rows.append({"Month": cod, "Phase": "At COD", "CAPEX (COP)": cod_total})

    df = pd.DataFrame(rows).sort_values("Month").reset_index(drop=True)
    df["Year"] = df["Month"].apply(lambda d: d.year)
    return df


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
        price_nom_base = [p0 if (i < term) else pm for i in range(op_years)]
    else:
        r = s.revenue2
        index_choice = r.indexation
        price_nom_base = [
            float(r.prices_constant_cop_per_kwh.get(i + 1, list(r.prices_constant_cop_per_kwh.values())[-1]))
            for i in range(op_years)
        ]

    idx_key = {"Colombia CPI": "col_cpi", "Colombia PPI": "col_ppi", "US CPI": "us_cpi", "Custom": "custom"}[index_choice]
    idx = annual_index_series(s.macro, base_year, years, idx_key)

    price_cop_per_kwh = np.array(price_nom_base) * idx.values
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
            "Index Level": idx.values,
        }
    )


# -----------------------------
# UI Helpers
# -----------------------------
def _metric_row(cols):
    c = st.columns(len(cols))
    for i, (label, val) in enumerate(cols):
        c[i].metric(label, val)


def _fmt_cop(x: float) -> str:
    return f"COP {x:,.0f}"


def _fmt_usd(x: float) -> str:
    return f"USD {x:,.0f}"


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


# Load scenario into session state
key = f"{project_name}::{scenario_name}"
if "loaded_key" not in st.session_state or st.session_state.loaded_key != key:
    s0 = _scenario_from_dict(proj["scenarios"][scenario_name])
    st.session_state.loaded_key = key
    st.session_state.scenario = s0

s: ScenarioInputs = st.session_state.scenario

tab_macro, tab_timeline, tab_gen, tab_capex, tab_rev, tab_compare = st.tabs(
    ["A) Macroeconomic", "B) Timeline", "C) Power Generation", "D) CAPEX", "E) Power Revenues", "Compare"]
)

# -----------------------------
# A) Macro
# -----------------------------
with tab_macro:
    st.subheader("Macroeconomic inputs (annual rates, %)")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        s.macro.col_cpi = st.number_input("Colombian CPI (%)", value=float(s.macro.col_cpi), step=0.1)
    with c2:
        s.macro.col_ppi = st.number_input("Colombian PPI (%)", value=float(s.macro.col_ppi), step=0.1)
    with c3:
        s.macro.us_cpi = st.number_input("US CPI (%)", value=float(s.macro.us_cpi), step=0.1)
    with c4:
        s.macro.custom_index_rate = st.number_input("Custom index (%)", value=float(s.macro.custom_index_rate), step=0.1)

    st.markdown("#### FX (COP per USD)")
    fx1, fx2, fx3 = st.columns([1.2, 1, 1])
    with fx1:
        s.macro.fx_cop_per_usd_start = st.number_input("Starting FX (COP/USD)", value=float(s.macro.fx_cop_per_usd_start), step=10.0)
    with fx2:
        s.macro.fx_method = st.selectbox("FX method", ["Inflation differential (PPP approx.)", "Flat"],
                                         index=0 if s.macro.fx_method != "Flat" else 1)
    with fx3:
        s.macro.fx_flat = st.number_input("Flat FX (if selected)", value=float(s.macro.fx_flat or s.macro.fx_cop_per_usd_start), step=10.0)

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
        s.timeline.dev_months = int(st.number_input("Development (months)", value=int(s.timeline.dev_months), step=1))
    with t3:
        s.timeline.capex_months = int(st.number_input("CAPEX / Construction (months)", value=int(s.timeline.capex_months), step=1))
    with t4:
        s.timeline.operation_years = int(st.number_input("Operation (years)", value=int(s.timeline.operation_years), step=1))

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

    st.caption("RTB marks the end of Development and start of construction CAPEX. COD marks the start of Operation and revenue generation.")


# -----------------------------
# C) Generation
# -----------------------------
with tab_gen:
    st.subheader("Power generation")
    g1, g2, g3, g4 = st.columns(4)
    with g1:
        s.generation.mwac = st.number_input("Capacity (MWac)", value=float(s.generation.mwac), step=0.1)
    with g2:
        s.generation.mwp = st.number_input("Capacity (MWp)", value=float(s.generation.mwp), step=0.1)
    with g3:
        s.generation.degradation_pct = st.number_input("Annual degradation (%)", value=float(s.generation.degradation_pct), step=0.05)
    with g4:
        s.generation.production_choice = st.selectbox("Production choice", ["P50", "P75", "P90"],
                                                      index=["P50", "P75", "P90"].index(s.generation.production_choice))

    p1, p2, p3 = st.columns(3)
    with p1:
        s.generation.p50_mwh_yr = st.number_input("P50 production (MWh/yr)", value=float(s.generation.p50_mwh_yr), step=1000.0)
    with p2:
        s.generation.p75_mwh_yr = st.number_input("P75 production (MWh/yr)", value=float(s.generation.p75_mwh_yr), step=1000.0)
    with p3:
        s.generation.p90_mwh_yr = st.number_input("P90 production (MWh/yr)", value=float(s.generation.p90_mwh_yr), step=1000.0)

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
# D) CAPEX
# -----------------------------
with tab_capex:
    st.subheader("CAPEX (COP) — line items, schedule, and unit economics")

    s.capex.distribution = st.selectbox(
        "Distribution for Development/Construction phases",
        CAPEX_DISTRIBUTIONS,
        index=CAPEX_DISTRIBUTIONS.index(s.capex.distribution) if s.capex.distribution in CAPEX_DISTRIBUTIONS else 0,
        help="Controls how Development and Construction CAPEX are spread monthly within their timeline windows.",
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
        num_rows="dynamic",  # <-- THIS is what lets user add extra CAPEX lines
        column_config={
            "Item": st.column_config.TextColumn("CAPEX line item"),
            "Amount_COP": st.column_config.NumberColumn("Amount (COP)", min_value=0.0, step=1_000_000.0, format="%.0f"),
            "Phase": st.column_config.SelectboxColumn("Phase", options=CAPEX_PHASES),
        },
    )

    # Persist back into scenario
    s.capex.lines = edited.to_dict(orient="records")

    # Totals + unit CAPEX
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
    st.dataframe(sched[["Month", "Phase", "CAPEX (COP)"]], use_container_width=True, hide_index=True)

    fig = px.bar(sched, x="Month", y="CAPEX (COP)", color="Phase")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Annual CAPEX (calendar years)")
    ann = sched.groupby("Year", as_index=False)["CAPEX (COP)"].sum()
    st.dataframe(ann, use_container_width=True, hide_index=True)
    fig2 = px.bar(ann, x="Year", y="CAPEX (COP)")
    fig2.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig2, use_container_width=True)

    st.caption("CAPEX is in COP and not escalated in V1. Next step can add inflation escalation and VAT/tax handling if you want.")


# -----------------------------
# E) Revenues
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
            r.ppa_price_cop_per_kwh = st.number_input("PPA price at COD (COP/kWh)", value=float(r.ppa_price_cop_per_kwh), step=1.0)
        with c2:
            r.ppa_term_years = int(st.number_input("PPA term (years)", value=int(r.ppa_term_years), step=1))
        with c3:
            r.merchant_price_cop_per_kwh = st.number_input("Post-term / merchant (COP/kWh)", value=float(r.merchant_price_cop_per_kwh), step=1.0)
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
        df_show["Revenue (COP)"] = df_show["Revenue (COP)"].round(0)
        st.dataframe(df_show, use_container_width=True, hide_index=True)

        st.markdown("#### Annual energy and revenue (COP)")
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
        df_show["Revenue (USD)"] = df_show["Revenue (USD)"].round(0)
        st.dataframe(df_show, use_container_width=True, hide_index=True)

        st.markdown("#### Annual energy and revenue (USD)")
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
# Compare
# -----------------------------
with tab_compare:
    st.subheader("Scenario comparison (revenues + CAPEX)")
    if not compare_scenarios:
        st.warning("Select scenarios to compare in the sidebar.")
    else:
        rows = []
        chart_rows = []
        for nm in compare_scenarios:
            sd = _scenario_from_dict(proj["scenarios"][nm])
            dfx = operating_year_table(sd)

            cap_df = pd.DataFrame(sd.capex.lines or [])
            total_cap = float(cap_df["Amount_COP"].fillna(0).sum()) if (not cap_df.empty and "Amount_COP" in cap_df.columns) else 0.0
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
                    "Total Revenue (COP)": float(dfx["Revenue (COP)"].sum()),
                    "Total Revenue (USD)": float(dfx["Revenue (USD)"].sum()),
                }
            )

            df2 = dfx[["Year", "Revenue (COP)", "Revenue (USD)"]].copy()
            df2["Scenario"] = nm
            chart_rows.append(df2)

        summary = pd.DataFrame(rows).sort_values("Scenario")
        st.dataframe(summary, use_container_width=True, hide_index=True)

        chart = pd.concat(chart_rows, ignore_index=True)
        cur = st.selectbox("Comparison chart currency", ["COP", "USD"], index=0)
        fig = px.line(chart, x="Year", y=("Revenue (COP)" if cur == "COP" else "Revenue (USD)"), color="Scenario")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)


# Persist scenario on each run
proj["scenarios"][scenario_name] = _scenario_to_dict(s)
_save_db(db)
