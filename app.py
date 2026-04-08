"""
InfraFund Financial Model - Interactive Investor Dashboard
==========================================================
A narrative-driven web application for non-technical investors.

Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from typing import Dict

from config import THEME, ASSUMPTION_TOOLTIPS, ModelAssumptions, get_default_assumptions
from calculations import (
    compute_income_statement, compute_valuation, compute_valuation_breakdown,
    compute_cap_table, compute_community_impact, apply_scenario,
)
from charts import (
    create_revenue_waterfall, create_scenario_comparison, create_cap_table_pie,
    create_runway_chart, create_community_impact_chart, create_energy_savings_chart,
)

st.set_page_config(
    page_title="InfraFund Financial Model",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
    html, body, [class*="st-"] {{ font-family: {THEME['font_family']}; }}
    [data-testid="stSidebar"] {{ background-color: {THEME['secondary_color']}; }}
    [data-testid="stSidebar"] * {{ color: white !important; }}
    [data-testid="stSidebar"] .stSlider label,
    [data-testid="stSidebar"] .stNumberInput label {{
        margin-bottom: 8px !important;
        display: block !important;
    }}
    [data-testid="stSidebar"] .stSlider,
    [data-testid="stSidebar"] .stNumberInput {{ margin-bottom: 20px !important; }}
    [data-testid="stMetric"] {{
        background-color: #f8f9fa;
        padding: 16px;
        border-radius: 8px;
        border-left: 4px solid {THEME['primary_color']};
    }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 24px; }}
    .stTabs [data-baseweb="tab"] {{ padding: 12px 24px; font-weight: 600; }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
</style>
""", unsafe_allow_html=True)


def init_session_state():
    if "assumptions" not in st.session_state:
        st.session_state.assumptions = get_default_assumptions()
    if "scenario" not in st.session_state:
        st.session_state.scenario = "Base"

init_session_state()


def render_sidebar() -> Dict:
    with st.sidebar:
        st.image("https://via.placeholder.com/200x60?text=InfraFund", width=200)
        st.markdown("---")

        st.subheader("🎯 Scenario")
        scenario = st.radio("Select forecast scenario:", options=["Worst", "Base", "Best"],
                            index=1, horizontal=True, help=ASSUMPTION_TOOLTIPS["scenario"])
        st.session_state.scenario = scenario
        st.markdown("---")

        st.subheader("💰 Fee Structure")
        with st.expander("ℹ️ Why these rates?", expanded=False):
            st.markdown(ASSUMPTION_TOOLTIPS["upfront_fee_rate"])
        st.markdown("<br>", unsafe_allow_html=True)
        upfront_fee = st.slider("Tokenisation Fee Rate (%)", min_value=0.5, max_value=5.0,
                                value=st.session_state.assumptions.upfront_fee_rate * 100, step=0.1, format="%.1f%%") / 100
        st.markdown("<br>", unsafe_allow_html=True)
        kwh_rate = st.slider("Assurance (kWh) Revenue Rate (%)", min_value=1.0, max_value=4.0,
                             value=st.session_state.assumptions.kwh_revenue_rate * 100, step=0.1, format="%.2f%%") / 100
        st.markdown("---")

        st.subheader("📈 Growth")
        st.markdown("<br>", unsafe_allow_html=True)
        year_1_funds = st.number_input("Year 1 Funds Raised (£M)", min_value=5.0, max_value=50.0,
                                       value=st.session_state.assumptions.base_funds_raised[0] / 1_000_000, step=1.0) * 1_000_000
        st.markdown("<br>", unsafe_allow_html=True)
        growth_rate = st.slider("Annual Growth Rate (Y1→Y5)", min_value=1.5, max_value=4.0, value=2.0, step=0.1, format="%.1fx")
        st.markdown("---")

        st.subheader("🧮 Valuation")
        with st.expander("ℹ️ Why this multiple?", expanded=False):
            st.markdown(ASSUMPTION_TOOLTIPS["ev_revenue_multiple"])
        st.markdown("<br>", unsafe_allow_html=True)
        ev_multiple = st.slider("Exit EV/Revenue Multiple", min_value=3.0, max_value=12.0,
                                value=st.session_state.assumptions.ev_revenue_multiple, step=0.5, format="%.1fx")

        st.session_state.assumptions.upfront_fee_rate = upfront_fee
        st.session_state.assumptions.kwh_revenue_rate = kwh_rate
        st.session_state.assumptions.ev_revenue_multiple = ev_multiple
        st.session_state.assumptions.base_funds_raised = [year_1_funds * (growth_rate ** i) for i in range(5)]

    return st.session_state.assumptions


def render_story_tab(results, valuation):
    st.markdown("""
    ## ⚡ The Operating System for Renewable Infrastructure
    
    **InfraFund** automates *Social License* — turning local opposition into community support
    via AI, tokenisation, and trustless smart-contract escrow. We deliver **25–50% energy bill
    savings** to households while removing the #1 bottleneck to grid expansion: NIMBYism.
    """)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Equity Value Today", f"£{valuation.equity_value_today / 1_000_000:.1f}M",
                  help="Present value of future exit, discounted at VC rates")
    with col2:
        st.metric("Year 5 Exit Valuation", f"£{valuation.exit_valuation / 1_000_000:.1f}M",
                  help=f"Based on {st.session_state.assumptions.ev_revenue_multiple}x revenue multiple")
    with col3:
        st.metric("Investor IRR", f"{valuation.investor_irr:.0%}",
                  delta="Strong" if valuation.investor_irr > 0.3 else "Moderate")
    with col4:
        st.metric("Cash-on-Cash", f"{valuation.cash_on_cash:.1f}x", help="Exit proceeds ÷ initial investment")

    st.markdown("---")
    st.subheader("📊 5-Year Financial Snapshot")
    df = results.to_dataframe()
    df_display = df[["Year", "Gross Revenue", "EBIT", "Net Income"]].copy()
    df_display["Gross Revenue"] = df_display["Gross Revenue"].apply(lambda x: f"£{x/1e6:.1f}M")
    df_display["EBIT"] = df_display["EBIT"].apply(lambda x: f"£{x/1e6:.1f}M")
    df_display["Net Income"] = df_display["Net Income"].apply(lambda x: f"£{x/1e6:.1f}M")
    df_transposed = df_display.set_index("Year").T
    st.dataframe(df_transposed, use_container_width=True)

    st.markdown("---")
    render_community_impact_section()


def render_community_impact_section():
    st.subheader("🌍 Community Impact")
    st.markdown("InfraFund translates capital deployment directly into household savings and clean energy production.")
    fum = st.slider("Platform Scale — Funds Under Management", min_value=10_000_000, max_value=1_000_000_000,
                    value=100_000_000, step=10_000_000, format="£%.0f", key="community_fum_slider")
    impact = compute_community_impact(fum)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Households Empowered", f"{impact.households_reached:,}")
    with c2:
        st.metric("Annual Savings / Household", f"£{impact.avg_annual_savings_low:,.0f}–£{impact.avg_annual_savings_high:,.0f}")
    with c3:
        st.metric("Total Community Savings", f"up to £{impact.total_savings_high/1e6:,.1f}M")
    with c4:
        st.metric("CO2 Avoided (tonnes/yr)", f"{impact.co2_avoided_tonnes:,.0f}")

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(create_community_impact_chart(impact), use_container_width=True)
    with col2:
        st.plotly_chart(create_energy_savings_chart(impact), use_container_width=True)


def render_unit_economics_tab(results):
    st.subheader("💵 Revenue Streams")
    col1, col2 = st.columns(2)
    with col1:
        fig = create_revenue_waterfall(results.years, results.upfront_fee, results.completion_fee, results.kwh_revenue)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown("""
        ### How We Make Money
        **1. Tokenisation Fee (2%)** — One-time fee when projects raise capital
        **2. Verification Fee (0.5%)** — Recognized over 3-year build period
        **3. Assurance Fee (1p/kWh)** — Recurring revenue from energy production *(starts Year 4)*
        """)
        y5_total = results.gross_revenue[4]
        if y5_total > 0:
            mix_data = {"Stream": ["Tokenisation", "Verification", "Assurance"],
                        "Amount": [results.upfront_fee[4], results.completion_fee[4], results.kwh_revenue[4]]}
            st.dataframe(pd.DataFrame(mix_data), use_container_width=True, hide_index=True)


def render_growth_tab(results, assumptions):
    st.subheader("📈 Growth & Projections")
    worst = apply_scenario(assumptions.base_funds_raised, "Worst")
    base = apply_scenario(assumptions.base_funds_raised, "Base")
    best = apply_scenario(assumptions.base_funds_raised, "Best")
    col1, col2 = st.columns(2)
    with col1:
        fig = create_scenario_comparison(
            assumptions.years,
            [w * assumptions.upfront_fee_rate for w in worst],
            [b * assumptions.upfront_fee_rate for b in base],
            [bt * assumptions.upfront_fee_rate for bt in best],
        )
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown(f"### Current Scenario: **{st.session_state.scenario} Case**\n| Year | Funds Raised |")
        st.markdown("|------|--------------|")
        for yr, fr in zip(assumptions.years, results.funds_raised):
            st.markdown(f"| {yr} | £{fr/1e6:.0f}M |")


def render_valuation_tab(valuation, assumptions, results):
    st.subheader("🧮 Valuation & Investor Returns")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("### VC Valuation Method\nWe use a **tiered discount rate** approach common in venture capital:")
        rate_df = pd.DataFrame({
            "Period": ["Years 1–3", "Years 4–6", "Years 7–10"],
            "Discount Rate": ["50%", "35%", "25%"],
            "Risk Profile": ["Early stage", "Growth stage", "Mature"],
        })
        st.dataframe(rate_df, use_container_width=True, hide_index=True)
        with st.expander("ℹ️ Why these rates?"):
            st.markdown(ASSUMPTION_TOOLTIPS["required_return"])
    with col2:
        st.markdown("### Key Outputs")
        st.metric("Exit EV/Revenue", f"{assumptions.ev_revenue_multiple}x")
        st.metric("Exit Proceeds (£250K)", f"£{valuation.exit_proceeds / 1000:.0f}K")

    st.markdown("---")
    st.subheader("📐 Valuation Breakdown")
    breakdown = compute_valuation_breakdown(
        results.gross_revenue, assumptions.ev_revenue_multiple,
        (assumptions.required_return_y1_3, assumptions.required_return_y3_6, assumptions.required_return_y6_10),
    )
    st.markdown(f"""
    **Step 1: Calculate Exit Enterprise Value**
    - Year {breakdown.exit_year} Revenue: £{breakdown.exit_year_revenue/1e6:.2f}M
    - EV Multiple: {breakdown.ev_multiple}x
    - **Exit EV = £{breakdown.exit_ev/1e6:.2f}M**
    
    **Step 2: Discount to Present Value**
    - Discount Factor: {breakdown.discount_factor:.2f}
    - **Equity Value Today = £{breakdown.exit_ev/1e6:.2f}M ÷ {breakdown.discount_factor:.2f} = £{breakdown.equity_value_today/1e6:.2f}M**
    
    **Step 3: Investor Returns**
    - Entry Stake: {breakdown.entry_stake:.2%} → Post-dilution: {breakdown.stake_at_exit:.2%}
    - Exit Proceeds: £{breakdown.exit_proceeds/1e6:.2f}M
    - **Cash-on-Cash: {breakdown.cash_on_cash:.1f}x | IRR: {breakdown.investor_irr:.0%}**
    """)


def render_cap_table_tab(assumptions):
    st.subheader("📊 Capitalization Table")
    cap_df = compute_cap_table(assumptions.founder_shares, assumptions.funding_rounds)
    col1, col2 = st.columns(2)
    with col1:
        fig = create_cap_table_pie(cap_df)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown("### Ownership Summary")
        cap_display = cap_df.copy()
        cap_display["Shares"] = cap_display["Shares"].apply(lambda x: f"{x:,.0f}")
        cap_display["Ownership %"] = cap_display["Ownership %"].apply(lambda x: f"{x:.1%}")
        st.dataframe(cap_display, use_container_width=True, hide_index=True)


def render_assumptions_tab(assumptions):
    st.subheader("🎯 Defensible Assumptions Dashboard")
    st.markdown("Every assumption in this model is grounded in market data and comparable benchmarks.")

    st.markdown("### 💰 Revenue Pillars")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**2% Tokenisation Fee**")
        st.slider("Tokenisation Fee (%)", 1.0, 4.0, 2.0, 0.1, key="tok_fee_demo", disabled=True)
        with st.expander("View Methodology & Proof"):
            st.markdown(ASSUMPTION_TOOLTIPS["tokenisation_fee"])
    with col2:
        st.markdown("**0.5% Verification Fee**")
        st.slider("Verification Fee (%)", 0.25, 1.0, 0.5, 0.05, key="ver_fee_demo", disabled=True)
        with st.expander("View Methodology & Proof"):
            st.markdown(ASSUMPTION_TOOLTIPS["verification_fee"])
    with col3:
        st.markdown("**1p/kWh Assurance Fee**")
        st.slider("Assurance Fee (p/kWh)", 0.5, 2.0, 1.0, 0.1, key="ass_fee_demo", disabled=True)
        with st.expander("View Methodology & Proof"):
            st.markdown(ASSUMPTION_TOOLTIPS["assurance_fee"])

    st.markdown("---")
    st.markdown("### 📊 Operational Targets")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**£18M Year 1 Target**")
        st.number_input("Year 1 FUM (£M)", value=18.0, disabled=True, key="y1_target_demo")
        with st.expander("View Methodology & Proof"):
            st.markdown(ASSUMPTION_TOOLTIPS["year1_target"])
    with col2:
        st.markdown("**Year 1 Operating Expenses**")
        y1_opex = assumptions.opex_y1_salaries + assumptions.opex_y1_marketing + assumptions.opex_y1_legal + assumptions.opex_y1_insurance + assumptions.opex_y1_rd
        st.metric("Total Y1 OpEx", f"£{y1_opex/1000:,.0f}K")
        with st.expander("View Breakdown"):
            st.markdown(ASSUMPTION_TOOLTIPS["opex_y1"])

    st.markdown("---")
    st.markdown("### 🔐 Smart Contract Audit Costs")
    audit_df = pd.DataFrame({
        "Year": assumptions.years,
        "Audit Cost (£)": [f"£{c:,.0f}" for c in assumptions.smart_contract_costs]
    })
    st.dataframe(audit_df, use_container_width=True, hide_index=True)
    with st.expander("View Audit Provider Benchmarks"):
        st.markdown(ASSUMPTION_TOOLTIPS["smart_contract_costs"])


def main():
    assumptions = render_sidebar()
    results = compute_income_statement(st.session_state.scenario, assumptions)
    valuation = compute_valuation(
        results.gross_revenue, assumptions.ev_revenue_multiple,
        (assumptions.required_return_y1_3, assumptions.required_return_y3_6, assumptions.required_return_y6_10),
    )

    st.title("⚡ InfraFund Financial Model")
    st.markdown(f"*Scenario: **{st.session_state.scenario} Case** | Last updated: April 2026*")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🏠 Our Story", "💵 Unit Economics", "📈 Growth", "🧮 Valuation", "📊 Cap Table", "🎯 Assumptions"
    ])

    with tab1:
        render_story_tab(results, valuation)
    with tab2:
        render_unit_economics_tab(results)
    with tab3:
        render_growth_tab(results, assumptions)
    with tab4:
        render_valuation_tab(valuation, assumptions, results)
    with tab5:
        render_cap_table_tab(assumptions)
    with tab6:
        render_assumptions_tab(assumptions)


if __name__ == "__main__":
    main()
