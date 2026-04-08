"""
Joule Financial Model — Interactive Investor Dashboard
========================================================
A narrative-driven web application for non-technical investors.

Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from typing import Dict

# Local imports
from config import (
    THEME, 
    ASSUMPTION_TOOLTIPS, 
    ModelAssumptions, 
    get_default_assumptions,
)
from calculations import (
    compute_income_statement,
    compute_valuation,
    compute_cap_table,
    apply_scenario,
)
from charts import (
    create_revenue_waterfall,
    create_scenario_comparison,
    create_cap_table_pie,
    create_runway_chart,
)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Joule Financial Model",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for premium aesthetic
st.markdown(f"""
<style>
    /* Global font */
    html, body, [class*="st-"] {{
        font-family: {THEME['font_family']};
    }}
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {{
        background-color: {THEME['secondary_color']};
    }}
    [data-testid="stSidebar"] * {{
        color: white !important;
    }}
    
    /* Metric cards */
    [data-testid="stMetric"] {{
        background-color: #f8f9fa;
        padding: 16px;
        border-radius: 8px;
        border-left: 4px solid {THEME['primary_color']};
    }}
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 24px;
    }}
    .stTabs [data-baseweb="tab"] {{
        padding: 12px 24px;
        font-weight: 600;
    }}
    
    /* Hide Streamlit branding */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

def init_session_state():
    """Initialize session state with default assumptions."""
    if "assumptions" not in st.session_state:
        st.session_state.assumptions = get_default_assumptions()
    if "scenario" not in st.session_state:
        st.session_state.scenario = "Base"

init_session_state()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — INPUT CONTROLS
# ══════════════════════════════════════════════════════════════════════════════

def render_sidebar() -> Dict:
    """Render sidebar controls and return updated assumptions."""
    
    with st.sidebar:
        st.image("https://via.placeholder.com/200x60?text=JOULE", width=200)  # Replace with logo
        st.markdown("---")
        
        # ── Scenario Selector ────────────────────────────────────────────────
        st.subheader("🎯 Scenario")
        scenario = st.radio(
            "Select forecast scenario:",
            options=["Worst", "Base", "Best"],
            index=1,  # Default to Base
            horizontal=True,
            help=ASSUMPTION_TOOLTIPS["scenario"],
        )
        st.session_state.scenario = scenario
        
        st.markdown("---")
        
        # ── Fee Structure ────────────────────────────────────────────────────
        st.subheader("💰 Fee Structure")
        
        with st.expander("ℹ️ Why these rates?", expanded=False):
            st.markdown(ASSUMPTION_TOOLTIPS["upfront_fee_rate"])
        
        upfront_fee = st.slider(
            "Upfront Fee Rate (%)",
            min_value=0.5, max_value=5.0, 
            value=st.session_state.assumptions.upfront_fee_rate * 100,
            step=0.1,
            format="%.1f%%",
        ) / 100
        
        kwh_rate = st.slider(
            "kWh Revenue Rate (%)",
            min_value=1.0, max_value=4.0,
            value=st.session_state.assumptions.kwh_revenue_rate * 100,
            step=0.1,
            format="%.2f%%",
        ) / 100
        
        st.markdown("---")
        
        # ── Growth Assumptions ───────────────────────────────────────────────
        st.subheader("📈 Growth")
        
        year_1_funds = st.number_input(
            "Year 1 Funds Raised (£M)",
            min_value=5.0, max_value=50.0,
            value=st.session_state.assumptions.base_funds_raised[0] / 1_000_000,
            step=1.0,
        ) * 1_000_000
        
        growth_rate = st.slider(
            "Annual Growth Rate (Y1→Y5)",
            min_value=1.5, max_value=4.0,
            value=2.0,  # Approximate from base case
            step=0.1,
            format="%.1fx",
        )
        
        st.markdown("---")
        
        # ── Valuation ────────────────────────────────────────────────────────
        st.subheader("🧮 Valuation")
        
        with st.expander("ℹ️ Why this multiple?", expanded=False):
            st.markdown(ASSUMPTION_TOOLTIPS["ev_revenue_multiple"])
        
        ev_multiple = st.slider(
            "Exit EV/Revenue Multiple",
            min_value=3.0, max_value=12.0,
            value=st.session_state.assumptions.ev_revenue_multiple,
            step=0.5,
            format="%.1fx",
        )
        
        # Update session state
        st.session_state.assumptions.upfront_fee_rate = upfront_fee
        st.session_state.assumptions.kwh_revenue_rate = kwh_rate
        st.session_state.assumptions.ev_revenue_multiple = ev_multiple
        # Recalculate funds raised based on Y1 and growth rate
        st.session_state.assumptions.base_funds_raised = [
            year_1_funds * (growth_rate ** i) for i in range(5)
        ]
    
    return st.session_state.assumptions


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT — TABS
# ══════════════════════════════════════════════════════════════════════════════

def render_story_tab(results, valuation):
    """Tab 1: Company narrative and headline metrics."""
    
    st.markdown("""
    ## ⚡ Powering the Green Energy Transition
    
    Joule is a **green energy infrastructure financing platform** that connects investors 
    with renewable energy projects. We earn fees from capital raised and ongoing kWh 
    production monitoring.
    """)
    
    # Headline metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Equity Value Today",
            f"£{valuation.equity_value_today / 1_000_000:.1f}M",
            help="Present value of future exit, discounted at VC rates",
        )
    
    with col2:
        st.metric(
            "Year 5 Exit Valuation",
            f"£{valuation.exit_valuation / 1_000_000:.1f}M",
            help=f"Based on {st.session_state.assumptions.ev_revenue_multiple}x revenue multiple",
        )
    
    with col3:
        st.metric(
            "Investor IRR",
            f"{valuation.investor_irr:.0%}",
            delta="Strong" if valuation.investor_irr > 0.3 else "Moderate",
        )
    
    with col4:
        st.metric(
            "Cash-on-Cash",
            f"{valuation.cash_on_cash:.1f}x",
            help="Exit proceeds ÷ initial investment",
        )
    
    st.markdown("---")
    
    # Quick P&L snapshot
    st.subheader("📊 5-Year Financial Snapshot")
    df = results.to_dataframe()
    df_display = df[["Year", "Gross Revenue", "EBIT", "Net Income"]].copy()
    df_display["Gross Revenue"] = df_display["Gross Revenue"].apply(lambda x: f"£{x/1e6:.1f}M")
    df_display["EBIT"] = df_display["EBIT"].apply(lambda x: f"£{x/1e6:.1f}M")
    df_display["Net Income"] = df_display["Net Income"].apply(lambda x: f"£{x/1e6:.1f}M")
    st.dataframe(df_display, use_container_width=True, hide_index=True)


def render_unit_economics_tab(results):
    """Tab 2: Revenue breakdown and unit economics."""
    
    st.subheader("💵 Revenue Streams")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = create_revenue_waterfall(
            results.years,
            results.upfront_fee,
            results.completion_fee,
            results.kwh_revenue,
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("""
        ### How We Make Money
        
        **1. Upfront Fee (2%)**  
        One-time fee when projects raise capital
        
        **2. Completion Fee (0.5%)**  
        Recognized over 3-year build period
        
        **3. kWh Revenue (2.12%)**  
        Recurring revenue from energy production  
        *(starts Year 4 after project completion)*
        """)
        
        # Revenue mix pie for Year 5
        y5_total = results.gross_revenue[4]
        if y5_total > 0:
            mix_data = {
                "Stream": ["Upfront", "Completion", "kWh"],
                "Amount": [results.upfront_fee[4], results.completion_fee[4], results.kwh_revenue[4]],
            }
            st.dataframe(pd.DataFrame(mix_data), use_container_width=True, hide_index=True)


def render_growth_tab(results, assumptions):
    """Tab 3: Growth projections and scenario analysis."""
    
    st.subheader("📈 Growth & Projections")
    
    # Compute all three scenarios for comparison
    worst = apply_scenario(assumptions.base_funds_raised, "Worst")
    base = apply_scenario(assumptions.base_funds_raised, "Base")
    best = apply_scenario(assumptions.base_funds_raised, "Best")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = create_scenario_comparison(
            assumptions.years,
            [w * assumptions.upfront_fee_rate for w in worst],  # Simplified: just upfront
            [b * assumptions.upfront_fee_rate for b in base],
            [bt * assumptions.upfront_fee_rate for bt in best],
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown(f"""
        ### Current Scenario: **{st.session_state.scenario} Case**
        
        | Year | Funds Raised |
        |------|--------------|
        """)
        for i, (yr, fr) in enumerate(zip(assumptions.years, results.funds_raised)):
            st.markdown(f"| {yr} | £{fr/1e6:.0f}M |")


def render_valuation_tab(valuation, assumptions):
    """Tab 4: Valuation deep dive."""
    
    st.subheader("🧮 Valuation & Investor Returns")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### VC Valuation Method
        
        We use a **tiered discount rate** approach common in venture capital:
        """)
        
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


def render_cap_table_tab(assumptions):
    """Tab 5: Cap table visualization."""
    
    st.subheader("📊 Capitalization Table")
    
    cap_df = compute_cap_table(
        assumptions.founder_shares,
        assumptions.funding_rounds,
    )
    
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


# ══════════════════════════════════════════════════════════════════════════════
# MAIN EXECUTION
# ══════════════════════════════════════════════════════════════════════════════

def main():
    """Main application entry point."""
    
    # Render sidebar and get updated assumptions
    assumptions = render_sidebar()
    
    # Compute model results (cached when inputs unchanged)
    results = compute_income_statement(
        st.session_state.scenario,
        assumptions,
    )
    
    valuation = compute_valuation(
        results.gross_revenue,
        assumptions.ev_revenue_multiple,
        (
            assumptions.required_return_y1_3,
            assumptions.required_return_y3_6,
            assumptions.required_return_y6_10,
        ),
    )
    
    # Header
    st.title("⚡ Joule Financial Model")
    st.markdown(f"*Scenario: **{st.session_state.scenario} Case** | Last updated: April 2026*")
    
    # Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Our Story",
        "💵 Unit Economics", 
        "📈 Growth",
        "🧮 Valuation",
        "📊 Cap Table",
    ])
    
    with tab1:
        render_story_tab(results, valuation)
    
    with tab2:
        render_unit_economics_tab(results)
    
    with tab3:
        render_growth_tab(results, assumptions)
    
    with tab4:
        render_valuation_tab(valuation, assumptions)
    
    with tab5:
        render_cap_table_tab(assumptions)


if __name__ == "__main__":
    main()
