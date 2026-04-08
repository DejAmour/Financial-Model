"""
InfraFund Financial Model - Chart Components
=============================================
Professional Plotly charts optimized for investor presentations.
"""

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
from typing import List
from config import CHART_COLORS, THEME


def create_revenue_waterfall(years: List[int], upfront: List[float], completion: List[float], kwh: List[float]) -> go.Figure:
    """Stacked bar chart showing revenue composition over time."""
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Tokenisation Fee", x=years, y=upfront, marker_color=CHART_COLORS["revenue"][0],
                         hovertemplate="Tokenisation Fee: £%{y:,.0f}<extra></extra>"))
    fig.add_trace(go.Bar(name="Verification Fee", x=years, y=completion, marker_color=CHART_COLORS["revenue"][1],
                         hovertemplate="Verification Fee: £%{y:,.0f}<extra></extra>"))
    fig.add_trace(go.Bar(name="Assurance Revenue", x=years, y=kwh, marker_color=CHART_COLORS["revenue"][2],
                         hovertemplate="Assurance Revenue: £%{y:,.0f}<extra></extra>"))
    fig.update_layout(
        barmode="stack",
        title={"text": "Revenue Streams Breakdown", "font": {"size": 18, "color": THEME["secondary_color"]}},
        xaxis_title="Year", yaxis_title="Revenue (£)", yaxis_tickformat="£,.0s",
        legend=dict(orientation="h", yanchor="bottom", y=1.02), template="plotly_white", height=400,
    )
    return fig


def create_scenario_comparison(years: List[int], worst: List[float], base: List[float], best: List[float]) -> go.Figure:
    """Line chart comparing revenue across scenarios."""
    fig = go.Figure()
    for name, data, dash in [("Worst Case", worst, "dot"), ("Base Case", base, "solid"), ("Best Case", best, "dash")]:
        fig.add_trace(go.Scatter(
            name=name, x=years, y=data, mode="lines+markers",
            line=dict(color=CHART_COLORS["scenarios"][name.split()[0]], width=3, dash=dash),
            marker=dict(size=8), hovertemplate=f"{name}: £%{{y:,.0f}}<extra></extra>",
        ))
    fig.update_layout(
        title={"text": "Scenario Comparison: Gross Revenue", "font": {"size": 18, "color": THEME["secondary_color"]}},
        xaxis_title="Year", yaxis_title="Gross Revenue (£)", yaxis_tickformat="£,.0s",
        legend=dict(orientation="h", yanchor="bottom", y=1.02), template="plotly_white", height=400,
    )
    return fig


def create_cap_table_pie(cap_table_df: pd.DataFrame) -> go.Figure:
    """Pie chart showing ownership distribution."""
    fig = go.Figure(go.Pie(
        labels=cap_table_df["Shareholder"], values=cap_table_df["Ownership %"],
        hole=0.4, marker_colors=CHART_COLORS["cap_table"],
        textinfo="label+percent", textposition="outside",
        hovertemplate="%{label}: %{percent:.1%}<extra></extra>",
    ))
    fig.update_layout(
        title={"text": "Cap Table: Final Ownership Distribution", "font": {"size": 18, "color": THEME["secondary_color"]}},
        showlegend=False, template="plotly_white", height=450,
        annotations=[dict(text="Ownership", x=0.5, y=0.5, font_size=14, showarrow=False)],
    )
    return fig


def create_runway_chart(years: List[int], cash_balance: List[float], burn_rate: List[float]) -> go.Figure:
    """Combined bar (cash) and line (burn rate) chart for runway analysis."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(name="Cash Balance", x=years, y=cash_balance, marker_color=THEME["primary_color"],
                         opacity=0.7, hovertemplate="Cash: £%{y:,.0f}<extra></extra>"), secondary_y=False)
    fig.add_trace(go.Scatter(name="Monthly Burn Rate", x=years, y=burn_rate, mode="lines+markers",
                             line=dict(color=THEME["secondary_color"], width=3), marker=dict(size=10),
                             hovertemplate="Burn: £%{y:,.0f}/mo<extra></extra>"), secondary_y=True)
    fig.update_layout(
        title={"text": "Runway Analysis: Cash & Burn Rate", "font": {"size": 18, "color": THEME["secondary_color"]}},
        template="plotly_white", height=400, legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    fig.update_yaxes(title_text="Cash Balance (£)", tickformat="£,.0s", secondary_y=False)
    fig.update_yaxes(title_text="Burn Rate (£/month)", tickformat="£,.0s", secondary_y=True)
    return fig


def create_community_impact_chart(impact) -> go.Figure:
    """Gauge-style indicator cards for community impact metrics."""
    fig = make_subplots(
        rows=1, cols=3,
        specs=[[{"type": "indicator"}, {"type": "indicator"}, {"type": "indicator"}]],
        subplot_titles=["Households Empowered", "CO2 Avoided (tonnes/yr)", "Total Annual Savings"],
    )
    fig.add_trace(go.Indicator(mode="number", value=impact.households_reached,
                               number={"valueformat": ",.0f", "font": {"color": THEME["primary_color"], "size": 40}}), row=1, col=1)
    fig.add_trace(go.Indicator(mode="number", value=impact.co2_avoided_tonnes,
                               number={"valueformat": ",.0f", "suffix": " t", "font": {"color": THEME["secondary_color"], "size": 40}}), row=1, col=2)
    fig.add_trace(go.Indicator(mode="number", value=impact.total_savings_high,
                               number={"valueformat": ",.0f", "prefix": "up to £", "font": {"color": THEME["accent_color"], "size": 34}}), row=1, col=3)
    fig.update_layout(template="plotly_white", height=250, margin=dict(t=60, b=20, l=20, r=20))
    return fig


def create_energy_savings_chart(impact) -> go.Figure:
    """Before/after energy bill comparison chart showing 25-50% savings range."""
    avg_bill = impact.avg_annual_savings_low / impact.bill_reduction_pct_low
    categories = ["Before InfraFund", "With InfraFund (Low)", "With InfraFund (High)"]
    values = [avg_bill, avg_bill - impact.avg_annual_savings_low, avg_bill - impact.avg_annual_savings_high]
    colors = ["#FF6B6B", THEME["accent_color"], THEME["primary_color"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=categories, y=values, marker_color=colors,
                         text=[f"£{v:,.0f}/yr" for v in values], textposition="outside",
                         hovertemplate="%{x}: £%{y:,.0f}/yr<extra></extra>"))
    fig.add_hline(y=avg_bill, line_dash="dash", line_color="#FF6B6B",
                  annotation_text="Baseline avg bill", annotation_position="right")
    fig.update_layout(
        title={"text": "Average Household Energy Bill: Before vs After InfraFund", "font": {"size": 16, "color": THEME["secondary_color"]}},
        yaxis_title="Annual Bill (£)", yaxis_tickformat="£,.0f",
        template="plotly_white", height=380, showlegend=False,
    )
    return fig
