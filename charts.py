"""
Joule Financial Model — Chart Components
=========================================
Professional Plotly charts optimized for investor presentations.
"""

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
from typing import List, Dict
from config import CHART_COLORS, THEME


def create_revenue_waterfall(
    years: List[int],
    upfront: List[float],
    completion: List[float],
    kwh: List[float],
) -> go.Figure:
    """Stacked bar chart showing revenue composition over time."""
    
    fig = go.Figure()
    
    # Stacked bars
    fig.add_trace(go.Bar(
        name="Upfront Fee",
        x=years,
        y=upfront,
        marker_color=CHART_COLORS["revenue"][0],
        hovertemplate="Upfront Fee: £%{y:,.0f}<extra></extra>",
    ))
    
    fig.add_trace(go.Bar(
        name="Completion Fee",
        x=years,
        y=completion,
        marker_color=CHART_COLORS["revenue"][1],
        hovertemplate="Completion Fee: £%{y:,.0f}<extra></extra>",
    ))
    
    fig.add_trace(go.Bar(
        name="kWh Revenue",
        x=years,
        y=kwh,
        marker_color=CHART_COLORS["revenue"][2],
        hovertemplate="kWh Revenue: £%{y:,.0f}<extra></extra>",
    ))
    
    fig.update_layout(
        barmode="stack",
        title={
            "text": "Revenue Streams Breakdown",
            "font": {"size": 18, "color": THEME["secondary_color"]},
        },
        xaxis_title="Year",
        yaxis_title="Revenue (£)",
        yaxis_tickformat="£,.0s",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        template="plotly_white",
        height=400,
    )
    
    return fig


def create_scenario_comparison(
    years: List[int],
    worst: List[float],
    base: List[float],
    best: List[float],
) -> go.Figure:
    """Line chart comparing revenue across scenarios."""
    
    fig = go.Figure()
    
    for name, data, dash in [
        ("Worst Case", worst, "dot"),
        ("Base Case", base, "solid"),
        ("Best Case", best, "dash"),
    ]:
        fig.add_trace(go.Scatter(
            name=name,
            x=years,
            y=data,
            mode="lines+markers",
            line=dict(
                color=CHART_COLORS["scenarios"][name.split()[0]],
                width=3,
                dash=dash,
            ),
            marker=dict(size=8),
            hovertemplate=f"{name}: £%{{y:,.0f}}<extra></extra>",
        ))
    
    fig.update_layout(
        title={
            "text": "Scenario Comparison: Gross Revenue",
            "font": {"size": 18, "color": THEME["secondary_color"]},
        },
        xaxis_title="Year",
        yaxis_title="Gross Revenue (£)",
        yaxis_tickformat="£,.0s",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        template="plotly_white",
        height=400,
    )
    
    return fig


def create_cap_table_pie(cap_table_df: pd.DataFrame) -> go.Figure:
    """Pie chart showing ownership distribution."""
    
    fig = go.Figure(go.Pie(
        labels=cap_table_df["Shareholder"],
        values=cap_table_df["Ownership %"],
        hole=0.4,
        marker_colors=CHART_COLORS["cap_table"],
        textinfo="label+percent",
        textposition="outside",
        hovertemplate="%{label}: %{percent:.1%}<extra></extra>",
    ))
    
    fig.update_layout(
        title={
            "text": "Cap Table: Final Ownership Distribution",
            "font": {"size": 18, "color": THEME["secondary_color"]},
        },
        showlegend=False,
        template="plotly_white",
        height=450,
        annotations=[dict(
            text="Ownership",
            x=0.5, y=0.5,
            font_size=14,
            showarrow=False,
        )],
    )
    
    return fig


def create_runway_chart(
    years: List[int],
    cash_balance: List[float],
    burn_rate: List[float],
) -> go.Figure:
    """Combined bar (cash) and line (burn rate) chart for runway analysis."""
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Bar(
            name="Cash Balance",
            x=years,
            y=cash_balance,
            marker_color=THEME["primary_color"],
            opacity=0.7,
            hovertemplate="Cash: £%{y:,.0f}<extra></extra>",
        ),
        secondary_y=False,
    )
    
    fig.add_trace(
        go.Scatter(
            name="Monthly Burn Rate",
            x=years,
            y=burn_rate,
            mode="lines+markers",
            line=dict(color=THEME["secondary_color"], width=3),
            marker=dict(size=10),
            hovertemplate="Burn: £%{y:,.0f}/mo<extra></extra>",
        ),
        secondary_y=True,
    )
    
    fig.update_layout(
        title={
            "text": "Runway Analysis: Cash & Burn Rate",
            "font": {"size": 18, "color": THEME["secondary_color"]},
        },
        template="plotly_white",
        height=400,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    
    fig.update_yaxes(title_text="Cash Balance (£)", tickformat="£,.0s", secondary_y=False)
    fig.update_yaxes(title_text="Burn Rate (£/month)", tickformat="£,.0s", secondary_y=True)
    
    return fig
