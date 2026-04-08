"""
Joule Financial Model — Configuration & Default Assumptions
============================================================
Centralizes all model parameters with metadata for tooltips.
"""

from dataclasses import dataclass, field
from typing import List, Dict

# ══════════════════════════════════════════════════════════════════════════════
# THEME CONFIGURATION
# ═══════════════════════���══════════════════════════════════════════════════════

THEME = {
    "primary_color": "#00B050",      # Joule green
    "secondary_color": "#1F4E79",    # Deep blue
    "accent_color": "#2E75B6",       # Medium blue
    "background": "#FAFAFA",
    "text_color": "#333333",
    "font_family": "Inter, sans-serif",
}

CHART_COLORS = {
    "revenue": ["#1F4E79", "#2E75B6", "#00B050"],  # Upfront, Completion, kWh
    "scenarios": {"Worst": "#FF6B6B", "Base": "#1F4E79", "Best": "#00B050"},
    "cap_table": ["#1F4E79", "#2E75B6", "#5B9BD5", "#A5C8E4", "#00B050", "#70AD47", "#FFC000"],
}


# ══════════════════════════════════════════════════════════════════════════════
# ASSUMPTION METADATA (for tooltips)
# ══════════════════════════════════════════════════════════════════════════════

ASSUMPTION_TOOLTIPS: Dict[str, str] = {
    "upfront_fee_rate": """
        **Upfront Fee (2% of Funds Raised)**  
        This is a one-time fee charged when a project raises capital through our platform.
        Industry standard for infrastructure financing platforms ranges from 1.5%–3%.
        We benchmark at 2% to remain competitive while ensuring platform sustainability.
    """,
    "completion_fee_rate": """
        **Completion Fee (0.5% of Funds Raised)**  
        Recognized straight-line over the 3-year project build period.
        This aligns our incentives with successful project delivery.
        Annual recognition = 0.5% ÷ 3 years = ~0.167% per year.
    """,
    "kwh_fee_rate": """
        **kWh Revenue (£0.01 per kWh)**  
        Post-completion recurring revenue from energy production monitoring.
        
        **Calculation Basis (UK DESNZ):**
        - Wind turbine: 2.5 MW capacity (midpoint 2–3 MW)
        - Annual output: 7,000 MWh per turbine
        - Cost per turbine: £3.3M (midpoint £2.6M–£4M)
        
        Formula: `Funds ÷ £3.3M × 7,000 MWh × £0.01 ≈ 2.12% of Funds p.a.`
    """,
    "scenario": """
        **Scenario Selection**  
        - **Worst Case (×0.6):** 40% reduction in funds raised — stress test
        - **Base Case (×1.0):** Central planning assumption
        - **Best Case (×1.4):** 40% uplift — optimistic market conditions
        
        Only Total Funds Raised is affected; all other assumptions remain constant.
    """,
    "ev_revenue_multiple": """
        **EV/Revenue Multiple (7×)**  
        Based on comparable green fintech / infrastructure platform transactions:
        - Octopus Energy (2024): 8–10× revenue
        - Ripple Energy (2023): 5–7× revenue
        
        We use 7× as a conservative mid-market assumption for Year 5 exit.
    """,
    "required_return": """
        **Tiered Discount Rates (VC Method)**  
        - Years 1–3: **50%** (highest risk — early stage)
        - Years 4–6: **35%** (growth stage, product-market fit)
        - Years 7–10: **25%** (more mature, lower execution risk)
        
        This reflects the declining risk profile as the company scales.
    """,
}


# ══════════════════════════════════════════════════════════════════════════════
# DEFAULT MODEL ASSUMPTIONS
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ModelAssumptions:
    """Encapsulates all financial model inputs with defaults from your Excel model."""
    
    # Timeline
    years: List[int] = field(default_factory=lambda: [2026, 2027, 2028, 2029, 2030])
    
    # Fee Structure
    upfront_fee_rate: float = 0.02              # 2%
    completion_fee_rate: float = 0.005          # 0.5%
    project_completion_years: int = 3
    kwh_fee_rate: float = 0.01                  # £0.01 per kWh
    kwh_revenue_rate: float = 0.0212            # 2.12% of funds raised
    kwh_contract_duration: int = 10
    
    # Funds Raised (Base Case)
    base_funds_raised: List[float] = field(default_factory=lambda: [
        18_000_000, 56_000_000, 150_000_000, 312_000_000, 576_000_000
    ])
    
    # Scenario Multipliers
    worst_case_multiplier: float = 0.6
    best_case_multiplier: float = 1.4
    
    # COGS
    smart_contract_costs: List[float] = field(default_factory=lambda: [
        185_000, 170_000, 105_000, 115_000, 130_000
    ])
    hosting_api_rate: float = 0.015             # 1.5% of gross revenue
    
    # OpEx Year 1 (hardcoded £)
    opex_y1_salaries: float = 225_000
    opex_y1_marketing: float = 50_000
    opex_y1_legal: float = 100_000
    opex_y1_insurance: float = 20_000
    opex_y1_rd: float = 50_000
    
    # OpEx Years 2–5 (% of gross revenue)
    opex_salaries_pct: List[float] = field(default_factory=lambda: [0.25, 0.25, 0.25, 0.25])
    opex_marketing_pct: List[float] = field(default_factory=lambda: [0.12, 0.07, 0.07, 0.07])
    opex_legal_pct: List[float] = field(default_factory=lambda: [0.08, 0.08, 0.08, 0.08])
    opex_insurance_pct: List[float] = field(default_factory=lambda: [0.02, 0.01, 0.01, 0.01])
    opex_rd_pct: List[float] = field(default_factory=lambda: [0.19, 0.11, 0.05, 0.05])
    
    # Tax
    small_profits_rate: float = 0.19
    main_rate: float = 0.25
    lower_profits_threshold: float = 50_000
    upper_profits_threshold: float = 250_000
    loss_carry_forward_limit: float = 5_000_000
    
    # Equity Raises
    equity_raises: List[float] = field(default_factory=lambda: [
        500_000, 1_500_000, 4_000_000, 0, 9_000_000
    ])
    
    # Working Capital
    dso_days: int = 45
    dpo_days: int = 30
    
    # Valuation
    ev_revenue_multiple: float = 7.0
    required_return_y1_3: float = 0.50
    required_return_y3_6: float = 0.35
    required_return_y6_10: float = 0.25
    revenue_cagr_y5_y10: float = 0.60
    
    # Cap Table
    founder_shares: int = 1_000_000
    funding_rounds: List[Dict] = field(default_factory=lambda: [
        {"name": "Pre-seed T1", "ownership_pct": 0.040652},
        {"name": "Pre-seed T2", "ownership_pct": 0.029269},
        {"name": "Seed", "ownership_pct": 0.137200},
        {"name": "Series A T1", "ownership_pct": 0.095278},
        {"name": "Series A T2", "ownership_pct": 0.049624},
        {"name": "Series B", "ownership_pct": 0.129228},
    ])


def get_default_assumptions() -> ModelAssumptions:
    """Factory function for default assumptions."""
    return ModelAssumptions()
