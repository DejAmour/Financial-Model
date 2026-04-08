"""
InfraFund Financial Model - Configuration & Default Assumptions
===============================================================
Centralizes all model parameters with metadata for tooltips.
"""

from dataclasses import dataclass, field
from typing import List, Dict

THEME = {    
    "primary_color": "#00B050",
    "secondary_color": "#1F4E79",
    "accent_color": "#2E75B6",
    "background": "#FAFAFA",
    "text_color": "#333333",
    "font_family": "Inter, sans-serif",
}

CHART_COLORS = {
    "revenue": ["#1F4E79", "#2E75B6", "#00B050"],
    "scenarios": {"Worst": "#FF6B6B", "Base": "#1F4E79", "Best": "#00B050"},
    "cap_table": ["#1F4E79", "#2E75B6", "#5B9BD5", "#A5C8E4", "#00B050", "#70AD47", "#FFC000"],
    "impact": ["#00B050", "#2E75B6", "#FFC000", "#FF6B6B"],
}



COMMUNITY_IMPACT_CONFIG = {
    "avg_turbine_cost_gbp": 3_300_000,
    "annual_mwh_per_turbine": 7_000,
    "household_kwh_per_year": 3_100,
    "avg_annual_bill_gbp": 1_800,
    "savings_rate_low": 0.25,
    "savings_rate_high": 0.50,
    "co2_tonnes_per_mwh": 0.193,
}

ASSUMPTION_TOOLTIPS: Dict[str, str] = {
    "upfront_fee_rate": """
**Tokenisation Fee (2% of Funds Raised)**
Our 2% tokenisation fee is 50% lower than traditional investment banking fees (3-5%),
making green infrastructure investment accessible at scale.

**Market Benchmark:** Traditional IB fees: 3-5%. Direct listing platforms: 1-2%.
    """,
    "completion_fee_rate": """
**Verification Fee (0.5% of Funds Raised)**
Recognized straight-line over the 3-year project build period.
Covers smart-contract audit and on-chain verification costs.
    """,
    "kwh_fee_rate": """
**Assurance Fee (1p/kWh)**
Post-completion recurring revenue from energy production monitoring.

**UK DESNZ Data:** Wind turbine 2.5 MW, 7,000 MWh/year, £3.3M cost.
Formula: Funds / £3.3M x 7,000 MWh x £0.01 = 2.12% of Funds p.a.
    """,
    "scenario": """
**Scenario Selection**
- **Worst Case (x0.6):** 40% reduction in funds raised
- **Base Case (x1.0):** Central planning assumption
- **Best Case (x1.4):** 40% uplift
    """,
    "ev_revenue_multiple": """
**EV/Revenue Multiple (7x)**
Based on comparable transactions:
- Octopus Energy (2024): 8-10x revenue
- Ripple Energy (2023): 5-7x revenue
    """,
    "required_return": """
**Tiered Discount Rates (VC Method)**
- Years 1-3: **50%** (early stage)
- Years 4-6: **35%** (growth stage)
- Years 7-10: **25%** (mature)
    """,
    "tokenisation_fee": """
**2% Tokenisation Fee**
Our 2% tokenisation fee is 50% lower than traditional investment banking fees (3-5%),
making green infrastructure investment accessible at scale.

**Market Benchmark:**
- Traditional IB fees: 3-5% (Goldman Sachs, Morgan Stanley)
- Direct listing platforms: 1-2%
- InfraFund: 2% (value-add midpoint)
    """,
    "verification_fee": """
**0.5% Verification Fee**
Covers smart-contract auditing and on-chain project verification.

**Market Benchmark:**
- Trail of Bits audit: £80K-£150K per engagement
- OpenZeppelin audit: £60K-£120K per engagement
    """,
    "assurance_fee": """
**1p/kWh Assurance Fee**
Recurring post-completion revenue tied to energy generation.

**UK DESNZ Data:**
- Onshore wind capacity factor: ~27%
- 2.5 MW turbine -> ~5,900-7,000 MWh/year
- At £0.01/kWh -> £59-£70 per kW installed per year
    """,
    "year1_target": """
**£18M Year 1 Funds-Under-Management Target**
Derived from a bottom-up pipeline of 3-4 community energy projects
at an average deal size of £4M-£6M.

**Comparable Platform Growth:**
- Ripple Energy: £12M in first 12 months
- Abundance Investment: £15M in Year 1
    """,
    "opex_y1": """
**Year 1 Operating Expenses - £445K**
- Salaries & Contractors: £225K (core team of 4)
- Marketing & BD: £50K (digital + events)
- Legal & Compliance: £100K (FCA authorisation, AML)
- Insurance: £20K (PI + D&O)
- R&D / Technology: £50K (smart-contract development)
    """,
    "smart_contract_costs": """
**Tokenomics & Smart Contract Audit Costs**
Annual audit costs reflect evolving smart-contract complexity.

**Security Audit Providers & Market Rates:**
- Trail of Bits: £80K-£200K per audit
- OpenZeppelin: £60K-£150K per audit
- ConsenSys Diligence: £50K-£100K per audit
    """,
}


@dataclass
class ModelAssumptions:
    """Encapsulates all financial model inputs."""
    years: List[int] = field(default_factory=lambda: [2026, 2027, 2028, 2029, 2030])
    upfront_fee_rate: float = 0.02
    completion_fee_rate: float = 0.005
    project_completion_years: int = 3
    kwh_fee_rate: float = 0.01
    kwh_revenue_rate: float = 0.0212
    kwh_contract_duration: int = 10
    base_funds_raised: List[float] = field(default_factory=lambda: [18_000_000, 56_000_000, 150_000_000, 312_000_000, 576_000_000])
    worst_case_multiplier: float = 0.6
    best_case_multiplier: float = 1.4
    smart_contract_costs: List[float] = field(default_factory=lambda: [185_000, 170_000, 105_000, 115_000, 130_000])
    hosting_api_rate: float = 0.015
    opex_y1_salaries: float = 225_000
    opex_y1_marketing: float = 50_000
    opex_y1_legal: float = 100_000
    opex_y1_insurance: float = 20_000
    opex_y1_rd: float = 50_000
    opex_salaries_pct: List[float] = field(default_factory=lambda: [0.25, 0.25, 0.25, 0.25])
    opex_marketing_pct: List[float] = field(default_factory=lambda: [0.12, 0.07, 0.07, 0.07])
    opex_legal_pct: List[float] = field(default_factory=lambda: [0.08, 0.08, 0.08, 0.08])
    opex_insurance_pct: List[float] = field(default_factory=lambda: [0.02, 0.01, 0.01, 0.01])
    opex_rd_pct: List[float] = field(default_factory=lambda: [0.19, 0.11, 0.05, 0.05])
    small_profits_rate: float = 0.19
    main_rate: float = 0.25
    lower_profits_threshold: float = 50_000
    upper_profits_threshold: float = 250_000
    loss_carry_forward_limit: float = 5_000_000
    equity_raises: List[float] = field(default_factory=lambda: [500_000, 1_500_000, 4_000_000, 0, 9_000_000])
    dso_days: int = 45
    dpo_days: int = 30
    ev_revenue_multiple: float = 7.0
    required_return_y1_3: float = 0.50
    required_return_y3_6: float = 0.35
    required_return_y6_10: float = 0.25
    revenue_cagr_y5_y10: float = 0.60
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
