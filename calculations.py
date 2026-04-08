"""
Joule Financial Model — Calculation Engine
============================================
Pure Python functions that replicate your Excel formulas.
All functions are stateless and cacheable.
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Dict, Tuple
from config import ModelAssumptions


@dataclass
class IncomeStatementResult:
    """Structured output from income statement computation."""
    years: List[int]
    funds_raised: List[float]
    upfront_fee: List[float]
    completion_fee: List[float]
    kwh_revenue: List[float]
    gross_revenue: List[float]
    cogs: List[float]
    gross_profit: List[float]
    opex: List[float]
    ebit: List[float]
    tax: List[float]
    net_income: List[float]
    
    def to_dataframe(self) -> pd.DataFrame:
        """Convert to pandas DataFrame for easy charting."""
        return pd.DataFrame({
            "Year": self.years,
            "Funds Raised": self.funds_raised,
            "Upfront Fee": self.upfront_fee,
            "Completion Fee": self.completion_fee,
            "kWh Revenue": self.kwh_revenue,
            "Gross Revenue": self.gross_revenue,
            "COGS": self.cogs,
            "Gross Profit": self.gross_profit,
            "Operating Expenses": self.opex,
            "EBIT": self.ebit,
            "Tax": self.tax,
            "Net Income": self.net_income,
        })


def apply_scenario(base_funds: List[float], scenario: str) -> List[float]:
    """Apply scenario multiplier to base funds raised."""
    multipliers = {"Worst": 0.6, "Base": 1.0, "Best": 1.4}
    mult = multipliers.get(scenario, 1.0)
    return [f * mult for f in base_funds]


def compute_revenue_streams(
    funds_raised: List[float],
    upfront_fee_rate: float,
    completion_fee_rate: float,
    project_completion_years: int,
    kwh_revenue_rate: float,
) -> Tuple[List[float], List[float], List[float]]:
    """
    Compute the three revenue streams.
    
    Returns:
        (upfront_fee, completion_fee, kwh_revenue) — each a list of 5 values
    """
    n = len(funds_raised)
    
    # 1. Upfront Fee: immediate recognition
    upfront_fee = [f * upfront_fee_rate for f in funds_raised]
    
    # 2. Completion Fee: straight-line over build period (3-year sliding window)
    completion_fee = []
    annual_rate = completion_fee_rate / project_completion_years
    for i in range(n):
        # Sum funds from cohorts active in this year's window
        start = max(0, i - project_completion_years + 1)
        cohort_sum = sum(funds_raised[start:i + 1])
        completion_fee.append(cohort_sum * annual_rate)
    
    # 3. kWh Revenue: starts 3 years after funds raised (lagged by build period)
    kwh_revenue = []
    for i in range(n):
        if i < project_completion_years:
            kwh_revenue.append(0)
        else:
            # Cumulative funds from completed cohorts
            completed_funds = sum(funds_raised[: i - project_completion_years + 1])
            kwh_revenue.append(completed_funds * kwh_revenue_rate)
    
    return upfront_fee, completion_fee, kwh_revenue


def compute_cogs(
    gross_revenue: List[float],
    smart_contract_costs: List[float],
    hosting_api_rate: float,
) -> List[float]:
    """Compute Cost of Goods Sold."""
    return [
        sc + (gr * hosting_api_rate)
        for sc, gr in zip(smart_contract_costs, gross_revenue)
    ]


def compute_opex(
    gross_revenue: List[float],
    assumptions: ModelAssumptions,
) -> List[float]:
    """Compute Operating Expenses (Year 1 hardcoded, Years 2–5 % of revenue)."""
    opex = []
    
    # Year 1 (index 0): hardcoded values
    y1_total = (
        assumptions.opex_y1_salaries +
        assumptions.opex_y1_marketing +
        assumptions.opex_y1_legal +
        assumptions.opex_y1_insurance +
        assumptions.opex_y1_rd
    )
    opex.append(y1_total)
    
    # Years 2–5 (indices 1–4): % of gross revenue
    for i in range(1, len(gross_revenue)):
        pct_idx = i - 1  # Index into the percentage lists
        total_pct = (
            assumptions.opex_salaries_pct[pct_idx] +
            assumptions.opex_marketing_pct[pct_idx] +
            assumptions.opex_legal_pct[pct_idx] +
            assumptions.opex_insurance_pct[pct_idx] +
            assumptions.opex_rd_pct[pct_idx]
        )
        opex.append(gross_revenue[i] * total_pct)
    
    return opex


def compute_tax(ebit: List[float], assumptions: ModelAssumptions) -> List[float]:
    """
    Compute UK Corporation Tax with loss carry-forward.
    Simplified: uses marginal relief logic from your Excel model.
    """
    tax = []
    loss_balance = 0.0
    
    for profit in ebit:
        if profit <= 0:
            # Accumulate losses
            loss_balance += abs(profit)
            tax.append(0)
        else:
            # Apply loss relief
            relief = min(loss_balance, assumptions.loss_carry_forward_limit, profit)
            taxable = max(0, profit - relief)
            loss_balance = max(0, loss_balance - relief)
            
            # Marginal relief calculation
            if taxable <= assumptions.lower_profits_threshold:
                t = taxable * assumptions.small_profits_rate
            elif taxable <= assumptions.upper_profits_threshold:
                t = (taxable * assumptions.main_rate) - \
                    ((assumptions.upper_profits_threshold - taxable) * 3 / 200)
            else:
                t = taxable * assumptions.main_rate
            
            tax.append(max(0, t))
    
    return tax


def compute_income_statement(
    scenario: str,
    assumptions: ModelAssumptions,
) -> IncomeStatementResult:
    """
    Full income statement computation.
    This is the main entry point called when inputs change.
    """
    # Apply scenario to funds raised
    funds_raised = apply_scenario(assumptions.base_funds_raised, scenario)
    
    # Revenue streams
    upfront_fee, completion_fee, kwh_revenue = compute_revenue_streams(
        funds_raised,
        assumptions.upfront_fee_rate,
        assumptions.completion_fee_rate,
        assumptions.project_completion_years,
        assumptions.kwh_revenue_rate,
    )
    
    # Gross revenue
    gross_revenue = [
        u + c + k for u, c, k in zip(upfront_fee, completion_fee, kwh_revenue)
    ]
    
    # COGS
    cogs = compute_cogs(
        gross_revenue,
        assumptions.smart_contract_costs,
        assumptions.hosting_api_rate,
    )
    
    # Gross profit
    gross_profit = [gr - c for gr, c in zip(gross_revenue, cogs)]
    
    # Operating expenses
    opex = compute_opex(gross_revenue, assumptions)
    
    # EBIT
    ebit = [gp - op for gp, op in zip(gross_profit, opex)]
    
    # Tax
    tax = compute_tax(ebit, assumptions)
    
    # Net income
    net_income = [e - t for e, t in zip(ebit, tax)]
    
    return IncomeStatementResult(
        years=assumptions.years,
        funds_raised=funds_raised,
        upfront_fee=upfront_fee,
        completion_fee=completion_fee,
        kwh_revenue=kwh_revenue,
        gross_revenue=gross_revenue,
        cogs=cogs,
        gross_profit=gross_profit,
        opex=opex,
        ebit=ebit,
        tax=tax,
        net_income=net_income,
    )


# ════════════════════════════════════════════════════════��═════════════════════
# VALUATION & CAP TABLE
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ValuationResult:
    """Valuation metrics output."""
    equity_value_today: float
    exit_valuation: float
    investor_irr: float
    cash_on_cash: float
    exit_proceeds: float


def compute_valuation(
    gross_revenue: List[float],
    ev_multiple: float,
    discount_rates: Tuple[float, float, float],
    exit_year: int = 5,
    investment: float = 250_000,
    entry_stake: float = 0.040652,
    dilution_factor: float = 0.75,  # Approximate dilution through Series B
) -> ValuationResult:
    """
    VC-method valuation calculation.
    """
    # Exit valuation (EV = Revenue × Multiple)
    exit_revenue = gross_revenue[exit_year - 1] if exit_year <= len(gross_revenue) else gross_revenue[-1]
    exit_ev = exit_revenue * ev_multiple
    
    # Simplified: assume exit equity ≈ exit EV (no debt)
    exit_equity = exit_ev
    
    # Tiered discount factor
    r1, r2, r3 = discount_rates
    discount_factor = (
        (1 + r1) ** min(exit_year, 3) *
        (1 + r2) ** max(0, min(exit_year, 6) - 3) *
        (1 + r3) ** max(0, exit_year - 6)
    )
    
    # Present value
    equity_value_today = exit_equity / discount_factor
    
    # Investor returns
    stake_at_exit = entry_stake * dilution_factor
    exit_proceeds = exit_equity * stake_at_exit
    cash_on_cash = exit_proceeds / investment if investment > 0 else 0
    
    # IRR calculation (simple approximation)
    if exit_proceeds > 0 and investment > 0:
        irr = (exit_proceeds / investment) ** (1 / exit_year) - 1
    else:
        irr = 0
    
    return ValuationResult(
        equity_value_today=equity_value_today,
        exit_valuation=exit_equity,
        investor_irr=irr,
        cash_on_cash=cash_on_cash,
        exit_proceeds=exit_proceeds,
    )


def compute_cap_table(
    founder_shares: int,
    funding_rounds: List[Dict],
) -> pd.DataFrame:
    """
    Compute cap table ownership distribution.
    """
    shareholders = ["Founders"]
    shares = [founder_shares]
    cumulative = founder_shares
    
    for rnd in funding_rounds:
        # New shares = cumulative / (1 - pct) - cumulative
        pct = rnd["ownership_pct"]
        new_shares = cumulative / (1 - pct) - cumulative
        shareholders.append(rnd["name"])
        shares.append(new_shares)
        cumulative += new_shares
    
    total = sum(shares)
    ownership = [s / total for s in shares]
    
    return pd.DataFrame({
        "Shareholder": shareholders,
        "Shares": shares,
        "Ownership %": ownership,
    })
