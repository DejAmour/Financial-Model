"""
InfraFund Financial Model - Calculation Engine
===============================================
Pure Python functions that replicate the Excel formulas.
All functions are stateless and cacheable.
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Dict, Tuple
from config import ModelAssumptions, COMMUNITY_IMPACT_CONFIG


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
    """Compute the three revenue streams."""
    n = len(funds_raised)
    upfront_fee = [f * upfront_fee_rate for f in funds_raised]

    completion_fee = []
    annual_rate = completion_fee_rate / project_completion_years
    for i in range(n):
        start = max(0, i - project_completion_years + 1)
        cohort_sum = sum(funds_raised[start:i + 1])
        completion_fee.append(cohort_sum * annual_rate)

    kwh_revenue = []
    for i in range(n):
        if i < project_completion_years:
            kwh_revenue.append(0)
        else:
            completed_funds = sum(funds_raised[: i - project_completion_years + 1])
            kwh_revenue.append(completed_funds * kwh_revenue_rate)

    return upfront_fee, completion_fee, kwh_revenue


def compute_cogs(
    gross_revenue: List[float],
    smart_contract_costs: List[float],
    hosting_api_rate: float,
) -> List[float]:
    """Compute Cost of Goods Sold."""
    return [sc + (gr * hosting_api_rate) for sc, gr in zip(smart_contract_costs, gross_revenue)]


def compute_opex(gross_revenue: List[float], assumptions: ModelAssumptions) -> List[float]:
    """Compute Operating Expenses."""
    opex = []
    y1_total = (
        assumptions.opex_y1_salaries + assumptions.opex_y1_marketing +
        assumptions.opex_y1_legal + assumptions.opex_y1_insurance + assumptions.opex_y1_rd
    )
    opex.append(y1_total)

    for i in range(1, len(gross_revenue)):
        pct_idx = i - 1
        total_pct = (
            assumptions.opex_salaries_pct[pct_idx] + assumptions.opex_marketing_pct[pct_idx] +
            assumptions.opex_legal_pct[pct_idx] + assumptions.opex_insurance_pct[pct_idx] +
            assumptions.opex_rd_pct[pct_idx]
        )
        opex.append(gross_revenue[i] * total_pct)
    return opex


def compute_tax(ebit: List[float], assumptions: ModelAssumptions) -> List[float]:
    """Compute UK Corporation Tax with loss carry-forward."""
    tax = []
    loss_balance = 0.0
    for profit in ebit:
        if profit <= 0:
            loss_balance += abs(profit)
            tax.append(0)
        else:
            relief = min(loss_balance, assumptions.loss_carry_forward_limit, profit)
            taxable = max(0, profit - relief)
            loss_balance = max(0, loss_balance - relief)
            if taxable <= assumptions.lower_profits_threshold:
                t = taxable * assumptions.small_profits_rate
            elif taxable <= assumptions.upper_profits_threshold:
                t = (taxable * assumptions.main_rate) - ((assumptions.upper_profits_threshold - taxable) * 3 / 200)
            else:
                t = taxable * assumptions.main_rate
            tax.append(max(0, t))
    return tax


def compute_income_statement(scenario: str, assumptions: ModelAssumptions) -> IncomeStatementResult:
    """Full income statement computation."""
    funds_raised = apply_scenario(assumptions.base_funds_raised, scenario)
    upfront_fee, completion_fee, kwh_revenue = compute_revenue_streams(
        funds_raised, assumptions.upfront_fee_rate, assumptions.completion_fee_rate,
        assumptions.project_completion_years, assumptions.kwh_revenue_rate,
    )
    gross_revenue = [u + c + k for u, c, k in zip(upfront_fee, completion_fee, kwh_revenue)]
    cogs = compute_cogs(gross_revenue, assumptions.smart_contract_costs, assumptions.hosting_api_rate)
    gross_profit = [gr - c for gr, c in zip(gross_revenue, cogs)]
    opex = compute_opex(gross_revenue, assumptions)
    ebit = [gp - op for gp, op in zip(gross_profit, opex)]
    tax = compute_tax(ebit, assumptions)
    net_income = [e - t for e, t in zip(ebit, tax)]

    return IncomeStatementResult(
        years=assumptions.years, funds_raised=funds_raised, upfront_fee=upfront_fee,
        completion_fee=completion_fee, kwh_revenue=kwh_revenue, gross_revenue=gross_revenue,
        cogs=cogs, gross_profit=gross_profit, opex=opex, ebit=ebit, tax=tax, net_income=net_income,
    )


@dataclass
class ValuationResult:
    """Valuation metrics output."""
    equity_value_today: float
    exit_valuation: float
    investor_irr: float
    cash_on_cash: float
    exit_proceeds: float


@dataclass
class ValuationBreakdown:
    """Step-by-step valuation breakdown for transparency UI."""
    exit_year_revenue: float
    ev_multiple: float
    exit_ev: float
    exit_equity: float
    discount_rate_y1_3: float
    discount_rate_y3_6: float
    discount_factor: float
    equity_value_today: float
    entry_stake: float
    dilution_factor: float
    stake_at_exit: float
    exit_proceeds: float
    investment: float
    cash_on_cash: float
    investor_irr: float
    exit_year: int


def compute_valuation(
    gross_revenue: List[float], ev_multiple: float, discount_rates: Tuple[float, float, float],
    exit_year: int = 5, investment: float = 250_000, entry_stake: float = 0.040652, dilution_factor: float = 0.75,
) -> ValuationResult:
    """VC-method valuation calculation."""
    exit_revenue = gross_revenue[exit_year - 1] if exit_year <= len(gross_revenue) else gross_revenue[-1]
    exit_ev = exit_revenue * ev_multiple
    exit_equity = exit_ev
    r1, r2, r3 = discount_rates
    discount_factor = (1 + r1) ** min(exit_year, 3) * (1 + r2) ** max(0, min(exit_year, 6) - 3) * (1 + r3) ** max(0, exit_year - 6)
    equity_value_today = exit_equity / discount_factor
    stake_at_exit = entry_stake * dilution_factor
    exit_proceeds = exit_equity * stake_at_exit
    cash_on_cash = exit_proceeds / investment if investment > 0 else 0
    irr = (exit_proceeds / investment) ** (1 / exit_year) - 1 if exit_proceeds > 0 and investment > 0 else 0
    return ValuationResult(equity_value_today=equity_value_today, exit_valuation=exit_equity,
                           investor_irr=irr, cash_on_cash=cash_on_cash, exit_proceeds=exit_proceeds)


def compute_valuation_breakdown(
    gross_revenue: List[float], ev_multiple: float, discount_rates: Tuple[float, float, float],
    exit_year: int = 5, investment: float = 250_000, entry_stake: float = 0.040652, dilution_factor: float = 0.75,
) -> ValuationBreakdown:
    """Returns a fully transparent, step-by-step breakdown of the VC valuation."""
    exit_revenue = gross_revenue[exit_year - 1] if exit_year <= len(gross_revenue) else gross_revenue[-1]
    exit_ev = exit_revenue * ev_multiple
    exit_equity = exit_ev
    r1, r2, r3 = discount_rates
    discount_factor = (1 + r1) ** min(exit_year, 3) * (1 + r2) ** max(0, min(exit_year, 6) - 3) * (1 + r3) ** max(0, exit_year - 6)
    equity_value_today = exit_equity / discount_factor
    stake_at_exit = entry_stake * dilution_factor
    exit_proceeds = exit_equity * stake_at_exit
    cash_on_cash = exit_proceeds / investment if investment > 0 else 0
    irr = (exit_proceeds / investment) ** (1 / exit_year) - 1 if exit_proceeds > 0 and investment > 0 else 0
    return ValuationBreakdown(
        exit_year_revenue=exit_revenue, ev_multiple=ev_multiple, exit_ev=exit_ev, exit_equity=exit_equity,
        discount_rate_y1_3=r1, discount_rate_y3_6=r2, discount_factor=discount_factor,
        equity_value_today=equity_value_today, entry_stake=entry_stake, dilution_factor=dilution_factor,
        stake_at_exit=stake_at_exit, exit_proceeds=exit_proceeds, investment=investment,
        cash_on_cash=cash_on_cash, investor_irr=irr, exit_year=exit_year,
    )


def compute_cap_table(founder_shares: int, funding_rounds: List[Dict]) -> pd.DataFrame:
    """Compute cap table ownership distribution."""
    shareholders = ["Founders"]
    shares = [founder_shares]
    cumulative = founder_shares
    for rnd in funding_rounds:
        pct = rnd["ownership_pct"]
        new_shares = cumulative / (1 - pct) - cumulative
        shareholders.append(rnd["name"])
        shares.append(new_shares)
        cumulative += new_shares
    total = sum(shares)
    ownership = [s / total for s in shares]
    return pd.DataFrame({"Shareholder": shareholders, "Shares": shares, "Ownership %": ownership})


@dataclass
class CommunityImpactResult:
    """Community impact metrics derived from funds under management."""
    funds_under_management: float
    num_turbines: float
    total_mwh_per_year: float
    households_reached: int
    avg_annual_savings_low: float
    avg_annual_savings_high: float
    total_savings_low: float
    total_savings_high: float
    co2_avoided_tonnes: float
    bill_reduction_pct_low: float
    bill_reduction_pct_high: float


def compute_community_impact(funds_under_management: float, cfg: dict = None) -> CommunityImpactResult:
    """Translate platform scale into community impact metrics."""
    if cfg is None:
        cfg = COMMUNITY_IMPACT_CONFIG
    num_turbines = funds_under_management / cfg["avg_turbine_cost_gbp"]
    total_mwh_per_year = num_turbines * cfg["annual_mwh_per_turbine"]
    total_kwh_per_year = total_mwh_per_year * 1_000
    households_reached = int(total_kwh_per_year / cfg["household_kwh_per_year"])
    avg_bill = cfg["avg_annual_bill_gbp"]
    avg_annual_savings_low = avg_bill * cfg["savings_rate_low"]
    avg_annual_savings_high = avg_bill * cfg["savings_rate_high"]
    total_savings_low = households_reached * avg_annual_savings_low
    total_savings_high = households_reached * avg_annual_savings_high
    co2_avoided_tonnes = total_mwh_per_year * cfg["co2_tonnes_per_mwh"]
    return CommunityImpactResult(
        funds_under_management=funds_under_management, num_turbines=num_turbines,
        total_mwh_per_year=total_mwh_per_year, households_reached=households_reached,
        avg_annual_savings_low=avg_annual_savings_low, avg_annual_savings_high=avg_annual_savings_high,
        total_savings_low=total_savings_low, total_savings_high=total_savings_high,
        co2_avoided_tonnes=co2_avoided_tonnes, bill_reduction_pct_low=cfg["savings_rate_low"],
        bill_reduction_pct_high=cfg["savings_rate_high"],
    )
