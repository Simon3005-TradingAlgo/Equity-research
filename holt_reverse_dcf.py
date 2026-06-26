"""
holt_reverse_dcf.py
===================

A HOLT-style economic-profit DCF with explicit competitive fade, plus a
reverse-DCF engine that extracts the *market-implied* CFROI / growth from the
current share price.

Design philosophy
-----------------
This is NOT a clone of Credit Suisse HOLT. The proprietary parts of HOLT --
the cross-country accounting normalisation over ~20k firms and the empirically
calibrated fade algorithms -- are not reproducible from public data. What IS
reproducible, and is the intellectually valuable core, is the *logic*:

    1. Treat the firm as one big project earning an economic return (CFROI)
       on a gross investment base.
    2. Let that return and the growth rate FADE toward the cost of capital
       (competition drives spreads to zero), instead of assuming a constant
       perpetuity growth.
    3. Value the resulting net cash-flow stream at a market-derived discount
       rate, with a terminal value anchored on the zero-spread identity
       (a firm earning exactly its cost of capital is worth its invested
       capital).
    4. INVERT the model: hold the structure fixed and solve for the CFROI the
       price is already paying for -- the "green dot" -- then test it against
       the firm's own track record for plausibility.

All rates are DECIMALS (0.10 == 10%). Monetary inputs share one currency unit.

Dependencies: numpy, pandas, scipy. (plotly only for the optional chart helper.)

Author: built for Simon's equity-research dashboard.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Optional, Sequence

import numpy as np
import pandas as pd
from scipy.optimize import brentq


FadeKind = Literal["linear", "exponential"]
SolveTarget = Literal["cfroi_terminal", "cfroi_current", "cfroi_shift", "growth_terminal"]


# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
@dataclass
class FirmInputs:
    """Economic state of the firm at t=0 plus the fade assumptions.

    Parameters
    ----------
    gross_investment : float
        Invested-capital / gross-investment base at t=0 (GI_0). In a full HOLT
        build this is inflation-adjusted gross investment; a defensible proxy is
        gross PP&E + net working capital + capitalised intangibles.
    cfroi : float
        Current economic return on the asset base (CFROI_0), decimal.
    asset_growth : float
        Current real asset / reinvestment growth rate (g_0), decimal.
    discount_rate : float
        Market-derived cost of capital (DR), decimal.
    net_debt : float
        Net debt to bridge enterprise -> equity value (debt - cash).
    shares_out : float
        Diluted shares outstanding.
    horizon : int
        Explicit fade horizon in years (HOLT default 5; 10 for high-persistence
        firms).
    cfroi_terminal : float | None
        CFROI level reached at the end of the fade. None -> DR (full fade to
        zero spread, the HOLT default for the typical firm).
    growth_terminal : float
        Long-run sustainable growth used for the terminal value. Must be < DR.
    fade : {"linear", "exponential"}
        Shape of the fade path between t=0 and the terminal level.
    fade_half_life : float | None
        Only used for exponential fade. Years to cover half the remaining gap.
        None -> horizon / 2.
    """

    gross_investment: float
    cfroi: float
    asset_growth: float
    discount_rate: float
    net_debt: float = 0.0
    shares_out: float = 1.0
    horizon: int = 5
    cfroi_terminal: Optional[float] = None
    growth_terminal: float = 0.02
    fade: FadeKind = "linear"
    fade_half_life: Optional[float] = None

    def __post_init__(self) -> None:
        if self.shares_out <= 0:
            raise ValueError("shares_out must be positive.")
        if self.horizon < 1:
            raise ValueError("horizon must be >= 1.")
        if self.growth_terminal >= self.discount_rate:
            raise ValueError(
                f"growth_terminal ({self.growth_terminal:.2%}) must be below the "
                f"discount_rate ({self.discount_rate:.2%}) for a finite terminal value."
            )

    @property
    def cfroi_term(self) -> float:
        """Terminal CFROI, defaulting to full fade to the discount rate."""
        return self.discount_rate if self.cfroi_terminal is None else self.cfroi_terminal


# --------------------------------------------------------------------------- #
# Fade paths
# --------------------------------------------------------------------------- #
def fade_path(start: float, end: float, n: int, kind: FadeKind = "linear",
              half_life: Optional[float] = None) -> np.ndarray:
    """Return the rate for years t = 1..n fading from `start` toward `end`.

    Linear hits `end` exactly at t=n. Exponential approaches it asymptotically.
    """
    t = np.arange(1, n + 1, dtype=float)
    if kind == "linear":
        return start + (end - start) * (t / n)
    if kind == "exponential":
        hl = (n / 2.0) if half_life is None else half_life
        if hl <= 0:
            raise ValueError("half_life must be positive.")
        return end + (start - end) * np.power(0.5, t / hl)
    raise ValueError(f"Unknown fade kind: {kind!r}")


# --------------------------------------------------------------------------- #
# Forward valuation
# --------------------------------------------------------------------------- #
@dataclass
class Valuation:
    """Output of a forward valuation run."""

    enterprise_value: float
    equity_value: float
    warranted_price: float
    pv_explicit: float
    pv_terminal: float
    projection: pd.DataFrame  # year-by-year economic engine

    def upside_vs(self, market_price: float) -> float:
        """Upside (decimal) of the warranted price over a given market price."""
        return self.warranted_price / market_price - 1.0


def value_firm(f: FirmInputs,
               cfroi_override: Optional[np.ndarray] = None,
               growth_override: Optional[np.ndarray] = None) -> Valuation:
    """Run the forward economic-profit DCF and return a full Valuation.

    Year t mechanics (t = 1..N), all on the *beginning-of-year* asset base:
        GCF_t   = CFROI_t * GI_{t-1}            # gross economic cash flow
        reinv_t = g_t     * GI_{t-1}            # reinvestment to grow the base
        FCF_t   = (CFROI_t - g_t) * GI_{t-1}    # free cash flow to the firm
        GI_t    = GI_{t-1} * (1 + g_t)

    Terminal value at N (zero-spread identity, generalised for a residual
    spread CFROI_term - DR growing at g_term):
        TV_N = GI_N + (CFROI_term - DR) * GI_N * (1 + g_term) / (DR - g_term)

    When CFROI_term == DR the residual term vanishes and TV_N == GI_N, i.e. a
    firm earning exactly its cost of capital is worth its invested capital.
    """
    n = f.horizon
    dr = f.discount_rate

    cfroi = (cfroi_override if cfroi_override is not None
             else fade_path(f.cfroi, f.cfroi_term, n, f.fade, f.fade_half_life))
    growth = (growth_override if growth_override is not None
              else fade_path(f.asset_growth, f.growth_terminal, n, f.fade, f.fade_half_life))

    if len(cfroi) != n or len(growth) != n:
        raise ValueError("Override arrays must have length == horizon.")

    gi_begin = np.empty(n)
    fcf = np.empty(n)
    gi_prev = f.gross_investment
    for i in range(n):
        gi_begin[i] = gi_prev
        fcf[i] = (cfroi[i] - growth[i]) * gi_prev
        gi_prev = gi_prev * (1.0 + growth[i])
    gi_end = gi_prev  # GI_N

    years = np.arange(1, n + 1)
    discount = (1.0 + dr) ** years
    pv_fcf = fcf / discount
    pv_explicit = float(pv_fcf.sum())

    residual_spread = (f.cfroi_term - dr)
    tv = gi_end + residual_spread * gi_end * (1.0 + f.growth_terminal) / (dr - f.growth_terminal)
    pv_terminal = float(tv / (1.0 + dr) ** n)

    ev = pv_explicit + pv_terminal
    equity = ev - f.net_debt
    price = equity / f.shares_out

    projection = pd.DataFrame({
        "year": years,
        "cfroi": cfroi,
        "asset_growth": growth,
        "spread_vs_dr": cfroi - dr,
        "gi_begin": gi_begin,
        "fcf": fcf,
        "discount_factor": 1.0 / discount,
        "pv_fcf": pv_fcf,
    })

    return Valuation(
        enterprise_value=ev,
        equity_value=equity,
        warranted_price=price,
        pv_explicit=pv_explicit,
        pv_terminal=pv_terminal,
        projection=projection,
    )


# --------------------------------------------------------------------------- #
# Reverse valuation  (the "green dot")
# --------------------------------------------------------------------------- #
@dataclass
class ImpliedExpectation:
    """Result of a reverse-DCF solve."""

    target: SolveTarget
    implied_value: float          # the solved rate (decimal)
    market_price: float
    base_value: float             # the corresponding input under the base case
    history: Optional[pd.Series] = None

    @property
    def vs_history_max(self) -> Optional[float]:
        """Implied terminal CFROI minus the firm's best historical CFROI."""
        if self.history is None or self.history.empty:
            return None
        return self.implied_value - float(np.nanmax(self.history.values))

    def plausibility(self) -> str:
        """Cheap qualitative flag against the firm's own track record."""
        if self.history is None or self.history.empty:
            return "no history supplied"
        hist = self.history.values
        hi, med = float(np.nanmax(hist)), float(np.nanmedian(hist))
        x = self.implied_value
        if x > hi:
            return ("AGGRESSIVE -- price implies a return above the firm's best "
                    "historical CFROI; the market is paying for an unproven step-up.")
        if x > med:
            return ("DEMANDING -- price implies an above-median return; sustainable "
                    "only if the firm holds the better end of its historical range.")
        return ("CONSERVATIVE -- price implies a return at/below the historical "
                "median; limited expectations embedded, watch for upside.")


def _price_for_target(f: FirmInputs, target: SolveTarget, x: float) -> float:
    """Warranted price when the chosen target rate is set to x."""
    n = f.horizon
    if target == "cfroi_terminal":
        cfroi = fade_path(f.cfroi, x, n, f.fade, f.fade_half_life)
        return value_firm(f, cfroi_override=cfroi).warranted_price
    if target == "cfroi_current":
        cfroi = fade_path(x, f.cfroi_term, n, f.fade, f.fade_half_life)
        return value_firm(f, cfroi_override=cfroi).warranted_price
    if target == "cfroi_shift":
        base = fade_path(f.cfroi, f.cfroi_term, n, f.fade, f.fade_half_life)
        return value_firm(f, cfroi_override=base + x).warranted_price
    if target == "growth_terminal":
        if x >= f.discount_rate:
            return np.inf  # keep the bracket well-behaved
        g = fade_path(f.asset_growth, x, n, f.fade, f.fade_half_life)
        f2 = FirmInputs(**{**f.__dict__, "growth_terminal": x})
        return value_firm(f2, growth_override=g).warranted_price
    raise ValueError(f"Unknown solve target: {target!r}")


def implied_expectation(f: FirmInputs,
                        market_price: float,
                        target: SolveTarget = "cfroi_terminal",
                        history: Optional[Sequence[float] | pd.Series] = None,
                        bracket: tuple[float, float] = (-0.20, 0.60)) -> ImpliedExpectation:
    """Solve for the rate that makes the warranted price equal the market price.

    Default target is the TERMINAL CFROI -- the most faithful analogue of HOLT's
    green dot ("what long-run return is priced in?"). Other targets let you ask
    "what current return", "what parallel shift to the whole path", or "what
    terminal growth" the market is implying.

    Uses Brent's method; widens the bracket once if no sign change is found.
    """
    base_map = {
        "cfroi_terminal": f.cfroi_term,
        "cfroi_current": f.cfroi,
        "cfroi_shift": 0.0,
        "growth_terminal": f.growth_terminal,
    }

    def objective(x: float) -> float:
        return _price_for_target(f, target, x) - market_price

    lo, hi = bracket
    try:
        root = brentq(objective, lo, hi, xtol=1e-8, maxiter=200)
    except ValueError:
        # No sign change in the default bracket -> widen once.
        lo2, hi2 = lo - 0.40, hi + 0.40
        if target == "growth_terminal":
            hi2 = min(hi2, f.discount_rate - 1e-4)
        try:
            root = brentq(objective, lo2, hi2, xtol=1e-8, maxiter=200)
        except ValueError as exc:
            raise ValueError(
                "Could not bracket a solution. The market price may be outside the "
                "range the model can produce by flexing this single input alone. "
                "Try a different target or check the inputs."
            ) from exc

    hist_series = None
    if history is not None:
        hist_series = pd.Series(history, dtype=float)

    return ImpliedExpectation(
        target=target,
        implied_value=float(root),
        market_price=market_price,
        base_value=base_map[target],
        history=hist_series,
    )


# --------------------------------------------------------------------------- #
# Relative Wealth Chart frame (history + forecast), dashboard-ready
# --------------------------------------------------------------------------- #
def relative_wealth_frame(f: FirmInputs,
                          history: Optional[pd.DataFrame] = None,
                          implied: Optional[ImpliedExpectation] = None) -> pd.DataFrame:
    """Assemble a tidy frame for a HOLT-style Relative Wealth Chart.

    `history` (optional) is a DataFrame indexed by year with columns
    ['cfroi', 'asset_growth'] for the realised past. The forecast rows carry the
    faded path; `implied` (optional) adds the market-implied terminal CFROI as a
    single point so you can drop the "green dot" onto the top panel.
    """
    val = value_firm(f)
    fc = val.projection[["year", "cfroi", "asset_growth"]].copy()
    fc["segment"] = "forecast"
    fc["discount_rate"] = f.discount_rate

    frames = []
    if history is not None:
        h = history.copy()
        if h.index.name != "year":
            h = h.reset_index().rename(columns={h.index.name or "index": "year"})
        h = h[["year", "cfroi", "asset_growth"]].copy()
        h["segment"] = "history"
        h["discount_rate"] = f.discount_rate
        frames.append(h)
    frames.append(fc)

    out = pd.concat(frames, ignore_index=True)

    if implied is not None and implied.target.startswith("cfroi"):
        green_dot = pd.DataFrame({
            "year": [int(fc["year"].max())],
            "cfroi": [implied.implied_value],
            "asset_growth": [np.nan],
            "segment": ["market_implied"],
            "discount_rate": [f.discount_rate],
        })
        out = pd.concat([out, green_dot], ignore_index=True)

    return out


def plot_relative_wealth(frame: pd.DataFrame, title: str = "Relative Wealth Chart"):
    """Optional plotly chart of the top two RWC panels. Returns a Figure.

    Kept import-guarded so the core engine has no plotly dependency.
    """
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
    except ImportError as exc:  # pragma: no cover
        raise ImportError("plot_relative_wealth requires plotly.") from exc

    hist = frame[frame.segment == "history"]
    fc = frame[frame.segment == "forecast"]
    dot = frame[frame.segment == "market_implied"]
    dr = float(frame["discount_rate"].iloc[0])

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        subplot_titles=("Economic Return (CFROI) vs Discount Rate",
                                        "Asset Growth"))

    fig.add_bar(x=hist.year, y=hist.cfroi * 100, name="CFROI (hist)",
                marker_color="#4F6D8E", row=1, col=1)
    fig.add_bar(x=fc.year, y=fc.cfroi * 100, name="CFROI (forecast)",
                marker_color="#D9A4AF", row=1, col=1)
    fig.add_scatter(x=frame.year, y=[dr * 100] * len(frame), name="Discount Rate",
                    mode="lines", line=dict(color="#2E8B57", width=2), row=1, col=1)
    if not dot.empty:
        fig.add_scatter(x=dot.year, y=dot.cfroi * 100, name="Market-implied CFROI",
                        mode="markers", marker=dict(color="#2E8B57", size=12,
                        symbol="circle"), row=1, col=1)

    fig.add_bar(x=hist.year, y=hist.asset_growth * 100, name="Asset Growth (hist)",
                marker_color="#A11C36", row=2, col=1)
    fig.add_bar(x=fc.year, y=fc.asset_growth * 100, name="Asset Growth (forecast)",
                marker_color="#D9A4AF", row=2, col=1)

    fig.update_layout(title=title, barmode="overlay", template="plotly_white",
                      legend=dict(orientation="h", y=-0.15))
    fig.update_yaxes(title_text="%", row=1, col=1)
    fig.update_yaxes(title_text="%", row=2, col=1)
    return fig


# --------------------------------------------------------------------------- #
# Demo / self-test
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    # Illustrative firm (round numbers, not a real company).
    f = FirmInputs(
        gross_investment=10_000.0,
        cfroi=0.14,            # currently earning 14% on the asset base
        asset_growth=0.06,     # growing assets 6% p.a.
        discount_rate=0.08,    # 8% cost of capital
        net_debt=1_500.0,
        shares_out=500.0,
        horizon=10,            # high-persistence -> 10y fade
        cfroi_terminal=None,   # full fade to DR
        growth_terminal=0.025,
        fade="linear",
    )

    val = value_firm(f)
    print("=== Forward valuation ===")
    print(f"Enterprise value : {val.enterprise_value:,.0f}")
    print(f"Equity value     : {val.equity_value:,.0f}")
    print(f"Warranted price  : {val.warranted_price:,.2f}")
    print(f"  PV explicit    : {val.pv_explicit:,.0f}")
    print(f"  PV terminal    : {val.pv_terminal:,.0f}")
    print(val.projection.round(4).to_string(index=False))

    # Round-trip test: feed the warranted price back in as the market price and
    # solve for terminal CFROI -> must recover f.cfroi_term (= DR here).
    print("\n=== Reverse round-trip (sanity check) ===")
    imp = implied_expectation(f, market_price=val.warranted_price,
                              target="cfroi_terminal")
    print(f"Recovered terminal CFROI : {imp.implied_value:.4%} "
          f"(input {f.cfroi_term:.4%})")
    assert abs(imp.implied_value - f.cfroi_term) < 1e-4, "round-trip failed"

    # Now a richer market scenario: the stock trades 25% above warranted.
    market = val.warranted_price * 1.25
    hist_cfroi = [0.11, 0.12, 0.13, 0.135, 0.14, 0.142, 0.14]  # track record
    imp2 = implied_expectation(f, market_price=market, target="cfroi_terminal",
                               history=hist_cfroi)
    print("\n=== Market-implied expectation (price 25% above warranted) ===")
    print(f"Market price            : {market:,.2f}")
    print(f"Implied terminal CFROI  : {imp2.implied_value:.2%}")
    print(f"vs best historical      : {imp2.vs_history_max:+.2%}")
    print(f"Plausibility            : {imp2.plausibility()}")

    print("\nAll checks passed.")
