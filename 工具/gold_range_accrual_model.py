# -*- coding: utf-8 -*-
# /// script
# requires-python = ">=3.11"
# dependencies = ["numpy"]
# ///
"""Reference model for the gold (XAUUSD) range accrual note, independent of the Excel workbook.

Note = principal-protected zero + one range-accrual coupon: each weekday fixing inside [lo, hi] accrues 1/n of the
annual coupon, a fixing outside accrues nothing. The coupon leg is a strip of European range digitals, one per
fixing, all paid at maturity, so the price depends only on the marginal
distribution of gold at each fixing date: flat (ATM term structure) and smile (piecewise-linear vol in strike from
Citi 10D/25D pillars) versions are both closed form. A Monte Carlo under the flat model adds the distribution of the
accrual fraction (P(full coupon), P(zero)) and checks the closed form.

Conventions mirror build_gold_range_accrual.py (keep the two in step; the build compares the workbook against
`price()` and fails above 1e-9 relative):
- fixings: every weekday after the pricing date up to and including pricing date + Days (holidays ignored).
- t = calendar days / 365. Pillar tenors: ON 1/365, 1W 7/365, 2W 14/365, 1M 1/12, 2M 2/12, 3M 0.25, 6M 0.5, 1Y 1.
- ATM: linear in total variance between pillars, flat vol outside. Smile: pillar vol minus ATM (vol points) linear
  in t between 1M..1Y, flat outside; pillar strikes from forward delta (premium not included, the USD-premium
  convention for XAUUSD; spot-vs-forward delta difference ignored).
- forward: carry b_k = ln(F_k/S)/tau_k from Citi outrights, linear in t, flat outside; F(t) = S exp(b(t) t).
- discount: SOFR par (ACT/360 simple) -> continuous zero ln(1+R tau 365/360)/tau, linear in t; before SOFR
  (2019-07) the USD 3M deposit rate flat. Payment at maturity, issuer funding spread added to the discount rate.

Usage: uv run gold_range_accrual_model.py --snapshot <csv> [--date YYYY-MM-DD] [terms...] [--mc 200000]
       uv run gold_range_accrual_model.py --snapshot <csv> --scan 0.03:0.15:0.01    # range half-width scan
"""
import argparse
import csv
import datetime as dt
import math
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

IV = 'COMMODITIES.IMPLIED_VOL.XAU.USD'
SPOT = 'FX.SPOT.XAU.USD.CITI'
ATM_T = [('ON', 1 / 365), ('1W', 7 / 365), ('2W', 14 / 365), ('1M', 1 / 12), ('2M', 2 / 12), ('3M', 0.25),
         ('6M', 0.5), ('1Y', 1.0)]
TEN_T = ATM_T[3:]                                   # 1M..1Y: forwards, SOFR, smile pillars
PILLARS = [('P10', 0.10, False), ('P25', 0.25, False), ('ATM', None, None), ('C25', 0.25, True), ('C10', 0.10, True)]
FWD = 'FX.FORWARD.FWD_OUTRIGHT.XAU.USD.{}.CITI'
SOFR = 'RATES.OIS.USD_SOFR.PAR.{}'
DEP = 'FX.DEPOSIT.USD.3M.CITI'
CHECK = [f'{IV}.25RR.3M', f'{IV}.25STR.3M']        # shown in the workbook as a consistency check only


def tags():
    """Citi tags in workbook order: (tag, 内容, 单位)."""
    out = [(SPOT, '黄金现价 XAUUSD', '美元/盎司')]
    out += [(FWD.format(k), f'黄金 {k} 远期', '美元/盎司') for k, _ in TEN_T]
    out += [(SOFR.format(k), f'SOFR {k} OIS', '%') for k, _ in TEN_T]
    out += [(DEP, '美元 3M 存款利率（2019-07 前替代 SOFR）', '%')]
    out += [(f'{IV}.ATM.{k}', f'{k} ATM 隐含波动率', '%') for k, _ in ATM_T]
    out += [(f'{IV}.{d}.{k}', f'{k} {d[1:]}D {"看跌" if d[0] == "P" else "看涨"}波动率', '%')
            for k, _ in TEN_T for d in ('P10', 'P25', 'C25', 'C10')]
    out += [(CHECK[0], '3M 25D 风险逆转（校验用）', '%'), (CHECK[1], '3M 25D 宽跨式（校验用）', '%')]
    return out


def ncdf(x):
    return 0.5 * math.erfc(-x / math.sqrt(2))


def npdf(x):
    return math.exp(-x * x / 2) / math.sqrt(2 * math.pi)


def ninv(p):
    """Acklam's inverse normal (|err| < 1.2e-9) refined by one Halley step -> ~1e-15."""
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02,
         -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01,
         -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00,
         4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00]
    if p < 0.02425:
        q = math.sqrt(-2 * math.log(p))
        x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    elif p > 1 - 0.02425:
        q = math.sqrt(-2 * math.log(1 - p))
        x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    else:
        q = p - 0.5
        r = q * q
        x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
    e = ncdf(x) - p
    u = e * math.sqrt(2 * math.pi) * math.exp(x * x / 2)
    return x - u / (1 + x * u / 2)


def interp(x, xs, ys):
    """Linear interpolation, flat outside [xs[0], xs[-1]]."""
    if x <= xs[0]:
        return ys[0]
    if x >= xs[-1]:
        return ys[-1]
    j = max(i for i in range(len(xs) - 1) if xs[i] <= x)
    return ys[j] + (x - xs[j]) * (ys[j + 1] - ys[j]) / (xs[j + 1] - xs[j])


class Market:
    def __init__(self, q, vol_shift=0.0):
        """q: {tag: value} in Citi units (prices, % rates, % vols); vol_shift in vol points."""
        self.S = q[SPOT]
        self.tt = [t for _, t in TEN_T]
        ok = lambda x: isinstance(x, float) and not math.isnan(x)
        # per pillar: SOFR par where Citi has it (2019-07 on), else the USD 3M deposit rate (bank credit, biased high)
        sofr = [q.get(SOFR.format(k)) for k, _ in TEN_T]
        self.r = [math.log(1 + (R if ok(R) else q[DEP]) / 100 * t * 365 / 360) / t for R, t in zip(sofr, self.tt)]
        self.r_src = 'SOFR' if all(ok(R) for R in sofr) else ('USD 3M deposit' if not any(ok(R) for R in sofr) else 'mixed')
        # carry from Citi outrights. Gaps (checked 2026-11-13): 2016-07-26..2017-02-10 and 2018-09-03..2018-11-02 for
        # every tenor, 2M also 2016-07-18..25. A missing 2M with 1M and 3M present takes their midpoint; any other
        # missing pillar takes the USD rate (gold lease rate 0).
        fwd = [q.get(FWD.format(k)) for k, _ in TEN_T]
        self.b = [math.log(F / self.S) / t if ok(F) else None for F, (_, t) in zip(fwd, TEN_T)]
        if self.b[1] is None and ok(fwd[0]) and ok(fwd[2]):
            self.b[1] = (self.b[0] + self.b[2]) / 2
        self.b = [r if b is None else b for b, r in zip(self.b, self.r)]
        self.at = [t for _, t in ATM_T]
        self.av = [(q[f'{IV}.ATM.{k}'] + vol_shift) / 100 for k, _ in ATM_T]
        self.aw = [v * v * t for v, t in zip(self.av, self.at)]
        self.off = {d: [(q[f'{IV}.{d}.{k}'] - q[f'{IV}.ATM.{k}']) / 100 for k, _ in TEN_T]
                    for d, _, _ in PILLARS if d != 'ATM'}

    def atm(self, t):
        if t <= self.at[0]:
            return self.av[0]
        if t >= self.at[-1]:
            return self.av[-1]
        return math.sqrt(interp(t, self.at, self.aw) / t)

    def fwd(self, t):
        return self.S * math.exp(interp(t, self.tt, self.b) * t)

    def rate(self, t):
        return interp(t, self.tt, self.r)

    def pillars(self, t):
        """[(K, vol)] for 10P, 25P, ATM, 25C, 10C at expiry t (ascending strikes)."""
        F, a, st = self.fwd(t), self.atm(t), math.sqrt(t)
        out = []
        for d, delta, call in PILLARS:
            v = a if d == 'ATM' else a + interp(t, self.tt, self.off[d])
            if d == 'ATM':
                K = F * math.exp(v * v * t / 2)
            elif call:
                K = F * math.exp(-v * st * ninv(delta) + v * v * t / 2)
            else:
                K = F * math.exp(v * st * ninv(delta) + v * v * t / 2)
            out.append((K, v))
        return out

    def smile(self, t, K):
        """(vol, dvol/dK): piecewise linear in strike, flat outside the 10D pillars."""
        p = self.pillars(t)
        if any(p[i][0] >= p[i + 1][0] for i in range(4)):
            raise ValueError(f'pillar strikes not ascending at t={t}: {p}')
        # same branch points as the workbook's MATCH(K, strikes, 1): K equal to the 10P strike uses segment 1
        if K < p[0][0]:
            return p[0][1], 0.0
        if K >= p[-1][0]:
            return p[-1][1], 0.0
        j = max(i for i in range(4) if p[i][0] <= K)
        (k0, v0), (k1, v1) = p[j], p[j + 1]
        s = (v1 - v0) / (k1 - k0)
        return v0 + s * (K - k0), s

    def p_above(self, t, K, smile):
        """Q(S_t > K) under the payment (T-forward = risk-neutral, deterministic rates) measure."""
        F = self.fwd(t)
        v, s = self.smile(t, K) if smile else (self.atm(t), 0.0)
        st = v * math.sqrt(t)
        d1 = (math.log(F / K) + st * st / 2) / st
        return ncdf(d1 - st) - F * math.sqrt(t) * npdf(d1) * s


def fixings(date, days):
    end = date + dt.timedelta(days=days)
    out, d = [], date + dt.timedelta(days=1)
    while d <= end:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def price(q, date, a, vol_shift=None):
    """Closed-form strip for the single-range note. a: N, days, basis, lo, hi, c, protect, spread (bp), vol_shift,
    issue, fee, target. Returns (dict of headline numbers, Market, fixing dates)."""
    m = Market(q, a.vol_shift if vol_shift is None else vol_shift)
    fx = fixings(date, a.days)
    n = len(fx)
    T = a.days / 365
    df = math.exp(-(m.rate(T) + a.spread / 1e4) * T)
    dcf = a.days / a.basis
    cost = a.N * (a.issue + a.fee)                      # client pays issue price + upfront fee
    res = dict(n=n, T=T, spot=m.S, fwdT=m.fwd(T), atmT=m.atm(T), r=m.rate(T), r_src=m.r_src, df=df, cost=cost)
    zero = a.N * a.protect * df
    unit = a.N * dcf * df                               # PV of the full coupon per unit of annual rate x avgP
    for model in ('flat', 'smile'):
        ps = [m.p_above((d - date).days / 365, a.lo, model == 'smile') -
              m.p_above((d - date).days / 365, a.hi, model == 'smile') for d in fx]
        avg = sum(ps) / n
        res[f'avgP_{model}'] = avg
        res[f'fv_{model}'] = zero + a.c * unit * avg
        res[f'margin_{model}'] = cost - res[f'fv_{model}']
        res[f'exp_cpn_{model}'] = a.c * avg
        # coupon that makes the issuer's total take (fee included) zero / equal to the target p.a.
        res[f'breakeven_c_{model}'] = (cost - zero) / (unit * avg)
        res[f'target_c_{model}'] = (cost - zero - a.target * a.N * dcf) / (unit * avg)
        # client's expected (risk-neutral) simple yield after the fee, on the coupon day-count basis
        res[f'client_yield_{model}'] = (a.N * a.protect + a.N * dcf * a.c * avg - cost) / cost * a.basis / a.days
    return res, m, fx


def monte_carlo(m, date, fx, a, paths, seed):
    """Flat model (ATM term structure, exact log-normal marginals): distribution of the accrual fraction."""
    import numpy as np
    rng = np.random.default_rng(seed)
    t = np.array([(d - date).days / 365 for d in fx])
    w = np.array([m.atm(x) ** 2 * x for x in t])
    dw = np.diff(np.concatenate([[0.0], w]))
    if (dw < -1e-12).any():
        raise ValueError('ATM total variance decreases between fixings; MC needs a non-decreasing variance')
    F = np.array([m.fwd(x) for x in t])
    frac = np.zeros(paths)
    chunk = 50000
    for s0 in range(0, paths, chunk):
        k = min(chunk, paths - s0)
        X = np.cumsum(rng.standard_normal((k, len(t))) * np.sqrt(np.maximum(dw, 0)), axis=1)
        S = F * np.exp(X - w / 2)
        frac[s0:s0 + k] = ((S >= a.lo) & (S <= a.hi)).mean(axis=1)
    cpn = a.c * frac
    out = dict(mc_avgP=float(frac.mean()), mc_avgP_se=float(frac.std(ddof=1) / math.sqrt(paths)),
               P_full=float((frac == 1).mean()), P_zero=float((frac == 0).mean()), mean_cpn=float(cpn.mean()),
               frac_q={q: float(np.quantile(frac, q)) for q in (0.1, 0.25, 0.5, 0.75, 0.9)},
               cpn_q={q: float(np.quantile(cpn, q)) for q in (0.1, 0.25, 0.5, 0.75, 0.9)})
    edges = [0, 0.2, 0.4, 0.6, 0.8, 1.0]
    out['frac_hist'] = {f'{lo:.0%}-{hi:.0%}': float(((frac >= lo) & (frac < hi)).mean()) for lo, hi in zip(edges, edges[1:])}
    out['frac_hist']['100%'] = out['P_full']
    return out


def load_snapshot(path):
    with open(path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    q = {}
    for r in rows:
        try:
            q[r['tag']] = float(r['value'])
        except ValueError:
            q[r['tag']] = float('nan')
    return q, dt.date.fromisoformat(rows[0]['data_date'])


def add_terms(ap):
    ap.add_argument('--N', type=float, default=1_000_000)
    ap.add_argument('--days', type=int, default=92)
    ap.add_argument('--basis', type=float, default=360)
    ap.add_argument('--lo', type=float, help='range lower bound (USD/oz)')
    ap.add_argument('--hi', type=float, help='range upper bound (USD/oz)')
    ap.add_argument('--c', type=float, default=0.05, help='annual coupon accrued on each fixing inside the range')
    ap.add_argument('--protect', type=float, default=1.0)
    ap.add_argument('--spread', type=float, default=0.0, help='issuer funding spread over SOFR, bp')
    ap.add_argument('--issue', type=float, default=1.0, help='issue price as a fraction of notional')
    ap.add_argument('--fee', type=float, default=0.0, help='upfront fee paid by the client, fraction of notional')
    ap.add_argument('--target', type=float, default=0.01, help='issuer target take p.a. (for the implied coupon)')
    ap.add_argument('--vol-shift', dest='vol_shift', type=float, default=0.0)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--snapshot', required=True, help='CSV tag,value,data_date')
    ap.add_argument('--date', help='pricing date (default: snapshot data_date)')
    add_terms(ap)
    ap.add_argument('--mc', type=int, default=0, help='Monte Carlo paths (0 = skip)')
    ap.add_argument('--seed', type=int, default=20261113)
    ap.add_argument('--scan', help='lo:hi:step half-widths of a range centred on spot')
    a = ap.parse_args()
    q, d0 = load_snapshot(a.snapshot)
    date = dt.date.fromisoformat(a.date) if a.date else d0
    S = q[SPOT]
    if a.scan:
        lo, hi, st = (float(x) for x in a.scan.split(':'))
        print(f'spot {S:.2f}  date {date}  days {a.days}  coupon {a.c:.2%}')
        print('half_width  lo  hi  avgP_flat  avgP_smile  margin_smile_pct  margin_pa  breakeven_c_smile')
        w = lo
        while w <= hi + 1e-12:
            a.lo, a.hi = S * (1 - w), S * (1 + w)
            r, _, _ = price(q, date, a)
            print(f'{w:6.1%} {a.lo:8.1f} {a.hi:8.1f} {r["avgP_flat"]:9.4f} {r["avgP_smile"]:9.4f} '
                  f'{r["margin_smile"] / a.N:9.3%} {r["margin_smile"] / a.N * a.basis / a.days:9.3%} {r["breakeven_c_smile"]:9.3%}')
            w += st
        return
    if a.lo is None or a.hi is None:
        ap.error('--lo and --hi are required (or use --scan)')
    r, m, fx = price(q, date, a)
    print('---- summary ----')
    for k, v in r.items():
        print(k, f'{v:.10g}' if isinstance(v, float) else v)
    if a.mc:
        out = monte_carlo(m, date, fx, a, a.mc, a.seed)
        z = (out['mc_avgP'] - r['avgP_flat']) / out['mc_avgP_se']
        print('mc_paths', a.mc, 'mc_avgP', f'{out["mc_avgP"]:.6f}', 'se', f'{out["mc_avgP_se"]:.2e}', 'z_vs_closed', f'{z:.2f}')
        for k in ('P_full', 'P_zero', 'frac_q', 'cpn_q', 'frac_hist'):
            print(k, out[k])


if __name__ == '__main__':
    main()
