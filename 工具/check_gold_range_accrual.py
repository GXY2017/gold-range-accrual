# -*- coding: utf-8 -*-
# /// script
# requires-python = ">=3.11"
# dependencies = ["pywin32", "numpy", "tzdata"]
# ///
"""Exercise the open gold range accrual workbook the way a user does: set the pricing date (and data source) on the
参数 sheet, wait for Citi CVTSHIST to re-pull, then recompute every headline number with gold_range_accrual_model.py
and scan for Excel errors. Restores the original date and source at the end (unless --keep).

Usage: uv run check_gold_range_accrual.py [--book <name or path>] [--dates 2025-06-13,2019-03-15,2016-08-15]
                                          [--source 自动] [--mode 按现价平移] [--timeout 90] [--keep]
Prints one key=value line per date and a summary; exit 1 if any date mismatches or shows errors.
"""
import argparse
import datetime as dt
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_gold_range_accrual as B  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8')


def serial(d):
    return (dt.datetime.combine(d, dt.time()) - dt.datetime(1899, 12, 30)).days


def wait_refresh(app, wb, d, timeout):
    """Wait until 行情 shows a Citi close dated within the 7-day look-back of d and the values stop changing."""
    mk = wb.Worksheets('行情')
    t0, prev = time.time(), None
    while time.time() - t0 < timeout:
        pythoncom.PumpWaitingMessages()
        time.sleep(3)
        try:
            b7 = mk.Range('B7').Value2
            vals = mk.Range(f'D{B.TAG_ROW0}:D{B.TAG_LAST}').Value
        except pythoncom.com_error:
            continue
        if not isinstance(b7, float) or app.CalculationState != 0:
            continue
        got = B.as_date(b7)
        if not (d - dt.timedelta(days=7) <= got <= d):
            continue
        if sum(isinstance(x[0], float) for x in vals) == 0:
            continue
        if vals == prev:
            return got
        prev = vals
    return None


def check_date(app, wb, d, source, timeout):
    wb.Names('ra_Source').RefersToRange.Value = source
    wb.Names('ra_Date').RefersToRange.Value2 = serial(d)
    got = wait_refresh(app, wb, d, timeout) if source != '快照' else d
    if got is None:
        return dict(date=d, ok=False, why='Citi did not refresh within timeout')
    app.Calculate()
    res, _, _, terms, q, bad = B.compare_with_model(wb)
    errs = B.scan_errors(wb)
    status = wb.Worksheets('总览').Range('A3').Value
    src = wb.Worksheets('总览').Range('A4').Value
    return dict(date=d, data_date=got, ok=not bad and not errs, spot=q[B.M.SPOT], r_src=res['r_src'],
                lo=terms.lo, hi=terms.hi, fv_smile=res['fv_smile'], margin_smile=res['margin_smile'],
                avgP_smile=res['avgP_smile'], breakeven=res['breakeven_c_smile'], target_c=res['target_c_smile'],
                bad=bad, errs=errs[:5],
                status=status, source_line=src)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--book', help='workbook name or full path (default: the open gold_range_accrual_*.xlsx)')
    ap.add_argument('--dates', default='2025-06-13,2019-03-15,2016-08-15')
    ap.add_argument('--source', default='自动', choices=B.SOURCES)
    ap.add_argument('--timeout', type=int, default=90)
    ap.add_argument('--mode', choices=['固定边界', '按现价平移'], help='区间模式 for the test (restored afterwards)')
    ap.add_argument('--set', action='append', default=[], metavar='NAME=VALUE',
                    help='set a defined name for the test, e.g. ra_Fee=0.005 ra_Spread=30 (restored afterwards)')
    ap.add_argument('--keep', action='store_true', help='leave the last tested date in place')
    a = ap.parse_args()
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject('Excel.Application')
    except Exception:
        B.fail('env', 'open the workbook in Excel first', 'no running Excel')
    books = [w for w in app.Workbooks if (a.book and (w.Name == a.book or w.FullName.lower() == a.book.lower()))
             or (not a.book and w.Name.startswith('gold_range_accrual_'))]
    if not books:
        B.fail('bad_input', 'open the workbook or pass --book', 'gold range accrual workbook not open')
    wb = books[-1]
    if app.Calculation != -4105:
        B.fail('env', 'switch Excel to automatic calculation (Formulas > Calculation Options)', 'manual calculation mode')
    d0, s0 = wb.Names('ra_Date').RefersToRange.Value2, wb.Names('ra_Source').RefersToRange.Value
    mode_cell = wb.Names('ra_Mode').RefersToRange
    m0 = mode_cell.Value
    if a.mode:
        mode_cell.Value = a.mode
    saved = {}
    for kv in a.set:
        k, v = kv.split('=', 1)
        saved[k] = wb.Names(k).RefersToRange.Value
        wb.Names(k).RefersToRange.Value = float(v)
    fails = 0
    try:
        for x in a.dates.split(','):
            d = dt.date.fromisoformat(x.strip())
            r = check_date(app, wb, d, a.source, a.timeout)
            fails += not r['ok']
            print(' '.join(f'{k}={v:.6g}' if isinstance(v, float) else f'{k}={v}' for k, v in r.items()))
    finally:
        if not a.keep:
            for k, v in saved.items():
                wb.Names(k).RefersToRange.Value = v
            mode_cell.Value = m0
            wb.Names('ra_Source').RefersToRange.Value = s0
            wb.Names('ra_Date').RefersToRange.Value2 = d0
            back = wait_refresh(app, wb, B.as_date(d0), a.timeout) if s0 != '快照' else B.as_date(d0)
            res, _, _, _, _, bad = B.compare_with_model(wb)
            print('restored', B.as_date(d0), s0, 'data_date', back, 'fv_smile', f"{res['fv_smile']:.2f}", 'bad', bad)
            fails += bool(bad)
            wb.Save()
    print('---- summary ----')
    print('book', wb.Name, 'dates', a.dates, 'source', a.source, 'mode', a.mode or m0, 'set', a.set, 'failures', fails)
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
