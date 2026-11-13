# -*- coding: utf-8 -*-
# /// script
# requires-python = ">=3.11"
# dependencies = ["pywin32", "numpy", "tzdata"]
# ///
"""Build gold_range_accrual_<YYYYMMDD>.xlsx: a gold (XAUUSD) range accrual note (principal-protected zero plus a coupon that
accrues on each weekday fixing inside one range) split layer by layer into basic trades and priced from Citi Velocity data
(CVTSHIST through the Excel add-in). Excel-COM helpers are in xlcom.py.

All pricing is native Excel formulas (no VBA): the coupon is a strip of European range digitals, one per weekday
fixing, so flat and smile prices are closed form. gold_range_accrual_model.py is the independent reference: the build
fails if the workbook and the model differ by more than 1e-9 (relative), and its Monte Carlo fills the static
risk-neutral distribution block.

Needs: Excel with the Citi Velocity add-in and a logged-in Velocity session (else --offline uses the newest snapshot
in ../data). Leaves the workbook open in Excel.

--share builds a copy for people outside: market inputs only from --snapshot, no Citi formula, no price history, and
author / save-path metadata removed from the file (the workbook is closed after the scrub).

Usage: uv run <this file> [--date YYYY-MM-DD (default: last London 16:00 close)] [--days 92] [--c 0.05]
                          [--lo X --hi Y] [--spread bp --issue 1.0 --fee 0 --target 0.01] [--out <path.xlsx>]
                          [--replace] [--offline] [--share --snapshot <csv>]
"""
import argparse
import csv
import datetime as dt
import math
import sys
import time
from pathlib import Path
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pythoncom
import win32com.client
from win32com.shell import shell, shellcon

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from xlcom import ADDIN, BLUE, GREEN, GREY, XL_ERR, Sheet, fail, name, retry, rgb  # noqa: E402
import gold_range_accrual_model as M  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8')
TOPIC = HERE.parent
SNAP_DIR = TOPIC / 'data'
TAGS = M.tags()                    # (tag, 内容, 单位)
NT = len(TAGS)
TAG_ROW0 = 11                      # 行情!A11..
TAG_LAST = TAG_ROW0 + NT - 1
SPILL_ROW = 95                     # 行情!A95 holds =CVTSHIST(...)
SOURCES = ['自动', 'Citi 收盘', '快照']
MAXFIX = 262                       # weekday fixings in at most 366 days
FIX0 = 8                           # first fixing row on 定盘 / 第2层 / 第3层 / 第4层
FIXL = FIX0 + MAXFIX - 1
WIDTHS = [round(0.03 + 0.01 * i, 4) for i in range(13)]   # 敏感性: ±3% .. ±15%
SENS0 = 5                          # 敏感性 table rows
HELP0 = 22                         # 敏感性 helper grid first row
HIST0 = 40                         # 历史 first data row
EARLIEST = dt.date(2016, 7, 18)    # Citi XAU outrights start (spot 2014-04-21, smile pillars 2014-05-05)
MC_PATHS = 200_000
BOUNDS = [('ra_Lo', '区间下限 L'), ('ra_Hi', '区间上限 U')]
G = {t: f"'行情'!$G${TAG_ROW0 + i}" for i, (t, _, _) in enumerate(TAGS)}


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- formula builders
# This Excel (16.0 build 19127, perpetual licence) has no LET, no Formula2 and no dynamic arrays (checked
# 2026-11-13): everything is plain formulas written through Range.Formula; the smile segment index j sits in its own
# helper column so the interpolation formulas stay short.
def lin(x, xs, ys):
    """Linear interpolation of ys over ascending xs at x, flat outside (matches M.interp). x: a cell or name."""
    j = f'MATCH({x},{xs},1)'
    return (f'IF({x}<=INDEX({xs},1),INDEX({ys},1),IF({x}>=INDEX({xs},ROWS({xs})),INDEX({ys},ROWS({ys})),'
            f'INDEX({ys},{j})+({x}-INDEX({xs},{j}))*(INDEX({ys},{j}+1)-INDEX({ys},{j}))/(INDEX({xs},{j}+1)-INDEX({xs},{j}))))')


def fx_refs(r):
    """定盘 row r: forward, t, ATM, pillar strikes, pillar vols."""
    return dict(F=f"'定盘'!$D${r}", t=f"'定盘'!$C${r}", A=f"'定盘'!$E${r}", KS=f"'定盘'!$K${r}:$O${r}",
                VS=f"'定盘'!$F${r}:$J${r}", B=f"'定盘'!$B${r}")


def seg_j(k, f):
    """Smile segment: 0 below the 10D put strike, 5 at/above the 10D call strike (same branch points as M.smile)."""
    return f"IFERROR(MATCH({k},{f['KS']},1),0)"


def seg_slope(j, f):
    ks, vs = f['KS'], f['VS']
    return f'IF(OR({j}=0,{j}>=5),0,(INDEX({vs},{j}+1)-INDEX({vs},{j}))/(INDEX({ks},{j}+1)-INDEX({ks},{j})))'


def seg_vol(j, k, f):
    ks, vs = f['KS'], f['VS']
    return (f'IF({j}=0,INDEX({vs},1),IF({j}>=5,INDEX({vs},5),INDEX({vs},{j})+({k}-INDEX({ks},{j}))*'
            f'(INDEX({vs},{j}+1)-INDEX({vs},{j}))/(INDEX({ks},{j}+1)-INDEX({ks},{j}))))')


def p_smile(k, v, s, f):
    """Q(S_t > k) with smile: N(d2) − F√t φ(d1) ∂σ/∂K; v, s are cells holding σ(k) and ∂σ/∂K."""
    F, t = f['F'], f['t']
    sd = f'({v}*SQRT({t}))'
    return (f'NORM.S.DIST((LN({F}/{k})-{sd}^2/2)/{sd},TRUE)'
            f'-{F}*SQRT({t})*NORM.S.DIST((LN({F}/{k})+{sd}^2/2)/{sd},FALSE)*{s}')


def black_call(k, v, f):
    """Undiscounted Black call on the fixing-date forward."""
    F, t = f['F'], f['t']
    sd = f'({v}*SQRT({t}))'
    return f'{F}*NORM.S.DIST((LN({F}/{k})+{sd}^2/2)/{sd},TRUE)-{k}*NORM.S.DIST((LN({F}/{k})-{sd}^2/2)/{sd},TRUE)'


def flat_p(k, f):
    return f"NORM.S.DIST((LN({f['F']}/{k})-{f['A']}^2*{f['t']}/2)/({f['A']}*SQRT({f['t']})),TRUE)"


def maxabs(rng):
    return f'MAX(MAX({rng}),-MIN({rng}))'


# ---------------------------------------------------------------- sheets
def build_terms(s, wb, date, source, a):
    s.title('A1', '参数（蓝色可改；Excel 自动计算时改完即更新，手动计算模式按 F9）')
    s.put(2, 1, [["='总览'!A3&\"；\"&'总览'!A4"]])
    s.fmt('A2', bold=True)
    rows = [
        ['项目', '数值', '说明'],
        ['一、产品条款', '', ''],
        ['名义本金（美元）', 1000000, '金额随此缩放'],
        ['期限（天）', a.days, '1–366 天；到期日 = 定价日 + 天数'],
        ['计息基准（天/年）', 360, '票息 = 名义本金 × 年化票息 × 区间内定盘次数 ÷ 总定盘次数 × 天数 ÷ 计息基准'],
        ['区间下限', 0, '黄金（美元/盎司）定盘价不低于下限、不高于上限的那天计息；区间外那天不计息，第二天回到区间内继续计息'],
        ['区间上限', 0, ''],
        ['票息（年化）', a.c, '每天都在区间内时拿到的年化票息'],
        ['到期保本比例', 1, '到期最低偿付本金的比例；承担发行方信用风险'],
        ['二、行情与情景', '', ''],
        ['定价日', '', '取 Citi London 16:00 收盘；当天没有数据时取此前 7 天内最近一天。最早 2016-07-18（Citi 黄金远期的起点）；'
         '2019-07 前没有 SOFR，美元利率改用 3M 存款利率；Citi 黄金远期在 2016-07-26 至 2017-02-10、2018-09-03 至 2018-11-02 缺数据，'
         '那几天持有成本取美元利率（租借利率按 0）'],
        ['数据来源', source, '自动：用定价日的 Citi 收盘，取不到用快照。Citi 收盘：只用定价日收盘。快照：只用建表时保存的数值，不需要 Citi'],
        ['ATM 波动率平移（波动率点）', 0, '整个波动率曲面平移，例如 1 表示每个点位 +1%；做情景分析用'],
        ['区间模式', '固定边界', '固定边界：用上面的边界。按现价平移：边界乘以 现价÷区间参考现价，换定价日时看同样形状的新产品'],
        ['区间参考现价', 0, '边界对应的现价，只在“按现价平移”时使用；建表时取定价日的现价'],
        ['三、发行方收费与模型', '', ''],
        ['发行方融资利差（bp）', a.spread, '加在 SOFR 上给票据全部美元现金流贴现（默认 0），代表发行方借钱比 SOFR 贵的部分；调高后零息部分变便宜，发行方收益变大'],
        ['发行价（占本金）', a.issue, '客户买票据付的价格；常见 100%，也有 99.x% 或 100% 以上'],
        ['认购费 / 前端费用（占本金）', a.fee, '客户在发行价之外另付给银行或分销商的费用（默认 0）'],
        ['发行方目标收益（年化，占本金）', a.target, '用来反推：发行方要赚这么多时，最多能给多少票息（第1层、总览、敏感性表）'],
        ['数字期权价差宽度 ε（美元/盎司）', 0.5, '第4层：数字期权用行权价 K±ε 的普通看涨期权价差近似'],
        ['', '', ''],
        ['单条行情报价要改，就在“行情”表 F 列“手工覆盖”里填，留空用 Citi 或快照。', '', ''],
    ]
    s.put(3, 1, rows)
    s.header('A3:C3')
    for r in (4, 12, 18):
        s.fmt(f'A{r}:C{r}', bold=True, fill=GREY)
    s.fmt('B5:B11', color=BLUE)
    s.fmt('B13:B17', color=BLUE)
    s.fmt('B19:B23', color=BLUE)
    s.fmt('B5', nf='#,##0')
    s.fmt('B8:B9', nf='#,##0.00')
    s.fmt('B10', nf='0.00%')
    s.fmt('B11', nf='0%')
    s.fmt('B13', nf='yyyy-mm-dd')
    s.fmt('B15', nf='+0.00;-0.00;0')
    s.fmt('B17', nf='#,##0.00')
    s.fmt('B19', nf='0')
    s.fmt('B20:B22', nf='0.00%')
    s.rng('B13').Value2 = (dt.datetime.combine(date, dt.time()) - dt.datetime(1899, 12, 30)).days
    for cell, lst in (('B14', ','.join(SOURCES)), ('B16', '固定边界,按现价平移')):
        s.rng(cell).Validation.Delete()
        s.rng(cell).Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=lst)
    for nm, ref in [('ra_N', 'B5'), ('ra_Days', 'B6'), ('ra_Basis', 'B7'), ('ra_LoIn', 'B8'), ('ra_HiIn', 'B9'),
                    ('ra_Cpn', 'B10'), ('ra_Protect', 'B11'), ('ra_Date', 'B13'), ('ra_Source', 'B14'),
                    ('ra_VolShift', 'B15'), ('ra_Mode', 'B16'), ('ra_RefSpot', 'B17'), ('ra_Spread', 'B19'),
                    ('ra_IssuePx', 'B20'), ('ra_Fee', 'B21'), ('ra_Target', 'B22'), ('ra_Eps', 'B23')]:
        name(wb, nm, f"'参数'!${ref[0]}${ref[1:]}")
    d0 = 28
    rows = [
        ['生效参数（公式，不用改）', '', ''],
        ['平移系数', '=IF(ra_Mode="按现价平移",ra_Spot/ra_RefSpot,1)', ''],
        ['区间下限（生效）', f'=ra_LoIn*B{d0 + 1}', ''],
        ['区间上限（生效）', f'=ra_HiIn*B{d0 + 1}', ''],
        ['定盘次数 n', '=NETWORKDAYS(ra_Date+1,ra_Date+ra_Days)', '定价日之后到到期日（含）的每个工作日（周一至周五）定盘一次，不扣节假日'],
        ['计息期限系数', '=ra_Days/ra_Basis', ''],
        ['满额票息（美元）', '=ra_N*ra_Cpn*ra_DCF', '每次定盘都在区间内时的票息'],
        ['贴现利率（连续复利）', '=ra_rUSD+ra_Spread/10000', 'SOFR OIS（插值到到期日）+ 发行方融资利差'],
        ['贴现因子', '=EXP(-ra_rPay*ra_T)', '票息和本金都在到期日支付'],
        ['客户付出（美元）', '=ra_N*(ra_IssuePx+ra_Fee)', '发行价 + 认购费；发行方总收益 = 客户付出 − 公允价值'],
    ]
    s.put(d0, 1, rows)
    s.header(f'A{d0}:C{d0}')
    s.fmt(f'B{d0 + 1}', nf='0.0000')
    s.fmt(f'B{d0 + 2}:B{d0 + 3}', nf='#,##0.00')
    s.fmt(f'B{d0 + 4}', nf='0')
    s.fmt(f'B{d0 + 5}', nf='0.000000')
    s.fmt(f'B{d0 + 6}', nf='#,##0.00')
    s.fmt(f'B{d0 + 7}', nf='0.0000%')
    s.fmt(f'B{d0 + 8}', nf='0.000000')
    s.fmt(f'B{d0 + 9}', nf='#,##0.00')
    for k, nm in enumerate(['ra_Lo', 'ra_Hi', 'ra_Fix', 'ra_DCF', 'ra_Amt', 'ra_rPay', 'ra_DF', 'ra_Cost'], start=2):
        name(wb, nm, f"'参数'!$B${d0 + k}")
    s.widths({'A': 32, 'B': 14, 'C': 110})


def build_market(s, wb):
    s.title('A1', '行情（Citi Velocity，绿色 = Citi 公式取数；定价日和数据来源在“参数”表改）')
    rows = [
        ['定价日', '=ra_Date', '在“参数”表改'],
        ['数据来源', '=ra_Source', '在“参数”表改'],
        ['回看起点', '=B3-7', ''],
        ['CVTSHIST 起止', '=TEXT(B5,"yyyymmdd")', '=TEXT(B3,"yyyymmdd")'],
        ['实际数据日期', '', 'Citi 返回的最近一个收盘日'],
        ['快照日期', '', '建表时 Citi 返回的数据日期'],
    ]
    s.put(3, 1, rows)
    for c in ('B3', 'B5', 'B7', 'B8'):
        s.fmt(c, nf='yyyy-mm-dd')
    rows = [['Citi 代码', '内容', '单位', 'Citi 收盘', '快照', '手工覆盖', '采用值']]
    for i, (tag, desc, unit) in enumerate(TAGS):
        r = TAG_ROW0 + i
        snap = f'IF(ISNUMBER(E{r}),E{r},NA())'
        # snapshot only when Citi returned nothing at all; a tag with no history on the chosen date stays #N/A and
        # the derived tables substitute (USD deposit for SOFR, 1M/3M midpoint for the 2M forward): never mix dates
        close_or_snap = f'IF(ISNUMBER(D{r}),D{r},IF(COUNT($D${TAG_ROW0}:$D${TAG_LAST})=0,{snap},NA()))'
        rows.append([tag, desc, unit, '', '', '',
                     f'=IF(F{r}<>"",F{r},IF($B$4="快照",{snap},IF($B$4="Citi 收盘",D{r},{close_or_snap})))'])
    s.put(TAG_ROW0 - 1, 1, rows)
    s.header(f'A{TAG_ROW0 - 1}:G{TAG_ROW0 - 1}')
    s.fmt(f'D{TAG_ROW0}:D{TAG_LAST}', color=GREEN, nf='0.0000')
    s.fmt(f'E{TAG_ROW0}:E{TAG_LAST}', nf='0.0000')
    s.fmt(f'F{TAG_ROW0}:F{TAG_LAST}', color=BLUE, nf='0.0000')
    s.fmt(f'G{TAG_ROW0}:G{TAG_LAST}', nf='0.0000', bold=True)
    top = TAG_LAST + 3                                   # 55
    rows = [
        ['推导参数', '数值', '说明'],
        ['现价 S', f'={G[M.SPOT]}', '上表采用值（来源见参数表）'],
        ['产品期限 T（年，ACT/365）', '=ra_Days/365', '波动率、利率、远期用的年化期限'],
        ['美元利率 r(T)（连续复利）', '=' + lin('ra_T', 'ra_TT', 'ra_R'), '下表 SOFR 零息利率按期限线性插值'],
        ['到期远期 F(T)', '=ra_Spot*EXP(' + lin('ra_T', 'ra_TT', 'ra_B') + '*ra_T)', '下表持有成本按期限线性插值'],
        ['到期 ATM 波动率 σ(T)', f"=INDEX('定盘'!$E${FIX0}:$E${FIXL},ra_Fix)", '最后一个定盘日的 ATM 波动率（总方差插值）'],
        ['持有成本 − 美元利率（≈ −黄金租借利率）', '=' + lin('ra_T', 'ra_TT', 'ra_B') + '-ra_rUSD', '远期隐含；正值表示远期升水高于美元利率'],
        ['美元 OIS 贴现因子（不含融资利差）', '=EXP(-ra_rUSD*ra_T)', ''],
    ]
    s.put(top, 1, rows)
    s.header(f'A{top}:C{top}')
    s.fmt(f'B{top + 1}', nf='#,##0.00')
    s.fmt(f'B{top + 2}', nf='0.0000')
    s.fmt(f'B{top + 3}', nf='0.0000%')
    s.fmt(f'B{top + 4}', nf='#,##0.00')
    s.fmt(f'B{top + 5}:B{top + 6}', nf='0.0000%')
    s.fmt(f'B{top + 7}', nf='0.000000')
    for nm, off in [('ra_Spot', 1), ('ra_T', 2), ('ra_rUSD', 3), ('ra_FwdT', 4), ('ra_VolT', 5)]:
        name(wb, nm, f"'行情'!$B${top + off}")
    # tenor pillars 1M..1Y: forwards, carry, SOFR, smile offsets
    p0 = top + 10                                         # header row
    rows = [['期限', 't（年）', '远期（美元/盎司）', '持有成本 b = ln(F/S)/t', 'SOFR OIS（%）', '美元零息利率 r（连续复利）',
             '10D 看跌 − ATM', '25D 看跌 − ATM', '25D 看涨 − ATM', '10D 看涨 − ATM', '来源']]
    for i, (k, t) in enumerate(M.TEN_T):
        r = p0 + 1 + i
        fwd = G[M.FWD.format(k)]
        # Citi outright gaps (2016-07-26..2017-02-10, 2018-09-03..2018-11-02; 2M also 2016-07-18..25): a missing 2M
        # with 1M and 3M present takes their midpoint, any other gap takes the USD rate (lease 0); same as M.Market
        b = f'=IF(ISNUMBER(C{r}),LN(C{r}/ra_Spot)/B{r},F{r})'
        why = '"；远期缺，持有成本取美元利率（租借利率按 0）"'
        if k == '2M':
            mid = f'AND(ISNUMBER(C{r - 1}),ISNUMBER(C{r + 1}))'
            b = f'=IF(ISNUMBER(C{r}),LN(C{r}/ra_Spot)/B{r},IF({mid},(D{r - 1}+D{r + 1})/2,F{r}))'
            why = f'IF({mid},"；远期缺，用 1M、3M 中点",{why})'
        offs = [f"=({G[f'{M.IV}.{d}.{k}']}-{G[f'{M.IV}.ATM.{k}']})/100" for d in ('P10', 'P25', 'C25', 'C10')]
        rows.append([k, f'={round(t, 12)}' if k not in ('2M',) else '=2/12', f'={fwd}', b, f"={G[M.SOFR.format(k)]}",
                     f'=LN(1+IF(ISNUMBER(E{r}),E{r},{G[M.DEP]})/100*B{r}*365/360)/B{r}', *offs,
                     f'=IF(ISNUMBER(E{r}),"SOFR","美元 3M 存款利率替代")&IF(ISNUMBER(C{r}),"",{why})'])
    rows[1][1] = '=1/12'
    s.put(p0, 1, rows)
    s.header(f'A{p0}:K{p0}')
    pe = p0 + len(M.TEN_T)
    s.fmt(f'B{p0 + 1}:B{pe}', nf='0.0000')
    s.fmt(f'C{p0 + 1}:C{pe}', nf='#,##0.00')
    s.fmt(f'D{p0 + 1}:D{pe}', nf='0.0000%')
    s.fmt(f'E{p0 + 1}:E{pe}', nf='0.0000')
    s.fmt(f'F{p0 + 1}:J{pe}', nf='0.0000%')
    for nm, c in [('ra_TT', 'B'), ('ra_B', 'D'), ('ra_R', 'F'), ('ra_OffP10', 'G'), ('ra_OffP25', 'H'),
                  ('ra_OffC25', 'I'), ('ra_OffC10', 'J')]:
        name(wb, nm, f"'行情'!${c}${p0 + 1}:${c}${pe}")
    a0 = pe + 3                                            # ATM table header
    rows = [['ATM 期限', 't（年）', 'ATM 波动率（含平移）', '总方差 σ²t']]
    for i, (k, t) in enumerate(M.ATM_T):
        r = a0 + 1 + i
        tt = {'ON': '=1/365', '1W': '=7/365', '2W': '=14/365', '1M': '=1/12', '2M': '=2/12', '3M': '=0.25',
              '6M': '=0.5', '1Y': '=1'}[k]
        rows.append([k, tt, f"=({G[f'{M.IV}.ATM.{k}']}+ra_VolShift)/100", f'=C{r}^2*B{r}'])
    s.put(a0, 1, rows)
    s.header(f'A{a0}:D{a0}')
    ae = a0 + len(M.ATM_T)
    s.fmt(f'B{a0 + 1}:B{ae}', nf='0.0000')
    s.fmt(f'C{a0 + 1}:C{ae}', nf='0.00%')
    s.fmt(f'D{a0 + 1}:D{ae}', nf='0.000000')
    for nm, c in [('ra_AT', 'B'), ('ra_AV', 'C'), ('ra_AW', 'D')]:
        name(wb, nm, f"'行情'!${c}${a0 + 1}:${c}${ae}")
    c0 = ae + 2
    rows = [['校验（3M）', '由点位算', 'Citi 报价'],
            ['25D 看涨 − 25D 看跌 对 25D 风险逆转', f"={G[M.IV + '.C25.3M']}-{G[M.IV + '.P25.3M']}", f'={G[M.CHECK[0]]}'],
            ['(25D 看涨 + 25D 看跌)/2 − ATM 对 25D 宽跨式', f"=({G[M.IV + '.C25.3M']}+{G[M.IV + '.P25.3M']})/2-{G[M.IV + '.ATM.3M']}",
             f'={G[M.CHECK[1]]}']]
    s.put(c0, 1, rows)
    s.header(f'A{c0}:C{c0}')
    s.fmt(f'B{c0 + 1}:C{c0 + 2}', nf='0.000')
    if c0 + 4 >= SPILL_ROW:
        fail('bad_input', 'raise SPILL_ROW', f'行情 tables reach row {c0 + 2}, spill row {SPILL_ROW}')
    s.put(SPILL_ROW - 1, 1, [['Citi 原始返回（下面的公式会自动展开，这一片不要输入）']])
    s.fmt(f'A{SPILL_ROW - 1}', bold=True)
    s.widths({'A': 40, 'B': 22, 'C': 20, 'D': 22, 'E': 14, 'F': 22, 'G': 14, 'H': 14, 'I': 14, 'J': 14, 'K': 30})


def build_fixings(s):
    s.title('A1', '定盘与曲面：每个定盘日的远期、ATM 波动率和五个微笑点位')
    s.put(2, 1, [['t = 定盘日距定价日的日历天数 ÷ 365。远期：持有成本按期限线性插值；ATM：总方差线性插值，两端平坦；'
                  '点位波动率 = ATM + 点位减 ATM 的差（1M–1Y 之间按期限线性插值，两端平坦）；点位行权价按远期 Delta（不含期权费）换算。']])
    s.put(3, 1, [['期限超过 366 天的部分不排定盘日。最后一列检查五个点位的行权价是否从低到高排列（插值要求）。']])
    hdr = ['序号', '定盘日', 't（年）', '远期 F(t)', 'ATM σ(t)', '10D 看跌 σ', '25D 看跌 σ', 'ATM σ', '25D 看涨 σ', '10D 看涨 σ',
           '10D 看跌 K', '25D 看跌 K', 'ATM K', '25D 看涨 K', '10D 看涨 K', '行权价升序']
    rows = [hdr]
    for i in range(1, MAXFIX + 1):
        r = FIX0 + i - 1
        e = f'IF($B{r}="","",'
        rows.append([
            i, f'=IF(A{r}<=ra_Fix,WORKDAY(ra_Date,A{r}),"")', f'=IF(B{r}="","",(B{r}-ra_Date)/365)',
            f'={e}ra_Spot*EXP({lin(f"$C{r}", "ra_TT", "ra_B")}*$C{r}))',
            f'={e}IF($C{r}<=INDEX(ra_AT,1),INDEX(ra_AV,1),IF($C{r}>=INDEX(ra_AT,ROWS(ra_AT)),INDEX(ra_AV,ROWS(ra_AV)),'
            f'SQRT({lin(f"$C{r}", "ra_AT", "ra_AW")}/$C{r}))))',
            f'={e}$E{r}+{lin(f"$C{r}", "ra_TT", "ra_OffP10")})',
            f'={e}$E{r}+{lin(f"$C{r}", "ra_TT", "ra_OffP25")})',
            f'={e}$E{r})',
            f'={e}$E{r}+{lin(f"$C{r}", "ra_TT", "ra_OffC25")})',
            f'={e}$E{r}+{lin(f"$C{r}", "ra_TT", "ra_OffC10")})',
            f'={e}$D{r}*EXP(F{r}*SQRT($C{r})*NORM.S.INV(0.1)+F{r}^2*$C{r}/2))',
            f'={e}$D{r}*EXP(G{r}*SQRT($C{r})*NORM.S.INV(0.25)+G{r}^2*$C{r}/2))',
            f'={e}$D{r}*EXP(H{r}^2*$C{r}/2))',
            f'={e}$D{r}*EXP(-I{r}*SQRT($C{r})*NORM.S.INV(0.25)+I{r}^2*$C{r}/2))',
            f'={e}$D{r}*EXP(-J{r}*SQRT($C{r})*NORM.S.INV(0.1)+J{r}^2*$C{r}/2))',
            f'={e}AND(K{r}<L{r},L{r}<M{r},M{r}<N{r},N{r}<O{r}))',
        ])
    s.put(FIX0 - 1, 1, rows)
    s.header(f'A{FIX0 - 1}:P{FIX0 - 1}')
    s.fmt(f'B{FIX0}:B{FIXL}', nf='yyyy-mm-dd ddd')
    s.fmt(f'C{FIX0}:C{FIXL}', nf='0.0000')
    s.fmt(f'D{FIX0}:D{FIXL}', nf='#,##0.00')
    s.fmt(f'E{FIX0}:J{FIXL}', nf='0.00%')
    s.fmt(f'K{FIX0}:O{FIXL}', nf='#,##0.0')
    s.widths({'A': 6, 'B': 16, 'C': 9, 'D': 11, 'E': 9, 'F': 11, 'G': 11, 'H': 9, 'I': 11, 'J': 11,
              'K': 11, 'L': 11, 'M': 11, 'N': 11, 'O': 11, 'P': 10})


def build_layer3(s):
    s.title('A1', '第3层：每天的区间数字期权 = 数字看涨（下限）− 数字看涨（上限）')
    s.put(2, 1, [['数字看涨 = 定盘价高于 K 的风险中性概率 Q(S_t > K)（到期支付，确定利率下即远期测度）。平坦：N(d2)，σ = 当天 ATM。'
                  '微笑：Q(S_t>K) = N(d2) − F·√t·φ(d1)·∂σ/∂K，σ 和 ∂σ/∂K 取当天五个点位在行权价上的线性插值（两端平坦、斜率 0）。']])
    s.put(3, 1, [['偏斜修正 = −F·√t·φ(d1)·∂σ/∂K：看跌一侧波动率更高（∂σ/∂K < 0）时，修正为正，即高于 K 的概率比平坦模型大。']])
    hdr1 = ['', '', '']
    hdr2 = ['序号', '定盘日', 't（年）']
    for _, lab in BOUNDS:
        hdr1 += [lab, '', '', '', '', '', '']
        hdr2 += ['微笑段号 j', 'σ(K)', '∂σ/∂K（每美元）', 'N(d2)', '偏斜修正', 'Q(S>K)·微笑', 'Q(S>K)·平坦']
    rows = [hdr1, hdr2]
    for i in range(1, MAXFIX + 1):
        r = FIX0 + i - 1
        f = fx_refs(r)
        e = f"IF({f['B']}=\"\",\"\","
        row = [i, f"={e}{f['B']})", f"={e}{f['t']})"]
        for bi, (bn, _) in enumerate(BOUNDS):
            cj, cv, cs, cn, ck = (f'{col(L3W * bi + 4 + j)}{r}' for j in range(5))
            sd = f"({cv}*SQRT({f['t']}))"
            row += [
                f'={e}{seg_j(bn, f)})',
                f'={e}{seg_vol(cj, bn, f)})',
                f'={e}{seg_slope(cj, f)})',
                f"={e}NORM.S.DIST((LN({f['F']}/{bn})-{sd}^2/2)/{sd},TRUE))",
                f"={e}-{f['F']}*SQRT({f['t']})*NORM.S.DIST((LN({f['F']}/{bn})+{sd}^2/2)/{sd},FALSE)*{cs})",
                f'={e}{cn}+{ck})',
                f'={e}{flat_p(bn, f)})',
            ]
        rows.append(row)
    s.put(FIX0 - 2, 1, rows)
    s.header(f'A{FIX0 - 2}:{col(3 + L3W * len(BOUNDS))}{FIX0 - 1}')
    s.fmt(f'B{FIX0}:B{FIXL}', nf='yyyy-mm-dd')
    s.fmt(f'C{FIX0}:C{FIXL}', nf='0.0000')
    for bi in range(len(BOUNDS)):
        c0 = 4 + L3W * bi
        s.fmt(f'{col(c0)}{FIX0}:{col(c0)}{FIXL}', nf='0')
        s.fmt(f'{col(c0 + 1)}{FIX0}:{col(c0 + 1)}{FIXL}', nf='0.00%')
        s.fmt(f'{col(c0 + 2)}{FIX0}:{col(c0 + 2)}{FIXL}', nf='0.0E+00')
        s.fmt(f'{col(c0 + 3)}{FIX0}:{col(c0 + 6)}{FIXL}', nf='0.000000')
    s.widths({'A': 6, 'B': 11, 'C': 8})


L3W = 7               # 第3层 columns per bound
L4W = 8               # 第4层 columns per bound
SW = 8                # 敏感性 helper columns per width


def l3(bi, kind):
    """第3层 column of bound bi: kind 'smile' or 'flat'."""
    return col(4 + L3W * bi + (5 if kind == 'smile' else 6))


def l4(bi, kind):
    """第4层 column of bound bi: 'rep' (call-spread digital) or 'diff' (minus 第3层)."""
    return col(3 + L4W * bi + (6 if kind == 'rep' else 7))


def build_layer2(s):
    s.title('A1', '第2层：区间累积 = n 份每天的区间数字期权（每份 1/n），全部在到期日支付')
    s.put(2, 1, [['某天定盘价落在区间内，这一天就计入 1/n 的满额票息；所以票息部分的价值 = 满额票息 × 贴现因子 × 各定盘日落在区间内的'
                  '风险中性概率的平均数。这一层是恒等式，不依赖模型；概率来自第3层。']])
    hdr = ['序号', '定盘日', 't（年）', 'Q(区间内)·平坦', 'Q(区间内)·微笑', '当天期望计息（年化，微笑）',
           '现值贡献·微笑（美元）', '现值贡献·平坦（美元）']
    rows = [['', '平均或合计', '', f'=AVERAGE(D{FIX0}:D{FIXL})', f'=AVERAGE(E{FIX0}:E{FIXL})', f'=AVERAGE(F{FIX0}:F{FIXL})',
             f'=SUM(G{FIX0}:G{FIXL})', f'=SUM(H{FIX0}:H{FIXL})'],
            ['', '校验：现值贡献合计 − 第1层区间累积现值', '', '', '', '', "=G4-'第1层'!G6", "=H4-'第1层'!F6"],
            [''] * 8, hdr]
    for i in range(1, MAXFIX + 1):
        r = FIX0 + i - 1
        e = f"IF('第3层'!$B{r}=\"\",\"\","
        p = lambda bi, k: f"'第3层'!{l3(bi, k)}{r}"
        rows.append([i, f"=IF('第3层'!B{r}=\"\",\"\",'第3层'!B{r})", f"=IF('第3层'!B{r}=\"\",\"\",'第3层'!C{r})",
                     f'={e}{p(0, "flat")}-{p(1, "flat")})', f'={e}{p(0, "smile")}-{p(1, "smile")})',
                     f'={e}ra_Cpn*E{r})', f'={e}ra_Amt*E{r}*ra_DF/ra_Fix)', f'={e}ra_Amt*D{r}*ra_DF/ra_Fix)'])
    s.put(4, 1, rows)
    s.header(f'A{FIX0 - 1}:H{FIX0 - 1}')
    s.fmt('A4:H5', bold=True)
    s.fmt(f'B{FIX0}:B{FIXL}', nf='yyyy-mm-dd')
    s.fmt(f'C{FIX0}:C{FIXL}', nf='0.0000')
    s.fmt(f'D4:E{FIXL}', nf='0.0000%')
    s.fmt(f'F4:F{FIXL}', nf='0.000%')
    s.fmt(f'G4:H{FIXL}', nf='#,##0.00')
    s.fmt('G5:H5', nf='0.0E+00')
    s.widths({'A': 6, 'B': 36, 'C': 8, 'D': 18, 'E': 18, 'F': 22, 'G': 20, 'H': 20})


def build_layer4(s):
    s.title('A1', '第4层：每份数字看涨 ≈ 看涨期权价差 [C(K−ε) − C(K+ε)] ÷ 2ε（黄金普通看涨期权，按微笑波动率定价）')
    s.put(2, 1, [['C = 远期口径（不贴现）的 Black 看涨期权价，波动率取该行权价在当天微笑上的插值。ε 在参数表改。ε→0 时价差收敛到第3层的'
                  'Q(S>K)·微笑，所以区间累积可以用一串普通期权静态复制；跳到点位行权价附近（微笑折点）时差额会大一些。']])
    hdr1, hdr2 = ['', ''], ['序号', '定盘日']
    for _, lab in BOUNDS:
        hdr1 += [lab, '', '', '', '', '', '', '']
        hdr2 += ['段号 j(K−ε)', 'σ(K−ε)', 'C(K−ε)', '段号 j(K+ε)', 'σ(K+ε)', 'C(K+ε)', '价差复制的数字期权', '− 第3层（微笑）']
    sm = ['最大绝对差', '']
    pv = ['价差复制的数字期权平均值', '']
    for bi in range(len(BOUNDS)):
        c, c2 = l4(bi, 'diff'), l4(bi, 'rep')
        sm += [''] * (L4W - 1) + [f'={maxabs(f"{c}{FIX0}:{c}{FIXL}")}']
        pv += [''] * (L4W - 2) + [f'=AVERAGE({c2}{FIX0}:{c2}{FIXL})', '']
    rows = [sm, pv, hdr1, hdr2]                    # rows 4..7; fixing rows start at FIX0 = 8 like the other layers
    for i in range(1, MAXFIX + 1):
        r = FIX0 + i - 1
        f = fx_refs(r)
        e = f"IF({f['B']}=\"\",\"\","
        row = [i, f"={e}{f['B']})"]
        for bi, (bn, _) in enumerate(BOUNDS):
            c0 = 3 + L4W * bi
            cells = [f'{col(c0 + j)}{r}' for j in range(L4W)]
            for sign, off in (('-', 0), ('+', 3)):
                k = f'({bn}{sign}ra_Eps)'
                cj, cv = cells[off], cells[off + 1]
                row += [f'={e}{seg_j(k, f)})', f'={e}{seg_vol(cj, k, f)})', f'={e}{black_call(k, cv, f)})']
            row += [f'={e}({cells[2]}-{cells[5]})/(2*ra_Eps))', f"={e}{cells[6]}-'第3层'!{l3(bi, 'smile')}{r})"]
        rows.append(row)
    s.put(4, 1, rows)
    s.header(f'A{FIX0 - 2}:{col(2 + L4W * len(BOUNDS))}{FIX0 - 1}')
    s.fmt('A4:A5', bold=True)
    s.fmt(f'B{FIX0}:B{FIXL}', nf='yyyy-mm-dd')
    for bi in range(len(BOUNDS)):
        c0 = 3 + L4W * bi
        for off in (0, 3):
            s.fmt(f'{col(c0 + off)}{FIX0}:{col(c0 + off)}{FIXL}', nf='0')
            s.fmt(f'{col(c0 + off + 1)}{FIX0}:{col(c0 + off + 1)}{FIXL}', nf='0.00%')
            s.fmt(f'{col(c0 + off + 2)}{FIX0}:{col(c0 + off + 2)}{FIXL}', nf='#,##0.0000')
        s.fmt(f'{l4(bi, "rep")}5:{l4(bi, "rep")}{FIXL}', nf='0.000000')
        s.fmt(f'{l4(bi, "diff")}4:{l4(bi, "diff")}{FIXL}', nf='0.0E+00')
    s.widths({'A': 22, 'B': 11})


def build_layer1(s):
    s.title('A1', '第1层：票据 = 保本零息票据 + 区间累积票息')
    s.put(2, 1, [['单位价值 = 满额支付每 1 美元的现值。客户按发行价买入票据并另付认购费（参数表）；票息的期权费来自零息票据少付的利息。'
                  '平坦 = 每个定盘日用当天 ATM 波动率；微笑 = 用 10D/25D 点位插值的波动率微笑（主结果）。']])
    rows = [
        ['组成', '客户方向', '满额支付（美元）', '单位价值·平坦', '单位价值·微笑', '现值·平坦（美元）', '现值·微笑（美元）', '占本金（微笑）'],
        ['保本零息票据', '买入', '=ra_N*ra_Protect', '=ra_DF', '=ra_DF', '=C5*D5', '=C5*E5', '=G5/ra_N'],
        ['="区间累积 "&TEXT(ra_Lo,"#,##0")&"–"&TEXT(ra_Hi,"#,##0")&"，年化 "&TEXT(ra_Cpn,"0.00%")', '买入',
         '=ra_Amt', "=ra_DF*'第2层'!D4", "=ra_DF*'第2层'!E4", '=C6*D6', '=C6*E6', '=G6/ra_N'],
        ['公允价值合计', '', '', '', '', '=SUM(F5:F6)', '=SUM(G5:G6)', '=G7/ra_N'],
        ['客户付出（发行价 + 认购费）', '', '', '', '', '=ra_Cost', '=ra_Cost', ''],
        ['发行方总收益（客户付出 − 公允价值）', '', '', '', '', '=F8-F7', '=G8-G7', ''],
        ['发行方总收益占本金', '', '', '', '', '=F9/ra_N', '=G9/ra_N', ''],
        ['发行方总收益折年化', '', '', '', '', '=F10*ra_Basis/ra_Days', '=G10*ra_Basis/ra_Days', ''],
    ]
    s.put(4, 1, rows)
    s.header('A4:H4')
    s.fmt('C5:C6', nf='#,##0.00')
    s.fmt('D5:E6', nf='0.000000')
    s.fmt('F5:G9', nf='#,##0.00')
    s.fmt('F10:G11', nf='0.000%')
    s.fmt('H5:H7', nf='0.000%')
    s.fmt('A7:H7', bold=True)
    s.fmt('A9:G9', bold=True)
    rows = [
        ['发行方收益拆分（微笑）', '美元', '占本金'],
        ['发行方融资利差（bp，参数表）', '=ra_Spread', ''],
        ['发行价（参数表）', '', '=ra_IssuePx'],
        ['认购费（参数表）', '', '=ra_Fee'],
        ['发行方总收益（客户付出 − 公允价值）', '=G9', '=K8/ra_N'],
        ['其中：认购费', '=ra_N*ra_Fee', '=K9/ra_N'],
        ['其中：融资利差部分', '=G7*(EXP(ra_Spread/10000*ra_T)-1)', '=K10/ra_N'],
        ['其中：结构化收益（发行价 − 按 SOFR 估值的公允价值）', '=ra_N*ra_IssuePx-K12', '=K11/ra_N'],
        ['按 SOFR 贴现的公允价值', '=G7*EXP(ra_Spread/10000*ra_T)', '=K12/ra_N'],
        ['校验：三项之和 − 总收益', '=K9+K10+K11-K8', ''],
    ]
    s.put(4, 10, rows)
    s.header('J4:L4')
    s.fmt('K5', nf='0')
    s.fmt('L6:L7', nf='0.00%')
    s.fmt('K8:K12', nf='#,##0.00')
    s.fmt('L8:L12', nf='0.000%')
    s.fmt('K13', nf='0.0E+00')
    s.fmt('J8:L8', bold=True)
    tgt = lambda c: f'=(ra_Cost-ra_N*ra_Protect*ra_DF-ra_Target*ra_N*ra_DCF)/(ra_N*ra_DCF*ra_DF*{c}15)'
    rows = [
        ['票息与概率', '平坦', '微笑', '说明'],
        ['平均计息概率', "='第2层'!D4", "='第2层'!E4", '各定盘日落在区间内的风险中性概率的平均数'],
        ['期望年化票息', '=ra_Cpn*B15', '=ra_Cpn*C15', '风险中性期望，不是对实际走势的预测；历史上的实际结果见“历史”表'],
        ['期望票息（美元，到期）', '=ra_Amt*B15', '=ra_Amt*C15', ''],
        ['盈亏平衡票息', '=(ra_Cost-ra_N*ra_Protect*ra_DF)/(ra_N*ra_DCF*ra_DF*B15)',
         '=(ra_Cost-ra_N*ra_Protect*ra_DF)/(ra_N*ra_DCF*ra_DF*C15)', '发行方总收益（含认购费）为零时能给的年化票息'],
        ['期权预算：客户付出减保本零息现值（美元）', '=ra_Cost-ra_N*ra_Protect*ra_DF', '=ra_Cost-ra_N*ra_Protect*ra_DF', ''],
        ['按发行方目标收益反推的票息', tgt('B'), tgt('C'), '="发行方总收益折年化等于 "&TEXT(ra_Target,"0.00%")&" 时能给的年化票息"'],
        ['客户期望年化收益（扣认购费，风险中性）', '=(ra_N*ra_Protect+B17-ra_Cost)/ra_Cost*ra_Basis/ra_Days',
         '=(ra_N*ra_Protect+C17-ra_Cost)/ra_Cost*ra_Basis/ra_Days', '="到期期望收回 ÷ 客户付出 − 1，按计息基准年化；对照 SOFR "&TEXT(ra_rUSD,"0.00%")'],
    ]
    s.put(14, 1, rows)
    s.header('A14:D14')
    s.fmt('B15:C15', nf='0.00%')
    s.fmt('B16:C16', nf='0.000%')
    s.fmt('B17:C17', nf='#,##0.00')
    s.fmt('B18:C18', nf='0.000%')
    s.fmt('B19:C19', nf='#,##0.00')
    s.fmt('B20:C21', nf='0.000%')
    b0 = 24
    rows = [['边界与现价的距离', '水平', '相对现价', '相对到期远期', '折合 σ√T 个数（对远期）']]
    for i, (lab, nm) in enumerate([('区间下限', 'ra_Lo'), ('现价', 'ra_Spot'), ('到期远期', 'ra_FwdT'), ('区间上限', 'ra_Hi')]):
        r = b0 + 1 + i
        rows.append([lab, f'={nm}', f'=B{r}/ra_Spot-1', f'=LN(B{r}/ra_FwdT)', f'=D{r}/(ra_VolT*SQRT(ra_T))'])
    s.put(b0, 1, rows)
    s.header(f'A{b0}:E{b0}')
    s.fmt(f'B{b0 + 1}:B{b0 + 4}', nf='#,##0.00')
    s.fmt(f'C{b0 + 1}:D{b0 + 4}', nf='0.00%')
    s.fmt(f'E{b0 + 1}:E{b0 + 4}', nf='0.00')
    s.widths({'A': 42, 'B': 12, 'C': 16, 'D': 16, 'E': 22, 'F': 18, 'G': 18, 'H': 14, 'I': 3, 'J': 50, 'K': 14, 'L': 10})


def build_sens(s):
    s.title('A1', '敏感性：区间宽度（以现价为中心对称）对价值的影响，票息取参数表的票息')
    s.put(2, 1, [['A 列半宽可以改。发行方收益 = 客户付出（发行价 + 认购费）− 公允价值。盈亏平衡票息 = 发行方收益为零时能给的年化票息；'
                  '目标收益票息 = 发行方按参数表的目标收益（年化）收取时能给的票息。实际报价一般低于盈亏平衡票息，差额就是发行方收益。']])
    rows = [['区间半宽（±，相对现价）', '下限', '上限', '平均计息概率·平坦', '平均计息概率·微笑', '公允价值·微笑（美元）',
             '发行方收益（美元）', '占本金', '折年化', '盈亏平衡票息·微笑', '目标收益票息·微笑']]
    for j, w in enumerate(WIDTHS):
        r = SENS0 + j
        cf, cs = col(2 + SW * j), col(2 + SW * j + SW - 1)
        rows.append([round(w, 4), f'=ra_Spot*(1-A{r})', f'=ra_Spot*(1+A{r})',
                     f'=AVERAGE({cf}{HELP0}:{cf}{HELP0 + MAXFIX - 1})', f'=AVERAGE({cs}{HELP0}:{cs}{HELP0 + MAXFIX - 1})',
                     f'=ra_N*ra_Protect*ra_DF+ra_Amt*ra_DF*E{r}', f'=ra_Cost-F{r}', f'=G{r}/ra_N',
                     f'=H{r}*ra_Basis/ra_Days', f'=(ra_Cost-ra_N*ra_Protect*ra_DF)/(ra_N*ra_DCF*ra_DF*E{r})',
                     f'=(ra_Cost-ra_N*ra_Protect*ra_DF-ra_Target*ra_N*ra_DCF)/(ra_N*ra_DCF*ra_DF*E{r})'])
    s.put(SENS0 - 1, 1, rows)
    end = SENS0 + len(WIDTHS) - 1
    s.header(f'A{SENS0 - 1}:K{SENS0 - 1}')
    s.fmt(f'A{SENS0}:A{end}', nf='0.0%', color=BLUE)
    s.fmt(f'B{SENS0}:C{end}', nf='#,##0.0')
    s.fmt(f'D{SENS0}:E{end}', nf='0.00%')
    s.fmt(f'F{SENS0}:G{end}', nf='#,##0.00')
    s.fmt(f'H{SENS0}:K{end}', nf='0.000%')
    s.put(HELP0 - 2, 1, [['辅助计算：每个定盘日落在各宽度区间内的概率（每个宽度 8 列：平坦、下限段号/斜率/波动率、上限段号/斜率/波动率、微笑），不用改']])
    s.fmt(f'A{HELP0 - 2}', bold=True)
    hdr = ['序号']
    for w in WIDTHS:
        hdr += [f'±{w:.0%} 平坦', 'j 下', '∂σ/∂K 下', 'σ 下', 'j 上', '∂σ/∂K 上', 'σ 上', f'±{w:.0%} 微笑']
    rows = [hdr]
    for i in range(1, MAXFIX + 1):
        r = HELP0 + i - 1
        f = fx_refs(FIX0 + i - 1)
        e = f"IF({f['B']}=\"\",\"\","
        row = [i]
        for j in range(len(WIDTHS)):
            lo, hi = f'$B${SENS0 + j}', f'$C${SENS0 + j}'
            c = [f'{col(2 + SW * j + x)}{r}' for x in range(SW)]
            row += [f'={e}{flat_p(lo, f)}-{flat_p(hi, f)})',
                    f'={e}{seg_j(lo, f)})', f'={e}{seg_slope(c[1], f)})', f'={e}{seg_vol(c[1], lo, f)})',
                    f'={e}{seg_j(hi, f)})', f'={e}{seg_slope(c[4], f)})', f'={e}{seg_vol(c[4], hi, f)})',
                    f'={e}{p_smile(lo, c[3], c[2], f)}-({p_smile(hi, c[6], c[5], f)}))']
        rows.append(row)
    s.put(HELP0 - 1, 1, rows)
    last_c = col(1 + SW * len(WIDTHS))
    s.header(f'A{HELP0 - 1}:{last_c}{HELP0 - 1}')
    s.fmt(f'B{HELP0}:{last_c}{HELP0 + MAXFIX - 1}', nf='0.0000')
    s.widths({'A': 26, 'B': 11, 'C': 11, 'D': 18, 'E': 18, 'F': 20, 'G': 18, 'H': 10, 'I': 10, 'J': 18, 'K': 18})


def build_history_frame(s):
    s.title('A1', '历史回看：过去每个交易日按当天现价发一张相对区间、期限都相同的票据，实际能计息多少')
    s.put(2, 1, [['边界 = 当天收盘价 × (当前边界 ÷ 当前现价)；之后 n 个交易日（n = 当前定盘次数）的收盘价落在区间内的比例就是计息比例。'
                  '窗口按日滚动，相邻窗口高度重叠，分布只作参考。这是实际走势（历史测度），和第1层的风险中性概率不是一回事。']])


def build_history_data(s, hist):
    """hist: [(date, close)] ascending. Writes data, per-window formulas and the summary block."""
    n = len(hist)
    last = HIST0 + n - 1
    rows = [['日期', '收盘（美元/盎司）', '区间下限', '区间上限', '计息比例', '年化票息']]
    base = dt.datetime(1899, 12, 30)
    for i, (d, v) in enumerate(hist):
        r = HIST0 + i
        rng = f'INDEX($B$1:$B${last},{r}+1):INDEX($B$1:$B${last},{r}+ra_Fix)'
        e = f'IF({r}+ra_Fix>{last},"",'
        rows.append([(dt.datetime.combine(d, dt.time()) - base).days, v,
                     f'=B{r}*ra_Lo/ra_Spot', f'=B{r}*ra_Hi/ra_Spot',
                     f'={e}COUNTIFS({rng},">="&C{r},{rng},"<="&D{r})/ra_Fix)',
                     f'={e}ra_Cpn*E{r})'])
    s.put(HIST0 - 1, 1, rows)
    s.header(f'A{HIST0 - 1}:F{HIST0 - 1}')
    s.fmt(f'A{HIST0}:A{last}', nf='yyyy-mm-dd')
    s.fmt(f'B{HIST0}:D{last}', nf='#,##0.00')
    s.fmt(f'E{HIST0}:F{last}', nf='0.0%')
    E = f'$E${HIST0}:$E${last}'
    rows = [
        ['项目', '数值', '说明'],
        ['数据', f'="Citi FX.SPOT.XAU.USD.CITI London 16:00 收盘，"&TEXT(A{HIST0},"yyyy-mm-dd")&" 至 "&TEXT(A{last},"yyyy-mm-dd")',
         '建表时取的静态数值；重建工作簿时更新'],
        ['完整窗口数', f'=COUNT({E})', ''],
        ['每个窗口的定盘次数 n', '=ra_Fix', '与当前产品相同'],
        ['区间（相对当天现价）', '=TEXT(ra_Lo/ra_Spot-1,"+0.0%;-0.0%")&" 至 "&TEXT(ra_Hi/ra_Spot-1,"+0.0%;-0.0%")', ''],
        ['平均计息比例', f'=AVERAGE({E})', ''],
        ['全部计息的窗口占比', f'=COUNTIF({E},1)/COUNT({E})', '整段期限每天都在区间内'],
        ['零计息的窗口占比', f'=COUNTIF({E},0)/COUNT({E})', ''],
        ['计息比例 10% 分位', f'=PERCENTILE.INC({E},0.1)', ''],
        ['计息比例中位数', f'=PERCENTILE.INC({E},0.5)', ''],
        ['计息比例 90% 分位', f'=PERCENTILE.INC({E},0.9)', ''],
        ['平均年化票息', f'=AVERAGE($F${HIST0}:$F${last})', ''],
        ['对照：风险中性平均计息概率（微笑）', "='第1层'!C15", '定价用的概率'],
        ['对照：风险中性期望年化票息（微笑）', "='第1层'!C16", ''],
    ]
    s.put(4, 1, rows)
    s.header('A4:C4')
    s.fmt('B6:B7', nf='0')
    s.fmt('B9:B17', nf='0.0%')
    rows = [['计息比例', '历史', '风险中性（Monte Carlo，建表时）']]
    bins = [(0, 0.2, '0–20%'), (0.2, 0.4, '20–40%'), (0.4, 0.6, '40–60%'), (0.6, 0.8, '60–80%'), (0.8, 1.0, '80–<100%')]
    for i, (lo, hi, lab) in enumerate(bins):
        rows.append([lab, f'=COUNTIFS({E},">={lo}",{E},"<{hi}")/COUNT({E})', f"='总览'!{MC_HIST_COL}{MC_ROW0 + 9 + i}"])
    rows.append(['100%', f'=COUNTIF({E},1)/COUNT({E})', f"='总览'!{MC_HIST_COL}{MC_ROW0 + 14}"])
    s.put(21, 1, rows)
    s.header('A21:C21')
    s.fmt('B22:C27', nf='0.0%')
    ch = s.ws.ChartObjects().Add(s.ws.Range('E4').Left, s.ws.Range('E4').Top, 460, 260).Chart
    ch.ChartType = 51              # xlColumnClustered
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    for c, nm in (('B', '历史（实际走势）'), ('C', '风险中性（Monte Carlo）')):
        sr = ch.SeriesCollection().NewSeries()
        sr.Name = nm
        sr.XValues = s.ws.Range('A22:A27')
        sr.Values = s.ws.Range(f'{c}22:{c}27')
    ch.HasTitle = True
    ch.ChartTitle.Text = '计息比例的分布'
    ch.Axes(2).TickLabels.NumberFormat = '0%'
    ch.HasLegend = True
    ch.Legend.Position = -4107
    s.widths({'A': 40, 'B': 22, 'C': 26, 'D': 12, 'E': 14, 'F': 14})
    return last


def build_history_none(s):
    """--share: no price history in the file; keeps B9/B15 (read by 总览) as text so nothing shows an error."""
    labels = ['完整窗口数', '每个窗口的定盘次数 n', '区间（相对当天现价）', '平均计息比例', '全部计息的窗口占比', '零计息的窗口占比',
              '计息比例 10% 分位', '计息比例中位数', '计息比例 90% 分位', '平均年化票息']
    rows = [['项目', '数值', '说明'],
            ['数据', '未包含', '分享版不含历史行情。有 Citi Velocity 时不加 --share 重建；或在 data/ 放 '
                             'gold_spot_history_<YYYYMMDD>.csv（列 date, close），用 --offline 重建']]
    rows += [[lab, '—', ''] for lab in labels]
    s.put(4, 1, rows)
    s.header('A4:C4')
    s.widths({'A': 40, 'B': 22, 'C': 100})


def scrub_xlsx(path, extra=()):
    """--share: drop the saved-from folder (x15ac:absPath) and the author fields from the closed .xlsx package, then
    return every personal string (Windows user, home folder, Excel user name) still found in any part of it."""
    import os
    import re
    import zipfile
    path = Path(path)
    tmp = path.with_name(path.name + '.scrub')
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                data = re.sub(rb'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15"[^>]*>\s*<x15ac:absPath[^>]*/>'
                              rb'\s*</mc:Choice>\s*</mc:AlternateContent>', b'', data)
            elif item.filename == 'docProps/core.xml':
                data = re.sub(rb'<(dc:creator|cp:lastModifiedBy)>[^<]*</\1>', rb'<\1></\1>', data)
            elif item.filename == 'docProps/app.xml':
                data = re.sub(rb'<(Company|Manager)>[^<]*</\1>', rb'<\1></\1>', data)
            zout.writestr(item, data)
    os.replace(tmp, path)
    home = str(Path.home())
    needles = {x.lower() for x in (os.environ.get('USERNAME', ''), home, home.replace('\\', '/'), *extra) if x}
    leaks = []
    with zipfile.ZipFile(path) as z:
        for nm in z.namelist():
            text = z.read(nm).decode('utf-8', 'ignore').lower()
            leaks += [f'{nm}: {n}' for n in sorted(needles) if n in text]
    return leaks


MC_ROW0 = 33          # 总览 static Monte Carlo block
MC_HIST_COL = 'C'


def build_overview(s):
    s.title('A1', '黄金区间累积票据：组成与定价')
    s.put(2, 1, [['="定价日 "&TEXT(ra_Date,"yyyy-mm-dd")&"（行情 "&TEXT(\'行情\'!B7,"yyyy-mm-dd")&"）；现价 "&TEXT(ra_Spot,"#,##0.00")'
                  '&"；到期 ATM 波动率 "&TEXT(ra_VolT,"0.00%")&"；"&ra_Days&" 天、"&ra_Fix&" 次定盘；每 "&TEXT(ra_N,"#,##0")&" 美元名义本金"']])
    a3 = lambda k: f"AVERAGE('第3层'!{l3(*k)}{FIX0}:{l3(*k)}{FIXL})"
    tree = [
        # 层, 组成, 方向, 满额支付, 现值平坦, 现值微笑, 价差复制(微笑), 所在表
        (0, '黄金区间累积票据（整张）', '买入', '', "='第1层'!F7", "='第1层'!G7", '', '第1层'),
        (1, '├ 保本零息票据', '买入', "='第1层'!C5", "='第1层'!F5", "='第1层'!G5", '', '第1层'),
        (1, '="└ 区间累积票息 "&TEXT(ra_Lo,"#,##0")&"–"&TEXT(ra_Hi,"#,##0")&"，年化 "&TEXT(ra_Cpn,"0.00%")', '买入',
         '=ra_Amt', "='第1层'!F6", "='第1层'!G6", '', '第1层'),
        (2, '="  └ "&ra_Fix&" 份每天的区间数字期权：定盘落在区间内，当天计入 1/"&ra_Fix', '买入', '=ra_Amt',
         "=ra_DF*ra_Amt*'第2层'!D4", "=ra_DF*ra_Amt*'第2层'!E4", '', '第2层'),
        (3, '="    ├ 数字看涨：定盘高于 "&TEXT(ra_Lo,"#,##0")&" 付 1/"&ra_Fix', '买入', '',
         f"=ra_DF*ra_Amt*{a3((0, 'flat'))}", f"=ra_DF*ra_Amt*{a3((0, 'smile'))}", f"=ra_DF*ra_Amt*'第4层'!{l4(0, 'rep')}5", '第3层'),
        (3, '="    └ 数字看涨：定盘高于 "&TEXT(ra_Hi,"#,##0")&" 付 1/"&ra_Fix', '卖出', '',
         f"=-ra_DF*ra_Amt*{a3((1, 'flat'))}", f"=-ra_DF*ra_Amt*{a3((1, 'smile'))}", f"=-ra_DF*ra_Amt*'第4层'!{l4(1, 'rep')}5", '第3层'),
    ]
    rows = [['层', '组成', '客户方向', '满额支付（美元）', '现值·平坦（美元）', '现值·微笑（美元）', '占本金（微笑）',
             '看涨价差复制·微笑（美元）', '所在表']]
    for i, (lv, lab, side, pay, pf, ps, rep, sheet) in enumerate(tree):
        rows.append([lv, lab, side, pay, pf, ps, f'=F{7 + i}/ra_N', rep, sheet])
    s.put(6, 1, rows)
    s.header('A6:I6')
    end = 6 + len(tree)
    s.fmt(f'D7:F{end}', nf='#,##0.00;[Red]-#,##0.00')
    s.fmt(f'G7:G{end}', nf='0.000%;[Red]-0.000%')
    s.fmt(f'H7:H{end}', nf='#,##0.00;[Red]-#,##0.00')
    for rr in (7, 8, 9):
        s.fmt(f'A{rr}:I{rr}', bold=True)
    c = end + 2                                            # 14
    rows = [
        ['', '校验：保本零息 + 区间累积 − 整张票据', '', '', '=E8+E9-E7', '=F8+F9-F7'],
        ['', '校验：区间累积 = 数字看涨（下限）− 数字看涨（上限）', '', '', '=E11+E12-E9', '=F11+F12-F9'],
        ['', '校验：第4层价差复制 − 第3层（两条边界的最大差，单位：概率）', '', '', '',
         '=MAX(' + ','.join(f"'第4层'!{l4(bi, 'diff')}4" for bi in range(len(BOUNDS))) + ')'],
        ['', '="客户付出（发行价 "&TEXT(ra_IssuePx,"0.00%")&" + 认购费 "&TEXT(ra_Fee,"0.00%")&"）"', '', '', '=ra_Cost', '=ra_Cost'],
        ['', '发行方总收益（客户付出 − 公允价值）', '', '', "='第1层'!F9", "='第1层'!G9"],
        ['', '　占本金', '', '', "='第1层'!F10", "='第1层'!G10"],
        ['', '　折年化', '', '', "='第1层'!F11", "='第1层'!G11"],
        ['', '　其中：认购费', '', '', '', "='第1层'!K9"],
        ['', '="　其中：融资利差部分（"&ra_Spread&" bp，微笑）"', '', '', '', "='第1层'!K10"],
        ['', '　其中：结构化收益（发行价 − 按 SOFR 估值的公允价值，微笑）', '', '', '', "='第1层'!K11"],
        ['', '期望年化票息（风险中性）', '', '', "='第1层'!B16", "='第1层'!C16"],
        ['', '平均计息概率', '', '', "='第1层'!B15", "='第1层'!C15"],
        ['', '盈亏平衡票息（发行方总收益为零）', '', '', "='第1层'!B18", "='第1层'!C18"],
        ['', '="按发行方目标收益（年化 "&TEXT(ra_Target,"0.00%")&"）反推的票息"', '', '', "='第1层'!B20", "='第1层'!C20"],
        ['', '客户期望年化收益（扣认购费，风险中性）', '', '', "='第1层'!B21", "='第1层'!C21"],
        ['', '历史：同样相对区间的平均计息比例（见“历史”表）', '', '', '', "='历史'!B9"],
        ['', '历史：平均年化票息', '', '', '', "='历史'!B15"],
    ]
    s.put(c, 1, rows)
    s.fmt(f'E{c}:F{c + 2}', nf='0.0E+00')
    s.fmt(f'E{c + 3}:F{c + 4}', nf='#,##0.00')
    s.fmt(f'E{c + 5}:F{c + 6}', nf='0.000%')
    s.fmt(f'E{c + 7}:F{c + 9}', nf='#,##0.00')
    s.fmt(f'E{c + 10}:F{c + 16}', nf='0.00%')
    s.fmt(f'B{c + 4}:F{c + 4}', bold=True)
    s.fmt(f'B{c + 10}:F{c + 10}', bold=True)
    s.fmt(f'B{c + 13}:F{c + 13}', bold=True)
    s.put(c - 1, 5, [['平坦', '微笑']])
    s.fmt(f'E{c - 1}:F{c - 1}', bold=True)
    checks = f'MAX(ABS(F{c}),ABS(F{c + 1}),ABS(E{c}),ABS(E{c + 1}))'
    s.put(3, 1, [[
        '=IF(ra_Fix<1,"注意：期限内没有定盘日",IF(ra_Date<DATE(2016,7,18),"注意：定价日早于 2016-07-18，Citi 没有黄金远期，结果不可用",'
        f"IF(COUNTIF('定盘'!P{FIX0}:P{FIXL},FALSE)>0,\"注意：有定盘日的微笑点位行权价不是升序，插值不可靠\","
        'IF(ra_Lo>=ra_Hi,"注意：区间下限不低于上限",'
        'IF(OR(ra_Spot<=ra_Lo,ra_Spot>=ra_Hi),"注意：现价已在区间外",'
        f'IF({checks}>0.000001,"注意：拆分校验没有归零","校验通过"))))))'],
        ['="行情："&IF(\'行情\'!B4="快照","快照（"&TEXT(\'行情\'!B8,"yyyy-mm-dd")&" 收盘）",'
         f"IF(COUNT('行情'!D{TAG_ROW0}:D{TAG_LAST})>0,\"Citi \"&TEXT('行情'!B7,\"yyyy-mm-dd\")&\" London 16:00 收盘\""
         f"&IF(COUNT('行情'!D{TAG_ROW0}:D{TAG_LAST})<{NT},\"；\"&({NT}-COUNT('行情'!D{TAG_ROW0}:D{TAG_LAST}))&\" 个代码那天没有数据，已用替代口径（见行情表）\",\"\"),"
         'IF(\'行情\'!B4="Citi 收盘","Citi 未取到数据（未登录或没有加载项），可改用快照",'
         '"Citi 未取到，已改用快照（"&TEXT(\'行情\'!B8,"yyyy-mm-dd")&" 收盘）")))']])
    s.fmt('A3:A4', bold=True)
    # static Monte Carlo block (filled after the build), with a staleness flag against the stamp in K
    m0 = MC_ROW0
    k = lambda i: f'K{m0 + i}'
    rows = [
        ['', '风险中性分布（Monte Carlo，建表时按右侧参数算好的静态数值，改参数后不更新）'],
        ['', f'="状态："&IF(AND(ra_Date={k(1)},ra_Days={k(2)},ROUND(ra_Lo,6)=ROUND({k(3)},6),ROUND(ra_Hi,6)=ROUND({k(4)},6),'
             f'ra_Cpn={k(5)},ra_VolShift={k(6)},ROUND(ra_Spot,6)=ROUND({k(7)},6)),"与当前参数一致","参数或行情已改，下面是建表时的结果")'],
        ['', '模型：平坦波动率（每个定盘日的 ATM 总方差），对数正态，路径数', '', ''],
        ['', '平均计息比例（应与第1层平坦一致）', '', ''],
        ['', '全部计息的概率', '', ''],
        ['', '零计息的概率', '', ''],
        ['', '年化票息：10% 分位 / 中位数 / 90% 分位', '', ''],
        ['', '年化票息：平均', '', ''],
        ['', '计息比例分布', '', ''],
    ]
    rows += [['', lab, ''] for lab in ('0–20%', '20–40%', '40–60%', '60–80%', '80–<100%', '100%')]
    s.put(m0, 1, rows)
    s.fmt(f'B{m0}:B{m0 + 1}', bold=True)
    s.put(m0, 10, [['Monte Carlo 用的参数'], ['定价日'], ['期限（天）'], ['区间下限'], ['区间上限'], ['票息'], ['波动率平移'], ['现价']])
    s.fmt(f'J{m0}', bold=True)
    # payoff chart data (L:O)
    pts = [('=ra_Lo-(ra_Hi-ra_Lo)*0.15', 0), ('=ra_Lo', 0), ('=ra_Lo', '=ra_Cpn'), ('=ra_Hi', '=ra_Cpn'), ('=ra_Hi', 0),
           ('=ra_Hi+(ra_Hi-ra_Lo)*0.15', 0)]
    rows = [['图表数据：黄金定盘价', '当天计息（年化）', '现价', '现价线高度']] + [[x, y, '', ''] for x, y in pts]
    rows[1][2], rows[2][2] = '=ra_Spot', '=ra_Spot'
    rows[1][3], rows[2][3] = 0, '=ra_Cpn*1.15'
    s.put(6, 12, rows)
    s.header('L6:O6')
    s.fmt('L7:L12', nf='#,##0')
    s.fmt('M7:M12', nf='0.00%')
    s.fmt('N7:N8', nf='#,##0')
    s.widths({'A': 5, 'B': 62, 'C': 9, 'D': 16, 'E': 16, 'F': 16, 'G': 12, 'H': 20, 'I': 8, 'J': 14, 'K': 14,
              'L': 18, 'M': 14, 'N': 10, 'O': 10})
    return c


def add_overview_charts(s, wb):
    ws = s.ws
    ch = ws.ChartObjects().Add(ws.Range('L16').Left, ws.Range('L16').Top, 480, 270).Chart
    ch.ChartType = 75          # xlXYScatterLinesNoMarkers
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    s1 = ch.SeriesCollection().NewSeries()
    s1.Name = '某天定盘价落在这里时，这一天的计息（年化）'
    s1.XValues = ws.Range('L7:L12')
    s1.Values = ws.Range('M7:M12')
    s2 = ch.SeriesCollection().NewSeries()
    s2.Name = '定价日现价'
    s2.XValues = ws.Range('N7:N8')
    s2.Values = ws.Range('O7:O8')
    s2.Format.Line.DashStyle = 4
    ch.HasTitle = True
    ch.ChartTitle.Text = '每天的计息：定盘价在区间内按年化票息计息，区间外当天不计息'
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.Axes(2).TickLabels.NumberFormat = '0%'
    # probability by fixing date (dynamic length via names)
    name(wb, 'ra_cX', f"OFFSET('第2层'!$B${FIX0},0,0,ra_Fix,1)")
    name(wb, 'ra_cF', f"OFFSET('第2层'!$D${FIX0},0,0,ra_Fix,1)")
    name(wb, 'ra_cS', f"OFFSET('第2层'!$E${FIX0},0,0,ra_Fix,1)")
    ch = ws.ChartObjects().Add(ws.Range('L36').Left, ws.Range('L36').Top, 480, 270).Chart
    ch.ChartType = 4           # xlLine
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    for nm, rngname in (('平坦（ATM）', 'ra_cF'), ('微笑', 'ra_cS')):
        sr = ch.SeriesCollection().NewSeries()
        sr.Name = nm
        sr.Values = f"='{wb.Name}'!{rngname}"
        sr.XValues = f"='{wb.Name}'!ra_cX"
    ch.HasTitle = True
    ch.ChartTitle.Text = '每个定盘日落在区间内的风险中性概率'
    ch.Axes(2).TickLabels.NumberFormat = '0%'
    ch.Axes(2).MinimumScale = 0
    ch.Axes(2).MaximumScale = 1
    ch.Axes(1).TickLabels.NumberFormat = 'mm-dd'
    ch.HasLegend = True
    ch.Legend.Position = -4107


def build_notes(s):
    s.title('A1', '说明')
    lines = [
        '用途：黄金（XAUUSD）区间累积票据定价。票据 = 保本零息票据 + 区间累积票息；工作簿把它一层层拆成基本交易，用 Citi Velocity 行情定价。',
        '条款：每个工作日定盘一次（定价日之后到到期日，含到期日，不扣节假日），共 n 次。定盘价落在区间内（含边界）的那天计息，区间外的那天不计息，'
        '第二天回到区间内继续计息。票息 = 名义本金 × 年化票息 × 区间内天数 ÷ n × 天数 ÷ 计息基准，到期和本金一起支付。',
        '第1层：票据 = 保本零息 + 区间累积。客户付出 = 发行价 + 认购费（参数表），发行方总收益 = 客户付出 − 公允价值，右侧拆成认购费、'
        '融资利差部分和结构化收益；下面给出期望票息、盈亏平衡票息、按目标收益反推的票息和客户扣费后的期望年化收益。',
        '第2层：区间累积 = n 份每天的区间数字期权，每份付 1/n 的满额票息（恒等式，不依赖模型）。',
        '第3层：每份区间数字期权 = 数字看涨（下限）− 数字看涨（上限）。数字看涨的价值是 Q(S_t > K) × 贴现因子；平坦模型用当天 ATM 波动率，'
        '微笑模型多一项偏斜修正 −F·√t·φ(d1)·∂σ/∂K。',
        '第4层：每份数字看涨 ≈ 一对行权价 K−ε、K+ε 的普通看涨期权价差 ÷ 2ε。票息只取决于每个定盘日的单日分布，所以整条票息可以用普通期权静态复制，'
        '价格直接含波动率微笑。',
        '行情：Citi CVTSHIST London 16:00 收盘：黄金现价、1M/2M/3M/6M/1Y 远期、SOFR OIS 同样五个期限、ATM 波动率（ON 到 1Y 八个期限）、'
        '10D/25D 看跌看涨波动率（1M 到 1Y），另有美元 3M 存款利率（SOFR 之前的替代）和 3M 25D 风险逆转、宽跨式（只作校验）。',
        '利率与远期：远期由 Citi 黄金远期直接给出（2016-07-26 至 2017-02-10、2018-09-03 至 2018-11-02 缺数据，那几天持有成本取美元利率，即租借利率按 0），'
        '持有成本 b = ln(F/S)/t 按期限线性插值；b 减美元利率约等于负的黄金租借利率。贴现用 SOFR OIS（ACT/360 单利换成连续复利，按期限线性插值），'
        '加参数表的发行方融资利差。',
        '波动率：ATM 按总方差线性插值；微笑点位以“点位减 ATM”的差在 1M–1Y 之间按期限插值、1M 以内沿用 1M；点位行权价按远期 Delta、'
        '不含期权费的惯例换算（XAUUSD 期权费以美元计），这个惯例是假设，Citi 没有标注。行权价上线性插值，10D 以外取平。',
        '模型局限：线性插值的微笑在点位处有折点，折点附近第4层的价差复制差额会大一些；不含跳空；节假日没有扣除；发行方信用风险不计；'
        '风险中性概率不是对实际走势的预测（“历史”表给实际走势的对照）。',
        '敏感性：以现价为中心的对称区间，半宽 3%–15%；给出公允价值、发行方收益、盈亏平衡票息和按目标收益反推的票息。A 列可以改。',
        '历史：Citi 黄金现价自 2014-04-21 的每日收盘。每个交易日按当天价格套用当前的相对区间，统计之后 n 个交易日落在区间内的比例。',
        '风险中性分布（总览下方）：建表时用 gold_range_accrual_model.py 跑 20 万条路径的 Monte Carlo（平坦波动率），数值是静态的；状态行会提示参数是否已改。',
        '校验：建表脚本用独立的 Python 模型（工具/gold_range_accrual_model.py）重算平坦和微笑价值、发行方收益、平均计息概率、各种票息和敏感性表的一行，'
        '相对差超过 1e-9 即报错。换定价日的测试用 工具/check_gold_range_accrual.py。',
        '颜色：蓝字 = 可以改的输入；绿字 = Citi 公式取数；黑字 = 公式。',
        '参数表（蓝字）可改：名义本金、期限、计息基准、区间上下限、票息、保本比例；定价日、数据来源（自动 / Citi 收盘 / 快照）、波动率平移、区间模式；'
        '发行方融资利差、发行价、认购费、发行方目标收益、数字期权价差宽度。单条行情可在行情表 F 列手工覆盖。',
        '换日期：改“参数”表的定价日（需 Excel 已登录 Velocity），数据来源用“自动”或“Citi 收盘”，几秒后整本自动重算。最早 2016-07-18。'
        '这是“那天新发一张同样条款的票据”的价格；固定边界离当时现价太远时，可把区间模式切到“按现价平移”。',
        '重建工作簿：uv run 工具/build_gold_range_accrual.py --replace'
        '（默认定价日为最近一个 London 16:00 收盘日，区间为现价 ±6% 取整到 50 美元，票息 5%）。需 Excel 已登录 Velocity，否则加 --offline 用快照；'
        '--share --snapshot <csv> 生成不含 Citi 公式、历史数据和作者信息的分享版。',
    ]
    s.put(3, 1, [[x] for x in lines])
    s.widths({'A': 150})
    s.fmt(f'A3:A{2 + len(lines)}', wrap=True)


# ---------------------------------------------------------------- Citi plumbing
def wait_table(ws, top, ncols, min_rows, timeout):
    """Wait for a CVTSHIST spill at row `top`: returns (header_row, values) once the header has ncols CLOSE columns,
    at least min_rows dated rows are present and two reads 2 s apart agree; (None, reason) otherwise."""
    t0, seen = time.time(), None
    while time.time() - t0 < timeout:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)
        a = retry(lambda: ws.Range(f'A{top}').Value)
        if isinstance(a, int) and a in XL_ERR:
            return None, XL_ERR[a]
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        v = retry(lambda: ws.Range(ws.Cells(top, 1), ws.Cells(max(last, top + 2), ncols + 1)).Value)
        for i, row in enumerate(v):
            if sum(isinstance(c, str) and c.endswith(' - CLOSE') for c in row) == ncols:
                data = [r for r in v[i + 1:] if r[0] is not None]
                if len(data) >= min_rows:
                    if seen == v:
                        return top + i, v
                    seen = v
                    time.sleep(2.0)
                break
    return None, 'timeout'


def as_date(x):
    if isinstance(x, (int, float)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(x))).date()
    return dt.date(x.year, x.month, x.day)


def save_snapshot(mk):
    d = as_date(mk.Range('B7').Value2)
    vals = [r[0] for r in mk.Range(f'D{TAG_ROW0}:D{TAG_LAST}').Value]
    out = SNAP_DIR / f'gold_ra_citi_{d:%Y%m%d}.csv'
    with open(out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['tag', 'value', 'data_date', 'source'])
        for (t, _, _), x in zip(TAGS, vals):
            w.writerow([t, x if isinstance(x, float) else '', f'{d:%Y-%m-%d}', 'Citi Velocity CVTSHIST CLOSE (London 16:00)'])
    print('snapshot saved', out.name)


def load_snapshot(path, date):
    files = [Path(path)] if path else [p for p in sorted(SNAP_DIR.glob('gold_ra_citi_*.csv'), reverse=True)
                                       if p.stem[-8:] <= f'{date:%Y%m%d}']
    for p in files:
        q, d = M.load_snapshot(p)
        if M.SPOT in q and not math.isnan(q[M.SPOT]):
            return q, d, p
    return None


def save_history(hist, d):
    out = SNAP_DIR / f'gold_spot_history_{d:%Y%m%d}.csv'
    with open(out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['date', 'close', 'source'])
        for x, v in hist:
            w.writerow([f'{x:%Y-%m-%d}', v, 'Citi Velocity FX.SPOT.XAU.USD.CITI CLOSE (London 16:00)'])
    print('history saved', out.name, len(hist), 'rows')


def load_history(date):
    for p in sorted(SNAP_DIR.glob('gold_spot_history_*.csv'), reverse=True):
        if p.stem[-8:] <= f'{date:%Y%m%d}':
            with open(p, encoding='utf-8-sig') as f:
                rows = list(csv.DictReader(f))
            return [(dt.date.fromisoformat(r['date']), float(r['close'])) for r in rows if dt.date.fromisoformat(r['date']) <= date], p
    return None, None


def pull_history(wb, end, timeout):
    """Gold spot closes 2014-01-01..end through CVTSHIST on a scratch sheet -> [(date, close)] ascending."""
    tmp = wb.Worksheets.Add()
    tmp.Name = '临时'
    tmp.Range('A1').Value = M.SPOT
    tmp.Range('A3').Formula = f'=CVTSHIST($A$1,"DAILY",,"20140101","{end:%Y%m%d}","CLOSE")'
    hr, v = wait_table(tmp, 3, 1, 2000, timeout)
    hist = None
    if hr is not None:
        data = v[hr - 3 + 1:]
        hist = sorted((as_date(r[0]), float(r[1])) for r in data if r[0] is not None and isinstance(r[1], float))
    wb.Application.DisplayAlerts = False
    tmp.Delete()
    return hist, v if hr is None else None


def scan_errors(wb):
    """Every Excel error value on the pricing sheets (行情 excluded: tags without history are #N/A by design)."""
    errs = []
    for sheet in ('总览', '参数', '定盘', '第1层', '第2层', '第3层', '第4层', '敏感性', '历史'):
        ur = wb.Worksheets(sheet).UsedRange
        r0, c0 = ur.Row, ur.Column
        vals = retry(lambda: ur.Value)
        for i, row in enumerate(vals):
            for j, c in enumerate(row):
                if isinstance(c, int) and c in XL_ERR:
                    errs.append(f'{sheet}!{col(c0 + j)}{r0 + i}: {XL_ERR[c]}')
    return errs


def nv(wb, nm):
    """Value of a workbook-level defined name."""
    return retry(lambda: wb.Names(nm).RefersToRange.Value)


def compare_with_model(wb, tol=1e-9):
    """Recompute the open workbook's headline numbers with gold_range_accrual_model.py (inputs read by defined name,
    market values from 行情 G). Returns (res, market, fixings, terms, q, bad); bad lists every cell off by more than
    tol (relative, floor 1). Used by the build and by check_gold_range_accrual.py (date switching)."""
    vals = retry(lambda: wb.Worksheets('行情').Range(f'G{TAG_ROW0}:G{TAG_LAST}').Value)
    q = {t: (x[0] if isinstance(x[0], float) else float('nan')) for (t, _, _), x in zip(TAGS, vals)}
    date = as_date(nv(wb, 'ra_Date'))
    g = lambda nm: nv(wb, nm)
    terms = SimpleNamespace(N=g('ra_N'), days=int(g('ra_Days')), basis=g('ra_Basis'), lo=g('ra_Lo'), hi=g('ra_Hi'),
                            c=g('ra_Cpn'), protect=g('ra_Protect'), spread=g('ra_Spread'), vol_shift=g('ra_VolShift'),
                            issue=g('ra_IssuePx'), fee=g('ra_Fee'), target=g('ra_Target'))
    res, m, fx = M.price(q, date, terms)
    gv = lambda sheet, addr: retry(lambda: wb.Worksheets(sheet).Range(addr).Value)
    got = {'fv_flat': gv('第1层', 'F7'), 'fv_smile': gv('第1层', 'G7'), 'margin_flat': gv('第1层', 'F9'),
           'margin_smile': gv('第1层', 'G9'), 'avgP_flat': gv('第2层', 'D4'), 'avgP_smile': gv('第2层', 'E4'),
           'df': g('ra_DF'), 'cost': g('ra_Cost'),
           'breakeven_c_flat': gv('第1层', 'B18'), 'breakeven_c_smile': gv('第1层', 'C18'),
           'target_c_flat': gv('第1层', 'B20'), 'target_c_smile': gv('第1层', 'C20'),
           'client_yield_flat': gv('第1层', 'B21'), 'client_yield_smile': gv('第1层', 'C21')}
    bad = [(k, x, res[k]) for k, x in got.items()
           if not isinstance(x, float) or abs(x - res[k]) > tol * max(1.0, abs(res[k]))]
    # one sensitivity row (±6%)
    j6 = WIDTHS.index(0.06)
    spot = q[M.SPOT]
    t1 = SimpleNamespace(**{**vars(terms), 'lo': spot * (1 - WIDTHS[j6]), 'hi': spot * (1 + WIDTHS[j6])})
    r6, _, _ = M.price(q, date, t1)
    for k, c in (('avgP_smile', 'E'), ('breakeven_c_smile', 'J'), ('target_c_smile', 'K')):
        x = gv('敏感性', f'{c}{SENS0 + j6}')
        if not isinstance(x, float) or abs(x - r6[k]) > tol * max(1.0, abs(r6[k])):
            bad.append((f'sens_{k}', x, r6[k]))
    if res['n'] != int(g('ra_Fix')):
        bad.append(('n', g('ra_Fix'), res['n']))
    return res, m, fx, terms, q, bad


def last_london_close():
    now = dt.datetime.now(ZoneInfo('Europe/London'))
    d = now.date()
    if now.time() < dt.time(16, 30):
        d -= dt.timedelta(days=1)
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def r50(x):
    return round(x / 50) * 50


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--date', help='pricing date YYYY-MM-DD (default: last London 16:00 close)')
    ap.add_argument('--days', type=int, default=92)
    ap.add_argument('--c', type=float, default=0.05, help='annual coupon accrued on each fixing inside the range')
    ap.add_argument('--spread', type=float, default=0.0, help='issuer funding spread over SOFR, bp')
    ap.add_argument('--issue', type=float, default=1.0, help='issue price, fraction of notional (1.0 = 100%%)')
    ap.add_argument('--fee', type=float, default=0.0, help='upfront fee paid by the client, fraction of notional')
    ap.add_argument('--target', type=float, default=0.01, help='issuer target take p.a., used to back out the coupon')
    ap.add_argument('--lo', type=float, help='range lower bound (default: spot x 0.94 rounded to 50)')
    ap.add_argument('--hi', type=float, help='range upper bound (default: spot x 1.06 rounded to 50)')
    ap.add_argument('--out', help='default: ../gold_range_accrual_<pricing date>.xlsx')
    ap.add_argument('--timeout', type=int, default=120, help='seconds to wait for each Citi spill')
    ap.add_argument('--offline', action='store_true', help='skip Citi; newest snapshot and history CSV on or before --date')
    ap.add_argument('--snapshot', help='snapshot CSV (tag,value,data_date) to use when Citi is unavailable')
    ap.add_argument('--source', default='自动', choices=SOURCES)
    ap.add_argument('--share', action='store_true',
                    help='copy for sharing: inputs only from --snapshot, no Citi formula, no history, metadata scrubbed')
    ap.add_argument('--replace', action='store_true', help='close the existing output if open and send it to the Recycle Bin')
    a = ap.parse_args()
    if a.share:
        if not a.snapshot:
            fail('bad_input', 'pass --snapshot <csv> with the inputs the shared copy should carry', '--share needs --snapshot')
        a.offline, a.source = True, '快照'
    if not 1 <= a.days <= 366:
        fail('bad_input', 'pass --days between 1 and 366', f'--days {a.days}')
    date = dt.date.fromisoformat(a.date) if a.date else last_london_close()
    if date < EARLIEST:
        fail('bad_input', f'pick a date on or after {EARLIEST}', f'Citi has no gold forwards before {EARLIEST}')
    out = Path(a.out) if a.out else TOPIC / f'gold_range_accrual_{date:%Y%m%d}.xlsx'
    if out.exists() and not a.replace:
        fail('bad_input', 'rerun with --replace (old file goes to the Recycle Bin) or pass --out', f'{out} already exists')

    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject('Excel.Application')
    except Exception:
        app = win32com.client.Dispatch('Excel.Application')
    app.Visible = True
    addin_path = None
    for ad in app.AddIns2:
        try:
            if ad.Name.lower() == ADDIN.lower():
                addin_path = ad.FullName
                if not ad.Installed:
                    ad.Installed = True
        except Exception:
            pass
    if not addin_path and not a.offline:
        fail('env', 'register the Citi Velocity add-in in Excel, or pass --offline', f'{ADDIN} not found in Excel AddIns2')
    for w in list(app.Workbooks):     # unsaved leftovers of an earlier failed run
        try:
            if not w.Path and {'总览', '定盘', '第4层'} <= {x.Name for x in w.Worksheets}:
                w.Close(SaveChanges=False)
        except Exception:
            pass
    if out.exists():
        for w in list(app.Workbooks):
            try:
                if w.FullName.lower() == str(out).lower():
                    w.Close(SaveChanges=False)
            except Exception:
                pass
        rc, _ = shell.SHFileOperation((0, shellcon.FO_DELETE, str(out), None,
                                       shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION |
                                       shellcon.FOF_SILENT | shellcon.FOF_NOERRORUI, None, None))
        if rc or out.exists():
            fail('env', 'close the workbook in Excel and rerun', f'could not recycle {out} (rc={rc})')
        print('old workbook sent to the Recycle Bin')
    wb = app.Workbooks.Add()
    app.DisplayAlerts = False
    try:
        build(app, wb, a, out, date, addin_path)
    except BaseException:
        for fn in (lambda: setattr(app, 'ScreenUpdating', True), lambda: setattr(app, 'Calculation', -4105),
                   lambda: wb.Close(SaveChanges=False)):
            try:
                fn()
            except Exception:
                pass
        raise


def build(app, wb, a, out, date, addin_path):
    order = ['总览', '参数', '行情', '定盘', '第1层', '第2层', '第3层', '第4层', '敏感性', '历史', '说明']
    wb.Worksheets(1).Name = order[0]
    while wb.Worksheets.Count > 1:
        wb.Worksheets(2).Delete()
    for nm in order[1:]:
        wb.Worksheets.Add().Name = nm
    for nm in reversed(order):
        wb.Worksheets(nm).Move(wb.Worksheets(1))
    got = [x.Name for x in wb.Worksheets]
    if got != order:
        fail('env', 'check sheet ordering code', f'sheet order {got}')
    sh = {nm: Sheet(wb.Worksheets(nm)) for nm in order}
    app.Calculation = -4135
    app.ScreenUpdating = False
    t0 = time.time()
    build_market(sh['行情'], wb)
    build_terms(sh['参数'], wb, date, a.source, a)
    build_fixings(sh['定盘'])
    build_layer3(sh['第3层'])
    build_layer2(sh['第2层'])
    build_layer4(sh['第4层'])
    build_layer1(sh['第1层'])
    build_sens(sh['敏感性'])
    build_history_frame(sh['历史'])
    chk = build_overview(sh['总览'])
    build_notes(sh['说明'])
    app.ScreenUpdating = True
    print(f'sheets written in {time.time() - t0:.0f}s')

    mk = sh['行情'].ws
    app.Calculation = -4105
    app.CalculateFull()
    formula = f'=CVTSHIST($A${TAG_ROW0}:$A${TAG_LAST},"DAILY",,$B$6,$C$6,"CLOSE")'
    live = False
    if not a.offline:
        mk.Range(f'A{SPILL_ROW}').Formula = formula
        hr, v = wait_table(mk, SPILL_ROW, NT, 1, a.timeout)
        if hr is None and v == '#NAME?':
            app.RegisterXLL(addin_path)
            mk.Range(f'A{SPILL_ROW}').Formula = formula
            hr, v = wait_table(mk, SPILL_ROW, NT, 1, a.timeout)
        live = hr is not None
    else:
        hr, v = None, 'skipped (--offline)'
    if not live:
        snap = load_snapshot(a.snapshot, date)
        if snap is None:
            fail('auth', 'log in to Citi Velocity in Excel, or pass --snapshot <csv>',
                 f'Citi unavailable ({v}) and no gold_ra_citi_*.csv on or before {date} in {SNAP_DIR}')
        snap_q, snap_date, snap_path = snap
        print(f'Citi unavailable ({v}); using snapshot {snap_path.name} ({snap_date})')
        if a.share:
            mk.Range(f'A{SPILL_ROW}').Value = f'分享版没有 Citi 公式。有 Citi Velocity 加载项时，在本格输入 {formula}'
        elif a.offline:
            mk.Range(f'A{SPILL_ROW}').Formula = formula   # keeps the workbook able to refresh once Citi is back
        hr = SPILL_ROW + 1
    d0, d1 = hr + 1, hr + 12
    last_c = col(NT + 1)
    mk.Range('B7').Formula = f'=MAX($A${d0}:$A${d1})'
    for i in range(NT):
        r = TAG_ROW0 + i
        cell = (f'INDEX($B${d0}:${last_c}${d1},MATCH($B$7,$A${d0}:$A${d1},0),'
                f'MATCH($A{r}&" - CLOSE",$B${hr}:${last_c}${hr},0))')
        mk.Range(f'D{r}').Formula = f'=IFERROR(IF({cell}="",NA(),{cell}),NA())'
    app.CalculateFull()
    if live:
        retry(lambda: setattr(mk.Range(f'E{TAG_ROW0}:E{TAG_LAST}'), 'Value', mk.Range(f'D{TAG_ROW0}:D{TAG_LAST}').Value))
        mk.Range('B8').Value2 = mk.Range('B7').Value2
        save_snapshot(mk)
        data_date = as_date(mk.Range('B7').Value2)
    else:
        mk.Range(f'E{TAG_ROW0}:E{TAG_LAST}').Value = [[None if math.isnan(snap_q.get(t, float('nan'))) else snap_q[t]]
                                                      for t, _, _ in TAGS]
        mk.Range('B8').Value2 = (dt.datetime.combine(snap_date, dt.time()) - dt.datetime(1899, 12, 30)).days
        with open(snap_path, encoding='utf-8-sig') as f:
            src = next(csv.DictReader(f)).get('source') or ''
        mk.Range('C8').Value = f'快照文件 {snap_path.name}' + (f'：{src}' if src else '')
        data_date = snap_date
    app.CalculateFull()
    spot = mk.Range(f'G{TAG_ROW0}').Value
    if not isinstance(spot, float):
        fail('bad_input', 'check the Citi spill / snapshot', f'no gold spot value: {spot}')
    lo = a.lo if a.lo is not None else r50(spot * 0.94)
    hi = a.hi if a.hi is not None else r50(spot * 1.06)
    for nm, x in (('ra_LoIn', lo), ('ra_HiIn', hi), ('ra_RefSpot', round(spot, 2))):
        wb.Names(nm).RefersToRange.Value = x
    app.CalculateFull()

    # history: Citi (static values) or the newest saved CSV
    hist = None
    if live:
        hist, why = pull_history(wb, date, a.timeout)
        if hist:
            save_history(hist, date)
        else:
            print('history pull failed:', why)
    if a.share:
        build_history_none(sh['历史'])
    else:
        if not hist:
            hist, hp = load_history(date)
            if not hist:
                fail('auth', 'log in to Citi Velocity, or keep a gold_spot_history_*.csv in data/', 'no gold history available')
            print('history from', hp.name)
        wb.Worksheets('历史').Activate()
        build_history_data(sh['历史'], hist)
    app.CalculateFull()

    # independent model check
    gv = lambda sheet, addr: retry(lambda: wb.Worksheets(sheet).Range(addr).Value)
    res, m, fx, terms, q, bad = compare_with_model(wb)
    if bad:
        fail('bad_input', 'fix the workbook formulas or the model so they agree', f'workbook vs model: {bad}')
    print('model check ok: workbook = gold_range_accrual_model.py within 1e-9')

    mc = M.monte_carlo(m, date, fx, terms, MC_PATHS, 20261113)
    z = (mc['mc_avgP'] - res['avgP_flat']) / mc['mc_avgP_se']
    ov = wb.Worksheets('总览')
    m0 = MC_ROW0
    qq = mc['cpn_q']
    ov.Range(f'C{m0 + 2}').Value = MC_PATHS
    ov.Range(f'C{m0 + 3}').Value = mc['mc_avgP']
    ov.Range(f'D{m0 + 3}').Value = f'第1层平坦 {res["avgP_flat"]:.4%}，差 {z:+.2f} 个标准误'
    ov.Range(f'C{m0 + 4}').Value = mc['P_full']
    ov.Range(f'C{m0 + 5}').Value = mc['P_zero']
    ov.Range(f'C{m0 + 6}:E{m0 + 6}').Value = [[qq[0.1], qq[0.5], qq[0.9]]]
    ov.Range(f'C{m0 + 7}').Value = mc['mean_cpn']
    hist_vals = list(mc['frac_hist'].values())
    ov.Range(f'C{m0 + 9}:C{m0 + 14}').Value = [[x] for x in hist_vals]
    ov.Range(f'C{m0 + 2}').NumberFormat = '#,##0'
    for addr in (f'C{m0 + 3}:C{m0 + 5}', f'C{m0 + 6}:E{m0 + 7}', f'C{m0 + 9}:C{m0 + 14}'):
        ov.Range(addr).NumberFormat = '0.00%'
    stamp = [(dt.datetime.combine(date, dt.time()) - dt.datetime(1899, 12, 30)).days, terms.days, terms.lo, terms.hi,
             terms.c, terms.vol_shift, q[M.SPOT]]
    ov.Range(f'K{m0 + 1}:K{m0 + len(stamp)}').Value = [[x] for x in stamp]
    ov.Range(f'K{m0 + 1}').NumberFormat = 'yyyy-mm-dd'
    if abs(z) > 4:
        fail('bad_input', 'check the Monte Carlo against the closed form', f'MC avgP z = {z:.2f}')
    app.CalculateFull()

    ov.Activate()
    if a.share:
        wb.RemovePersonalInformation = True
    wb.SaveAs(str(out), FileFormat=51)
    add_overview_charts(sh['总览'], wb)
    app.CalculateFull()
    wb.Save()

    errs = scan_errors(wb)
    if errs:
        fail('bad_input', 'fix the formulas listed', f'{len(errs)} Excel errors {errs[:10]}')
    status = gv('总览', 'A3')
    p = lambda x, f='.6f': format(x, f) if isinstance(x, (int, float)) else str(x)
    print('---- summary ----')
    print('pricing_date', date, 'citi_data_date', data_date, 'live_citi', live)
    print('spot', p(spot, ',.2f'), 'fwdT', p(gv('行情', 'B59'), ',.2f'), 'atm_volT', p(gv('行情', 'B60'), '.4%'),
          'rUSD', p(gv('行情', 'B58'), '.4%'), 'carry_minus_r', p(gv('行情', 'B61'), '.4%'))
    print('terms', f'days={terms.days} n={res["n"]} range={terms.lo:g}-{terms.hi:g} coupon={terms.c:.2%}')
    print('fair_value flat/smile', p(res['fv_flat'], ',.2f'), p(res['fv_smile'], ',.2f'))
    print('issuer_margin flat/smile', p(res['margin_flat'], ',.2f'), p(res['margin_smile'], ',.2f'),
          'smile_pct', p(res['margin_smile'] / terms.N, '.3%'), 'smile_pa', p(res['margin_smile'] / terms.N * terms.basis / terms.days, '.3%'))
    print('avgP flat/smile', p(res['avgP_flat'], '.4%'), p(res['avgP_smile'], '.4%'),
          'exp_cpn_smile', p(res['exp_cpn_smile'], '.3%'), 'breakeven_c_smile', p(res['breakeven_c_smile'], '.3%'))
    print('fees', f'issue={terms.issue:.2%} fee={terms.fee:.2%} spread_bp={terms.spread:g} target_pa={terms.target:.2%}',
          'target_c_smile', p(res['target_c_smile'], '.3%'), 'client_yield_smile', p(res['client_yield_smile'], '.3%'))
    print('mc', f"avgP={mc['mc_avgP']:.4%} z={z:+.2f} P_full={mc['P_full']:.2%} P_zero={mc['P_zero']:.3%} "
                f"cpn_q10/50/90={qq[0.1]:.2%}/{qq[0.5]:.2%}/{qq[0.9]:.2%}")
    print('layer4_max_abs_diff', [p(gv('第4层', f'{l4(bi, "diff")}4'), '.1e') for bi in range(len(BOUNDS))])
    print('overview_checks', [p(x[0], '.1e') for x in gv('总览', f'F{chk}:F{chk + 1}')])
    if hist and not a.share:
        print('history', f'{hist[0][0]}..{hist[-1][0]} rows={len(hist)}', 'mean_frac', p(gv('历史', 'B9'), '.2%'),
              'full', p(gv('历史', 'B10'), '.2%'), 'zero', p(gv('历史', 'B11'), '.2%'), 'mean_cpn', p(gv('历史', 'B15'), '.3%'))
    else:
        print('history', 'not included (--share)')
    print('sens', [(round(gv('敏感性', f'A{SENS0 + j}'), 3), p(gv('敏感性', f'I{SENS0 + j}'), '.2%'),
                    p(gv('敏感性', f'J{SENS0 + j}'), '.2%')) for j in range(len(WIDTHS))])
    print('status', status, '|', gv('总览', 'A4'))
    print('excel_errors', len(errs))
    if a.share:
        wb.Close(SaveChanges=False)
        leaks = scrub_xlsx(out, [app.UserName])
        if leaks:
            fail('bad_input', 'remove the listed strings from the shared workbook', f'personal strings left: {leaks}')
        print('share_scrub ok: no author, save path or user name in the file; workbook closed')
    print('saved', 'file:///' + str(out).replace('\\', '/').replace(' ', '%20'))


if __name__ == '__main__':
    main()
