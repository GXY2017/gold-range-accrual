# -*- coding: utf-8 -*-
"""Excel-COM helpers for build_gold_range_accrual.py: colours, error codes, retry on a busy Excel, block writes."""
import sys
import time

import pythoncom

ADDIN = 'CitiVelocity-Ribbon-AddIn.xll'


XL_ERR = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
          -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!'}


def rgb(r, g, b):
    return r + g * 256 + b * 65536


BLUE, GREEN, GREY, DARK = rgb(0, 0, 255), rgb(0, 128, 0), rgb(242, 242, 242), rgb(64, 64, 64)


def fail(kind, nxt, msg):
    print(msg)
    print(f'ERROR kind={kind} next={nxt}')
    sys.exit(1)


def retry(fn, tries=30, wait=1.0):
    """Excel rejects COM calls while the add-in or VBA is busy; wait and retry."""
    for i in range(tries):
        try:
            return fn()
        except pythoncom.com_error as e:
            if e.hresult in (-2146777998, -2147418111, -2147417848) and i < tries - 1:
                time.sleep(wait)
                continue
            raise


class Sheet:
    def __init__(self, ws):
        self.ws = ws

    def rng(self, addr):
        return self.ws.Range(addr)

    def put(self, row, col, data):
        """Write a block (list of rows) starting at (row, col); strings starting with '=' become formulas."""
        nc = max(len(r) for r in data)
        data = [list(r) + [''] * (nc - len(r)) for r in data]
        r = self.ws.Range(self.ws.Cells(row, col), self.ws.Cells(row + len(data) - 1, col + nc - 1))
        retry(lambda: setattr(r, 'Formula', data))
        return r

    def fmt(self, addr, nf=None, color=None, bold=None, fill=None, wrap=None):
        r = self.ws.Range(addr)
        if nf is not None:
            r.NumberFormat = nf
        if color is not None:
            r.Font.Color = color
        if bold is not None:
            r.Font.Bold = bold
        if fill is not None:
            r.Interior.Color = fill
        if wrap is not None:
            r.WrapText = wrap

    def header(self, addr):
        self.fmt(addr, bold=True, fill=GREY)

    def title(self, addr, text):
        self.ws.Range(addr).Value = text
        self.ws.Range(addr).Font.Bold = True
        self.ws.Range(addr).Font.Size = 13

    def widths(self, spec):
        for col, w in spec.items():
            self.ws.Columns(col).ColumnWidth = w


def name(wb, nm, ref):
    wb.Names.Add(Name=nm, RefersTo='=' + ref)
