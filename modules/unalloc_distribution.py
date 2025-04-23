# modules/unalloc_distribution.py

import datetime
from typing import List, Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

#----- CONFIG & CONSTANTS --------------------------------------------------------------------
LIGHT_GRAY = PatternFill(fill_type="solid", fgColor="DDDDDD")
CURRENCY_FMT = '$#,##0.00'
NUMBER_FMT   = '#,##0.00'
#---------------------------------------------------------------------------------------------

def run_unalloc_distribution(workbook_path: str, month_start: datetime.date, month_end: datetime.date):
    """
    Read P.VM sheets (even if hidden/dashed) and Database/Main_Combo,
    compute unallocated cost distributions, and write "Unalloc_Distribution" sheet.
    """
    import sys, traceback
    print(f"[DEBUG] ▶ Starting unalloc distribution for {workbook_path}", flush=True)
    print(f"[DEBUG]    month_start={month_start}, month_end={month_end}", flush=True)

    #----- load workbook with values (even if sheet is hidden) ------------------------------------
    wb = load_workbook(workbook_path, data_only=True)
    print(f"[DEBUG]    loaded sheets: {wb.sheetnames}", flush=True)

    #----- helper: locate sheet by keywords (case-insensitive) ------------------------------------
    def find_sheet(keywords: List[str]) -> str:
        for name in wb.sheetnames:
            lname = name.lower()
            if all(k.lower() in lname for k in keywords):
                return name
        raise KeyError(f"No sheet matching keywords {keywords}")

    #----- helper: sheet -> DataFrame ---------------------------------------------------------------
    def sheet_to_df(ws) -> pd.DataFrame:
        it = ws.values
        hdr = next(it)
        cols = [str(h).strip() for h in hdr]
        return pd.DataFrame(it, columns=cols)

    #----- read Main_Combo from "Database" sheet -------------------------------------------------------
    df_main = pd.read_excel(
        workbook_path,
        sheet_name='Database',
        header=1,
        usecols='A:P',
        engine='openpyxl'
    ).dropna(how='all', subset=['Pad No'])
    print(f"[DEBUG] df_main: shape={df_main.shape}", flush=True)
    print(f"[DEBUG] df_main columns: {df_main.columns.tolist()}", flush=True)

    # normalize Main_Combo headers
    df_main.columns = [c.strip() for c in df_main.columns]
    df_main.rename(columns={'PAD START': 'Pad Start', 'PAD END': 'Pad End'}, inplace=True)

    # compute pad_days (clamp into month window)
    df_main['Pad Start'] = pd.to_datetime(df_main['Pad Start'])
    df_main['Pad End']   = pd.to_datetime(df_main['Pad End'])
    ms = pd.Timestamp(month_start)
    me = pd.Timestamp(month_end)
    df_main['pad_days'] = (
        df_main['Pad End'].clip(upper=me)
        - df_main['Pad Start'].clip(lower=ms)
    ).dt.days.clip(lower=0)
    print(f"[DEBUG] pad_days range: min={df_main['pad_days'].min()}, max={df_main['pad_days'].max()}", flush=True)

    #----- read & normalize P.VM sheets ---------------------------------------------------------------
    df_unalloc = sheet_to_df(wb[find_sheet(['p. vm', 'unalloc'])])
    df_adjust  = sheet_to_df(wb[find_sheet(['p. vm', 'adjustments'])])
    df_current = sheet_to_df(wb[find_sheet(['p. vm', 'current'])])
    print(f"[DEBUG] df_unalloc: shape={df_unalloc.shape}", flush=True)
    print(f"[DEBUG] df_unalloc cols: {df_unalloc.columns.tolist()}", flush=True)
    print(f"[DEBUG] df_current: shape={df_current.shape}", flush=True)
    print(f"[DEBUG] df_current cols: {df_current.columns.tolist()}", flush=True)

    #──────────────────────── header normalisation ────────────────────────
    rename_map = {
        'ENG BASIN R1':              'LBRT BASIN',
        'Chemical and Gel cost':     'Chem Cost',
        'Mat and Containment Costs': 'Mat Cost',
        'Other Pad Costs':           'Other Pad Cost',
        'Allocation VM':             'Alloc VM Cost',
    }
    for _df in (df_unalloc, df_adjust, df_current):
        _df.columns = [str(c).strip() for c in _df.columns]   # trim whitespace
        _df.rename(columns=rename_map, inplace=True)

    print(f"[DEBUG] df_current columns AFTER rename: {df_current.columns.tolist()}", flush=True)

    # filter only blank or non-6-digit project numbers
    mask = ~df_unalloc['Project Number'].astype(str).str.match(r'^\d{6}$')
    df_unalloc = df_unalloc.loc[mask]
    df_adjust  = df_adjust.loc[mask]

    #----- aggregate unalloc by basin -----------------------------------------------------------------
    grp_u = df_unalloc.groupby('LBRT BASIN')
    sand_unalloc   = grp_u['Prop Cost'].sum()
    handle_unalloc = grp_u['Truck Cost'].sum()
    # debug: show unalloc shapes and basin list
    print(f"[DEBUG] sand_unalloc  shape={sand_unalloc.shape}, basins={list(sand_unalloc.index)}", flush=True)
    print(f"[DEBUG] handle_unalloc shape={handle_unalloc.shape}, basins={list(handle_unalloc.index)}", flush=True)

    # sum the four daily-cost columns per pad then basin-sum
    row_daily     = df_unalloc[['Fuel Cost','Mat Cost','Other Pad Cost','Alloc VM Cost']].sum(axis=1)
    daily_unalloc = row_daily.groupby(df_unalloc['LBRT BASIN']).sum()
    print(f"[DEBUG] daily_unalloc shape={daily_unalloc.shape}, basins={list(daily_unalloc.index)}", flush=True)

    chem_unalloc = df_current.groupby('LBRT BASIN')['Chem Cost'].sum()
    print(f"[DEBUG] chem_unalloc shape={chem_unalloc.shape}, basins={list(chem_unalloc.index)}", flush=True)

    #----- aggregate denominators by basin ------------------------------------------------------------
    grp_m      = df_main.groupby('LBRT BASIN')
    prop_total = grp_m['Prop TN'].sum()
    day_total  = grp_m['pad_days'].sum()
    chem_total = df_current.groupby('LBRT BASIN')['Chem Cost'].sum()

    #─────────────────────────────────────────────────────────────────────────
    # 🔑  ALIGN **all** numerator & denominator Series to one master index
    #─────────────────────────────────────────────────────────────────────────
    basin_union = (
        sand_unalloc.index
        .union(handle_unalloc.index)
        .union(daily_unalloc.index)
        .union(chem_unalloc.index)
        .union(prop_total.index)
        .union(day_total.index)
        .union(chem_total.index)
    )

    def _fx(s):  # helper: reindex & fill NaN with 0
        return s.reindex(basin_union, fill_value=0)

    sand_unalloc   = _fx(sand_unalloc)
    handle_unalloc = _fx(handle_unalloc)
    daily_unalloc  = _fx(daily_unalloc)
    chem_unalloc   = _fx(chem_unalloc)

    prop_total = _fx(prop_total)
    day_total  = _fx(day_total)
    chem_total = _fx(chem_total)

    print("[DEBUG] aligned shapes →", 
          {k: v.shape for k, v in {
              'sand_unalloc': sand_unalloc,
              'handle_unalloc': handle_unalloc,
              'daily_unalloc': daily_unalloc,
              'chem_unalloc': chem_unalloc,
              'prop_total': prop_total,
              'day_total': day_total,
              'chem_total': chem_total
          }.items()}, flush=True)

    # helper / safe ratio
    def compute_ratio(unalloc, denom):
        r = unalloc.div(denom).replace([pd.NA, pd.NaT, float('inf')], 0)
        return r.fillna(0)

    ratio_sand   = compute_ratio(sand_unalloc, prop_total)
    ratio_handle = compute_ratio(handle_unalloc, prop_total)
    ratio_chem   = compute_ratio(chem_unalloc, chem_total)
    ratio_daily  = compute_ratio(daily_unalloc, day_total)

    #----- sprinkle zero-activity basins ----------------------------------------------------------------
    def sprinkle(ratio, unalloc, denom):
        zero_mask = (denom == 0) & (unalloc > 0)
        pool = unalloc[zero_mask].sum()
        valid = denom[(denom > 0) & (ratio.index != 'CA')]
        spr = pool / valid.sum() if valid.sum() else 0
        out = ratio.copy().reindex(denom.index, fill_value=0)
        for b in denom.index:
            out[b] += spr if (denom[b] > 0 and b != 'CA') else 0
        return out

    final_sand   = sprinkle(ratio_sand,   sand_unalloc,   prop_total)
    final_handle = sprinkle(ratio_handle, handle_unalloc, prop_total)
    final_chem   = sprinkle(ratio_chem,   chem_unalloc,   chem_total)
    final_daily  = sprinkle(ratio_daily,  daily_unalloc,  day_total)

    #----- DEBUG: shapes of final series before alignment ------------------------------------------------
    print(f"[DEBUG] final_sand:   shape={final_sand.shape}, index={list(final_sand.index)}", flush=True)
    print(f"[DEBUG] final_handle: shape={final_handle.shape}, index={list(final_handle.index)}", flush=True)
    print(f"[DEBUG] final_chem:   shape={final_chem.shape}, index={list(final_chem.index)}", flush=True)
    print(f"[DEBUG] final_daily:  shape={final_daily.shape}, index={list(final_daily.index)}", flush=True)

    #----- ALIGN all numerator & ratio series to the same basin list -----------------------------------
    basin_index = prop_total.index
    sand_unalloc   = sand_unalloc.reindex(basin_index, fill_value=0)
    handle_unalloc = handle_unalloc.reindex(basin_index, fill_value=0)
    chem_unalloc   = chem_unalloc.reindex(basin_index, fill_value=0)
    daily_unalloc  = daily_unalloc.reindex(basin_index, fill_value=0)
    final_sand     = final_sand.reindex(basin_index,   fill_value=0)
    final_handle   = final_handle.reindex(basin_index, fill_value=0)
    final_chem     = final_chem.reindex(basin_index,   fill_value=0)
    final_daily    = final_daily.reindex(basin_index,  fill_value=0)
    print(f"[DEBUG] AFTER reindex all shapes: sand={sand_unalloc.shape}, handle={handle_unalloc.shape}, chem={chem_unalloc.shape}, daily={daily_unalloc.shape}", flush=True)
    print(f"[DEBUG] RATIOS reindexed: sand={final_sand.shape}, handle={final_handle.shape}, chem={final_chem.shape}, daily={final_daily.shape}", flush=True)

    # now every series has length = len(basin_index) = 9
    #----- build summary DataFrame --------------------------------------------------------------------
    df_summary = pd.DataFrame({
        'Basin':         basin_index,
        'SandUnalloc':   sand_unalloc,
        'PropTotal':     prop_total,
        'RatioSand':     final_sand,
        'HandleUnalloc': handle_unalloc,
        'RatioHandle':   final_handle,
        'ChemUnalloc':   chem_unalloc,
        'RatioChem':     final_chem,
        'DailyUnalloc':  daily_unalloc,
        'DayTotal':      day_total,
        'RatioDaily':    final_daily,
    }).fillna(0).reset_index(drop=True)

    # add totals row
    totals = df_summary[['SandUnalloc','PropTotal','HandleUnalloc','ChemUnalloc','DailyUnalloc','DayTotal']].sum()
    df_summary.loc[len(df_summary)] = ['TOTAL',
                                      totals['SandUnalloc'],
                                      totals['PropTotal'],
                                      '',  # ratios not summed
                                      totals['HandleUnalloc'],
                                      '',
                                      totals['ChemUnalloc'],
                                      '',
                                      totals['DailyUnalloc'],
                                      totals['DayTotal'],
                                      '']

    #----- compute pad-level distributions -------------------------------------------------------------
    df_main = df_main.copy()
    df_main['Unalloc_Sand']   = df_main['Prop TN']   * df_main['LBRT BASIN'].map(final_sand)
    df_main['Unalloc_Handle'] = df_main['Prop TN']   * df_main['LBRT BASIN'].map(final_handle)
    df_main['Unalloc_Chem']   = df_main['Chem Cost'] * df_main['LBRT BASIN'].map(final_chem)
    df_main['Unalloc_Daily']  = df_main['pad_days']  * df_main['LBRT BASIN'].map(final_daily)

    #----- write to new sheet --------------------------------------------------------------------------
    if 'Unalloc_Distribution' in wb.sheetnames:
        del wb['Unalloc_Distribution']
    ws = wb.create_sheet('Unalloc_Distribution')

    # write summary
    for r_idx, row in enumerate(df_summary.itertuples(index=False), start=1):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(r_idx, c_idx, val)
            # formatting totals row bold
            if row.Basin == 'TOTAL':
                cell.font = Font(bold=True)
            # currency columns
            if c_idx in (2,3,5,7,9,10):
                cell.number_format = CURRENCY_FMT
            # non-currency numeric
            if c_idx == 11:
                cell.number_format = NUMBER_FMT

    # leave a blank row
    pad_start_row = len(df_summary) + 3

    # write pad-level header
    for c_idx, header in enumerate(df_main.columns.tolist(), start=1):
        cell = ws.cell(pad_start_row, c_idx, header)
        cell.font = Font(bold=True)

    # write pad-level data
    for r_off, row in enumerate(df_main.itertuples(index=False), start=1):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(pad_start_row + r_off, c_idx, val)
            # format pad_days
            if df_main.columns[c_idx-1]=='pad_days':
                cell.number_format = NUMBER_FMT
                cell.alignment = Alignment(horizontal='right')
            # distribution cols get gray fill + currency
            if df_main.columns[c_idx-1].startswith('Unalloc_'):
                cell.fill = LIGHT_GRAY
                cell.number_format = CURRENCY_FMT

    #----- save workbook --------------------------------------------------------------------------------
    wb.save(workbook_path)
