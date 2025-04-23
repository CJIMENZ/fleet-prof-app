# modules/unalloc_distribution.py

import datetime
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
    Read P.VM sheets and Database/Main_Combo, compute unallocated cost distributions
    and write a new "Unalloc_Distribution" sheet with summary + pad-level results.
    """
    #----- load workbook --------------------------------------------------------------------------------
    wb = load_workbook(workbook_path)
    #----------------------------------------------------------------------------------------------------

    #----- read Main_Combo from "Database" sheet -------------------------------------------------------
    df_main = pd.read_excel(
        workbook_path,
        sheet_name='Database',
        header=1,             # assume headers on row 2
        usecols='A:P',        # cols 1–16
        engine='openpyxl'
    ).dropna(how='all', subset=['Pad No'])  # keep only real pads

    # compute pad_days
    df_main['Pad Start'] = pd.to_datetime(df_main['Pad Start'])
    df_main['Pad End']   = pd.to_datetime(df_main['Pad End'])
    df_main['pad_days'] = (
        (df_main[['Pad End', pd.Timestamp(month_end)]].min(axis=1)
         - df_main[['Pad Start', pd.Timestamp(month_start)]].max(axis=1))
        .dt.days.clip(lower=0)
    )

    #----- read P.VM sheets --------------------------------------------------------------------------
    df_unalloc = pd.read_excel(workbook_path, sheet_name='P. VM – Unalloc', engine='openpyxl')
    df_adjust  = pd.read_excel(workbook_path, sheet_name='P. VM – Adjustments', engine='openpyxl')
    df_current = pd.read_excel(workbook_path, sheet_name='P. VM – Current', engine='openpyxl')

    # filter only blank or non-6-digit project numbers
    mask = ~df_unalloc['Project Number'].astype(str).str.match(r'^\d{6}$')
    df_unalloc = df_unalloc.loc[mask]
    df_adjust  = df_adjust.loc[mask]

    #----- aggregate unalloc by basin -----------------------------------------------------------------
    grp_u = df_unalloc.groupby('LBRT BASIN')
    sand_unalloc   = grp_u['Prop Cost'].sum()
    handle_unalloc = grp_u['Truck Cost'].sum()
    daily_unalloc  = grp_u[['Fuel Cost','Mat Cost','Other Pad Cost','Alloc VM Cost']].sum(axis=1).groupby(df_unalloc['LBRT BASIN']).sum()

    chem_unalloc = df_current.groupby('LBRT BASIN')['Chem Cost'].sum()

    #----- aggregate denominators by basin ------------------------------------------------------------
    grp_m = df_main.groupby('LBRT BASIN')
    prop_total = grp_m['Prop TN'].sum()
    chem_total = grp_m['Chem Cost'].sum()
    day_total  = grp_m['pad_days'].sum()

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

    #----- build summary DataFrame --------------------------------------------------------------------
    df_summary = pd.DataFrame({
        'Basin': prop_total.index,
        'SandUnalloc': sand_unalloc,
        'PropTotal':   prop_total,
        'RatioSand':   final_sand,
        'HandleUnalloc': handle_unalloc,
        'RatioHandle':  final_handle,
        'ChemUnalloc': chem_unalloc,
        'RatioChem':    final_chem,
        'DailyUnalloc': daily_unalloc,
        'DayTotal':     day_total,
        'RatioDaily':   final_daily,
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
