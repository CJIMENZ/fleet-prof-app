#!/usr/bin/env python3
# modules/unalloc_distribution.py
"""
Distribute unallocated Sand, Handling, Chemical, and Daily costs
----------------------------------------------------------------
Steps
1.  Pull *unallocated* lines from “P. VM - Unalloc” **and** the
    non-6-digit rows (the true unalloc-txn lines) in “P. VM - Adjustments”.
    ➜  Combined DF = NUMERATOR   (printed & written)
2.  Pull activity metrics from Database/Main-Combo + Chem Cost totals
    from “P. VM - Current” (only projects that exist in Main-Combo).
    ➜  Basin-level METRICS table = DENOMINATOR (printed & written)
3.  Build allocation ratios per basin, detect “orphans” where the
    denominator is 0, sprinkle those across active basins (≠ CA),
    and print three debug tables:
        3-a  orphan costs
        3-b  orphan ratios
        3-c  final basin ratios            (printed & written)
4.  Copy Main-Combo and append the four allocated cost columns
    (Unalloc_Sand / _Handle / _Chem / _Daily).               (printed & written)

Every DF shown in the console is also dropped into the worksheet
in the same order, separated by a blank row and a bold section title.
"""

from __future__ import annotations
import datetime, re, sys, traceback
from pathlib import Path
from typing import List, Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# ── Formatting ─────────────────────────────────────────────────────────────
LIGHT_GRAY   = PatternFill(fill_type="solid", fgColor="DDDDDD")
HEADER_FILL  = PatternFill(fill_type="solid", fgColor="C0C0C0")
CURRENCY_FMT = '$#,##0.00'
NUMBER_FMT   = '#,##0.00'

# ── Column header normalisation ────────────────────────────────────────────
RENAME_MAP = {
    'ENG BASIN R1':              'LBRT BASIN',
    'Chemical and Gel cost':     'Chem Cost',
    'Mat and Containment Costs': 'Mat Cost',
    'Other Pad Costs':           'Other Pad Cost',
    'Allocation VM':             'Alloc VM Cost',
    # sometimes Current sheet header shows “Chemical cost”
    'Chemical cost':             'Chem Cost',
}

# ═══════════════════════════════════════════════════════════════════════════
def _find_sheet_name(wb, keywords: List[str]) -> str:
    """Return the first sheet whose name contains *all* keywords (case-insensitive)."""
    for name in wb.sheetnames:
        lname = name.lower()
        if all(k.lower() in lname for k in keywords):
            return name
    raise KeyError(f"No sheet name contains keywords: {keywords!r}")

# -------------------------------------------------------------------------- 
def _read_pvm_body(workbook_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Read any P. VM sheet whose *real* data begin in row 4 (header row 1,
    grand-total row 2, blank row 3).  Returns a cleaned DataFrame.
    """
    df = pd.read_excel(
        workbook_path,
        sheet_name=sheet_name,
        header=0,            # row 1 is header
        engine="openpyxl"
    )
    df = df.iloc[2:]        # drop GT + blank
    df = df.dropna(how="all")          # strip empty rows at bottom
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns=RENAME_MAP, inplace=True)
    return df.reset_index(drop=True)

# -------------------------------------------------------------------------- 
def _read_pvm_adjustments(workbook_path: str, sheet_name: str) -> pd.DataFrame:
    """
    “P. VM - Adjustments” has summary stuff first; real header is on row 18,
    data start on row 19.
    """
    df = pd.read_excel(
        workbook_path,
        sheet_name=sheet_name,
        header=17,           # Excel row 18
        engine="openpyxl"
    )
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns=RENAME_MAP, inplace=True)
    return df.reset_index(drop=True)

# -------------------------------------------------------------------------- 
def _read_main_combo(workbook_path: str) -> pd.DataFrame:
    df = pd.read_excel(
        workbook_path,
        sheet_name="Database",
        header=1,            # Excel row 2
        usecols="A:P",
        engine="openpyxl"
    )
    df = df.dropna(how="all", subset=["Pad No"])
    df.columns = [c.strip() for c in df.columns]
    df.rename(columns={"PAD START":"Pad Start", "PAD END":"Pad End"}, inplace=True)
    return df.reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════
def run_unalloc_distribution(
        workbook_path: str,
        month_start: datetime.date,
        month_end:   datetime.date
) -> None:

    print(f"\n[INFO] ► Unalloc Distribution run for: {workbook_path}")
    print(f"       Window: {month_start} → {month_end}\n")

    try:
        wb = load_workbook(workbook_path, data_only=True)   # keep open for write-back
    except Exception:
        traceback.print_exc(); sys.exit(1)

    # ── Pull ALL raw data first ───────────────────────────────────────────
    sheet_unalloc  = _find_sheet_name(wb, ["p. vm", "unalloc"])
    sheet_current  = _find_sheet_name(wb, ["p. vm", "current"])
    sheet_adj      = _find_sheet_name(wb, ["p. vm", "adjust"])

    df_unalloc_raw = _read_pvm_body(workbook_path, sheet_unalloc)
    df_current_raw = _read_pvm_body(workbook_path, sheet_current)
    df_adjust_raw  = _read_pvm_adjustments(workbook_path, sheet_adj)
    df_main_raw    = _read_main_combo(workbook_path)

    # ---- print raw pulls -------------------------------------------------
    def _dbg(df, tag):
        print(f"[DEBUG] {tag}: shape={df.shape}")
        print(df.head(6).to_string(index=False), "\n")

    _dbg(df_unalloc_raw, "P. VM – Unalloc  (raw)")
    _dbg(df_adjust_raw,  "P. VM – Adjustments (raw)")
    _dbg(df_current_raw, "P. VM – Current  (raw)")
    _dbg(df_main_raw,    "Main_Combo (raw)")

    # ╔════════ STEP 1 ═══════════════════════════════════════════════════╗
    # Combine unallocated lines  (numerator)
    non_six = lambda s: ~s.astype(str).str.match(r"^\d{6}$")

    df_unalloc_num = df_unalloc_raw.copy()                                 # already non-6-digit
    df_adj_num     = df_adjust_raw.loc[non_six(df_adjust_raw["Project Number"])]
    df_num = pd.concat([df_unalloc_num, df_adj_num], ignore_index=True)

    _dbg(df_num, "STEP 1 ► COMBINED NUMERATOR")

    # ╔════════ STEP 2 ═══════════════════════════════════════════════════╗
    # Build denominator metrics
    df_main = df_main_raw.copy()
    # Pad-day calc (clamped to month)
    df_main["Pad Start"] = pd.to_datetime(df_main["Pad Start"])
    df_main["Pad End"]   = pd.to_datetime(df_main["Pad End"])
    ms, me = pd.Timestamp(month_start), pd.Timestamp(month_end)
    df_main["pad_days"] = (
        df_main["Pad End"].clip(upper=me) -
        df_main["Pad Start"].clip(lower=ms)
    ).dt.days.clip(lower=0)

    # Chem cost (only pads present in Main-Combo)
    df_current = df_current_raw.copy()
    chem_by_pad = df_current.groupby("Project Number")["Chem Cost"].sum()
    df_main["Chem Cost"] = df_main["Pad No"].map(chem_by_pad).fillna(0)

    # Basin-level denominator
    grp_m      = df_main.groupby("LBRT BASIN")
    prop_total = grp_m["Prop TN"].sum()
    day_total  = grp_m["pad_days"].sum()
    chem_total = grp_m["Chem Cost"].sum()

    df_den = pd.DataFrame({
        "Basin":     prop_total.index,
        "PropTotal": prop_total.values,
        "DayTotal":  day_total.reindex(prop_total.index, fill_value=0).values,
        "ChemTotal": chem_total.reindex(prop_total.index, fill_value=0).values
    })

    _dbg(df_den, "STEP 2 ► DENOMINATOR (Metrics)")

    # ╔════════ STEP 3 ═══════════════════════════════════════════════════╗
    # Build unalloc totals by basin  (numerator   ↓)
    grp_u  = df_num.groupby("LBRT BASIN")
    sand_u = grp_u["Prop Cost"].sum()
    hand_u = grp_u["Truck Cost"].sum()
    daily_u = (
        df_num[["Fuel Cost","Mat Cost","Other Pad Cost","Alloc VM Cost"]]
        .sum(axis=1)
        .groupby(df_num["LBRT BASIN"]).sum()
    )
    chem_u = grp_u["Chem Cost"].sum()

    # Prepare aligned series
    basins = df_den["Basin"].unique()
    def _al(s):
        return s.reindex(basins, fill_value=0)

    sand_u, hand_u, daily_u, chem_u = map(_al, (sand_u, hand_u, daily_u, chem_u))
    prop_total, day_total, chem_total = map(_al, (prop_total, day_total, chem_total))

    # -- orphan helper ----------------------------------------------------
    def _ratio(unalloc, denom):
        with pd.option_context("mode.use_inf_as_na", True):
            r = (unalloc / denom).fillna(0)
        return r

    def _sprinkle(unalloc, denom, base_ratio):
        """Return final ratio after sprinkling orphans (denom == 0)."""
        zero_mask = (denom == 0) & (unalloc > 0)
        orphan_cost = unalloc[zero_mask].sum()
        valid_mask  = (denom > 0) & (base_ratio.index != "CA")
        pool = denom[valid_mask].sum()
        orphan_ratio = orphan_cost / pool if pool else 0
        final = base_ratio.copy()
        final.loc[valid_mask] += orphan_ratio
        return orphan_cost, orphan_ratio, final

    # -- Compute base ratios ---------------------------------------------
    ratio_sand   = _ratio(sand_u,  prop_total)
    ratio_handle = _ratio(hand_u,  prop_total)
    ratio_daily  = _ratio(daily_u, day_total)
    ratio_chem   = _ratio(chem_u,  chem_total)

    # -- Sprinkle ---------------------------------------------------------
    orphan_sand,   orph_r_sand,   final_sand   = _sprinkle(sand_u,  prop_total, ratio_sand)
    orphan_handle, orph_r_handle, final_handle = _sprinkle(hand_u,  prop_total, ratio_handle)
    orphan_daily,  orph_r_daily,  final_daily  = _sprinkle(daily_u, day_total,  ratio_daily)
    orphan_chem,   orph_r_chem,   final_chem   = _sprinkle(chem_u,  chem_total, ratio_chem)

    # -- Debug tables -----------------------------------------------------
    df_orphans = pd.DataFrame({
        "Metric": ["Sand","Handle","Daily","Chem"],
        "OrphanCost": [orphan_sand, orphan_handle, orphan_daily, orphan_chem]
    })
    _dbg(df_orphans, "STEP 3-a ► ORPHAN COSTS")

    df_orphan_ratio = pd.DataFrame({
        "Metric": ["Sand","Handle","Daily","Chem"],
        "OrphanRatio": [orph_r_sand, orph_r_handle, orph_r_daily, orph_r_chem]
    })
    _dbg(df_orphan_ratio, "STEP 3-b ► ORPHAN RATIOS  (added to active basins ≠ CA)")

    df_ratios = pd.DataFrame({
        "Basin": final_sand.index,
        "RatioSand":   final_sand.values,
        "RatioHandle": final_handle.values,
        "RatioDaily":  final_daily.values,
        "RatioChem":   final_chem.values
    })
    _dbg(df_ratios, "STEP 3-c ► FINAL BASIN RATIOS")

    # ╔════════ STEP 4 ═══════════════════════════════════════════════════╗
    # Allocate per-pad
    df_out = df_main.copy()
    df_out["Unalloc_Sand"]   = df_out["Prop TN"]   * df_out["LBRT BASIN"].map(final_sand)
    df_out["Unalloc_Handle"] = df_out["Prop TN"]   * df_out["LBRT BASIN"].map(final_handle)
    df_out["Unalloc_Chem"]   = df_out["Chem Cost"] * df_out["LBRT BASIN"].map(final_chem)
    df_out["Unalloc_Daily"]  = df_out["pad_days"]  * df_out["LBRT BASIN"].map(final_daily)

    _dbg(df_out.head(), "STEP 4 ► PAD-LEVEL ALLOCATIONS (first rows)")

    # ╔════════ WRITE TO EXCEL ═══════════════════════════════════════════╗
    if "Unalloc_Distribution" in wb.sheetnames:
        del wb["Unalloc_Distribution"]
    ws = wb.create_sheet("Unalloc_Distribution")

    def _write_section(title: str, df: pd.DataFrame, start_row: int) -> int:
        """Write a DF preceded by a bold title, return next free row index."""
        cell = ws.cell(start_row, 1, title)
        cell.font = Font(bold=True, size=12)
        cell.fill = HEADER_FILL
        start_row += 2
        for r_idx, row in enumerate(
                dataframe_to_rows(df, index=False, header=True), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(r_idx, c_idx, value)
                # header row styling
                if r_idx == start_row:
                    cell.font = Font(bold=True)
                    cell.fill = HEADER_FILL
                # Currency formatting heuristics
                if isinstance(value, (int,float)) and ("Cost" in df.columns[c_idx-1] or "Unalloc" in df.columns[c_idx-1]):
                    cell.number_format = CURRENCY_FMT
        return r_idx + 2   # blank row after table

    row = 1
    row = _write_section("RAW – P. VM Unalloc",      df_unalloc_raw, row)
    row = _write_section("RAW – P. VM Adjustments",  df_adjust_raw,  row)
    row = _write_section("RAW – P. VM Current",      df_current_raw, row)
    row = _write_section("RAW – Main_Combo",         df_main_raw,    row)
    row = _write_section("STEP 1 – Numerator Combined", df_num,      row)
    row = _write_section("STEP 2 – Denominator Metrics", df_den,     row)
    row = _write_section("STEP 3-a – Orphan Costs",      df_orphans, row)
    row = _write_section("STEP 3-b – Orphan Ratios",     df_orphan_ratio, row)
    row = _write_section("STEP 3-c – Final Basin Ratios", df_ratios,  row)
    _ = _write_section("STEP 4 – Pad-level Allocations", df_out, row)

    wb.save(workbook_path)
    print(f"[INFO] ✔ Unalloc_Distribution sheet written & workbook saved\n")

# ── CLI helper ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python -m modules.unalloc_distribution "
              "<workbook_path> <YYYY-MM-DD start> <YYYY-MM-DD end>")
        sys.exit(1)
    run_unalloc_distribution(
        sys.argv[1],
        datetime.date.fromisoformat(sys.argv[2]),
        datetime.date.fromisoformat(sys.argv[3])
    )
