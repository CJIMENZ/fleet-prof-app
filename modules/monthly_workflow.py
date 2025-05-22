"""
modules/monthly_workflow.py

Orchestrates:
 1) create_new_workbook (CAD)
 2) calculate_fx_n5 => CAD->USD
 3) calculate_fx_n5_aud => AUD->USD (if oracle_aud_path is provided)
 4) clean_cad_data
 5) integrate_tableau => 'PnL_CAN_GL'
 6) create_comparison_sheet => merges into 'Comparison'
"""

import logging
import sys

from modules.fx_operations import (
    create_new_workbook,
    calculate_fx_n5,
    calculate_fx_n5_aud,
    clean_cad_data
)
from modules.tableau_operations import integrate_tableau
from modules.comparison_operations import create_comparison_sheet

def run_fx_and_comparison(
    config_parser,
    oracle_usd_path,
    oracle_cad_path,
    ref_file_path,
    oracle_aud_path=None
):
    """
    If oracle_aud_path is provided, also compute AUD->USD from 'LOS Management Report IS29'.
    """
    logging.info("Starting run_fx_and_comparison workflow.")

    # 1) create new workbook (USD + CA)
    new_wb = create_new_workbook(oracle_usd_path, oracle_cad_path)

    # 2) CAD->USD
    fx_val, latest_month = calculate_fx_n5(new_wb, ref_file_path)
    logging.info(f"CAD->USD = {fx_val}, month={latest_month}")

    # 2b) If we have an AUD file, do the AUD->USD logic
    if oracle_aud_path:
        logging.info("AUD file provided. Calculating AUD->USD FX.")
        # We pass the same 'latest_month' so both share the same date label
        calculate_fx_n5_aud(oracle_usd_path, oracle_aud_path, ref_file_path, existing_month=latest_month)
    else:
        logging.info("No AUD file provided, skipping AUD->USD calc.")

    # 3) Clean CAD => 'Data Sort CAD'
    new_wb, latest_month = clean_cad_data(new_wb, latest_month)

    # 4) integrate Tableau => 'PnL_CAN_GL'
    new_wb = integrate_tableau(config_parser, new_wb, latest_month)

    # 5) create comparison => merges to "Comparison"
    new_wb = create_comparison_sheet(new_wb, latest_month, ref_file_path)

    logging.info("run_fx_and_comparison workflow completed successfully.")
    return new_wb

if __name__ == "__main__":
    import configparser
    import os
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 4:
        print("Usage: python monthly_workflow.py <oracle_usd_file> <oracle_cad_file> <ref_file> [oracle_aud_file]")
        sys.exit(1)

    config = configparser.ConfigParser()
    config_file = os.path.join(os.path.dirname(__file__), "config.ini")
    if os.path.exists(config_file):
        config.read(config_file)

    usd_file = sys.argv[1]
    cad_file = sys.argv[2]
    ref_file = sys.argv[3]
    aud_file = sys.argv[4] if len(sys.argv) > 4 else None

    run_fx_and_comparison(config, usd_file, cad_file, ref_file, oracle_aud_path=aud_file)
