from flask import Flask, render_template, request, jsonify
import logging
from datetime import date
import os

from settings_manager import load_config, save_config
from modules.monthly_workflow import run_fx_and_comparison
from modules.cks_pivot_operations import pivot_cks_data_to_ref
from modules.view_download_operations import download_all_views
from modules.report_generation import build_monthly_database
from modules.pnl_pivot_operations import generate_pnl_pivot
from modules.project_vm_adjustment import generate_project_vm_adj
from modules.unalloc_distribution import run_unalloc_distribution
from modules.logging_setup import log_user_action, log_file_operation, log_operation_result, log_error

app = Flask(__name__)
config = load_config()

@app.route('/')
def index():
    log_user_action("Page View", "Home Page")
    return render_template('index.html')

@app.route('/update_fx')
def update_fx_page():
    return render_template('update_fx.html')

@app.route('/pivot_ck')
def pivot_ck_page():
    return render_template('pivot_ck.html')

@app.route('/build_db')
def build_db_page():
    return render_template('build_db.html')

@app.route('/pnl_pivot')
def pnl_pivot_page():
    return render_template('pnl_pivot.html')

@app.route('/vm_adj')
def vm_adj_page():
    return render_template('vm_adj.html')

@app.route('/unalloc_dist')
def unalloc_dist_page():
    return render_template('unalloc_dist.html')

@app.route('/download_views')
def download_views_page():
    return render_template('download_views.html')

@app.route('/settings')
def settings_page():
    return render_template('settings.html', config=config)

@app.route('/logs')
def logs_page():
    """Display application logs."""
    return render_template('logs.html')

@app.route('/api/update_fx_compare', methods=['POST'])
def update_fx_compare():
    try:
        data = request.json
        log_user_action("Update FX/Compare", f"Files: USD={data.get('oracle_usd')}, CAD={data.get('oracle_cad')}, AUD={data.get('oracle_aud')}")
        
        # Log file operations
        if data.get('oracle_usd'):
            log_file_operation("Read", data['oracle_usd'])
        if data.get('oracle_cad'):
            log_file_operation("Read", data['oracle_cad'])
        if data.get('oracle_aud'):
            log_file_operation("Read", data['oracle_aud'])

        ref_file = config['files'].get('ref_data_path', '')
        new_wb, latest_month = run_fx_and_comparison(
            config_parser=config,
            oracle_usd_path=data['oracle_usd'],
            oracle_cad_path=data['oracle_cad'],
            ref_file_path=ref_file,
            oracle_aud_path=data.get('oracle_aud')
        )

        config['files']['latest_month'] = latest_month
        save_config(config)
        
        log_operation_result("Update FX/Compare", "Success")
        return jsonify({"status": "success", "latest_month": latest_month})
    except Exception as e:
        log_error(e, "Update FX/Compare")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/pivot_cks_data', methods=['POST'])
def pivot_cks_data():
    try:
        data = request.json
        log_user_action("Pivot CK Data", f"File: {data.get('finance_file')}, Month: {data.get('month')}")
        
        if data.get('finance_file'):
            log_file_operation("Read", data['finance_file'])

        pivot_cks_data_to_ref(
            finance_file=data['finance_file'],
            ref_file=config['files']['ref_data_path'],
            target_month_str=data['month']
        )
        
        log_operation_result("Pivot CK Data", "Success")
        return jsonify({"status": "success"})
    except Exception as e:
        log_error(e, "Pivot CK Data")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/download_views', methods=['POST'])
def download_views():
    try:
        data = request.json
        log_user_action("Download Views", f"Save directory: {data.get('save_dir')}")
        
        latest_month = config['files'].get('latest_month')
        output = download_all_views(config, data['save_dir'], latest_month)
        log_operation_result("Download Views", "Success", f"Output: {output}")
        return jsonify({"status": "success", "output": output})
    except Exception as e:
        log_error(e, "Download Views")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/build_monthly_db', methods=['POST'])
def build_monthly_db():
    try:
        data = request.json
        log_user_action(
            "Build Monthly Database",
            f"Tableau exports: {data.get('exports_file')}, "
            f"Output dir: {data.get('output_dir')}"
        )
        
        # Log file operations
        if data.get('exports_file'):
            log_file_operation("Read", data['exports_file'])
        if data.get('output_dir'):
            log_file_operation("Write", data['output_dir'])

        output_path = build_monthly_database(
            tableau_exports_path=data['exports_file'],
            output_dir=data['output_dir'],
            config=config
        )

        log_operation_result("Build Monthly Database", "Success", output_path)
        return jsonify({"status": "success"})
    except Exception as e:
        log_error(e, "Build Monthly Database")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/generate_pnl_pivot', methods=['POST'])
def generate_pnl_pivot_endpoint():
    try:
        data = request.json
        month_file = data.get('month_file') or data.get('month_data_file')
        log_user_action("Generate PnL Pivot", f"File: {month_file}")

        if month_file:
            log_file_operation("Read", month_file)

        generate_pnl_pivot(month_file)
        log_operation_result("Generate PnL Pivot", "Success")
        return jsonify({"status": "success"})
    except Exception as e:
        log_error(e, "Generate PnL Pivot")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/generate_project_vm_adj', methods=['POST'])
def generate_project_vm_adj_endpoint():
    try:
        data = request.json
        workbook = data.get('workbook') or data.get('month_data_file')
        log_user_action("Generate Project VM Adjustment", f"File: {workbook}")

        if workbook:
            log_file_operation("Read", workbook)

        generate_project_vm_adj(workbook)
        log_operation_result("Generate Project VM Adjustment", "Success")
        return jsonify({"status": "success"})
    except Exception as e:
        log_error(e, "Generate Project VM Adjustment")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/create_unalloc_distributions', methods=['POST'])
def create_unalloc_distributions():
    try:
        data = request.json
        start = data.get('start') or data.get('month_start')
        end = data.get('end') or data.get('month_end')
        log_user_action(
            "Create Unallocated Distributions",
            f"File: {data.get('workbook')}, Start: {start}, End: {end}"
        )

        workbook = data.get('workbook')
        if workbook:
            log_file_operation("Read", workbook)

        run_unalloc_distribution(workbook, start, end)

        log_operation_result("Create Unallocated Distributions", "Success")
        return jsonify({"status": "success"})
    except Exception as e:
        log_error(e, "Create Unallocated Distributions")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/save_settings', methods=['POST'])
def save_settings_api():
    data = request.get_json()
    try:
        if 'tableau_online' not in config:
            config['tableau_online'] = {}
        config['tableau_online']['personal_access_token_name'] = data.get('token_name', '')
        config['tableau_online']['personal_access_token_secret'] = data.get('token_secret', '')

        if 'files' not in config:
            config['files'] = {}
        config['files']['ref_data_path'] = data.get('ref_data_path', '')
        config['files']['master_file_path'] = data.get('master_file_path', '')

        if 'appearance' not in config:
            config['appearance'] = {}
        config['appearance']['theme'] = data.get('theme', 'light')

        save_config(config)
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/logs')
def get_logs():
    """Return the last 200 lines from the newest log file."""
    try:
        log_dir = 'logs'
        files = [os.path.join(log_dir, f) for f in os.listdir(log_dir) if f.endswith('.log')]
        if not files:
            return jsonify(lines=[])
        latest = max(files, key=os.path.getmtime)
        with open(latest, 'r') as fh:
            lines = [
                l for l in fh.readlines()[-200:]
                if 'GET /api/logs' not in l
            ]
        return jsonify(lines=[l.rstrip('\n') for l in lines])
    except Exception as e:
        logging.error(e)
        return jsonify(lines=[], error=str(e)), 500 