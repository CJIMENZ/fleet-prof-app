from flask import Flask, render_template, request, jsonify
import logging

from datetime import date

from settings_manager import load_config, save_config
from modules.monthly_workflow import run_fx_and_comparison
from modules.cks_pivot_operations import pivot_cks_data_to_ref
from modules.view_download_operations import download_all_views
from modules.report_generation import build_monthly_database
from modules.pnl_pivot_operations import generate_pnl_pivot
from modules.project_vm_adjustment import generate_project_vm_adj
from modules.unalloc_distribution import run_unalloc_distribution

app = Flask(__name__)
config = load_config()

@app.route('/')
def index():
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

@app.route('/api/update_fx_compare', methods=['POST'])
def update_fx_compare():
    data = request.get_json()
    try:
        ref_file = config.get('files', {}).get('ref_data_path', '')
        run_fx_and_comparison(
            config,
            data['oracle_usd'],
            data['oracle_cad'],
            ref_file,
            oracle_aud_path=data.get('oracle_aud') or None
        )
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/pivot_cks_data', methods=['POST'])
def pivot_cks_data():
    data = request.get_json()
    try:
        ref_file = config.get('files', {}).get('ref_data_path', '')
        pivot_cks_data_to_ref(
            data['finance_file'],
            ref_file,
            data['month']
        )
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/download_views', methods=['POST'])
def download_views():
    data = request.get_json()
    try:
        out = download_all_views(config, data['save_dir'])
        return jsonify(status='success', output=out)
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/build_monthly_database', methods=['POST'])
def build_monthly_db():
    data = request.get_json()
    try:
        build_monthly_database(
            data['exports_file'],
            data['ref_data_file'],
            data['output_file']
        )
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/generate_pnl_pivot', methods=['POST'])
def generate_pnl():
    data = request.get_json()
    try:
        generate_pnl_pivot(data['month_file'])
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/generate_vm_adj', methods=['POST'])
def generate_vm_adj():
    data = request.get_json()
    try:
        generate_project_vm_adj(data['workbook'])
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

@app.route('/api/run_unalloc_distribution', methods=['POST'])
def unalloc_distribution():
    data = request.get_json()
    try:
        run_unalloc_distribution(
            data['workbook'],
            date.fromisoformat(data['start']),
            date.fromisoformat(data['end'])
        )
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500

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