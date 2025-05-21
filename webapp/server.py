from flask import Flask, render_template, request, jsonify
import logging

from settings_manager import load_config, save_config
from modules.monthly_workflow import run_fx_and_comparison
from modules.cks_pivot_operations import pivot_cks_data_to_ref
from modules.view_download_operations import download_all_views

app = Flask(__name__)
config = load_config()

@app.route('/')
def index():
    return render_template('index.html')

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
        config['appearance']['theme'] = data.get('theme', 'journal')

        save_config(config)
        return jsonify(status='success')
    except Exception as e:
        logging.error(e)
        return jsonify(status='error', message=str(e)), 500 