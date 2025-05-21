from flask import Flask, render_template, request, jsonify
import logging

from settings_manager import load_config
from modules.monthly_workflow import run_fx_and_comparison
from modules.cks_pivot_operations import pivot_cks_data_to_ref
from modules.view_download_operations import download_all_views

app = Flask(__name__)
config = load_config()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/update_fx_compare', methods=['POST'])
def update_fx_compare():
    data = request.get_json()
    try:
        run_fx_and_comparison(
            config,
            data['oracle_usd'],
            data['oracle_cad'],
            data['ref_file'],
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
        pivot_cks_data_to_ref(
            data['finance_file'],
            data['ref_file'],
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