# /routes/main.py
import io
import re
import uuid
import threading
from flask import (Blueprint, render_template, request, jsonify,
                   send_file, current_app, url_for)
from openpyxl import Workbook
from utils import allowed_file

# Локальные импорты
from excel_processor import process_excel_hybrid
import glob, json, os

main_bp = Blueprint('main', __name__)

def get_task_statuses():
    return current_app.config['TASK_STATUSES']

@main_bp.route('/')
def index():
    template_files = glob.glob(os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], '*.json'))
    templates_data = []
    for f in template_files:
        try:
            with open(f, 'r', encoding='utf-8') as file:
                data = json.load(file)
                data['id'] = os.path.basename(f).replace('.json', '')
                templates_data.append(data)
        except Exception as e:
            print(f"Ошибка чтения шаблона {f}: {e}")
    return render_template('index.html', saved_templates=templates_data)

@main_bp.route('/process', methods=['POST'])
def process_files():
    try:
        source_file = request.files.get('source_file')
        if not (source_file and source_file.filename and allowed_file(source_file.filename)):
            return jsonify({'error': 'Исходный файл .xlsx или .xlsm должен быть загружен.'}), 400

        source_file_obj = io.BytesIO(source_file.read())
        original_template_filename = ''
        template_rules, selected_template_id = [], request.form.get('saved_template')

        if selected_template_id:
            json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{selected_template_id}.json")
            with open(json_path, 'r', encoding='utf-8') as f:
                template_info = json.load(f)
            excel_filename = template_info.get('excel_file')
            original_template_filename = template_info.get('original_filename', 'template.xlsx')
            template_file_path = os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename)
            if excel_filename and os.path.exists(template_file_path):
                 with open(template_file_path, 'rb') as f:
                    template_file_obj = io.BytesIO(f.read())
            else:
                template_wb = Workbook()
                template_file_obj = io.BytesIO()
                template_wb.save(template_file_obj)
                template_file_obj.seek(0)
            t_start_cell = template_info['header_start_cell']
            template_rules = template_info.get('rules', [])
        else:
            template_file = request.files.get('template_file')
            if not (template_file and template_file.filename and allowed_file(template_file.filename)):
                return jsonify({'error': 'Если шаблон не выбран, его нужно загрузить вручную (.xlsx или .xlsm).'}), 400
            original_template_filename = template_file.filename
            template_file_obj = io.BytesIO(template_file.read())
            t_start_cell = request.form.get('template_range_start').upper()

        ranges = {
            's_start_row': int(re.search(r'\d+', request.form.get('source_range_start').upper()).group()),
            't_start_row': int(re.search(r'\d+', t_start_cell).group())
        }
        private_rules = []
        s_cols, t_cols = request.form.getlist('manual_source_col'), request.form.getlist('manual_template_col')
        for i in range(len(s_cols)):
            if s_cols[i] and t_cols[i]: private_rules.append({'s_col': s_cols[i].upper(), 't_col': t_cols[i].upper()})

        post_function = request.form.get('post_processing_function', 'none')
        visible_rows_only = request.form.get('visible_rows_only') == 'true'
        task_id = str(uuid.uuid4())
        task_statuses = get_task_statuses()

        thread = threading.Thread(target=process_excel_hybrid, args=(
            task_id, source_file_obj, template_file_obj, ranges, template_rules,
            private_rules, post_function, original_template_filename, task_statuses, visible_rows_only
        ))
        thread.start()
        return jsonify({'task_id': task_id})
    except Exception as e:
        return jsonify({'error': f'Ошибка на сервере: {e}'}), 500

@main_bp.route('/status/<task_id>')
def task_status(task_id):
    task_statuses = get_task_statuses()
    status_info = task_statuses.get(task_id, {})
    if status_info.get('result_file'):
        template_filename = status_info.get('template_filename', 'template.xlsx')
        _, file_extension = os.path.splitext(template_filename)
        file_name = f"processed_{task_id}{file_extension}"
        return jsonify({'status': status_info['status'], 'progress': status_info['progress'], 'result_file': file_name})
    return jsonify(status_info)

@main_bp.route('/download/<filename>')
def download_file(filename):
    task_statuses = get_task_statuses()
    task_id_with_ext = filename.replace('processed_', '')
    task_id, file_extension = os.path.splitext(task_id_with_ext)
    status_info = task_statuses.get(task_id)
    if status_info and status_info.get('result_file'):
        file_obj = status_info['result_file']
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        if file_extension.lower() == '.xlsm':
            mimetype = 'application/vnd.ms-excel.sheet.macroEnabled.12'
        return send_file(file_obj, as_attachment=True, download_name=filename, mimetype=mimetype)
    return "Файл не найден или обработка еще не завершена.", 404