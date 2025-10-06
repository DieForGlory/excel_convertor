# /routes/templates.py
import os
import glob
import json
import uuid
from flask import (Blueprint, render_template, request, flash, redirect,
                   url_for, current_app, send_from_directory)
from werkzeug.utils import secure_filename
from utils import allowed_file

templates_bp = Blueprint('templates', __name__, url_prefix='/templates')

@templates_bp.route('/')
def list():
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
    return render_template('templates_list.html', templates=templates_data)

@templates_bp.route('/new')
def new():
    return render_template('create_template.html')

@templates_bp.route('/create', methods=['POST'])
def create():
    try:
        template_name = request.form.get('template_name')
        header_start_cell = request.form.get('header_start_cell').upper()
        excel_file = request.files.get('excel_file')
        if not (template_name and header_start_cell and excel_file and excel_file.filename):
            flash("Ошибка: Все поля должны быть заполнены.", "error")
            return redirect(url_for('templates.new'))

        template_id = str(uuid.uuid4())
        _, file_extension = os.path.splitext(excel_file.filename)
        saved_excel_filename = f"{template_id}{file_extension}"
        excel_file.save(os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], saved_excel_filename))

        rules = []
        source_cells, template_cols = request.form.getlist('source_cell'), request.form.getlist('template_col')
        for i in range(len(source_cells)):
            if source_cells[i] and template_cols[i]:
                rules.append({"source_cell": source_cells[i].upper(), "template_col": template_cols[i].upper()})

        template_data = {
            "template_name": template_name, "excel_file": saved_excel_filename,
            "original_filename": excel_file.filename, "header_start_cell": header_start_cell, "rules": rules
        }
        with open(os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json"), 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=4)
        flash(f"Шаблон '{template_name}' успешно создан!", "success")
        return redirect(url_for('templates.list'))
    except Exception as e:
        flash(f"Произошла ошибка: {e}", "error")
        return redirect(url_for('templates.new'))

@templates_bp.route('/edit/<template_id>', methods=['GET', 'POST'])
def edit(template_id):
    json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json")
    if not os.path.exists(json_path):
        flash("Шаблон не найден.", "error")
        return redirect(url_for('templates.list'))
    if request.method == 'POST':
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            template_data['template_name'] = request.form.get('template_name')
            template_data['header_start_cell'] = request.form.get('header_start_cell').upper()
            new_excel_file = request.files.get('excel_file')
            if new_excel_file and new_excel_file.filename:
                if allowed_file(new_excel_file.filename):
                    old_excel_path = os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], template_data.get('excel_file', ''))
                    if os.path.exists(old_excel_path):
                        os.remove(old_excel_path)
                    _, file_extension = os.path.splitext(new_excel_file.filename)
                    saved_excel_filename = f"{template_id}{file_extension}"
                    new_excel_file.save(os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], saved_excel_filename))
                    template_data['excel_file'] = saved_excel_filename
                    template_data['original_filename'] = new_excel_file.filename
                else:
                    flash("Недопустимый формат файла.", "error")
                    return redirect(url_for('templates.edit', template_id=template_id))
            rules = []
            source_cells, template_cols = request.form.getlist('source_cell'), request.form.getlist('template_col')
            for i in range(len(source_cells)):
                if source_cells[i] and template_cols[i]:
                    rules.append({"source_cell": source_cells[i].upper(), "template_col": template_cols[i].upper()})
            template_data['rules'] = rules
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, ensure_ascii=False, indent=4)
            flash("Шаблон успешно обновлен!", "success")
            return redirect(url_for('templates.list'))
        except Exception as e:
            flash(f"Ошибка при обновлении: {e}", "error")
            return redirect(url_for('templates.edit', template_id=template_id))
    with open(json_path, 'r', encoding='utf-8') as f:
        template_data = json.load(f)
    return render_template('edit_template.html', template=template_data, template_id=template_id)

@templates_bp.route('/download/<template_id>')
def download(template_id):
    json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json")
    if not os.path.exists(json_path):
        return redirect(url_for('templates.list'))
    with open(json_path, 'r', encoding='utf-8') as f:
        template_data = json.load(f)
    excel_filename = template_data.get('excel_file')
    original_filename = template_data.get('original_filename', 'template.xlsx')
    if not excel_filename or not os.path.exists(os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename)):
        flash("Файл Excel для этого шаблона не найден.", "error")
        return redirect(url_for('templates.edit', template_id=template_id))
    return send_from_directory(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename, as_attachment=True, download_name=original_filename)

@templates_bp.route('/delete/<template_id>', methods=['POST'])
def delete(template_id):
    try:
        json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{secure_filename(template_id)}.json")
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            excel_filename = template_data.get('excel_file')
            if excel_filename:
                excel_path = os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename)
                if os.path.exists(excel_path):
                    os.remove(excel_path)
            os.remove(json_path)
            flash("Шаблон успешно удален.", "success")
        else:
            flash("Шаблон для удаления не найден.", "error")
    except Exception as e:
        flash(f"Ошибка при удалении: {e}", "error")
    return redirect(url_for('templates.list'))