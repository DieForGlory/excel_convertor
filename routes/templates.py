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
def list_templates():  # <<< ИЗМЕНЕНИЕ 1: Функция переименована
    """Отображает список всех шаблонов."""
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
    """Показывает форму создания нового шаблона."""
    return render_template('create_template.html')


@templates_bp.route('/create', methods=['POST'])
def create():
    """Обрабатывает создание нового шаблона."""
    try:
        # ... (код создания шаблона без изменений) ...
        template_name = request.form.get('template_name')
        s_start_row = int(request.form.get('s_start_row'))
        t_start_row = int(request.form.get('t_start_row'))
        excel_file = request.files.get('excel_file')

        if not all([template_name, s_start_row, t_start_row, excel_file]):
            flash('Все поля, включая файл, обязательны для заполнения.', 'error')
            return redirect(url_for('templates.new'))

        if not allowed_file(excel_file.filename):
            flash('Недопустимый формат файла. Разрешены только .xlsx и .xlsm.', 'error')
            return redirect(url_for('templates.new'))

        template_id = str(uuid.uuid4())
        excel_filename = f"{template_id}{os.path.splitext(excel_file.filename)[1]}"
        excel_path = os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename)
        excel_file.save(excel_path)

        template_data = {
            'name': template_name,
            's_start_row': s_start_row,
            't_start_row': t_start_row,
            'excel_file': excel_filename,
            'original_filename': secure_filename(excel_file.filename),
            'rules': [],
            'private_rules': [],
            'post_function': 'none'
        }
        json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=4)

        flash(f'Шаблон "{template_name}" успешно создан!', 'success')
        # <<< ИЗМЕНЕНИЕ 2: Обновлен вызов url_for
        return redirect(url_for('templates.list_templates'))
    except Exception as e:
        flash(f'Произошла ошибка при создании шаблона: {e}', 'error')
        return redirect(url_for('templates.new'))


@templates_bp.route('/edit/<template_id>', methods=['GET', 'POST'])
def edit(template_id):
    json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{secure_filename(template_id)}.json")
    if not os.path.exists(json_path):
        flash('Шаблон не найден.', 'error')
        return redirect(url_for('templates.list_templates'))

    with open(json_path, 'r', encoding='utf-8') as f:
        template_data = json.load(f)

    if request.method == 'POST':
        # --- НАЧАЛО ИСПРАВЛЕНИЯ ---

        # 1. Получаем данные из формы
        s_start_row_str = request.form.get('s_start_row')
        t_start_row_str = request.form.get('t_start_row')
        template_name = request.form.get('template_name')

        # 2. Проверяем, что обязательные поля не пустые
        if not s_start_row_str or not t_start_row_str or not template_name:
            flash('Ошибка: Название шаблона и строки заголовков должны быть заполнены.', 'error')
            return redirect(url_for('templates.edit', template_id=template_id))

        # 3. Если все хорошо, обновляем данные
        template_data['name'] = template_name
        template_data['s_start_row'] = int(s_start_row_str)
        template_data['t_start_row'] = int(t_start_row_str)
        template_data['post_function'] = request.form.get('post_function')

        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

        # ... (остальной код для обработки правил без изменений) ...
        rules = []
        s_cols = request.form.getlist('s_col[]')
        t_cols = request.form.getlist('t_col[]')
        for i in range(len(s_cols)):
            s_col, t_col = s_cols[i], t_cols[i]
            if s_col and t_col:
                rules.append({'s_col': s_col, 't_col': t_col})
        template_data['rules'] = rules

        private_rules = []
        s_cells = request.form.getlist('source_cell[]')
        t_private_cols = request.form.getlist('template_col[]')
        for i in range(len(s_cells)):
            s_cell, t_col = s_cells[i], t_private_cols[i]
            if s_cell and t_col:
                private_rules.append({'source_cell': s_cell, 'template_col': t_col})
        template_data['private_rules'] = private_rules

        # Сохраняем изменения в файл
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, ensure_ascii=False, indent=4)
            flash('Шаблон успешно обновлен!', 'success')
        except Exception as e:
            flash(f'Ошибка при сохранении шаблона: {e}', 'error')

        return redirect(url_for('templates.edit', template_id=template_id))

    return render_template('edit_template.html', template=template_data, template_id=template_id)


@templates_bp.route('/download/<template_id>')
def download(template_id):
    """Скачивание Excel-файла шаблона."""
    # ... (код без изменений) ...
    json_path = os.path.join(current_app.config['TEMPLATES_DB_FOLDER'], f"{secure_filename(template_id)}.json")
    if not os.path.exists(json_path):
        return "Template not found", 404

    with open(json_path, 'r', encoding='utf-8') as f:
        template_data = json.load(f)

    excel_filename = template_data.get('excel_file')
    original_filename = template_data.get('original_filename', 'template.xlsx')
    if not excel_filename or not os.path.exists(
            os.path.join(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename)):
        flash("Файл Excel для этого шаблона не найден.", "error")
        return redirect(url_for('templates.edit', template_id=template_id))
    return send_from_directory(current_app.config['TEMPLATE_EXCEL_FOLDER'], excel_filename, as_attachment=True,
                               download_name=original_filename)


@templates_bp.route('/delete/<template_id>', methods=['POST'])
def delete(template_id):
    """Удаление шаблона."""
    try:
        # ... (код удаления без изменений) ...
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
            flash("Шаблон не найден.", "error")
    except Exception as e:
        flash(f"Ошибка при удалении шаблона: {e}", "error")
    # <<< ИЗМЕНЕНИЕ 4: Обновлен вызов url_for
    return redirect(url_for('templates.list_templates'))