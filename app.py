import os
import re
import uuid
import threading
import json
import glob
from flask import Flask, render_template, request, send_from_directory, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from thefuzz import fuzz
from dadata import Dadata

import dictionary_matcher

# --- 1. Конфигурация приложения ---
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
TEMPLATES_DB_FOLDER = 'templates_db'
ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}  # <--- ИЗМЕНЕНО: Добавлена поддержка .xlsm

# --- НОВАЯ КОНФИГУРАЦИЯ DADATA ---
DADATA_API_KEY = "ВАШ_API_КЛЮЧ"
DADATA_SECRET_KEY = "ВАШ_СЕКРЕТНЫЙ_КЛЮЧ"

app = Flask(__name__)
app.config.from_mapping(
    UPLOAD_FOLDER=UPLOAD_FOLDER,
    PROCESSED_FOLDER=PROCESSED_FOLDER,
    TEMPLATES_DB_FOLDER=TEMPLATES_DB_FOLDER,
    SECRET_KEY='your-super-secret-key-change-it-for-production'
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(TEMPLATES_DB_FOLDER, exist_ok=True)

# --- 2. Глобальное хранилище статусов задач ---
task_statuses = {}


# --- 3. Вспомогательные функции (без изменений) ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_header(header):
    if not isinstance(header, str):
        header = str(header)
    return re.sub(r'[\s\W_]+', '', header.lower())


def get_col_from_cell(cell_coord):
    if not cell_coord: return None
    match = re.match(r"([A-Z]+)", cell_coord.upper())
    return match.group(1) if match else None


def find_column_indices(worksheet, start_row, headers_to_find):
    indices = {}
    header_row = worksheet[start_row]
    normalized_map = {normalize_header(cell.value): cell.column for cell in header_row if cell.value}
    for key, target_header in headers_to_find.items():
        normalized_target = normalize_header(target_header)
        found_col = normalized_map.get(normalized_target)
        if found_col:
            indices[key] = found_col
    return indices


# --- ОБНОВЛЕННАЯ ФУНКЦИЯ ПОСТ-ОБРАБОТКИ С DADATA (без изменений) ---
def apply_post_processing(task_id, workbook, start_row, function_name):
    if function_name == 'none' or not function_name:
        return

    worksheet = workbook.active
    dadata = Dadata(DADATA_API_KEY, DADATA_SECRET_KEY)
    total_rows = worksheet.max_row - start_row

    if function_name == 'coords_to_address':
        cols = find_column_indices(worksheet, start_row, {'lat': 'Широта', 'lon': 'Долгота', 'addr': 'Адрес'})
        if not all(k in cols for k in ['lat', 'lon', 'addr']):
            raise ValueError("Не найдены обязательные столбцы: 'Широта', 'Долгота', 'Адрес'")

        coords_to_process, rows_to_update = [], []
        for i, row_cells in enumerate(worksheet.iter_rows(min_row=start_row + 1)):
            lat_val, lon_val = row_cells[cols['lat'] - 1].value, row_cells[cols['lon'] - 1].value
            if lat_val and lon_val:
                coords_to_process.append({"lat": lat_val, "lon": lon_val})
                rows_to_update.append(start_row + 1 + i)

        if coords_to_process:
            try:
                task_statuses[task_id]['status'] = f"Геокодирование (координаты->адрес): {len(coords_to_process)} адресов..."
                results = dadata.geolocate(name="address", queries=coords_to_process)
                for i, result in enumerate(results):
                    if result and result['suggestions']:
                        address = result['suggestions'][0]['value']
                        worksheet.cell(row=rows_to_update[i], column=cols['addr']).value = address
            except Exception as e:
                print(f"Ошибка геокодирования DaData: {e}")

    elif function_name == 'address_to_coords':
        cols = find_column_indices(worksheet, start_row, {'lat': 'Широта', 'lon': 'Долгота', 'addr': 'Адрес'})
        if not all(k in cols for k in ['lat', 'lon', 'addr']):
            raise ValueError("Не найдены обязательные столбцы: 'Широта', 'Долгота', 'Адрес'")

        addresses_to_process, rows_to_update = [], []
        for i, row_cells in enumerate(worksheet.iter_rows(min_row=start_row + 1)):
            addr_val = row_cells[cols['addr'] - 1].value
            if addr_val:
                addresses_to_process.append(str(addr_val))
                rows_to_update.append(start_row + 1 + i)

        if addresses_to_process:
            try:
                task_statuses[task_id]['status'] = f"Геокодирование (адрес->координаты): {len(addresses_to_process)} адресов..."
                results = dadata.clean(name="address", source=addresses_to_process)
                for i, result in enumerate(results):
                    if result and result['geo_lat'] and result['geo_lon']:
                        worksheet.cell(row=rows_to_update[i], column=cols['lat']).value = float(result['geo_lat'])
                        worksheet.cell(row=rows_to_update[i], column=cols['lon']).value = float(result['geo_lon'])
            except Exception as e:
                print(f"Ошибка геокодирования DaData: {e}")

# --- Функции рефакторинга (без изменений) ---
# ... (вставьте сюда ваш рефакторингованный код из предыдущего шага)
def _apply_manual_rules(source_ws, template_ws, rules, s_start_row, t_start_row, used_source_cols, used_template_cols):
    """Применяет правила, заданные вручную (частные и из шаблона)."""
    s_end_row = source_ws.max_row
    for rule in rules:
        s_col_letter = rule.get('s_col') or get_col_from_cell(rule.get('source_cell'))
        t_col_letter = rule.get('t_col') or rule.get('template_col')

        if not s_col_letter or not t_col_letter:
            continue

        s_col_idx, t_col_idx = column_index_from_string(s_col_letter), column_index_from_string(t_col_letter)

        if s_col_idx in used_source_cols or t_col_idx in used_template_cols:
            continue

        for i, row in enumerate(source_ws.iter_rows(min_row=s_start_row + 1, max_row=s_end_row, min_col=s_col_idx, max_col=s_col_idx)):
            source_cell = row[0]
            target_cell = template_ws.cell(row=t_start_row + 1 + i, column=t_col_idx)

            # --- НОВАЯ ЛОГИКА ---
            target_cell.value = source_cell.value
            if source_cell.hyperlink:
                target_cell.hyperlink = source_cell.hyperlink.target
                target_cell.style = "Hyperlink"
            # --- КОНЕЦ НОВОЙ ЛОГИКИ ---

        used_source_cols.add(s_col_idx)
        used_template_cols.add(t_col_idx)


def _apply_dictionary_matching(source_ws, template_ws, s_start_row, t_start_row, used_source_cols, used_template_cols):
    """Сопоставляет столбцы на основе словаря синонимов."""
    reverse_dictionary = dictionary_matcher.get_reverse_dictionary()
    s_headers = {c.column: normalize_header(c.value) for c in source_ws[s_start_row] if c.value}
    t_headers = {c.column: normalize_header(c.value) for c in template_ws[t_start_row] if c.value}
    s_end_row = source_ws.max_row

    for s_col_idx, s_norm_h in s_headers.items():
        if s_col_idx in used_source_cols: continue
        canonical_name = reverse_dictionary.get(s_norm_h)
        if not canonical_name: continue
        normalized_canonical = normalize_header(canonical_name)

        for t_col_idx, t_norm_h in t_headers.items():
            if t_col_idx in used_template_cols: continue
            if t_norm_h == normalized_canonical:
                for i, row in enumerate(
                        source_ws.iter_rows(min_row=s_start_row + 1, max_row=s_end_row, min_col=s_col_idx,
                                            max_col=s_col_idx)):
                    source_cell = row[0]
                    target_cell = template_ws.cell(row=t_start_row + 1 + i, column=t_col_idx)

                    # --- НОВАЯ ЛОГИКА ---
                    target_cell.value = source_cell.value
                    if source_cell.hyperlink:
                        target_cell.hyperlink = source_cell.hyperlink.target
                        target_cell.style = "Hyperlink"
                    # --- КОНЕЦ НОВОЙ ЛОГИКИ ---

                used_source_cols.add(s_col_idx)
                used_template_cols.add(t_col_idx)
                break


def _apply_auto_matching(source_ws, template_ws, s_start_row, t_start_row, used_source_cols, used_template_cols,
                         task_id):
    """Автоматически сопоставляет столбцы по схожести названий."""
    s_headers = {c.column: normalize_header(c.value) for c in source_ws[s_start_row] if c.value}
    t_headers = {c.column: normalize_header(c.value) for c in template_ws[t_start_row] if c.value}
    s_end_row = source_ws.max_row

    auto_source_headers = {k: v for k, v in s_headers.items() if k not in used_source_cols}
    auto_template_headers = {k: v for k, v in t_headers.items() if k not in used_template_cols}
    auto_matches = {}

    for s_col, s_norm in auto_source_headers.items():
        best_match, best_score = None, 0
        for t_col, t_norm in auto_template_headers.items():
            score = fuzz.ratio(s_norm, t_norm)
            if score > best_score:
                best_score, best_match = score, t_col
        if best_score > 75:
            auto_matches[s_col] = best_match
            auto_template_headers = {k: v for k, v in auto_template_headers.items() if k != best_match}

    rows_to_process = list(source_ws.iter_rows(min_row=s_start_row + 1, max_row=s_end_row))
    for i, row in enumerate(rows_to_process):
        for s_col, t_col in auto_matches.items():
            source_cell = row[s_col - 1]
            target_cell = template_ws.cell(row=t_start_row + 1 + i, column=t_col)

            # --- НОВАЯ ЛОГИКА ---
            target_cell.value = source_cell.value
            if source_cell.hyperlink:
                target_cell.hyperlink = source_cell.hyperlink.target
                target_cell.style = "Hyperlink"
            # --- КОНЕЦ НОВОЙ ЛОГИКИ ---

        if (i + 1) % 100 == 0:
            task_statuses[task_id]['status'] = f'Автоматическое копирование: строка {i + 1}'


@app.route('/templates/edit/<template_id>', methods=['GET', 'POST'])
def edit_template(template_id):
    # --- ДИАГНОСТИКА 1 ---
    print(f"--- ВХОД В ФУНКЦИЮ edit_template ---")
    print(f"Метод запроса: {request.method}")
    print(f"ID шаблона из URL: '{template_id}'")

    json_path = os.path.join(app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json")
    if not os.path.exists(json_path):
        print(f"!!! ШАБЛОН НЕ НАЙДЕН по пути: '{json_path}'")
        flash("Шаблон не найден.", "error")
        return redirect(url_for('templates_list'))

    if request.method == 'POST':
        # --- ДИАГНОСТИКА 2 ---
        print(f"--> Зашли в блок POST для ID: '{template_id}'")
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)

            # Обновляем данные из формы
            template_data['template_name'] = request.form.get('template_name')
            template_data['header_start_cell'] = request.form.get('header_start_cell').upper()

            # Обновляем правила
            rules = []
            source_cells = request.form.getlist('source_cell')
            template_cols = request.form.getlist('template_col')
            for i in range(len(source_cells)):
                if source_cells[i] and template_cols[i]:
                    rules.append({
                        "source_cell": source_cells[i].upper(),
                        "template_col": template_cols[i].upper()
                    })
            template_data['rules'] = rules

            # Сохраняем обновленные данные в JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, ensure_ascii=False, indent=4)

            print(f"--> POST-запрос успешен. Перенаправляем на templates_list.")
            flash(f"Шаблон '{template_data['template_name']}' успешно обновлен!", "success")
            return redirect(url_for('templates_list'))
        except Exception as e:
            print(f"!!! ОШИБКА в блоке POST: {e}")
            flash(f"Произошла ошибка при обновлении: {e}", "error")
            return redirect(url_for('edit_template', template_id=template_id))

    # --- ДИАГНОСТИКА 3 ---
    print(f"--> Зашли в блок GET. Готовим страницу для ID: '{template_id}'")
    # При GET-запросе просто показываем форму с текущими данными
    with open(json_path, 'r', encoding='utf-8') as f:
        template_data = json.load(f)

    return render_template('edit_template.html', template=template_data, template_id=template_id)



# --- 4. Основная функция обработки (без изменений) ---
def process_excel_hybrid(task_id, source_path, template_path, ranges, template_rules, private_rules, post_function):
    try:
        task_statuses[task_id] = {'progress': 5, 'status': 'Подготовка...'}
        source_wb = load_workbook(source_path)
        source_ws = source_wb.active
        template_wb = load_workbook(template_path)
        template_ws = template_wb.active

        s_start_row, t_start_row = ranges['s_start_row'], ranges['t_start_row']
        used_source_cols, used_template_cols = set(), set()

        task_statuses[task_id]['status'] = 'Выполняю частные правила...'
        _apply_manual_rules(source_ws, template_ws, private_rules, s_start_row, t_start_row, used_source_cols, used_template_cols)
        task_statuses[task_id]['status'] = 'Применяю правила из шаблона...'
        _apply_manual_rules(source_ws, template_ws, template_rules, s_start_row, t_start_row, used_source_cols, used_template_cols)
        task_statuses[task_id]['status'] = 'Проверяю по словарю...'
        _apply_dictionary_matching(source_ws, template_ws, s_start_row, t_start_row, used_source_cols, used_template_cols)
        task_statuses[task_id]['status'] = 'Ищу автоматические совпадения...'
        _apply_auto_matching(source_ws, template_ws, s_start_row, t_start_row, used_source_cols, used_template_cols, task_id)
        task_statuses[task_id]['status'] = 'Запускаю пост-обработку...'
        apply_post_processing(task_id, template_wb, t_start_row, post_function)

        task_statuses[task_id]['status'] = 'Сохраняю результат...'
        processed_filename = f"{task_id}.xlsx"
        template_wb.save(os.path.join(app.config['PROCESSED_FOLDER'], processed_filename))
        source_wb.close()
        task_statuses[task_id].update({'progress': 100, 'status': 'Готово!', 'result_file': processed_filename})

    except Exception as e:
        task_statuses[task_id].update({'progress': 100, 'status': f"Ошибка: {e}", 'result_file': None})


# --- Роуты Flask с изменениями ---

@app.route('/templates')
def templates_list():
    # ... (код без изменений)
    template_files = glob.glob(os.path.join(app.config['TEMPLATES_DB_FOLDER'], '*.json'))
    templates_data = []
    for f in template_files:
        try:
            with open(f, 'r', encoding='utf-8') as file:
                data = json.load(file);
                data['id'] = os.path.basename(f).replace('.json', '');
                templates_data.append(data)
        except Exception as e:
            print(f"Ошибка чтения шаблона {f}: {e}")
    return render_template('templates_list.html', templates=templates_data)

@app.route('/templates/new')
def new_template_form():
    return render_template('create_template.html')


@app.route('/templates/create', methods=['POST'])
def create_template():
    try:
        template_name = request.form.get('template_name')
        header_start_cell = request.form.get('header_start_cell').upper()
        excel_file = request.files.get('excel_file')
        if not (template_name and header_start_cell and excel_file and allowed_file(excel_file.filename)):
            # <--- ИЗМЕНЕНО: Обновлен текст ошибки
            flash("Ошибка: все поля должны быть заполнены, и файл должен быть .xlsx или .xlsm", "error");
            return redirect(url_for('new_template_form'))
        # ... (остальной код без изменений)
        template_id = str(uuid.uuid4());
        excel_filename = f"{template_id}.xlsx"
        excel_file.save(os.path.join(app.config['TEMPLATES_DB_FOLDER'], excel_filename))
        rules = []
        source_cells, template_cols = request.form.getlist('source_cell'), request.form.getlist('template_col')
        for i in range(len(source_cells)):
            if source_cells[i] and template_cols[i]: rules.append(
                {"source_cell": source_cells[i].upper(), "template_col": template_cols[i].upper()})
        template_data = {"template_name": template_name, "excel_file": excel_filename,
                         "header_start_cell": header_start_cell, "rules": rules}
        with open(os.path.join(app.config['TEMPLATES_DB_FOLDER'], f"{template_id}.json"), 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=4)
        flash(f"Шаблон '{template_name}' успешно создан!", "success");
        return redirect(url_for('templates_list'))
    except Exception as e:
        flash(f"Произошла ошибка: {e}", "error");
        return redirect(url_for('new_template_form'))


@app.route('/dictionary')
def dictionary_ui():
    # ... (код без изменений)
    return render_template('dictionary.html', dictionary=dictionary_matcher.load_dictionary())

@app.route('/dictionary/add', methods=['POST'])
def add_to_dictionary():
    # ... (код без изменений)
    canonical_name, synonyms = request.form.get('canonical_name'), request.form.get('synonyms')
    if canonical_name and synonyms: dictionary_matcher.add_entry(canonical_name, synonyms)
    return redirect(url_for('dictionary_ui'))


@app.route('/dictionary/delete', methods=['POST'])
def delete_from_dictionary():
    # ... (код без изменений)
    canonical_name = request.form.get('canonical_name')
    if canonical_name: dictionary_matcher.delete_entry(canonical_name)
    return redirect(url_for('dictionary_ui'))


@app.route('/')
def index():
    # ... (код без изменений)
    template_files = glob.glob(os.path.join(app.config['TEMPLATES_DB_FOLDER'], '*.json'))
    templates_data = []
    for f in template_files:
        try:
            with open(f, 'r', encoding='utf-8') as file:
                data = json.load(file);
                data['id'] = os.path.basename(f).replace('.json', '');
                templates_data.append(data)
        except Exception as e:
            print(f"Ошибка чтения шаблона {f}: {e}")
    return render_template('index.html', saved_templates=templates_data)


@app.route('/process', methods=['POST'])
def process_files():
    try:
        source_file = request.files.get('source_file')
        if not (source_file and source_file.filename and allowed_file(source_file.filename)):
            # <--- ИЗМЕНЕНО: Обновлен текст ошибки
            return jsonify({'error': 'Исходный файл .xlsx или .xlsm должен быть загружен.'}), 400
        source_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(source_file.filename));
        source_file.save(source_path)

        template_rules, selected_template_id = [], request.form.get('saved_template')
        if selected_template_id:
            json_path = os.path.join(app.config['TEMPLATES_DB_FOLDER'], f"{selected_template_id}.json")
            with open(json_path, 'r', encoding='utf-8') as f:
                template_info = json.load(f)
            template_path = os.path.join(app.config['TEMPLATES_DB_FOLDER'], template_info['excel_file'])
            template_rules, t_start_cell = template_info.get('rules', []), template_info['header_start_cell']
        else:
            template_file = request.files.get('template_file')
            if not (template_file and template_file.filename and allowed_file(template_file.filename)):
                 # <--- ИЗМЕНЕНО: Обновлен текст ошибки
                return jsonify({'error': 'Если шаблон не выбран, его нужно загрузить вручную (.xlsx или .xlsm).'}), 400
            template_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(template_file.filename));
            template_file.save(template_path)
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
        task_id = str(uuid.uuid4())
        thread = threading.Thread(target=process_excel_hybrid, args=(
            task_id, source_path, template_path, ranges, template_rules, private_rules, post_function))
        thread.start()
        return jsonify({'task_id': task_id})
    except Exception as e:
        return jsonify({'error': f'Ошибка на сервере: {e}'}), 500


@app.route('/status/<task_id>')
def task_status(task_id):
    return jsonify(task_statuses.get(task_id, {}))

def get_cell_content(cell):
    """
    Извлекает содержимое ячейки. Если есть гиперссылка,
    возвращает строку формата "Значение (URL)".
    """
    if cell.hyperlink and cell.hyperlink.target:
        return f"{cell.value} ({cell.hyperlink.target})"
    return cell.value
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=5012)