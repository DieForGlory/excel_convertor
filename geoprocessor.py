# geoprocessor.py

from openpyxl.utils import column_index_from_string
from local_geocoder import get_coordinates_from_local_db, get_address_from_local_db

def find_column_by_header(worksheet, start_row, header_name):
    """Находит индекс колонки по имени заголовка."""
    for cell in worksheet[start_row]:
        if cell.value and header_name.lower() in str(cell.value).lower():
            return cell.column
    return None

def process_geodata(task_id, wb, start_row, function_name, task_statuses):
    """
    Главная функция пост-обработки.
    Ищет адреса/координаты и заполняет пустые ячейки, используя локальную базу.
    """
    ws = wb.active

    # Определяем, что нужно сделать: адрес->координаты или наоборот
    if function_name == 'get_coords_from_address':
        # --- Логика для "Адрес -> Координаты" ---

        # Ищем колонки по стандартным названиям
        address_col_idx = find_column_by_header(ws, start_row, 'адрес')
        lat_col_idx = find_column_by_header(ws, start_row, 'широта')
        lon_col_idx = find_column_by_header(ws, start_row, 'долгота')

        if not all([address_col_idx, lat_col_idx, lon_col_idx]):
            task_statuses[task_id]['status'] = 'Ошибка: не найдены колонки "Адрес", "Широта" или "Долгота"'
            return

        for i, row in enumerate(ws.iter_rows(min_row=start_row + 1)):
            task_statuses[task_id]['status'] = f'Обрабатываю строку {i+1} (Адрес -> Координаты)'

            address_cell = row[address_col_idx - 1]
            lat_cell = row[lat_col_idx - 1]
            lon_cell = row[lon_col_idx - 1]

            # Запускаем поиск, только если есть адрес, а координат нет
            if address_cell.value and not lat_cell.value and not lon_cell.value:
                lat, lon = get_coordinates_from_local_db(str(address_cell.value))
                if lat and lon:
                    lat_cell.value = lat
                    lon_cell.value = lon

    elif function_name == 'get_address_from_coords':
        # --- Логика для "Координаты -> Адрес" ---

        address_col_idx = find_column_by_header(ws, start_row, 'адрес')
        lat_col_idx = find_column_by_header(ws, start_row, 'широта')
        lon_col_idx = find_column_by_header(ws, start_row, 'долгота')

        if not all([address_col_idx, lat_col_idx, lon_col_idx]):
            task_statuses[task_id]['status'] = 'Ошибка: не найдены колонки "Адрес", "Широта" или "Долгота"'
            return

        for i, row in enumerate(ws.iter_rows(min_row=start_row + 1)):
            task_statuses[task_id]['status'] = f'Обрабатываю строку {i+1} (Координаты -> Адрес)'

            address_cell = row[address_col_idx - 1]
            lat_cell = row[lat_col_idx - 1]
            lon_cell = row[lon_col_idx - 1]

            # Запускаем, только если есть координаты, а адреса нет
            if lat_cell.value and lon_cell.value and not address_cell.value:
                try:
                    lat = float(lat_cell.value)
                    lon = float(lon_cell.value)
                    address = get_address_from_local_db(lat, lon)
                    if address:
                        address_cell.value = address
                except (ValueError, TypeError):
                    continue # Игнорируем, если в ячейках не числа