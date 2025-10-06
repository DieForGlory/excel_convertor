from dadata import Dadata
from utils import find_column_indices

# --- Конфигурация Dadata ---
DADATA_API_KEY = "ВАШ_API_КЛЮЧ"
DADATA_SECRET_KEY = "ВАШ_СЕКРЕТНЫЙ_КЛЮЧ"

def apply_post_processing(task_id, workbook, start_row, function_name, task_statuses):
    """
    Применяет функции пост-обработки (геокодирование) с использованием Dadata.
    """
    if function_name == 'none' or not function_name:
        return

    worksheet = workbook.active
    dadata = Dadata(DADATA_API_KEY, DADATA_SECRET_KEY)

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