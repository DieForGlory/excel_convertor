import pandas as pd
from scipy.spatial import cKDTree
import numpy as np

# --- Секция для прямого геокодирования (адрес -> координаты) ---

try:
    df = pd.read_csv('geobase.csv', encoding='utf-8')
    # Готовим данные для поиска адресов
    address_map = {row['address']: (row['latitude'], row['longitude']) for index, row in df.iterrows()}
except FileNotFoundError:
    df = pd.DataFrame(columns=['address', 'latitude', 'longitude'])
    address_map = {}


def get_coordinates_from_local_db(address: str):
    """Ищет точное совпадение адреса в локальной базе."""
    return address_map.get(address, (None, None))


# --- Секция для ОБРАТНОГО геокодирования (координаты -> адрес) ---

# Готовим данные для быстрого поиска по координатам
# Проверяем, что база не пуста и в ней есть нужные колонки
if not df.empty and 'latitude' in df.columns and 'longitude' in df.columns:
    # Создаем "дерево" для быстрого поиска ближайших соседей
    # Это работает в тысячи раз быстрее, чем перебор всех строк
    coords = np.array(list(zip(df['latitude'].values, df['longitude'].values)))
    kdtree = cKDTree(coords)
else:
    kdtree = None


def get_address_from_local_db(latitude: float, longitude: float):
    """
    Находит ближайший адрес в базе по заданным координатам.
    """
    if kdtree is None:
        print("База для поиска по координатам пуста или некорректна.")
        return None

    # Ищем в дереве 1 ближайшую точку к нашим координатам
    distance, idx = kdtree.query([latitude, longitude], k=1)

    # Получаем адрес найденной точки по ее индексу
    nearest_address = df.iloc[idx]['address']

    return nearest_address