import json
import os
import re

# Определяем имя файла как константу для удобства
DICTIONARY_FILE = 'dictionary.json'


def load_dictionary():
    """
    Загружает словарь из JSON-файла.
    Если файл не найден или содержит ошибку, возвращает пустой словарь.
    """
    if not os.path.exists(DICTIONARY_FILE):
        return {}
    try:
        with open(DICTIONARY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        # В случае ошибки чтения или парсинга, возвращаем пустые данные
        return {}


def save_dictionary(data):
    """Сохраняет данные словаря в JSON-файл с красивым форматированием."""
    with open(DICTIONARY_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def get_reverse_dictionary(data=None):
    """
    Создает "обратный" словарь для быстрого поиска: {синоним: каноничное_имя}.
    Все ключи (синонимы) приводятся к нормализованному виду (нижний регистр, без спецсимволов).
    """
    if data is None:
        data = load_dictionary()

    reverse_map = {}
    for canonical_name, synonyms in data.items():
        # Добавляем и само каноничное имя в список вариантов для сопоставления
        all_variants = synonyms + [canonical_name]
        for variant in all_variants:
            normalized_variant = _normalize(variant)
            reverse_map[normalized_variant] = canonical_name
    return reverse_map


def add_entry(canonical_name, synonyms_str):
    """Добавляет или обновляет запись в словаре."""
    dictionary = load_dictionary()
    # --- ИЗМЕНЕНИЕ ЗДЕСЬ ---
    # Разделяем строку по вашему разделителю, а не по запятой
    synonyms = [s.strip() for s in synonyms_str.split('@1!') if s.strip()]
    dictionary[canonical_name] = synonyms
    save_dictionary(dictionary)


def delete_entry(canonical_name):
    """Удаляет запись (каноничное имя и все его синонимы) из словаря."""
    dictionary = load_dictionary()
    # Проверяем наличие ключа перед удалением, чтобы избежать ошибок
    if canonical_name in dictionary:
        del dictionary[canonical_name]
        save_dictionary(dictionary)


def _normalize(text):
    """
    Внутренняя функция для приведения текста к единому виду:
    нижний регистр, без пробелов и не-буквенно-цифровых символов.
    """
    if not isinstance(text, str):
        text = str(text)
    return re.sub(r'[\s\W_]+', '', text.lower())