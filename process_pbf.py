import osmium
import csv


class AddressHandler(osmium.SimpleHandler):
    def __init__(self, writer):
        super(AddressHandler, self).__init__()
        self.writer = writer
        self.processed_nodes = 0

    def node(self, n):
        """
        Эта функция вызывается для каждого точечного объекта (node) в PBF-файле.
        """
        # Проверяем, есть ли у объекта тег с названием улицы.
        # Это главный фильтр, чтобы отсеять ненужные объекты.
        if 'addr:street' in n.tags:
            try:
                # Собираем полный адрес из частей.
                # Функция get() вернет None, если тега нет, что безопасно.
                city = n.tags.get('addr:city', '')
                street = n.tags.get('addr:street', '')
                housenumber = n.tags.get('addr:housenumber', '')

                # Собираем части в одну строку, убирая лишние запятые и пробелы
                full_address = ", ".join(filter(None, [city, street, housenumber]))

                # Получаем координаты
                lon, lat = n.location.lon, n.location.lat

                # Записываем результат в наш CSV-файл
                self.writer.writerow([full_address, lat, lon])

                # Выводим прогресс в консоль, чтобы было видно, что скрипт работает
                self.processed_nodes += 1
                if self.processed_nodes % 1000 == 0:
                    print(f"Найдено и обработано адресов: {self.processed_nodes}")

            except osmium.InvalidLocationError:
                # Игнорируем объекты без корректных координат
                pass


def process_pbf_to_csv(pbf_file, csv_file):
    """
    Главная функция, которая запускает весь процесс.
    """
    print(f"Открываем {csv_file} для записи...")
    with open(csv_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Записываем заголовки в CSV
        writer.writerow(['address', 'latitude', 'longitude'])

        print(f"Начинаем обработку файла {pbf_file}...")
        # Создаем наш обработчик
        handler = AddressHandler(writer)

        # Запускаем обработку PBF-файла
        # `locations=True` обязательно для получения координат
        handler.apply_file(pbf_file, locations=True)

    print(f"\nГотово! Все адреса сохранены в {csv_file}")
    print(f"Всего найдено адресов: {handler.processed_nodes}")


# --- ЗАПУСК СКРИПТА ---
if __name__ == '__main__':
    # Укажите имя вашего PBF-файла
    pbf_input_file = 'russia-251005.osm.pbf'

    # Имя файла, в который сохраним результат
    csv_output_file = 'geobase.csv'

    try:
        process_pbf_to_csv(pbf_input_file, csv_output_file)
    except FileNotFoundError:
        print(f"Ошибка: файл '{pbf_input_file}' не найден!")
        print("Пожалуйста, убедитесь, что он лежит в той же папке, что и скрипт, и имя указано верно.")