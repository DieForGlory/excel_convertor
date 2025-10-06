# app.py

from flask import Flask
import os

# Импортируем функцию-регистратор из нашего пакета routes
from routes import register_routes

def create_app():
    """
    Создает и конфигурирует экземпляр приложения Flask.
    """
    app = Flask(__name__)

    # 1. Загружаем конфигурацию из файла config.py
    app.config.from_pyfile('config.py')

    # 2. Создаем папки, если их нет
    os.makedirs(app.config['TEMPLATES_DB_FOLDER'], exist_ok=True)
    os.makedirs(app.config['TEMPLATE_EXCEL_FOLDER'], exist_ok=True)

    # 3. Глобальное хранилище статусов задач
    app.config['TASK_STATUSES'] = {}

    # 4. Регистрируем все роуты из папки /routes
    register_routes(app)

    return app

# 5. Запуск приложения
if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, port=5012)