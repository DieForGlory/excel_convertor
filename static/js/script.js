document.addEventListener('DOMContentLoaded', function() {

    // --- ЛОГИКА ВАЛИДАЦИИ ФОРМЫ ---
    function validateForm() {
        const errors = [];
        const cellRegex = /^[A-Z]+[1-9][0-9]*$/i; // Регулярное выражение для ячеек типа A1, B12
        const allowedExtensions = ['xlsx', 'xlsm'];

        const sourceFile = document.getElementById('source_file').files[0];
        const sourceRange = document.getElementById('source_range_start').value;
        const savedTemplate = document.getElementById('saved_template').value;
        const templateFile = document.getElementById('template_file').files[0];
        const templateRange = document.getElementById('template_range_start').value;

        // 1. Проверка исходного файла
        if (!sourceFile) {
            errors.push('Необходимо загрузить исходный файл.');
        } else {
            const fileExt = sourceFile.name.split('.').pop().toLowerCase();
            if (!allowedExtensions.includes(fileExt)) {
                errors.push(`Недопустимый формат исходного файла. Разрешены только .${allowedExtensions.join(', .')}.`);
            }
        }

        // 2. Проверка ячейки исходного файла
        if (!sourceRange) {
            errors.push('Необходимо указать начальную ячейку для исходного файла.');
        } else if (!cellRegex.test(sourceRange)) {
            errors.push('Некорректный формат начальной ячейки для исходного файла (пример: A1).');
        }

        // 3. Проверка шаблона (если не выбран сохраненный)
        if (!savedTemplate) {
            if (!templateFile) {
                errors.push('Если не выбран сохраненный шаблон, необходимо загрузить файл шаблона.');
            } else {
                 const fileExt = templateFile.name.split('.').pop().toLowerCase();
                if (!allowedExtensions.includes(fileExt)) {
                    errors.push(`Недопустимый формат файла шаблона. Разрешены только .${allowedExtensions.join(', .')}.`);
                }
            }

            if (!templateRange) {
                errors.push('Необходимо указать начальную ячейку для файла шаблона.');
            } else if (!cellRegex.test(templateRange)) {
                errors.push('Некорректный формат начальной ячейки для файла шаблона (пример: A1).');
            }
        }

        return errors;
    }

    // --- ОСНОВНАЯ ЛОГИКА ---
    const form = document.getElementById('process-form');
    if (form) {
        form.addEventListener('submit', function(event) {
            event.preventDefault(); // Всегда останавливаем отправку сначала

            const errorContainer = document.getElementById('error-messages');
            const errors = validateForm();

            if (errors.length > 0) {
                // Если есть ошибки, показываем их
                errorContainer.innerHTML = '<strong>Обнаружены ошибки:</strong><ul>' + errors.map(e => `<li>${e}</li>`).join('') + '</ul>';
                errorContainer.style.display = 'block';
                window.scrollTo(0, 0); // Прокрутить наверх, чтобы увидеть ошибки
                return; // Прекращаем выполнение
            } else {
                // Если ошибок нет, скрываем контейнер и продолжаем отправку
                errorContainer.style.display = 'none';
            }

            // Старая логика отправки формы через AJAX
            const formData = new FormData(form);
            const progressBar = document.getElementById('progress-bar');
            const progressContainer = document.getElementById('progress-container');
            const statusText = document.getElementById('status-text');
            const downloadLink = document.getElementById('download-link');
            const startButton = form.querySelector('button[type="submit"]');

            progressContainer.style.display = 'block';
            progressBar.style.width = '0%';
            statusText.textContent = 'Загрузка файлов на сервер...';
            downloadLink.style.display = 'none';
            startButton.disabled = true;

            fetch('/process', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    statusText.textContent = `Ошибка: ${data.error}`;
                    startButton.disabled = false;
                    return;
                }
                statusText.textContent = 'Файлы в очереди на обработку...';
                pollStatus(data.task_id);
            })
            .catch(error => {
                statusText.textContent = `Ошибка сети: ${error}`;
                startButton.disabled = false;
            });
        });
    }

    function pollStatus(taskId) {
        const progressBar = document.getElementById('progress-bar');
        const statusText = document.getElementById('status-text');
        const downloadLink = document.getElementById('download-link');
        const startButton = form.querySelector('button[type="submit"]');

        const interval = setInterval(() => {
            fetch(`/status/${taskId}`)
            .then(response => response.json())
            .then(data => {
                progressBar.style.width = `${data.progress || 0}%`;
                statusText.textContent = data.status || 'Ожидание...';

                if ((data.progress && data.progress >= 100) || data.result_file) {
                    clearInterval(interval);
                    if (data.result_file) {
                        downloadLink.href = `/download/${data.result_file}`;
                        downloadLink.style.display = 'block';
                        statusText.textContent = 'Обработка завершена!';
                    }
                    startButton.disabled = false;
                }
            })
            .catch(error => {
                clearInterval(interval);
                statusText.textContent = `Ошибка при проверке статуса: ${error}`;
                startButton.disabled = false;
            });
        }, 2000);
    }

    // Логика для ручных правил и скрытия/показа полей шаблона
    const savedTemplateSelect = document.getElementById('saved_template');
    const newTemplateFields = document.getElementById('new-template-fields');

    if (savedTemplateSelect && newTemplateFields) {
         savedTemplateSelect.addEventListener('change', function() {
            if (this.value) {
                newTemplateFields.style.display = 'none';
            } else {
                newTemplateFields.style.display = 'block';
            }
        });
        // Изначальная проверка при загрузке страницы
        if (savedTemplateSelect.value) {
            newTemplateFields.style.display = 'none';
        }
    }

    const manualRulesContainer = document.getElementById('manual-rules-container');
    if(manualRulesContainer) {
        document.getElementById('add-manual-rule').addEventListener('click', function() {
            const newRule = document.createElement('div');
            newRule.className = 'manual-rule-row';
            newRule.innerHTML = `
                <input type="text" name="manual_source_col" placeholder="Колонка в исходнике (напр. A)">
                <span>→</span>
                <input type="text" name="manual_template_col" placeholder="Колонка в шаблоне (напр. C)">
                <button type="button" class="remove-rule-btn">🗑️</button>
            `;
            manualRulesContainer.appendChild(newRule);
        });

        manualRulesContainer.addEventListener('click', function(e) {
            if (e.target.classList.contains('remove-rule-btn')) {
                e.target.parentElement.remove();
            }
        });
    }
});