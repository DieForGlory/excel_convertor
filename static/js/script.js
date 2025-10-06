// static/js/script.js

document.addEventListener('DOMContentLoaded', function() {

    const form = document.getElementById('process-form');
    const savedTemplateSelect = document.getElementById('saved_template');
    const newTemplateFields = document.getElementById('new-template-fields');
    const addManualRuleBtn = document.getElementById('add-manual-rule');
    const manualRulesContainer = document.getElementById('manual-rules-container');

    // Показываем/скрываем поля для загрузки шаблона вручную
    if (savedTemplateSelect) {
        savedTemplateSelect.addEventListener('change', function() {
            newTemplateFields.style.display = this.value ? 'none' : 'block';
        });
        // Initial check
        newTemplateFields.style.display = savedTemplateSelect.value ? 'none' : 'block';
    }

    // Добавление ручных правил
    if (addManualRuleBtn) {
        addManualRuleBtn.addEventListener('click', function() {
            const ruleRow = document.createElement('div');
            ruleRow.className = 'rule-row';

            // --- ИЗМЕНЕНИЕ: Новая HTML-структура для правила ---
            ruleRow.innerHTML = `
                <div class="rule-input-group">
                    <label>Из колонки</label>
                    <input type="text" name="manual_source_col" placeholder="A" required>
                </div>
                <div class="rule-arrow">→</div>
                <div class="rule-input-group">
                    <label>В колонку</label>
                    <input type="text" name="manual_template_col" placeholder="B" required>
                </div>
                <button type="button" class="btn btn-danger btn-sm remove-rule-btn" style="align-self: center; margin-top: 1rem;">Удалить</button>
            `;
            manualRulesContainer.appendChild(ruleRow);
        });
    }

    // Удаление ручных правил
    if (manualRulesContainer) {
        manualRulesContainer.addEventListener('click', function(e) {
            if (e.target && e.target.classList.contains('remove-rule-btn')) {
                e.target.closest('.rule-row').remove();
            }
        });
    }

    // Обработка отправки формы
    if (form) {
        form.addEventListener('submit', function(e) {
            e.preventDefault();

            const formData = new FormData(form);
            const errorContainer = document.getElementById('error-messages');
            const progressContainer = document.getElementById('progress-container');
            const statusBar = document.getElementById('progress-bar');
            const statusText = document.getElementById('status-text');
            const downloadLink = document.getElementById('download-link');

            // Сбрасываем все перед новым запуском
            errorContainer.style.display = 'none';
            errorContainer.textContent = '';
            progressContainer.style.display = 'block';
            downloadLink.style.display = 'none';
            statusBar.style.width = '0%';
            statusBar.textContent = '';
            statusText.textContent = 'Отправляю файлы на сервер...';

            fetch(form.action, {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    throw new Error(data.error);
                }
                statusText.textContent = 'Файлы получены, начинаю обработку...';
                pollStatus(data.task_id);
            })
            .catch(error => {
                progressContainer.style.display = 'none';
                errorContainer.textContent = `Произошла ошибка: ${error.message}`;
                errorContainer.style.display = 'block';
            });
        });
    }

    // Функция опроса статуса задачи
    function pollStatus(taskId) {
        const statusUrl = `/status/${taskId}`;
        const interval = setInterval(() => {
            fetch(statusUrl)
            .then(response => response.json())
            .then(data => {
                const statusBar = document.getElementById('progress-bar');
                const statusText = document.getElementById('status-text');

                statusText.textContent = data.status || 'Обработка...';
                const progress = data.progress || 0;
                statusBar.style.width = `${progress}%`;
                statusBar.textContent = `${progress}%`;

                if (progress >= 100) {
                    clearInterval(interval);
                    if (data.result_file) {
                        const downloadLink = document.getElementById('download-link');
                        downloadLink.href = `/download/${data.result_file}`;
                        downloadLink.style.display = 'inline-block';
                        statusText.textContent = 'Готово! Ваш файл можно скачать.';
                    } else if(data.status && data.status.startsWith('Ошибка')) {
                        statusBar.style.backgroundColor = 'var(--error-color)';
                    }
                }
            })
            .catch(() => {
                clearInterval(interval);
                document.getElementById('status-text').textContent = 'Ошибка получения статуса.';
            });
        }, 1500);
    }
});