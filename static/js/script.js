document.addEventListener('DOMContentLoaded', function() {
    // --- Инициализация для главной страницы ---
    if (document.getElementById('upload-form')) {
        initIndexPage();
    }

    // --- Инициализация для страницы создания/редактирования шаблона ---
    if (document.querySelector('form[action*="/templates/"]')) {
        initTemplatePage();
    }
});

/**
 * Инициализирует все скрипты для главной страницы (index.html)
 */
function initIndexPage() {
    const savedTemplateSelect = document.getElementById('saved_template');
    const uploadForm = document.getElementById('upload-form');
    const addRuleButton = document.querySelector('#rules-container + .btn-secondary');

    savedTemplateSelect.addEventListener('change', toggleManualTemplateFields);
    toggleManualTemplateFields();

    if (addRuleButton) {
        addRuleButton.addEventListener('click', () => addRule('index'));
    }

    uploadForm.addEventListener('submit', handleFormSubmit);
}

/**
 * Инициализирует скрипты для страниц создания и редактирования шаблонов.
 */
function initTemplatePage() {
    const addRuleButton = document.querySelector('#rules-container + .btn-secondary');
    if (addRuleButton) {
        addRuleButton.addEventListener('click', () => addRule('template'));

        // --- ИСПРАВЛЕНИЕ ЗДЕСЬ ---
        // Добавляем пустое правило, только если это страница СОЗДАНИЯ (/new)
        // и на ней еще нет ни одного правила.
        if (window.location.pathname.includes('/new') && !document.querySelector('.rule')) {
            addRule('template');
        }
        // --- КОНЕЦ ИСПРАВЛЕНИЯ ---
    }

    // Этот код корректно навешивает обработчики на уже существующие кнопки "X" на странице редактирования
    document.querySelectorAll('.rule .btn-danger').forEach(button => {
        const ruleDiv = button.closest('.rule');
        if(ruleDiv) {
            button.addEventListener('click', () => removeRule(ruleDiv.id));
        }
    });
}


/**
 * Переключает видимость полей для ручной загрузки ШАБЛОНА на главной странице.
 */
function toggleManualTemplateFields() {
    const savedTemplateSelect = document.getElementById('saved_template');
    const manualTemplateContainer = document.getElementById('manual-template-container');
    const templateRangeContainer = document.getElementById('template-range-container');
    const templateFileInput = document.getElementById('template_file');
    const templateRangeInput = document.getElementById('template_range_start');

    if (savedTemplateSelect.value) {
        manualTemplateContainer.style.display = 'none';
        templateRangeContainer.style.display = 'none';
        templateFileInput.removeAttribute('required');
        templateRangeInput.removeAttribute('required');
    } else {
        manualTemplateContainer.style.display = 'block';
        templateRangeContainer.style.display = 'block';
        templateFileInput.setAttribute('required', 'required');
        templateRangeInput.setAttribute('required', 'required');
    }
}

/**
 * Добавляет новое поле для правила сопоставления.
 * @param {string} pageType - 'index' или 'template' для определения типа полей.
 */
function addRule(pageType = 'index') {
    const container = document.getElementById('rules-container');
    if (!container) return;

    const ruleId = `rule-${Date.now()}`;
    const ruleDiv = document.createElement('div');
    ruleDiv.className = 'rule';
    ruleDiv.id = ruleId;

    let ruleHtml = '';
    if (pageType === 'template') {
        ruleHtml = `
            <div class="rule-inputs" style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
                <div>
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Ячейка в исходнике</label>
                    <input type="text" name="source_cell" placeholder="напр. A1" required>
                </div>
                <div>
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Столбец в шаблоне</label>
                    <input type="text" name="template_col" placeholder="напр. C" required>
                </div>
            </div>
            <button type="button" class="btn btn-danger">X</button>
        `;
    } else {
        ruleHtml = `
            <div class="rule-inputs">
                <input type="text" name="manual_source_col" placeholder="Столбец в исходнике (напр. A)" required>
                <input type="text" name="manual_template_col" placeholder="Столбец в шаблоне (напр. C)" required>
            </div>
            <button type="button" class="btn btn-danger">X</button>
        `;
    }

    ruleDiv.innerHTML = ruleHtml;
    ruleDiv.querySelector('.btn-danger').addEventListener('click', () => removeRule(ruleId));

    container.appendChild(ruleDiv);
}

/**
 * Удаляет поле для правила сопоставления.
 * @param {string} id - Уникальный ID элемента правила.
 */
function removeRule(id) {
    const ruleElement = document.getElementById(id);
    if (ruleElement) {
        ruleElement.remove();
    }
}


// --- Логика отправки формы и опроса статуса (без изменений) ---
let pollingInterval;
async function handleFormSubmit(event) {
    event.preventDefault();
    const form = event.target;
    const submitButton = form.querySelector('button[type="submit"]');
    const progressBarContainer = document.getElementById('progress-container');
    const formData = new FormData(form);
    progressBarContainer.style.display = 'block';
    submitButton.disabled = true;
    submitButton.textContent = 'Обработка...';
    updateProgress(0, 'Загрузка файлов на сервер...');
    try {
        const response = await fetch('/process', { method: 'POST', body: formData });
        const result = await response.json();
        if (response.ok && result.task_id) {
            updateProgress(5, 'Файлы загружены, начинаю обработку...');
            pollingInterval = setInterval(() => pollStatus(result.task_id), 2000);
        } else {
            throw new Error(result.error || 'Неизвестная ошибка сервера.');
        }
    } catch (error) {
        updateProgress(0, `Ошибка отправки: ${error.message}`, true);
        submitButton.disabled = false;
        submitButton.textContent = 'Обработать и скачать';
    }
}
async function pollStatus(taskId) {
    const submitButton = document.querySelector('#upload-form button[type="submit"]');
    const progressBarContainer = document.getElementById('progress-container');
    try {
        const statusResponse = await fetch(`/status/${taskId}`);
        const statusData = await statusResponse.json();
        updateProgress(statusData.progress, statusData.status, statusData.status.startsWith('Ошибка:'));
        if (statusData.progress >= 100) {
            clearInterval(pollingInterval);
            submitButton.disabled = false;
            submitButton.textContent = 'Обработать и скачать';
            if (statusData.result_file) {
                updateProgress(100, "Готово! Загрузка начинается...");
                window.location.href = `/download/${statusData.result_file}`;
                setTimeout(() => { progressBarContainer.style.display = 'none'; }, 5000);
            }
        }
    } catch (error) {
        updateProgress(0, `Ошибка опроса статуса: ${error.message}`, true);
        clearInterval(pollingInterval);
        submitButton.disabled = false;
        submitButton.textContent = 'Обработать и скачать';
    }
}
function updateProgress(percentage, statusText, isError = false) {
    const progressBar = document.getElementById('progress-bar');
    const progressStatus = document.getElementById('progress-status');
    percentage = Math.round(percentage) || 0;
    progressBar.style.width = `${percentage}%`;
    progressBar.textContent = `${percentage}%`;
    progressStatus.textContent = `Статус: ${statusText}`;
    if (isError) {
        progressBar.style.backgroundColor = '#dc3545';
    } else {
        progressBar.style.backgroundColor = '#007bff';
    }
}