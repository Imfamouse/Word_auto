# Руководство пользователя

## Сценарий 1: Создание документа

1. Откройте `DocumentControl.xlsm`.
2. Один раз запустите `modMain.AppInitialize` (создаст необходимые листы, заголовки и базовые справочники).
3. Нажмите кнопку **Новая карточка** (`modMain.OpenDocumentCard`).
4. Заполните реквизиты документа.
5. Выберите `document_type` (Repair Instruction или Engineering Analysis).
6. Нажмите **Создать Word из шаблона** (`modMain.CreateWordDocument`).

## Сценарий 2: Ручное написание инженерного содержания

1. Откройте созданный DOCX.
2. Заполняйте только инженерные разделы вручную.
3. Не удаляйте служебные заголовки и обязательные секции.

## Сценарий 3: Предвыпускная проверка

1. Нажмите **Проверить документ** (`modMain.ValidateCurrentDocument`).
2. Изучите отчёт в форме `frmValidationReport` или листе `validation_issues`.
3. Исправьте ошибки в Word/Excel.
4. Запустите проверку повторно.

## Сценарий 4: Выпуск

1. После статуса проверки без ошибок нажмите **Экспорт PDF** (`modMain.ExportCurrentToPdf`).
2. DOCX и PDF пути записываются в карточку документа.

## Типовые ошибки пользователя

- Остались маркеры `{{...}}` в колонтитуле.
- В тексте есть `TBD`/`XXX`/`???`/`sample`/`draft`.
- В карточке не заполнены обязательные поля.
- Отсутствует обязательный раздел для выбранного типа документа.


### Кнопки в `frmDocumentCard`

- **Save Card** — сохранить карточку в `doc_cards`.
- **Create DOCX** — создать Word-документ из шаблона и записать `word_doc_path`.
- **Validate** — запустить проверки и открыть `frmValidationReport`.
- **Export PDF** — экспортировать DOCX в PDF и записать `pdf_path`.
- **Close** — закрыть форму.

### Что вводить в каждое поле `frmDocumentCard`

- `Document ID` — уникальный номер документа (пример: `RI-2026-001`).
- `Document Type` — тип документа: `Repair Instruction` или `Engineering Analysis`.
- `Title` — краткий технический заголовок документа.
- `Aircraft Model` — тип/модель ВС (пример: `A320`).
- `Aircraft Number` — бортовой номер.
- `MSN` — заводской серийный номер.
- `Assembly Number` — номер узла/сборки.
- `Part Number` — номер детали.
- `Component Name` — название компонента.
- `Applicability` — применимость (к каким бортам/условиям применимо).
- `Revision` — ревизия документа.
- `Date` — дата документа в формате `YYYY-MM-DD`.
- `Author` — автор (инженер).
- `Checker` — проверяющий.
- `Approver` — утверждающий.
- `Related Analysis #` — связанный номер Engineering Analysis (если есть).
- `Related Instruction #` — связанный номер Repair Instruction (если есть).
- `References` — ссылки на нормативные документы/руководства.
- `Attachments` — перечень приложений/файлов.
- `Remarks` — примечания.
- `Status` — статус (`Draft`, `In Review`, `Released`).
- `Word Doc Path` — путь к сгенерированному DOCX (обычно заполняется системой).
- `PDF Path` — путь к сгенерированному PDF (обычно заполняется системой).

### Что делают кнопки в форме

- `Save Card` — сохраняет/обновляет карточку на листе `doc_cards`.
- `Create DOCX` — создаёт Word-документ из шаблона и записывает `Word Doc Path`.
- `Validate` — запускает проверки и показывает `frmValidationReport`.
- `Export PDF` — экспортирует DOCX в PDF и записывает `PDF Path`.
- `Close` — закрывает форму.
- `Field Help` — показывает встроенную подсказку по всем полям и кнопкам.

### Если `Create DOCX` не создаёт Word-файл

Проверьте по шагам:

1. В `cfg_app` задан корректный `templates_path`.
2. В `ref_templates` есть строка для вашего `Document Type`.
3. Файл шаблона реально существует в `templates_path`.
4. Для `Document Type` допустимы значения `Repair Instruction` / `Engineering Analysis` и алиасы `RI`/`EA`/`РИ`.
5. Word установлен и доступен через COM Automation.

Теперь сообщение ошибки возвращает первопричину (не только "DOCX creation failed").
