# Архитектура решения

## Этап 1 — анализ и проектирование

### ASSUMPTIONS

- Используется одна Excel-книга `DocumentControl.xlsm` как центральный файл.
- Шаблоны Word хранятся локально на файловом ресурсе, доступном пользователю.
- Политики ИБ допускают COM-автоматизацию Word из Excel.
- Пользователи работают с русской локалью Windows и офисными форматами OOXML (`.dotx`, `.docx`).
- Версионирование записей карточки документа выполняется полем `revision`.

### Риски

- Риск «висящих» процессов WINWORD.EXE при неаккуратной очистке COM-объектов.
- Риск хардкода путей (решается конфигурационным листом `cfg_app`).
- Риск смешивания UI и правил (решается разнесением модулей).
- Риск роста сложности проверок при добавлении новых типов документов.
- Риск пользовательских ошибок ручного редактирования шаблонов.

## Логическая архитектура

- **UI слой**: `frmDocumentCard`, `frmValidationReport`, кнопки на листе `ui_dashboard`.
- **Data Access слой**: `modExcelData`.
- **Business Rules слой**: `modValidation`, справочные листы `ref_*`, `rules_*`.
- **Word Automation слой**: `modWordAutomation`.
- **Reporting слой**: `modReport`.
- **Logging слой**: `modLogging` + лист `log_actions`.

## Структура Excel-книги

- `ui_dashboard` — кнопки сценариев и статус.
- `doc_cards` — карточки документов.
- `cfg_app` — конфигурация путей, имён файлов, поведения.
- `ref_document_types` — типы документов.
- `ref_templates` — сопоставление типа документа и шаблона.
- `ref_statuses` — допустимые статусы.
- `ref_users` — автор/проверяющий/утверждающий.
- `rules_required_fields` — обязательные поля по типу документа.
- `rules_required_sections` — обязательные разделы.
- `rules_filename` — шаблоны именования.
- `ea_clause_matrix` — матрица покрытия базиса для Engineering Analysis.
- `ri_section_matrix` — контрольная матрица разделов Repair Instruction.
- `validation_issues` — результаты последней проверки.
- `log_actions` — журнал действий.

### Колонки `doc_cards`

`document_id, document_type, title, aircraft_model, aircraft_number, msn, assembly_number, part_number, component_name, applicability, revision, date, author, checker, approver, related_analysis_number, related_instruction_number, references, attachments, remarks, status, word_doc_path, pdf_path`.

## Архитектура Word-шаблонов

- Шаблоны:
  - `templates/RepairInstruction.dotx`
  - `templates/EngineeringAnalysis.dotx`
- Реквизиты через маркеры вида `{{MarkerName}}` в титуле/колонтитулах/служебных блоках.
- Обязательные разделы оформлены заголовками `Heading 1`/`Heading 2`.
- Свободные зоны для инженерного текста явно отмечены как `<<Engineer writes manually>>`.
- Валидация ищет:
  - незаменённые `{{...}}`,
  - мусорные заглушки,
  - наличие обязательных заголовков.

## Архитектура VBA-проекта

### Modules

- `modConstants` — константы листов, статусов, кодов.
- `modConfig` — чтение/проверка конфигурации.
- `modExcelData` — CRUD карточек, справочники, матрицы.
- `modWordAutomation` — создание документа из шаблона, подстановка реквизитов, экспорт PDF.
- `modValidation` — проверки перед выпуском.
- `modReport` — публикация отчёта в `validation_issues` + форма отчёта.
- `modLogging` — запись действий в `log_actions`.
- `modMain` — точки входа пользовательских сценариев.

### Classes

- `clsDocumentCard` — объект карточки документа.
- `clsValidationIssue` — объект ошибки/предупреждения.
- `clsTemplateRule` — правило шаблона/раздела.

### Forms

- `frmDocumentCard` — ввод/редактирование карточки.
- `frmValidationReport` — просмотр отчёта ошибок.

## Этапы реализации

1. **MVP**: карточка, справочники, 2 шаблона, создание DOCX/PDF, базовая валидация, журнал.
2. **Усиление проверок**: маркеры, заглушки, разделы, матрицы EA/RI.
3. **UX/поддержка**: формы, dashboard, сообщения, dev-docs.
4. **Тесты и приёмка**: тест-план, чек-лист, known limitations.

