# Архитектура решения

## ASSUMPTIONS

- Используется одна Excel-книга `DocumentControl.xlsm` как центральный файл.
- Шаблоны Word хранятся локально на файловом ресурсе, доступном пользователю.
- Политики ИБ допускают COM-автоматизацию Word из Excel.
- Пользователи работают с русской локалью Windows и офисными форматами OOXML (`.dotx`, `.docx`).
- Версионирование записей карточки документа выполняется полем `revision`.

## Риски

- Риск «висящих» процессов `WINWORD.EXE` при неаккуратной очистке COM-объектов.
- Риск хардкода путей (решается конфигурационным листом `cfg_app`).
- Риск смешивания UI и правил (решается разнесением модулей).
- Риск роста сложности проверок при добавлении новых типов документов.
- Риск пользовательских ошибок ручного редактирования шаблонов.

---

## Логическая архитектура

- **UI слой**: `frmDocumentCard`, `frmValidationReport`, кнопки на листе `ui_dashboard`.
- **Data Access слой**: `modExcelData`.
- **Business Rules слой**: `modValidation`, справочные листы `ref_*`, `rules_*`.
- **Word Automation слой**: `modWordAutomation`.
- **Reporting слой**: `modReport`.
- **Logging слой**: `modLogging` + лист `log_actions`.

---

## Структура Excel-книги

| Лист | Назначение |
|---|---|
| `ui_dashboard` | Кнопки сценариев и статус |
| `doc_cards` | Карточки документов (31 колонка) |
| `cfg_app` | Конфигурация путей, имён файлов, поведения |
| `ref_document_types` | Типы документов |
| `ref_templates` | Сопоставление типа документа и шаблона |
| `ref_statuses` | Допустимые статусы |
| `ref_users` | Автор/проверяющий/утверждающий |
| `rules_required_fields` | Обязательные поля по типу документа |
| `rules_required_sections` | Обязательные разделы |
| `rules_filename` | Шаблоны именования |
| `ea_clause_matrix` | Матрица покрытия базиса для Engineering Analysis |
| `ri_section_matrix` | Контрольная матрица разделов Repair Instruction |
| `validation_issues` | Результаты последней проверки |
| `log_actions` | Журнал действий |

### Колонки `doc_cards` (31 колонка)

```
document_id, document_type, title,
aircraft_model, aircraft_variant, aircraft_number, msn,
aircraft_manufacture_date, aircraft_hours, aircraft_cycles,
assembly_number, part_number, component_name,
component_sn, component_hours, component_cycles, component_manufacture_date,
applicability, revision, date, author, checker, approver,
related_analysis_number, related_instruction_number,
references, attachments, remarks, status,
word_doc_path, pdf_path
```

---

## Архитектура Word-шаблонов

### Формат

Все шаблоны хранятся как **валидные OOXML-файлы** (`.dotx` = zip с сигнатурой `PK`).  
Плоский текст недопустим — `modWordAutomation` требует OOXML; при иных форматах выбрасывает ошибку.

### Откуда берётся шаблон RI

`templates/RepairInstruction.dotx` построен на основе реального документа  
`A.095.04.00000.078.0699_example.docx`, хранящегося в `templates/`.  
Структура, стили, таблицы и колонтитулы полностью сохранены — заменены только конкретные значения.

### Маркеры `{{...}}`

Word-шаблоны содержат маркеры в служебных зонах (титул, таблица подписей, таблица применяемости, колонтитул).  
`modWordAutomation.ReplaceAllMarkers` выполняет замену через Word Find & Replace.  
Маркеры `{{...}}` ставятся единым run-ом, чтобы Word нашёл их без разрывов.

#### Маркеры `RepairInstruction.dotx`

| Зона документа | Маркеры |
|---|---|
| Титульная страница | `{{AircraftModel}}`, `{{Title}}`, `{{DocumentID}}`, `{{Revision}}` |
| Таблица подписей | `{{Author}}`, `{{Checker}}`, `{{Approver}}` |
| Лист регистрации изменений | `{{Revision}}`, `{{DocDate}}` |
| Таблица применяемости (4.2) | `{{AircraftModel}}`, `{{AircraftVariant}}`, `{{MSN}}`, `{{AircraftNumber}}`, `{{AircraftManufactureDate}}`, `{{AircraftHours}}`, `{{AircraftCycles}}`, `{{ComponentName}}`, `{{AssemblyNumber}}`, `{{ComponentSN}}`, `{{ComponentHours}}`, `{{ComponentCycles}}`, `{{ComponentManufactureDate}}` |
| Раздел 7 (Уведомление) | `{{DocumentID}}`, `{{Title}}` |
| Колонтитул | `{{DocumentID}}`, `{{Revision}}`, `{{DocDate}}` |

Всего: **20 уникальных маркеров**.

### Плейсхолдеры инженерного содержания

Секции, которые инженер заполняет вручную (3.1, 3.2, 5.1, 5.2 и т.д.),  
содержат текст `<<Заполняется инженером>>`.  
Валидатор обнаруживает этот текст через проверку `ContainsTrashPlaceholder` (`InStr(..., "<<")`).

### `EngineeringAnalysis.dotx`

Шаблон EA в текущей версии является плоским текстовым файлом-заглушкой.  
**Требует замены на полноценный OOXML-шаблон** (аналогично RI).

---

## Архитектура VBA-проекта

### Modules

| Модуль | Назначение |
|---|---|
| `modConstants` | Константы листов, статусов, кодов ошибок |
| `modConfig` | Чтение/проверка конфигурации из `cfg_app` |
| `modExcelData` | CRUD карточек, справочники, матрицы |
| `modWordAutomation` | Создание DOCX из шаблона, замена маркеров, экспорт PDF |
| `modValidation` | Проверки перед выпуском (14 правил) |
| `modReport` | Публикация отчёта в `validation_issues` |
| `modLogging` | Запись действий в `log_actions` |
| `modWorkbookSetup` | Инициализация структуры книги и справочников |
| `modMain` | Точки входа пользовательских сценариев |

### Classes

| Класс | Назначение |
|---|---|
| `clsDocumentCard` | Объект карточки документа (31 поле) |
| `clsValidationIssue` | Объект ошибки/предупреждения |
| `clsTemplateRule` | Правило шаблона/раздела |

### Forms

| Форма | Назначение |
|---|---|
| `frmDocumentCard` | Ввод/редактирование карточки (31 поле, 6 кнопок) |
| `frmValidationReport` | Просмотр отчёта: список замечаний + счётчики |

---

## Этапы реализации

| Этап | Статус | Содержание |
|---|---|---|
| 1. Анализ и архитектура | ✅ завершён | Проектирование структуры, модулей, шаблонов |
| 2. MVP | ✅ завершён | Карточка, 2 шаблона (RI полный, EA заглушка), DOCX/PDF, журнал |
| 3. Усиление проверок | ✅ завершён | 14 правил валидации, матрицы EA/RI, отчёт с кнопкой Close |
| 4. UX и поддерживаемость | 🔄 в работе | Форма расширена до 31 поля, загрузка по имени колонки |
| 5. Тесты и выпуск | ⬜ не начат | Тест-план, чек-лист приёмки, приложение примеров |
