# VALIDATION_RULES.md

## Уровни проверок

### 1. Ошибки-блокеры (ERROR)

При наличии блокирующих ошибок выпуск документа в PDF не должен считаться завершённым.

Текущий перечень:
- не заполнен `document_id`
- не заполнен `document_type`
- не заполнен `title`
- не заполнен `revision`
- не заполнена `date`
- не заполнен `aircraft_model`
- не заполнен `component_name`
- не заполнен `author`
- `word_doc_path` пуст (DOCX ещё не создан)
- файл по `word_doc_path` не найден на диске
- в документе остались маркеры вида `{{...}}`
- отсутствуют обязательные разделы (из `rules_required_sections`)
- для applicable clauses в Engineering Analysis отсутствует `status`
- для applicable clauses отсутствует `means_of_compliance`
- для applicable clauses отсутствует `covered_in_section`
- в RI-матрице обязательный раздел отмечен как отсутствующий

### 2. Предупреждения (WARNING)

Не блокируют выпуск, но требуют внимания.

Текущий перечень:
- не заполнен `aircraft_number`
- не заполнен `msn`
- не заполнен `assembly_number`
- не заполнен `checker`
- не заполнен `approver`
- в тексте документа найдены: `TBD`, `XXX`, `???`, `<<` (незаполненные секции инженера)
- в тексте найдено слово `sample` или `draft` как **целое слово** (не в пути к файлу)

---

## Проверка незаполненных инженерных секций

Шаблон `RepairInstruction.dotx` содержит плейсхолдеры `<<Заполняется инженером>>` в секциях, которые инженер заполняет вручную.  
Валидатор обнаруживает их через проверку на `<<` (подстрока).  
**Перед выпуском инженер обязан заменить все `<<...>>` реальным содержанием.**

---

## Критичные реквизиты для сверки (блокирующие ошибки)

```
document_id
document_type
title
revision
date
aircraft_model
component_name
author
```

---

## Специальные проверки для Engineering Analysis

Для каждой строки матрицы `ea_clause_matrix`, где `applicability_flag = YES`:
- `status` обязателен
- `means_of_compliance` обязателен
- `covered_in_section` обязателен
- статус не должен быть пустым

## Специальные проверки для Repair Instruction

Для каждой строки матрицы `ri_section_matrix`, где `mandatory_flag = YES` и `document_id` совпадает:
- `present_flag` должен быть `YES`

---

## Реализация в коде

| Функция | Файл | Что делает |
|---|---|---|
| `ValidateRequiredFields` | `modValidation.bas` | 14 проверок обязательных полей карточки |
| `ValidateWordDocument` | `modValidation.bas` | Открывает DOCX, ищет `{{...}}`, `<<`, TBD, обязательные разделы |
| `ValidateEAClauseMatrix` | `modValidation.bas` | Проверка матрицы EA |
| `ValidateRISectionMatrix` | `modValidation.bas` | Проверка матрицы RI |
| `ContainsTrashPlaceholder` | `modValidation.bas` | Умный поиск заглушек (не срабатывает на "draft" в путях) |
| `MatchWholeWord` | `modValidation.bas` | Вспомогательная функция: слово целиком |

---

## Формат отчёта (`validation_issues`)

Каждая запись содержит:
- `document_id` — номер документа
- `severity` — `ERROR` или `WARNING`
- `code` — код проверки (`CARD_REQUIRED`, `WORD_DOC`, `UNRESOLVED_MARKER`, `TRASH_PLACEHOLDER`, `MISSING_SECTION`, `EA_MATRIX`, `RI_MATRIX`)
- `message` — описание проблемы
- `timestamp` — время проверки

Форма `frmValidationReport` показывает список + счётчики `Errors: N / Warnings: N` в заголовке.
