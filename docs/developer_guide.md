# Руководство разработчика

## Принципы

- Не генерировать инженерное содержание автоматически.
- Не смешивать UI с бизнес-правилами.
- Все листы, маркеры и коды — через константы в `modConstants`.
- Новые типы документов добавляются через справочники и правила, без изменения ядра.

---

## Структура карточки документа (`clsDocumentCard`)

Класс содержит 31 поле. Порядок полей жёстко соответствует порядку колонок в листе `doc_cards`.  
`LoadFromRow` / `SaveToRow` читают/пишут по **позиционному индексу** (колонка 1 = `DocumentID`, ..., колонка 31 = `PdfPath`).

Если добавляете поле — добавляйте его **в конец** (колонки 30, 31...) и обновляйте обе процедуры одновременно с `modWorkbookSetup.EnsureWorkbookStructure`.

### Текущий порядок колонок

```
1  document_id               → DocumentID
2  document_type             → DocumentType
3  title                     → Title
4  aircraft_model            → AircraftModel
5  aircraft_variant          → AircraftVariant
6  aircraft_number           → AircraftNumber
7  msn                       → MSN
8  aircraft_manufacture_date → AircraftManufactureDate
9  aircraft_hours            → AircraftHours
10 aircraft_cycles           → AircraftCycles
11 assembly_number           → AssemblyNumber
12 part_number               → PartNumber
13 component_name            → ComponentName
14 component_sn              → ComponentSN
15 component_hours           → ComponentHours
16 component_cycles          → ComponentCycles
17 component_manufacture_date→ ComponentManufactureDate
18 applicability             → Applicability
19 revision                  → Revision
20 date                      → DocDate
21 author                    → Author
22 checker                   → Checker
23 approver                  → Approver
24 related_analysis_number   → RelatedAnalysisNumber
25 related_instruction_number→ RelatedInstructionNumber
26 references                → References
27 attachments               → Attachments
28 remarks                   → Remarks
29 status                    → Status
30 word_doc_path             → WordDocPath
31 pdf_path                  → PdfPath
```

---

## Архитектура Word-шаблонов

### Требования к файлу шаблона

`IsOoxmlFile` в `modWordAutomation` проверяет сигнатуру `PK` (первые 2 байта).  
Если файл не является ZIP/OOXML — процедура выбрасывает ошибку.  
**Никакого fallback на plain-text нет.**

### Как устроен `RepairInstruction.dotx`

Шаблон построен на основе `A.095.04.00000.078.0699_example.docx` с помощью Python-скрипта:
- взята исходная XML-структура (стили, таблицы, колонтитул, нумерация страниц),
- конкретные значения заменены на `{{Marker}}` единым XML-ран-ом,
- секции инженерного содержания заменены на `<<Заполняется инженером>>`,
- Content-Type изменён на `template.main+xml` (тип `.dotx`).

Word открывает `.dotx` через `Documents.Add(templatePath)` — создаётся **новый документ** (шаблон не изменяется).

### Добавление нового маркера

1. Вставьте маркер `{{NewMarker}}` в нужное место шаблона как единый текстовый run (не разбитый по нескольким runs).
2. Добавьте соответствующее поле в `clsDocumentCard`.
3. Добавьте строку в `modWordAutomation.ReplaceAllMarkers`:
   ```vba
   ReplaceText wordDoc, "{{NewMarker}}", card.NewField
   ```
4. Добавьте поле в `frmDocumentCard.BuildFields` и `ReadCardFromForm`.
5. Добавьте колонку в `modWorkbookSetup.EnsureWorkbookStructure` (секция `SHEET_DOC_CARDS`).

### `EngineeringAnalysis.dotx`

Текущий файл является плоским текстовым файлом (ASCII) и **не работает**.  
Требуется создать полноценный OOXML-шаблон аналогично RI, на основе реального примера EA.

---

## Как расширить на новый тип документа

1. Добавить тип в `ref_document_types`.
2. Добавить запись в `ref_templates` (имя файла шаблона).
3. Создать `.dotx` шаблон и положить в `templates/`.
4. Добавить обязательные разделы в `rules_required_sections`.
5. При необходимости добавить специализированную проверку в `modValidation`.

---

## Проверки валидации

`modValidation.ValidateBeforeRelease` выполняет 4 группы проверок:

1. **`ValidateRequiredFields`** — 14 правил: обязательные поля карточки.
2. **`ValidateWordDocument`** — открывает DOCX через Word COM и проверяет:
   - наличие незаменённых маркеров `{{...}}`,
   - наличие `<<` (незаполненных секций) или `TBD`/`XXX`/`???`,
   - наличие обязательных заголовков из `rules_required_sections`.
3. **`ValidateEAClauseMatrix`** — для EA: проверка матрицы покрытия.
4. **`ValidateRISectionMatrix`** — для RI: проверка контрольной матрицы.

`ContainsTrashPlaceholder` использует `MatchWholeWord` — не срабатывает на `"draft"` в пути к файлу.

---

## Импорт VBA

1. Открыть `DocumentControl.xlsm`, Alt+F11 → VBA Editor.
2. Модули: `File → Import File` из `vba/modules/*.bas`.
3. Классы: `File → Import File` из `vba/classes/*.cls`.
4. Формы: создать пустые `frmDocumentCard`, `frmValidationReport`, вставить код из `.frm` (см. `docs/forms_import.md`).
5. `Debug → Compile VBAProject`.

---

## Обработка ошибок и COM

- `modWordAutomation` использует **late binding** (нет ссылки на Word Object Library).
- Каждая процедура с Word имеет блок `CleanUp:` и `Set ... = Nothing`.
- Шаблон открывается через `Documents.Add(templatePath)` — исходный файл не изменяется.
- Процессы `WINWORD.EXE` завершаются через `wordApp.Quit` в `CleanUp`.
