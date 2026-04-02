# Word_auto — Excel + Word + VBA document-control system

Локальная система подготовки инженерных документов авиационной компании в парадигме:

- **Excel** = карточка документа (31 поле) + справочники + правила валидации,
- **Word** = ручное написание инженерного содержания,
- **VBA** = автоматизация шаблонов, проверок, журналирования и выпуска.

> Система **не** генерирует инженерный текст автоматически. Инженер пишет содержательную часть вручную в Word.

## Что в репозитории

```
vba/
  modules/    — 9 VBA-модулей (логика, Word-автоматизация, валидация)
  classes/    — 3 класса (DocumentCard, ValidationIssue, TemplateRule)
  forms/      — 2 формы (карточка документа, отчёт валидации)
templates/
  RepairInstruction.dotx              — полноценный OOXML-шаблон RI,
                                        20 маркеров, структура из A.095.04.00000.078.0699
  EngineeringAnalysis.dotx            — заглушка (требует замены на OOXML)
  A.095.04.00000.078.0699_example.docx — исходный пример документа RI
sample_data/
  doc_cards_sample.csv                — пример данных карточки (31 колонка)
  ea_clause_matrix_sample.csv         — пример матрицы EA
  ri_section_matrix_sample.csv        — пример матрицы RI
docs/
  architecture.md        — архитектура системы
  user_guide.md          — инструкция пользователя
  developer_guide.md     — инструкция разработчика
  forms_import.md        — восстановление UserForm из .frm
  test_plan.md           — тест-план и чек-лист приёмки
  limitations_and_roadmap.md — ограничения и направления развития
```

## Быстрый старт

1. Создайте Excel-книгу `DocumentControl.xlsm`.
2. Импортируйте файлы из `vba/modules`, `vba/classes`, `vba/forms` (подробнее: `docs/forms_import.md`).
3. Запустите `modMain.AppInitialize` — автоматически создаст 14 листов с заголовками и базовые справочники.
4. В листе `cfg_app` укажите `templates_path` и `output_path`.
5. Скопируйте файлы из `templates/` в папку `templates_path`.
6. Запустите `modMain.OpenDocumentCard` и начните работу.

## Шаблон Repair Instruction

`RepairInstruction.dotx` построен на основе реального документа `A.095.04.00000.078.0699_example.docx`.  
При нажатии **Create DOCX** система подставляет 20 маркеров: тип ВС, номер документа, реквизиты компонента, подписи, дату — во все нужные места документа включая колонтитул.

## Совместимость

- Windows + Excel 2016+
- Word 2016+
- VBA (late binding к Word COM)
