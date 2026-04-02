# Word_auto — Excel + Word + VBA document-control system

Локальная система подготовки инженерных документов авиационной компании в парадигме:

- **Excel** = карточка документа + справочники + правила валидации,
- **Word** = ручное написание инженерного содержания,
- **VBA** = автоматизация шаблонов, проверок, журналирования и выпуска.

> Система **не** генерирует инженерный текст автоматически. Инженер пишет содержательную часть вручную в Word.

## Что в репозитории

- `docs/architecture.md` — архитектура (этапы 1–4).
- `docs/user_guide.md` — инструкция пользователя.
- `docs/developer_guide.md` — инструкция разработчика и расширения.
- `docs/test_plan.md` — тест-план и чек-лист приёмки.
- `docs/limitations_and_roadmap.md` — ограничения и направления развития.
- `vba/` — исходники VBA (модули, классы, формы).
- `templates/` — два Word-шаблона с маркерами.
- `sample_data/` — пример тестовых данных.

## Быстрый старт

1. Создайте Excel-книгу `DocumentControl.xlsm`.
2. Импортируйте файлы из `vba/modules`, `vba/classes`, `vba/forms` в VBA Editor.
3. Создайте листы согласно `docs/architecture.md`.
4. Скопируйте файлы из `templates/` в рабочую папку шаблонов.
5. Запустите `modMain.AppInitialize`, затем `modMain.OpenDocumentCard`.

## Совместимость

- Windows + Excel 2016+
- Word 2016+
- VBA (late binding к Word COM)

