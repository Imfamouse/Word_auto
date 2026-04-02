# DOCUMENT_TYPES.md

## 1. Repair Instruction

### Назначение
Документ описывает ремонтное решение, применимость, материалы/инструменты, последовательность выполнения, проверки и ограничения.

### Важное ограничение
Содержательная часть ремонта может сильно различаться от случая к случаю.
Поэтому система не должна пытаться автоматически писать repair procedure.

### Что автоматизируется
- реквизиты документа;
- объект ремонта;
- aircraft / MSN / applicability;
- assembly number / part number / component name;
- титул, шапка, колонтитулы;
- наличие обязательных разделов;
- ссылки на приложения и рисунки;
- проверка на незаполненные маркеры и мусорные заглушки.

### Типовые обязательные разделы
- Purpose / Scope
- Identification / Applicability
- References
- Description of damage or condition
- Materials / Tools / Equipment
- Repair procedure
- Inspection after repair
- Limitations / Notes
- Record of accomplishment / if applicable

---

## 2. Engineering Analysis

### Назначение
Документ обосновывает соответствие конструкции, ремонта или решения применимым пунктам базиса.

### Важное ограничение
Логика анализа, расчёты, применимость и набор пунктов базиса могут существенно меняться.
Поэтому система не должна пытаться автоматически писать reasoning или расчётную часть.

### Что автоматизируется
- реквизиты документа;
- общая структура;
- перечень clauses;
- матрица покрытия базиса;
- наличие статусов по applicable clauses;
- наличие means of compliance;
- наличие ссылок на разделы и подтверждающие материалы;
- проверка полноты покрытия и согласованности.

### Типовые обязательные разделы
- Objective / Scope
- Description of design / repair
- Inputs / assumptions
- Certification basis / applicable clauses
- Means of compliance
- Analysis
- Conclusion
- Limitations / remarks
- References / attachments

---

## Общий принцип
Word-шаблон задаёт каркас.
Инженер пишет техническое содержание вручную.
Excel хранит карточку документа, правила и матрицы проверки.
