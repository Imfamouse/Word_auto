# Импорт UserForm в репозитории без бинарных `.frx`

## Почему так

Некоторые PR-платформы не принимают бинарные файлы. Файлы UserForm `.frx` бинарные, поэтому в репозитории хранятся только текстовые `.frm` (code-behind).

## Как корректно восстановить формы в Excel VBA

1. В VBE создайте пустые формы вручную:
   - `frmDocumentCard`
   - `frmValidationReport`
2. Откройте файлы из репозитория:
   - `vba/forms/frmDocumentCard.frm`
   - `vba/forms/frmValidationReport.frm`
3. Скопируйте код **ниже `Option Explicit`** в окно кода соответствующей формы.
4. Сохраните проект и выполните `Debug -> Compile VBAProject`.

После вставки кода форма не останется пустой: элементы и кнопки (**Save Card / Create DOCX / Validate / Export PDF / Close / Field Help**) создаются в `UserForm_Initialize` динамически.

## Примечание

Это осознанное ограничение поставки в Git: бинарные `.frx` исключены для совместимости с PR-инструментами.
