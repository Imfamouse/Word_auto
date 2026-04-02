# CODEX_TASK.md

Ниже стартовый запрос, который нужно дать Codex в корне проекта.

---

You are working as a senior solution architect, senior VBA developer, and senior Word automation engineer.

Project goal:
Build a local Excel + Word + VBA document-control system for engineering documents in an aviation company.

Critical rule:
Do not try to auto-write engineering content.
The engineer writes the technical content manually in Word.
The software manages metadata, templates, consistency, validation, and release checks.

Read these root files first:
- AGENTS.md
- PROJECT_BRIEF.md
- DOCUMENT_TYPES.md
- VALIDATION_RULES.md
- IMPLEMENTATION_PLAN.md
- NAMING_RULES.md
- README.md

Then execute Stage 1 only.

Stage 1 output:
1. assumptions.md
2. architecture.md
3. workbook_structure.md
4. word_template_strategy.md
5. vba_module_design.md
6. validation_strategy.md
7. implementation_backlog.md

Rules:
- Do not implement full code yet.
- Make reasonable assumptions and mark them explicitly as ASSUMPTIONS.
- Keep the solution modular and maintainable.
- Separate UI, business rules, Excel data access, Word automation, validation, and logging.
- Target Excel 2016+ and a typical corporate Windows environment.

After Stage 1:
- summarize what was done,
- list created files,
- list open risks,
- propose the next implementation step.
