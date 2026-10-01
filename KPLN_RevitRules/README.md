# KPLN_RevitRules

Общее хранилище знаний отдельно от KPLN_RevitMcpBridge, по принципу
KPLN_NavisClashRules. Python 3.10+, без внешних зависимостей для записей анализа.

- `rules/revit-policy.md` — общие правила работы; инженерные пороги не заданы.
- `rules/nevantom-kt-family-mapping.md` — проверенное сопоставление подборов
  Nevatom KT с параметрами универсального одноуровневого семейства.
- `knowledge/proposals/review-lessons.md` — наблюдения из проверенных задач.
- `skills/kpln-revit-bridge/SKILL.md` — тонкий навык, читающий этот checkout.
- `src/create_record.py` — записи анализа с происхождением данных и правил.
- `src/nevantom_selection.py` — автономный разбор текста подбора и сравнение с
  JSON-снимком семейства; операций записи не создаёт.
- `local/` — игнорируемые записи конкретных проектов.

Установить навык: `./scripts/install-skill.ps1`. Он создаёт резервную копию
существующего навыка и сохраняет указатель на этот checkout. Обновление правил
не требует пересборки DLL или MCP; после обновления навыка открыть новую задачу.

Пример записи анализа:

```powershell
python -B ./src/create_record.py ./local/evidence.json --context ./local/context.json --output ./local/review-001.json
python -B -m unittest discover -s tests -v
```

Контекст — JSON с идентификатором проекта, источниками требований, session_id,
document_id и revision. Решения и разрешённые действия добавить в запись
после анализа; скрипт их не придумывает. Для обмена знаниями использовать
обычную командную работу с Git и проверку изменений. Remote не создаётся
автоматически; реальные модели и их идентификаторы остаются в local/.
